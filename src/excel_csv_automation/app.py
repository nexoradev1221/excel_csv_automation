"""Russian local operator workflow for auditable branch report consolidation."""

import hashlib
from dataclasses import replace
from io import BytesIO
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook

from excel_csv_automation.exports import export_exceptions_csv, export_report_xlsx
from excel_csv_automation.importers import import_csv, import_xlsx
from excel_csv_automation.mapping import map_import
from excel_csv_automation.models import CANONICAL_FIELDS, MappingProfile
from excel_csv_automation.pipeline import BatchFile, process_batch
from excel_csv_automation.profiles import (
    ProfileError,
    list_profile_ids,
    load_profile,
    save_profile,
)

st.set_page_config(page_title="Обработка Excel и CSV", page_icon="📊", layout="wide")
st.title("Обработка Excel и CSV")
st.caption("Локальная обработка: исходные файлы остаются на этом компьютере.")

_SAMPLE_PROFILES = Path("config/sample_profiles")
_LOCAL_PROFILES = Path("profiles")
_ENCODINGS = ("utf-8-sig", "utf-8", "cp1251", "utf-16")
_FIELD_NAMES = {
    "date": "Дата",
    "branch": "Филиал",
    "sku": "Артикул",
    "quantity": "Количество",
    "unit_price": "Цена за единицу",
    "transaction_id": "Номер операции",
}
_CATEGORY_NAMES = {
    "file": "Ошибка файла",
    "validation": "Ошибка строки",
    "duplicate": "Дубликат",
    "duplicate_conflict": "Конфликт дубля",
}
_REASON_NAMES = {
    "unsupported_file_type": "Неподдерживаемый формат файла.",
    "unknown_encoding": "Неизвестная кодировка CSV.",
    "decode_error": "Не удалось прочитать CSV в выбранной кодировке.",
    "binary_or_wrong_encoding": "Данные похожи на двоичный файл или выбрана неверная кодировка.",
    "missing_header": "Не найдена строка заголовков.",
    "blank_header": "У каждого столбца должен быть заголовок.",
    "duplicate_header": "Заголовки столбцов должны быть уникальными.",
    "row_width_mismatch": "Число значений в строке не соответствует заголовкам.",
    "csv_parse_error": "Не удалось разобрать структуру CSV.",
    "file_read_error": "Не удалось прочитать файл. Проверьте доступ к нему.",
    "sheet_not_selected": "Выберите лист Excel перед импортом.",
    "sheet_not_found": "Выбранный лист не найден в книге Excel.",
    "xlsx_read_error": "Не удалось прочитать книгу Excel.",
    "missing_mapped_columns": "Не найдены столбцы из профиля сопоставления.",
    "missing_required": "Заполните обязательное значение.",
    "formula_not_supported": "Формулы в датах и числах не поддерживаются.",
    "invalid_date": "Дата не соответствует выбранному формату.",
    "invalid_number": "Введите число с выбранным десятичным разделителем.",
    "nonfinite_number": "Число должно быть конечным.",
    "excess_precision": "Слишком много знаков после десятичного разделителя.",
    "negative_unit_price": "Цена не может быть отрицательной.",
    "duplicate": "Повторяется ключ операции; оставлена первая строка.",
    "duplicate_conflict": "Ключ операции повторяется с другими значениями.",
}


def _profile_catalog():
    catalog = {}
    labels = {}
    for source, directory, label in (
        ("sample", _SAMPLE_PROFILES, "пример"),
        ("local", _LOCAL_PROFILES, "локальный"),
    ):
        try:
            profile_ids = list_profile_ids(directory)
        except ProfileError as error:
            st.warning(f"Не удалось прочитать профили {directory}: {error}")
            continue
        for profile_id in profile_ids:
            key = f"{source}:{profile_id}"
            try:
                catalog[key] = load_profile(profile_id, directory)
                labels[key] = f"{profile_id} ({label})"
            except ProfileError as error:
                st.warning(f"Профиль {profile_id} пропущен: {error}")
    return catalog, labels


def _sheet_names(data: bytes) -> tuple[str, ...]:
    book = load_workbook(BytesIO(data), read_only=True, data_only=False)
    try:
        return tuple(book.sheetnames)
    finally:
        book.close()


def _import_preview(upload, profile: MappingProfile, input_index: int):
    data = upload.getvalue()
    if Path(upload.name).suffix.lower() == ".xlsx":
        return import_xlsx(data, upload.name, profile, input_index=input_index)
    return import_csv(data, upload.name, profile, input_index=input_index)


def _render_file_errors(errors):
    for error in errors:
        code = error.reasons[0]
        st.error(
            f"{error.source.filename}: {_REASON_NAMES.get(code, 'Проверьте настройки файла.')}"
        )


def _exception_rows(result):
    rows = []
    for error in result.exceptions:
        rows.append(
            {
                "Тип": _CATEGORY_NAMES.get(error.category, error.category),
                "Файл": error.source.filename,
                "Лист": error.source.sheet_name or "",
                "Строка": error.source.row_number or "",
                "Причины": " | ".join(
                    _REASON_NAMES.get(code, "Проверьте значение.") for code in error.reasons[::2]
                ),
                "Сохранённая строка": (
                    f"{error.retained_source.filename}, {error.retained_source.row_number}"
                    if error.retained_source
                    else ""
                ),
            }
        )
    return rows


uploads = st.file_uploader(
    "Загрузите файлы филиалов",
    type=("csv", "xlsx"),
    accept_multiple_files=True,
    help="Поддерживаются CSV и XLSX. Файлы не отправляются на внешние сервисы.",
)

if not uploads:
    st.info("Выберите один или несколько файлов CSV/XLSX, чтобы увидеть предварительный просмотр.")
    st.session_state.pop("batch_signature", None)
    st.session_state.pop("batch_outcome", None)
else:
    catalog, profile_labels = _profile_catalog()
    if not catalog:
        st.error("Не найдены профили. Проверьте config/sample_profiles и папку profiles.")
        st.stop()

    batch_files = []
    for input_index, upload in enumerate(uploads):
        file_bytes = upload.getvalue()
        extension = Path(upload.name).suffix.lower()
        st.subheader(f"{input_index + 1}. {upload.name}")

        compatible = [
            key
            for key, profile in catalog.items()
            if (extension == ".xlsx") == (profile.sheet_name is not None)
        ]
        options = compatible or list(catalog)
        default_index = 0
        selector = st.selectbox(
            "Профиль сопоставления",
            options,
            format_func=lambda key: profile_labels[key],
            index=default_index,
            key=f"profile_{input_index}",
        )
        selected = catalog[selector]
        sheet_name = selected.sheet_name

        if extension == ".xlsx":
            try:
                sheets = _sheet_names(file_bytes)
            except Exception as error:  # Workbook errors are displayed as file-level import errors.
                sheets = ()
                st.warning(f"Список листов недоступен: {error}")
            if sheets:
                sheet_index = sheets.index(sheet_name) if sheet_name in sheets else 0
                sheet_name = st.selectbox(
                    "Лист Excel",
                    sheets,
                    index=sheet_index,
                    key=f"sheet_{input_index}",
                )

        with st.expander("Настройки формата", expanded=False):
            date_format = st.text_input(
                "Формат даты",
                value=selected.date_format,
                key=f"date_format_{input_index}",
            )
            decimal_separator = st.selectbox(
                "Десятичный разделитель",
                (".", ","),
                index=(0 if selected.decimal_separator == "." else 1),
                key=f"decimal_separator_{input_index}",
            )
            quantity_places = st.number_input(
                "Знаков после запятой: количество",
                min_value=0,
                max_value=12,
                value=min(selected.quantity_decimal_places, 12),
                key=f"quantity_places_{input_index}",
            )
            price_places = st.number_input(
                "Знаков после запятой: цена",
                min_value=0,
                max_value=12,
                value=min(selected.price_decimal_places, 12),
                key=f"price_places_{input_index}",
            )
            if extension == ".csv":
                delimiter = st.text_input(
                    "Разделитель CSV",
                    value=selected.csv_delimiter,
                    key=f"delimiter_{input_index}",
                    max_chars=1,
                )
                encoding_options = tuple(dict.fromkeys((*_ENCODINGS, selected.csv_encoding)))
                encoding = st.selectbox(
                    "Кодировка CSV",
                    encoding_options,
                    index=encoding_options.index(selected.csv_encoding),
                    key=f"encoding_{input_index}",
                )
            else:
                delimiter = selected.csv_delimiter
                encoding = selected.csv_encoding

        try:
            prepared = replace(
                selected,
                sheet_name=sheet_name,
                date_format=date_format,
                decimal_separator=decimal_separator,
                quantity_decimal_places=int(quantity_places),
                price_decimal_places=int(price_places),
                csv_delimiter=delimiter,
                csv_encoding=encoding,
            )
        except (TypeError, ValueError) as error:
            st.error(f"Проверьте настройки профиля: {error}")
            continue

        imported = _import_preview(upload, prepared, input_index)
        if imported.errors:
            _render_file_errors(imported.errors)
            mapped_profile = prepared
        else:
            st.caption("Заголовки: " + " · ".join(imported.headers))
            preview = [dict(row.values) for row in imported.preview(5)]
            if preview:
                st.dataframe(preview, hide_index=True, use_container_width=True)
            mapping = {}
            for start in range(0, len(CANONICAL_FIELDS), 2):
                columns = st.columns(2)
                for offset, field in enumerate(CANONICAL_FIELDS[start : start + 2]):
                    headers = list(imported.headers)
                    current_header = prepared.column_map[field]
                    selected_index = (
                        headers.index(current_header) if current_header in headers else 0
                    )
                    mapping[field] = columns[offset].selectbox(
                        _FIELD_NAMES[field],
                        headers,
                        index=selected_index,
                        key=f"map_{input_index}_{field}",
                    )
            try:
                mapped_profile = replace(prepared, column_map=mapping)
                mapped = map_import(imported, mapped_profile)
                if mapped.errors:
                    _render_file_errors(mapped.errors)
                elif mapped.rows:
                    st.caption("Предварительный просмотр после сопоставления")
                    st.dataframe(
                        [dict(row.values) for row in mapped.rows[:5]],
                        hide_index=True,
                        use_container_width=True,
                    )
            except (TypeError, ValueError) as error:
                st.error(f"Сопоставьте шесть полей с разными заголовками: {error}")
                continue

        save_name = st.text_input(
            "Сохранить профиль под именем",
            value=mapped_profile.profile_id,
            key=f"save_name_{input_index}",
        )
        if st.button("Сохранить профиль", key=f"save_profile_{input_index}"):
            try:
                saved = save_profile(replace(mapped_profile, profile_id=save_name))
                st.success(f"Профиль сохранён: {saved}")
                st.rerun()
            except (OSError, ProfileError, TypeError, ValueError) as error:
                st.error(f"Профиль не сохранён: {error}")

        batch_files.append(BatchFile(upload.name, file_bytes, mapped_profile))

    if len(batch_files) != len(uploads):
        st.warning("Исправьте настройки профилей для всех файлов перед обработкой.")
    duplicate_key = st.multiselect(
        "Ключ поиска дублей для всего пакета",
        CANONICAL_FIELDS,
        default=("branch", "transaction_id"),
        format_func=lambda field: _FIELD_NAMES[field],
        key="duplicate_key",
    )
    allow_partial = st.checkbox(
        "Продолжить обработку доступных файлов, если у некоторых есть ошибки",
        value=False,
        help="Без этого выбора при ошибке файла отчёт не формируется.",
        key="allow_partial",
    )

    signature_builder = hashlib.sha256()
    for item in batch_files:
        signature_builder.update(item.filename.encode("utf-8"))
        signature_builder.update(item.data)
        signature_builder.update(repr(item.profile).encode("utf-8"))
    signature_builder.update(repr(tuple(duplicate_key)).encode("utf-8"))
    signature_builder.update(str(allow_partial).encode("ascii"))
    signature = signature_builder.hexdigest()
    if st.session_state.get("batch_signature") != signature:
        st.session_state["batch_signature"] = signature
        st.session_state.pop("batch_outcome", None)
        st.session_state.pop("report_xlsx", None)
        st.session_state.pop("exceptions_csv", None)

    if st.button(
        "Обработать пакет",
        type="primary",
        disabled=(len(batch_files) != len(uploads) or not duplicate_key),
    ):
        with st.status("Обработка файлов", expanded=True) as status:
            st.write("Чтение, сопоставление и проверка строк…")
            outcome = process_batch(
                batch_files,
                tuple(duplicate_key),
                allow_partial=allow_partial,
            )
            st.session_state["batch_outcome"] = outcome
            if outcome.result is None:
                status.update(label="Нужен выбор оператора", state="error")
            else:
                st.session_state["report_xlsx"] = export_report_xlsx(outcome.result)
                st.session_state["exceptions_csv"] = export_exceptions_csv(outcome.result)
                status.update(label="Обработка завершена", state="complete")

    outcome = st.session_state.get("batch_outcome")
    if outcome is not None:
        if outcome.result is None:
            st.warning("Отчёт не создан. Исправьте ошибки либо явно разрешите частичную обработку.")
            _render_file_errors(outcome.file_errors)
        else:
            result = outcome.result
            st.header("Результаты обработки")
            columns = st.columns(5)
            for column, label, value in zip(
                columns,
                ("Строк прочитано", "Принято", "Отклонено", "Дубли", "Ошибки файлов"),
                (
                    result.counts.parsed,
                    result.counts.accepted,
                    result.counts.rejected,
                    result.counts.duplicates,
                    result.counts.file_errors,
                ),
                strict=True,
            ):
                column.metric(label, value)
            st.metric("Итого по принятым строкам", f"{result.grand_total:,.2f}")
            st.dataframe(
                [
                    {"Филиал": row.branch, "Месяц": row.month, "Сумма": str(row.line_total)}
                    for row in result.summaries
                ],
                hide_index=True,
                use_container_width=True,
            )
            if result.exceptions:
                st.subheader("Исключения и дубли для проверки")
                st.dataframe(_exception_rows(result), hide_index=True, use_container_width=True)
            left, right = st.columns(2)
            left.download_button(
                "Скачать consolidated_sales.xlsx",
                data=st.session_state["report_xlsx"],
                file_name="consolidated_sales.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            right.download_button(
                "Скачать exceptions.csv",
                data=st.session_state["exceptions_csv"],
                file_name="exceptions.csv",
                mime="text/csv",
            )
