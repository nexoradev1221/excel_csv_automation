"""Safe in-memory XLSX and exception CSV report generation."""

import csv
import io
import json
from datetime import date, datetime
from decimal import Decimal

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from excel_csv_automation.models import ProcessingException, RunResult

_MONEY_FORMAT = "#,##0.00;[Red]-#,##0.00;0.00"
_EXCEPTION_COLUMNS = (
    "category",
    "source_file",
    "source_sheet",
    "source_row",
    "input_index",
    "reason_codes",
    "reasons",
    "retained_file",
    "retained_sheet",
    "retained_row",
    "retained_input_index",
    "original_values",
)


def _json_value(value):
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    return value


def _exception_record(error: ProcessingException) -> dict[str, str | int]:
    reasons = error.reasons
    codes = reasons[::2]
    messages = reasons[1::2]
    retained = error.retained_source
    return {
        "category": error.category,
        "source_file": error.source.filename,
        "source_sheet": error.source.sheet_name or "",
        "source_row": error.source.row_number or "",
        "input_index": error.source.input_index,
        "reason_codes": "; ".join(codes),
        "reasons": " | ".join(messages or reasons),
        "retained_file": retained.filename if retained else "",
        "retained_sheet": retained.sheet_name or "" if retained else "",
        "retained_row": retained.row_number or "" if retained else "",
        "retained_input_index": retained.input_index if retained else "",
        "original_values": json.dumps(
            {key: _json_value(value) for key, value in error.original_values.items()},
            ensure_ascii=False,
            sort_keys=True,
        ),
    }


def _literal_csv(value: object) -> str:
    """Prefix spreadsheet formula-like text so CSV consumers treat it literally."""
    text = "" if value is None else str(value)
    probe = text.lstrip(" \t\r\n")
    if probe.startswith(("=", "+", "-", "@")):
        return "'" + text
    return text


def export_exceptions_csv(result: RunResult) -> bytes:
    """Return a UTF-8 BOM CSV exception report with spreadsheet-safe text cells."""
    stream = io.StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=_EXCEPTION_COLUMNS, lineterminator="\r\n")
    writer.writeheader()
    for error in result.exceptions:
        writer.writerow(
            {key: _literal_csv(value) for key, value in _exception_record(error).items()}
        )
    return stream.getvalue().encode("utf-8-sig")


def _set_literal(cell, value) -> None:
    cell.value = value
    if isinstance(value, str):
        cell.data_type = "s"


def _format_table(sheet, *, currency_columns: tuple[int, ...] = ()) -> None:
    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="245A73")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 30
    for column in range(1, sheet.max_column + 1):
        values = [sheet.cell(row, column).value for row in range(1, sheet.max_row + 1)]
        width = max((len(str(value)) for value in values if value is not None), default=10)
        sheet.column_dimensions[get_column_letter(column)].width = min(max(width + 2, 12), 48)
    for column in currency_columns:
        for row in range(2, sheet.max_row + 1):
            sheet.cell(row, column).number_format = _MONEY_FORMAT


def export_report_xlsx(result: RunResult) -> bytes:
    """Build a formatted report workbook entirely in memory."""
    book = Workbook()
    cleaned = book.active
    cleaned.title = "Cleaned_Data"
    cleaned_headers = (
        "date",
        "branch",
        "sku",
        "quantity",
        "unit_price",
        "transaction_id",
        "line_total",
        "source_file",
        "source_sheet",
        "source_row",
        "input_index",
    )
    cleaned.append(cleaned_headers)
    for item in result.accepted:
        row = item.row
        values = (
            row.date,
            row.branch,
            row.sku,
            row.quantity,
            row.unit_price,
            row.transaction_id,
            item.line_total,
            row.source.filename,
            row.source.sheet_name or "",
            row.source.row_number or "",
            row.source.input_index,
        )
        output_row = cleaned.max_row + 1
        for column, value in enumerate(values, start=1):
            _set_literal(cleaned.cell(output_row, column), value)
    _format_table(cleaned, currency_columns=(5, 7))
    for row_number in range(2, cleaned.max_row + 1):
        cleaned.cell(row_number, 1).number_format = "yyyy-mm-dd"
        cleaned.cell(row_number, 4).number_format = "0.###"

    summary = book.create_sheet("Summary")
    summary.append(("branch", "month", "line_total"))
    for item in result.summaries:
        summary.append((item.branch, item.month, item.line_total))
    _format_table(summary, currency_columns=(3,))

    exceptions = book.create_sheet("Exceptions")
    exceptions.append(_EXCEPTION_COLUMNS)
    for error in result.exceptions:
        record = _exception_record(error)
        output_row = exceptions.max_row + 1
        for column, name in enumerate(_EXCEPTION_COLUMNS, start=1):
            _set_literal(exceptions.cell(output_row, column), record[name])
    _format_table(exceptions)

    run_info = book.create_sheet("Run_Info")
    run_info.append(("field", "value"))
    for name, value in (
        ("run_id", result.run_id),
        ("parsed", result.counts.parsed),
        ("accepted", result.counts.accepted),
        ("rejected", result.counts.rejected),
        ("duplicates", result.counts.duplicates),
        ("file_errors", result.counts.file_errors),
        ("grand_total", result.grand_total),
        ("rounding_policy", "ROUND_HALF_UP per accepted line to 0.01"),
        (
            "row_reconciliation",
            "OK"
            if result.counts.parsed
            == result.counts.accepted + result.counts.rejected + result.counts.duplicates
            else "ERROR",
        ),
    ):
        _set_literal(run_info.cell(run_info.max_row + 1, 1), name)
        _set_literal(run_info.cell(run_info.max_row, 2), value)
    for name, value in sorted(result.run_metadata.items()):
        output_row = run_info.max_row + 1
        _set_literal(run_info.cell(output_row, 1), name)
        _set_literal(run_info.cell(output_row, 2), value)
    _format_table(run_info, currency_columns=(2,))

    output = io.BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()
