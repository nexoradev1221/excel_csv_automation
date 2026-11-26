"""Report sheet structure, formatting, and formula-safe exception text."""

import csv
import io
from datetime import date
from decimal import Decimal

from openpyxl import load_workbook

from excel_csv_automation.calculations import calculate_run
from excel_csv_automation.exports import export_exceptions_csv, export_report_xlsx
from excel_csv_automation.models import (
    DeduplicationResult,
    NormalizedRow,
    ProcessingException,
    SourceReference,
    ValidationResult,
)


def sample_run():
    source = SourceReference("=branch.csv", 2, sheet_name="Sales", input_index=0)
    row = NormalizedRow(
        date=date(2026, 11, 4),
        branch="North",
        sku="=DEMO()",
        quantity=Decimal("2"),
        unit_price=Decimal("1.005"),
        transaction_id="0001",
        source=source,
        original_values={"sku": "=DEMO()"},
    )
    duplicate = ProcessingException(
        category="duplicate_conflict",
        source=SourceReference("later.csv", 3, input_index=1),
        reasons=("duplicate_conflict", "Different values"),
        original_values={"sku": "=DEMO()", "quantity": "-2"},
        retained_source=source,
    )
    rejection = ProcessingException(
        category="validation",
        source=SourceReference("formula.xlsx", 2, sheet_name="Sales", input_index=2),
        reasons=("formula_not_supported", "Formula cells cannot be used as quantity."),
        original_values={"quantity": "=SUM(A1:A2)"},
    )
    file_error = ProcessingException(
        category="file",
        source=SourceReference("broken.xlsx", None, input_index=3),
        reasons=("xlsx_read_error", "Cannot read workbook."),
    )
    validations = (
        ValidationResult(
            source=source,
            accepted=(row, row),
            rejected=(rejection,),
            file_errors=(file_error,),
        ),
    )
    return calculate_run(
        "export-test",
        validations,
        DeduplicationResult(retained=(row,), duplicates=(duplicate,)),
    )


def test_xlsx_report_has_all_sheets_formats_and_literal_untrusted_text():
    result = sample_run()
    book = load_workbook(io.BytesIO(export_report_xlsx(result)), data_only=False)
    assert book.sheetnames == ["Cleaned_Data", "Summary", "Exceptions", "Run_Info"]

    cleaned = book["Cleaned_Data"]
    assert cleaned.freeze_panes == "A2"
    assert cleaned.auto_filter.ref == "A1:K2"
    assert cleaned["A2"].number_format == "yyyy-mm-dd"
    assert cleaned["E2"].number_format == "#,##0.00;[Red]-#,##0.00;0.00"
    assert cleaned["G2"].value == 2.01
    assert cleaned["G2"].number_format == "#,##0.00;[Red]-#,##0.00;0.00"
    assert cleaned["C2"].value == "=DEMO()"
    assert cleaned["C2"].data_type == "s"
    assert cleaned["H2"].value == "=branch.csv"
    assert cleaned["H2"].data_type == "s"
    assert book["Summary"]["C2"].value == 2.01
    assert book["Exceptions"].max_row == 4
    assert all(
        cell.data_type != "f"
        for sheet in book.worksheets
        for row in sheet.iter_rows()
        for cell in row
    )
    book.close()


def test_exceptions_csv_uses_utf8_bom_and_neutralizes_formula_prefixes():
    data = export_exceptions_csv(sample_run())
    assert data.startswith(b"\xef\xbb\xbf")
    rows = list(csv.DictReader(io.StringIO(data.decode("utf-8-sig"), newline="")))
    assert len(rows) == 3
    duplicate = rows[0]
    assert duplicate["source_file"] == "later.csv"
    assert duplicate["retained_file"] == "'=branch.csv"
    assert duplicate["retained_row"] == "2"
    assert '"sku": "=DEMO()"' in duplicate["original_values"]
    validation = rows[1]
    assert validation["original_values"] == '{"quantity": "=SUM(A1:A2)"}'

    injected = ProcessingException(
        category="file",
        source=SourceReference(" =cmd|' /C calc'!A0", None),
        reasons=("bad", "value"),
    )
    result = calculate_run(
        "unsafe-name",
        [ValidationResult(source=injected.source, file_errors=(injected,))],
        DeduplicationResult(),
    )
    injected_csv = list(
        csv.DictReader(io.StringIO(export_exceptions_csv(result).decode("utf-8-sig"), newline=""))
    )
    assert injected_csv[0]["source_file"].startswith("' =")


def test_empty_result_still_has_downloadable_report_sheets_and_csv_header():
    source = SourceReference("empty.csv", None)
    result = calculate_run("empty", [ValidationResult(source=source)], DeduplicationResult())
    book = load_workbook(io.BytesIO(export_report_xlsx(result)), data_only=False)
    assert book.sheetnames == ["Cleaned_Data", "Summary", "Exceptions", "Run_Info"]
    assert book["Cleaned_Data"].max_row == 1
    assert book["Summary"].max_row == 1
    assert book["Exceptions"].max_row == 1
    assert Decimal(str(book["Run_Info"]["B8"].value)) == Decimal("0.00")
    book.close()
    csv_rows = list(
        csv.reader(io.StringIO(export_exceptions_csv(result).decode("utf-8-sig"), newline=""))
    )
    assert len(csv_rows) == 1
    assert csv_rows[0][0] == "category"
