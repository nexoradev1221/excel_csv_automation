"""Workbook selection, raw cell preservation, and actionable file-level failures."""

import json
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pytest
from openpyxl import Workbook

from excel_csv_automation.importers import import_xlsx, import_xlsx_file
from excel_csv_automation.models import CANONICAL_FIELDS, MappingProfile

ROOT = Path(__file__).resolve().parents[2]


def make_profile(**settings):
    return MappingProfile(
        profile_id="test", column_map={field: field for field in CANONICAL_FIELDS}, **settings
    )


def workbook_bytes():
    book = Workbook()
    notes = book.active
    notes.title = "Notes"
    notes.append(["Not transaction data"])
    sheet = book.create_sheet("Sales")
    sheet.append([" date ", "branch", "sku", "quantity", "unit_price", "transaction_id"])
    sheet.append([datetime(2026, 11, 1), "North", "0007", 2, 1.25, "00001"])
    sheet.append([None, None, None, None, None, None])
    sheet.append([datetime(2026, 11, 2), "North", "=DEMO()", 1, 4, "00002"])
    output = BytesIO()
    book.save(output)
    book.close()
    return output.getvalue()


def test_selected_sheet_keeps_raw_types_formula_text_and_physical_rows():
    result = import_xlsx(
        workbook_bytes(), "orders.XLSX", make_profile(sheet_name="Sales"), input_index=3
    )
    assert not result.errors
    assert result.headers == ("date", "branch", "sku", "quantity", "unit_price", "transaction_id")
    assert [row.source.row_number for row in result.rows] == [2, 4]
    assert [row.source.sheet_name for row in result.rows] == ["Sales", "Sales"]
    assert all(row.source.input_index == 3 for row in result.rows)
    assert result.rows[0].values["date"] == datetime(2026, 11, 1)
    assert result.rows[0].values["sku"] == "0007"
    assert result.rows[0].values["transaction_id"] == "00001"
    assert result.rows[1].values["sku"] == "=DEMO()"
    assert result.preview(1) == result.rows[:1]


def test_sheet_must_be_selected_and_named_sheet_must_exist():
    missing_selection = import_xlsx(workbook_bytes(), "orders.xlsx", make_profile())
    assert missing_selection.errors[0].reasons[0] == "sheet_not_selected"
    absent = import_xlsx(workbook_bytes(), "orders.xlsx", make_profile(sheet_name="Missing"))
    assert absent.errors[0].reasons[0] == "sheet_not_found"
    assert "Notes, Sales" in absent.errors[0].reasons[1]
    assert not absent.rows


@pytest.mark.parametrize(
    "build,code",
    [
        (lambda sheet: sheet.append([None, None]), "missing_header"),
        (lambda sheet: sheet.append(["id", None]), "blank_header"),
        (lambda sheet: sheet.append([" id ", "id"]), "duplicate_header"),
    ],
)
def test_invalid_headers_exclude_entire_workbook(build, code):
    book = Workbook()
    sheet = book.active
    sheet.title = "Sales"
    build(sheet)
    output = BytesIO()
    book.save(output)
    book.close()
    result = import_xlsx(output.getvalue(), "invalid.xlsx", make_profile(sheet_name="Sales"))
    assert result.errors[0].reasons[0] == code
    assert not result.rows


@pytest.mark.parametrize(
    "filename,data,profile,code",
    [
        ("wrong.csv", b"not a workbook", "selected", "unsupported_file_type"),
        ("broken.xlsx", b"not a workbook", "selected", "xlsx_read_error"),
        ("broken.xlsx", b"not a workbook", "none", "sheet_not_selected"),
    ],
)
def test_wrong_or_unreadable_workbook_has_a_file_error(filename, data, profile, code):
    selected = make_profile(sheet_name="Sales") if profile == "selected" else make_profile()
    result = import_xlsx(data, filename, selected)
    assert result.errors[0].reasons[0] == code
    assert not result.rows


def test_demo_xlsx_workbooks_import_the_configured_sales_sheets():
    for filename, profile_id in (
        ("01_north.xlsx", "canonical_xlsx"),
        ("03_east.xlsx", "alias_xlsx"),
        ("05_center.xlsx", "canonical_xlsx"),
    ):
        settings = json.loads(
            (ROOT / "config" / "sample_profiles" / f"{profile_id}.json").read_text(encoding="utf-8")
        )
        settings["duplicate_key"] = tuple(settings["duplicate_key"])
        result = import_xlsx_file(ROOT / "sample_data" / filename, MappingProfile(**settings))
        assert not result.errors
        assert len(result.rows) == 5
        assert all(row.source.sheet_name == settings["sheet_name"] for row in result.rows)
        assert all(row.source.row_number == index for index, row in enumerate(result.rows, start=2))
