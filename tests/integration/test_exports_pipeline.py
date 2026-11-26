"""Reopen generated reports and compare every exported result to the oracle."""

import csv
import io
import json
from collections import Counter
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from excel_csv_automation.calculations import calculate_run
from excel_csv_automation.deduplication import deduplicate_rows
from excel_csv_automation.exports import export_exceptions_csv, export_report_xlsx
from excel_csv_automation.importers import import_csv_file, import_xlsx_file
from excel_csv_automation.mapping import map_import
from excel_csv_automation.models import MappingProfile
from excel_csv_automation.validation import validate_rows

ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "sample_data"


def profile(profile_id):
    settings = json.loads(
        (ROOT / "config" / "sample_profiles" / f"{profile_id}.json").read_text(encoding="utf-8")
    )
    settings["duplicate_key"] = tuple(settings["duplicate_key"])
    return MappingProfile(**settings)


def full_run():
    batch = json.loads((SAMPLES / "batch.json").read_text(encoding="utf-8"))
    validations = []
    rows = []
    profiles = []
    for input_index, item in enumerate(batch["files"]):
        selected_profile = profile(item["profile_id"])
        profiles.append(selected_profile)
        path = SAMPLES / item["filename"]
        imported = (
            import_xlsx_file(path, selected_profile, input_index=input_index)
            if path.suffix == ".xlsx"
            else import_csv_file(path, selected_profile, input_index=input_index)
        )
        result = validate_rows(map_import(imported, selected_profile), selected_profile)
        validations.append(result)
        rows.extend(result.accepted)
    deduplication = deduplicate_rows(rows, profiles[0])
    return calculate_run("export-oracle", validations, deduplication)


def test_workbook_sheets_values_and_exception_csv_match_full_oracle():
    result = full_run()
    oracle = json.loads((SAMPLES / "expected_results.json").read_text(encoding="utf-8"))
    workbook_bytes = export_report_xlsx(result)
    book = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    assert book.sheetnames == ["Cleaned_Data", "Summary", "Exceptions", "Run_Info"]

    cleaned = book["Cleaned_Data"]
    assert cleaned.max_row == 19
    assert cleaned.freeze_panes == "A2"
    assert cleaned.auto_filter.ref == "A1:K19"
    cleaned_values = {
        (cleaned.cell(row, 8).value, cleaned.cell(row, 10).value): Decimal(
            str(cleaned.cell(row, 7).value)
        )
        for row in range(2, cleaned.max_row + 1)
    }
    expected_lines = {
        (item["file"], item["row"]): Decimal(item["line_total"]) for item in oracle["accepted"]
    }
    assert cleaned_values == expected_lines
    assert all(
        cell.data_type != "f"
        for sheet in book.worksheets
        for row in sheet.iter_rows()
        for cell in row
    )

    summary = book["Summary"]
    summaries = {
        (summary.cell(row, 1).value, summary.cell(row, 2).value): Decimal(
            str(summary.cell(row, 3).value)
        )
        for row in range(2, summary.max_row + 1)
    }
    expected_summaries = {
        (item["branch"], item["month"]): Decimal(item["line_total"]) for item in oracle["summaries"]
    }
    assert summaries == expected_summaries
    assert book["Exceptions"].max_row == 13
    run_info = {
        book["Run_Info"].cell(row, 1).value: book["Run_Info"].cell(row, 2).value
        for row in range(2, book["Run_Info"].max_row + 1)
    }
    assert run_info["grand_total"] == float(oracle["grand_total"])
    assert run_info["row_reconciliation"] == "OK"
    assert run_info["accepted"] == 18
    book.close()

    exception_data = export_exceptions_csv(result)
    assert exception_data.startswith(b"\xef\xbb\xbf")
    exceptions = list(csv.DictReader(io.StringIO(exception_data.decode("utf-8-sig"), newline="")))
    assert len(exceptions) == 12
    assert Counter(item["category"] for item in exceptions) == {
        "validation": 7,
        "duplicate": 4,
        "duplicate_conflict": 1,
    }
