"""Compare production batch calculations to the hand-specified financial oracle."""

import json
from decimal import Decimal
from pathlib import Path

from excel_csv_automation.calculations import calculate_run
from excel_csv_automation.deduplication import deduplicate_rows
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


def test_full_batch_counts_totals_and_summaries_match_independent_oracle():
    batch = json.loads((SAMPLES / "batch.json").read_text(encoding="utf-8"))
    oracle = json.loads((SAMPLES / "expected_results.json").read_text(encoding="utf-8"))
    validations = []
    valid_rows = []
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
        assert not result.file_errors
        validations.append(result)
        valid_rows.extend(result.accepted)

    assert len({item.duplicate_key for item in profiles}) == 1
    duplicate_result = deduplicate_rows(valid_rows, profiles[0])
    run = calculate_run("synthetic-oracle", validations, duplicate_result)

    assert run.counts.parsed == oracle["counts"]["parsed"]
    assert run.counts.accepted == oracle["counts"]["accepted"]
    assert run.counts.rejected == oracle["counts"]["rejected"]
    assert run.counts.duplicates == oracle["counts"]["duplicates"]
    assert run.counts.file_errors == oracle["counts"]["file_errors"]
    assert run.counts.parsed == run.counts.accepted + run.counts.rejected + run.counts.duplicates
    assert run.grand_total == Decimal(oracle["grand_total"])

    accepted_totals = {
        (item.row.source.filename, item.row.source.row_number): item.line_total
        for item in run.accepted
    }
    expected_lines = {
        (item["file"], item["row"]): Decimal(item["line_total"]) for item in oracle["accepted"]
    }
    assert accepted_totals == expected_lines
    summaries = {(item.branch, item.month): item.line_total for item in run.summaries}
    expected_summaries = {
        (item["branch"], item["month"]): Decimal(item["line_total"]) for item in oracle["summaries"]
    }
    assert summaries == expected_summaries
