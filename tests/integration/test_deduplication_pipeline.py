"""Reconcile duplicate retention to the independent six-file demo oracle."""

import json
from pathlib import Path

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


def test_all_synthetic_rows_match_duplicate_oracle_and_conflict_policy():
    batch = json.loads((SAMPLES / "batch.json").read_text(encoding="utf-8"))
    oracle = json.loads((SAMPLES / "expected_results.json").read_text(encoding="utf-8"))
    validated = []
    rejected = []
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
        validated.extend(result.accepted)
        rejected.extend(result.rejected)

    rejected_refs = {(error.source.filename, error.source.row_number) for error in rejected}
    expected_rejected = {(row["file"], row["row"]) for row in oracle["rejected"]}
    assert rejected_refs == expected_rejected

    keys = {item.duplicate_key for item in profiles}
    assert len(keys) == 1, "A batch must use one consistent duplicate key."
    result = deduplicate_rows(validated, profiles[0])
    accepted_refs = {(row.source.filename, row.source.row_number) for row in result.retained}
    expected_accepted = {(row["file"], row["row"]) for row in oracle["accepted"]}
    assert accepted_refs == expected_accepted

    duplicate_refs = {
        (error.source.filename, error.source.row_number) for error in result.duplicates
    }
    expected_duplicates = {(row["file"], row["row"]) for row in oracle["duplicates"]}
    assert duplicate_refs == expected_duplicates
    retained_by_reference = {
        (row.source.filename, row.source.row_number): row.source for row in result.retained
    }
    for error, expected in zip(result.duplicates, oracle["duplicates"], strict=True):
        retained_reference = expected["retained_file"], expected["retained_row"]
        assert error.retained_source == retained_by_reference[retained_reference]
        assert (error.category == "duplicate_conflict") is expected["conflict"]
    assert len(result.retained) == 18
    assert len(result.duplicates) == 5
