"""Validate all synthetic branch rows through production imports and mappings."""

import json
from pathlib import Path

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


def test_all_synthetic_rows_partition_into_23_valid_and_7_rejected():
    batch = json.loads((SAMPLES / "batch.json").read_text(encoding="utf-8"))
    oracle = json.loads((SAMPLES / "expected_results.json").read_text(encoding="utf-8"))
    accepted = []
    rejected = []
    for item in batch["files"]:
        selected_profile = profile(item["profile_id"])
        path = SAMPLES / item["filename"]
        imported = (
            import_xlsx_file(path, selected_profile)
            if path.suffix == ".xlsx"
            else import_csv_file(path, selected_profile)
        )
        mapped = map_import(imported, selected_profile)
        validated = validate_rows(mapped, selected_profile)
        assert not validated.file_errors
        accepted.extend(validated.accepted)
        rejected.extend(validated.rejected)

    assert len(accepted) == 23
    assert len(rejected) == 7
    references = {(item.source.filename, item.source.row_number) for item in rejected}
    expected = {(item["file"], item["row"]) for item in oracle["rejected"]}
    assert references == expected
    assert sum(len(item.reasons) // 2 for item in rejected) >= len(rejected)
