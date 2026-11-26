"""Batch orchestration, audit counts, and explicit partial processing consent."""

import json
from pathlib import Path

import pytest

from excel_csv_automation.models import MappingProfile
from excel_csv_automation.pipeline import BatchFile, process_batch

ROOT = Path(__file__).resolve().parents[2]


def profile(profile_id):
    settings = json.loads(
        (ROOT / "config" / "sample_profiles" / f"{profile_id}.json").read_text(encoding="utf-8")
    )
    settings["duplicate_key"] = tuple(settings["duplicate_key"])
    return MappingProfile(**settings)


def test_batch_with_file_error_requires_explicit_partial_opt_in():
    valid_profile = profile("russian_csv")
    valid_path = ROOT / "sample_data" / "02_south.csv"
    files = [
        BatchFile("bad.csv", b"\xff", profile("canonical_csv")),
        BatchFile(valid_path.name, valid_path.read_bytes(), valid_profile),
    ]

    blocked = process_batch(files, valid_profile.duplicate_key, run_id="blocked")
    assert blocked.result is None
    assert len(blocked.file_errors) == 1
    assert blocked.file_errors[0].source.filename == "bad.csv"
    assert blocked.file_errors[0].source.input_index == 0

    partial = process_batch(
        files, valid_profile.duplicate_key, allow_partial=True, run_id="partial"
    )
    assert partial.result is not None
    assert partial.result.run_id == "partial"
    assert partial.result.counts.file_errors == 1
    assert partial.result.counts.parsed == 5
    assert partial.result.counts.accepted == 2
    assert partial.result.counts.rejected == 2
    assert partial.result.counts.duplicates == 1


def test_empty_batch_and_empty_duplicate_key_are_rejected():
    with pytest.raises(ValueError, match="at least one"):
        process_batch([], ("branch", "transaction_id"))
    item = BatchFile("a.csv", b"", profile("canonical_csv"))
    with pytest.raises(ValueError, match="Duplicate key"):
        process_batch([item], ())


def test_unsupported_extension_is_reported_as_file_error_without_reading():
    selected = profile("canonical_csv")
    outcome = process_batch([BatchFile("legacy.xls", b"ignored", selected)], selected.duplicate_key)
    assert outcome.result is None
    assert outcome.file_errors[0].reasons[0] == "unsupported_file_type"
