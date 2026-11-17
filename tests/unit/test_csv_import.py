"""Observable CSV import behavior, including physical provenance and atomic errors."""

import json
from pathlib import Path

import pytest

from excel_csv_automation.importers import import_csv, import_csv_file
from excel_csv_automation.models import CANONICAL_FIELDS, MappingProfile

ROOT = Path(__file__).resolve().parents[2]


def make_profile(**settings):
    return MappingProfile(
        profile_id="test", column_map={field: field for field in CANONICAL_FIELDS}, **settings
    )


@pytest.mark.parametrize(
    "filename,profile_id,expected_header,expected_id",
    [
        ("02_south.csv", "russian_csv", "Филиал", "Номер"),
        ("04_west.csv", "alias_csv", "Office", "Txn"),
        ("06_port.csv", "canonical_csv", "branch", "transaction_id"),
    ],
)
def test_demo_csv_files_import_with_explicit_profile_and_original_text(
    filename, profile_id, expected_header, expected_id
):
    settings = json.loads(
        (ROOT / "config" / "sample_profiles" / f"{profile_id}.json").read_text(encoding="utf-8")
    )
    settings["duplicate_key"] = tuple(settings["duplicate_key"])
    result = import_csv_file(
        ROOT / "sample_data" / filename, MappingProfile(**settings), input_index=2
    )
    assert not result.errors
    assert len(result.rows) == 5
    assert expected_header in result.headers
    assert result.rows[0].values[expected_id].startswith("0")
    assert [row.source.row_number for row in result.rows] == [2, 3, 4, 5, 6]
    assert all(
        row.source.filename == filename and row.source.input_index == 2 for row in result.rows
    )
    assert all(isinstance(value, str) for row in result.rows for value in row.values.values())
    assert result.preview(2) == result.rows[:2]


def test_quoted_delimiters_multiline_and_blank_lines_keep_physical_provenance():
    content = b'id,description\r\n0001,"two, fields\r\non two lines"\r\n\r\n0002,end\r\n'
    result = import_csv(content, "upload.csv", make_profile())
    assert not result.errors
    assert [row.source.row_number for row in result.rows] == [2, 5]
    assert result.rows[0].values == {"id": "0001", "description": "two, fields\r\non two lines"}
    assert result.rows[1].values["id"] == "0002"


def test_empty_cells_and_formula_like_identifiers_are_not_coerced():
    result = import_csv(b" id , sku ,qty\n0001,=DEMO(),\n,,\n", "upload.csv", make_profile())
    assert result.headers == ("id", "sku", "qty")
    assert result.rows[0].values == {"id": "0001", "sku": "=DEMO()", "qty": ""}
    assert result.rows[1].values == {"id": "", "sku": "", "qty": ""}


@pytest.mark.parametrize(
    "content,code",
    [
        (b"", "missing_header"),
        (b"\n", "missing_header"),
        (b"id,,qty\n1,2,3", "blank_header"),
        (b"id, id \n1,2", "duplicate_header"),
        (b'id,note\n1,"unfinished', "csv_parse_error"),
        (b'id,note\n1,"closed"trailing', "csv_parse_error"),
        (b"id,note\n1,ok\n2,extra,cell", "row_width_mismatch"),
        (b"id,note\n1,ok\n2", "row_width_mismatch"),
    ],
)
def test_malformed_files_are_actionable_and_never_partially_accepted(content, code):
    result = import_csv(content, "broken.csv", make_profile(), input_index=4)
    assert not result.rows
    assert len(result.errors) == 1
    error = result.errors[0]
    assert error.category == "file"
    assert error.reasons[0] == code
    assert error.reasons[1]
    assert error.source.filename == "broken.csv"
    assert error.source.input_index == 4


def test_malformed_record_identifies_its_first_physical_line():
    result = import_csv(b'id,note\n1,ok\n2,"multi\nline",extra', "bad.csv", make_profile())
    assert result.errors[0].source.row_number == 3
    assert result.errors[0].original_values == {
        "column_1": "2",
        "column_2": "multi\nline",
        "column_3": "extra",
    }
    assert not result.rows


def test_header_only_is_successful_empty_import():
    result = import_csv(b"id,note\n", "empty.csv", make_profile())
    assert result.headers == ("id", "note")
    assert not result.rows and not result.errors
    assert result.preview() == ()
    with pytest.raises(ValueError):
        result.preview(-1)
