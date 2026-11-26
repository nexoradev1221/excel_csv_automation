"""Column mapping preserves raw values and source references."""

import json
from pathlib import Path

from excel_csv_automation.importers import import_csv_file, import_xlsx_file
from excel_csv_automation.mapping import map_import
from excel_csv_automation.models import (
    CANONICAL_FIELDS,
    FileImport,
    MappingProfile,
    SourceReference,
)

ROOT = Path(__file__).resolve().parents[2]


def make_profile(**overrides):
    values = {
        "profile_id": "test",
        "column_map": {field: field for field in CANONICAL_FIELDS},
    }
    values.update(overrides)
    return MappingProfile(**values)


def load_sample_profile(profile_id):
    settings = json.loads(
        (ROOT / "config" / "sample_profiles" / f"{profile_id}.json").read_text(encoding="utf-8")
    )
    settings["duplicate_key"] = tuple(settings["duplicate_key"])
    return MappingProfile(**settings)


def test_russian_csv_mapping_returns_canonical_fields_without_coercion():
    profile = load_sample_profile("russian_csv")
    imported = import_csv_file(ROOT / "sample_data" / "02_south.csv", profile, input_index=2)
    result = map_import(imported, profile)
    assert not result.errors
    assert tuple(result.rows[0].values) == CANONICAL_FIELDS
    assert result.rows[0].values["transaction_id"] == "0101"
    assert result.rows[0].values["date"] == "03.10.2026"
    assert result.rows[0].source == imported.rows[0].source
    assert result.source == imported.source


def test_alias_xlsx_mapping_preserves_native_dates_sheet_and_physical_rows():
    profile = load_sample_profile("alias_xlsx")
    imported = import_xlsx_file(ROOT / "sample_data" / "03_east.xlsx", profile, input_index=4)
    result = map_import(imported, profile)
    assert not result.errors
    assert tuple(result.rows[0].values) == CANONICAL_FIELDS
    assert result.rows[0].values["branch"] == "Восток"
    assert result.rows[0].source.sheet_name == "Orders"
    assert result.rows[0].source.row_number == 2
    assert result.rows[0].source.input_index == 4


def test_missing_mapping_columns_produce_file_error_without_partial_rows():
    profile = make_profile(
        column_map={
            field: ("Absent SKU" if field == "sku" else field) for field in CANONICAL_FIELDS
        }
    )
    imported = FileImport(
        source=SourceReference("input.csv", None),
        headers=CANONICAL_FIELDS,
        rows=(),
    )
    result = map_import(imported, profile)
    assert not result.rows
    assert result.errors[0].category == "file"
    assert result.errors[0].reasons[0] == "missing_mapped_columns"
    assert "Absent SKU" in result.errors[0].reasons[1]
    assert result.errors[0].source.row_number is None


def test_import_file_errors_are_preserved_without_mapping():
    profile = make_profile()
    imported = import_csv_file(ROOT / "sample_data" / "missing.csv", profile)
    result = map_import(imported, profile)
    assert result.errors == imported.errors
    assert not result.rows
