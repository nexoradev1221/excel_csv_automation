"""Profile and provenance boundaries, independent of future parsing services."""

import pytest

from excel_csv_automation.models import CANONICAL_FIELDS, MappingProfile, SourceReference


def make_profile(**overrides):
    values = {"profile_id": "demo", "column_map": {name: name for name in CANONICAL_FIELDS}}
    values.update(overrides)
    return MappingProfile(**values)


@pytest.mark.parametrize(
    "overrides",
    [
        {"schema_version": 2},
        {"column_map": {"date": "Date"}},
        {"column_map": {name: "Same header" for name in CANONICAL_FIELDS}},
        {"duplicate_key": ()},
        {"duplicate_key": ("branch", "branch")},
        {"duplicate_key": ("unknown",)},
        {"csv_delimiter": "||"},
        {"csv_delimiter": "\n"},
        {"decimal_separator": ":"},
        {"quantity_decimal_places": -1},
    ],
)
def test_invalid_profile_cannot_enter_pipeline(overrides):
    with pytest.raises(ValueError):
        make_profile(**overrides)


def test_custom_business_key_and_csv_locale_are_preserved():
    profile = make_profile(
        duplicate_key=("branch", "transaction_id", "sku"),
        csv_encoding="cp1251",
        csv_delimiter=";",
        decimal_separator=",",
        date_format="%d.%m.%Y",
    )
    assert profile.duplicate_key == ("branch", "transaction_id", "sku")
    assert profile.csv_encoding == "cp1251"
    assert profile.date_format == "%d.%m.%Y"


def test_sources_disambiguate_same_named_uploads_and_support_file_errors():
    first = SourceReference("sales.csv", 2, input_index=0)
    second = SourceReference("sales.csv", 2, input_index=1)
    assert first != second
    assert SourceReference("broken.csv", None).row_number is None


@pytest.mark.parametrize("row,index", [(0, 0), (-1, 0), (2, -1)])
def test_invalid_source_coordinates_are_rejected(row, index):
    with pytest.raises(ValueError):
        SourceReference("sales.csv", row, input_index=index)
