"""Safe profile names and versioned JSON persistence behavior."""

import json

import pytest

from excel_csv_automation.models import CANONICAL_FIELDS, MappingProfile
from excel_csv_automation.profiles import (
    ProfileError,
    list_profile_ids,
    load_profile,
    save_profile,
)


def make_profile(**overrides):
    values = {
        "profile_id": "russian_sales",
        "column_map": {
            field: f"Поле {index}" for index, field in enumerate(CANONICAL_FIELDS, start=1)
        },
        "csv_delimiter": ";",
        "csv_encoding": "cp1251",
        "date_format": "%d.%m.%Y",
        "decimal_separator": ",",
        "sheet_name": "Продажи",
        "duplicate_key": ("branch", "transaction_id"),
        "quantity_decimal_places": 2,
        "price_decimal_places": 3,
    }
    values.update(overrides)
    return MappingProfile(**values)


def test_profile_round_trip_preserves_unicode_and_all_settings(tmp_path):
    expected = make_profile()
    saved_path = save_profile(expected, tmp_path)
    loaded = load_profile(expected.profile_id, tmp_path)
    assert saved_path == tmp_path / "russian_sales.json"
    assert loaded == expected
    assert "Продажи" in saved_path.read_text(encoding="utf-8")
    assert json.loads(saved_path.read_text(encoding="utf-8"))["schema_version"] == 1


@pytest.mark.parametrize("profile_id", ["../outside", "sub/name", ".hidden", "é", "a" * 65])
def test_profile_ids_cannot_escape_storage_directory(profile_id, tmp_path):
    with pytest.raises(ProfileError):
        load_profile(profile_id, tmp_path)
    with pytest.raises(ProfileError):
        save_profile(make_profile(profile_id=profile_id), tmp_path)
    assert list(tmp_path.iterdir()) == []


def test_blank_profile_id_is_rejected_by_the_profile_contract(tmp_path):
    with pytest.raises(ValueError, match="cannot be blank"):
        make_profile(profile_id="")
    with pytest.raises(ProfileError):
        load_profile("", tmp_path)
    assert list(tmp_path.iterdir()) == []


def test_load_rejects_schema_corruption_unknown_fields_and_filename_mismatch(tmp_path):
    save_profile(make_profile(), tmp_path)
    path = tmp_path / "russian_sales.json"
    value = json.loads(path.read_text(encoding="utf-8"))

    value["schema_version"] = 2
    path.write_text(json.dumps(value), encoding="utf-8")
    with pytest.raises(ProfileError, match="invalid or unsupported"):
        load_profile("russian_sales", tmp_path)

    value["schema_version"] = 1
    value["unexpected"] = True
    path.write_text(json.dumps(value), encoding="utf-8")
    with pytest.raises(ProfileError, match="invalid or unsupported"):
        load_profile("russian_sales", tmp_path)

    del value["unexpected"]
    value["profile_id"] = "different"
    path.write_text(json.dumps(value), encoding="utf-8")
    with pytest.raises(ProfileError, match="does not match"):
        load_profile("russian_sales", tmp_path)


def test_profile_listing_is_stable_and_ignores_unrecognized_files(tmp_path):
    save_profile(make_profile(profile_id="West"), tmp_path)
    save_profile(make_profile(profile_id="east"), tmp_path)
    (tmp_path / "../outside.json").write_text("{}", encoding="utf-8")
    (tmp_path / "README.md").write_text("not a profile", encoding="utf-8")
    assert list_profile_ids(tmp_path) == ("east", "West")
    assert list_profile_ids(tmp_path / "missing") == ()


def test_save_replaces_a_profile_atomically_and_loads_the_latest_version(tmp_path):
    first = make_profile()
    second = make_profile(csv_delimiter=",")
    save_profile(first, tmp_path)
    save_profile(second, tmp_path)
    assert load_profile("russian_sales", tmp_path) == second
    assert sorted(path.name for path in tmp_path.iterdir()) == ["russian_sales.json"]
