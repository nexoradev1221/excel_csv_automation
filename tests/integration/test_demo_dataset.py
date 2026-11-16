"""Verify the fixtures and independent oracle before production processing exists.

These tests read fixture cells directly; they do not claim application acceptance.
Later integration tests must run the production pipeline against this same oracle.
"""

import csv
import json
import runpy
from collections import defaultdict
from datetime import datetime
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from pathlib import Path

import pytest
from openpyxl import load_workbook

from excel_csv_automation.models import MappingProfile

ROOT = Path(__file__).resolve().parents[2]
SAMPLES = ROOT / "sample_data"


@pytest.fixture(scope="module")
def oracle():
    return json.loads((SAMPLES / "expected_results.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="module")
def batch_rows():
    batch = json.loads((SAMPLES / "batch.json").read_text(encoding="utf-8"))
    result = {}
    assert len(batch["files"]) == 6
    assert len({source["branch"] for source in batch["files"]}) == 6
    for source in batch["files"]:
        settings = json.loads(
            (ROOT / "config" / "sample_profiles" / f"{source['profile_id']}.json").read_text(
                encoding="utf-8"
            )
        )
        settings["duplicate_key"] = tuple(settings["duplicate_key"])
        profile = MappingProfile(**settings)
        path = SAMPLES / source["filename"]
        if path.suffix == ".xlsx":
            book = load_workbook(path, read_only=True, data_only=False)
            try:
                rows = list(book[profile.sheet_name].values)
            finally:
                book.close()
        else:
            with path.open(encoding=profile.csv_encoding, newline="") as stream:
                rows = list(csv.reader(stream, delimiter=profile.csv_delimiter))
        headers, *data = rows
        assert set(headers) == set(profile.column_map.values())
        assert len(data) == 5
        for row_number, values in enumerate(data, start=2):
            source_values = dict(zip(headers, values, strict=True))
            canonical = {name: source_values[header] for name, header in profile.column_map.items()}
            result[(path.name, row_number)] = (canonical, profile)
    return result


def reference(item):
    return item["file"], item["row"]


def number(value, profile):
    return Decimal(str(value).replace(profile.decimal_separator, "."))


def date(value, profile):
    if isinstance(value, datetime):
        return value.date()
    return datetime.strptime(value, profile.date_format).date()


def normalized(values, profile):
    return (
        date(values["date"], profile),
        values["branch"],
        values["sku"],
        number(values["quantity"], profile),
        number(values["unit_price"], profile),
        values["transaction_id"],
    )


def test_oracle_partitions_every_physical_source_row_once(batch_rows, oracle):
    categories = [oracle[name] for name in ("accepted", "rejected", "duplicates")]
    references = [reference(item) for category in categories for item in category]
    assert len(set(references)) == len(references) == 30
    assert set(references) == set(batch_rows)
    assert oracle["counts"] == {
        "parsed": 30,
        "accepted": 18,
        "rejected": 7,
        "duplicates": 5,
        "file_errors": 0,
    }
    assert [len(items) for items in categories] == [18, 7, 5]


def test_accepted_source_values_support_hand_specified_amounts_and_months(batch_rows, oracle):
    branch_month = defaultdict(lambda: Decimal("0.00"))
    months = defaultdict(lambda: Decimal("0.00"))
    for expected in oracle["accepted"]:
        values, profile = batch_rows[reference(expected)]
        amount = number(values["quantity"], profile) * number(values["unit_price"], profile)
        amount = amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        assert amount == Decimal(expected["line_total"])
        assert values["branch"] == expected["branch"]
        assert values["transaction_id"] == expected["transaction_id"]
        assert date(values["date"], profile).strftime("%Y-%m") == expected["month"]
        assert str(values["transaction_id"]).startswith("0")
        branch_month[expected["branch"], expected["month"]] += amount
        months[expected["month"]] += amount
    assert dict(branch_month) == {
        (item["branch"], item["month"]): Decimal(item["line_total"]) for item in oracle["summaries"]
    }
    assert dict(months) == {
        month: Decimal(value) for month, value in oracle["month_totals"].items()
    }
    assert sum(branch_month.values()) == Decimal(oracle["grand_total"]) == Decimal("159.53")


def test_rejected_fixture_rows_contain_the_documented_errors(batch_rows, oracle):
    for expected in oracle["rejected"]:
        values, profile = batch_rows[reference(expected)]
        match expected["reason"]:
            case "invalid_date":
                with pytest.raises(ValueError):
                    date(values["date"], profile)
            case "invalid_unit_price":
                with pytest.raises(InvalidOperation):
                    number(values["unit_price"], profile)
            case "missing_sku":
                assert not values["sku"]
            case "missing_quantity":
                assert not values["quantity"]
            case "nonfinite_quantity":
                assert not number(values["quantity"], profile).is_finite()
            case "negative_unit_price":
                assert number(values["unit_price"], profile) < 0
            case _:
                pytest.fail(f"Unexpected oracle reason: {expected['reason']}")


def test_duplicate_references_and_conflicts_match_raw_inputs(batch_rows, oracle):
    accepted = {reference(item) for item in oracle["accepted"]}
    for duplicate in oracle["duplicates"]:
        retained = duplicate["retained_file"], duplicate["retained_row"]
        assert retained in accepted
        raw, profile = batch_rows[reference(duplicate)]
        first, first_profile = batch_rows[retained]
        assert (raw["branch"], raw["transaction_id"]) == (first["branch"], first["transaction_id"])
        assert (normalized(raw, profile) != normalized(first, first_profile)) == duplicate[
            "conflict"
        ]
    assert ("05_center.xlsx", 4) not in accepted
    assert ("05_center.xlsx", 5) in accepted
    assert ("06_port.csv", 6) in accepted  # Same ID, different branch must survive.


def test_encoding_sheet_selection_and_literal_identifiers(batch_rows):
    south = (SAMPLES / "02_south.csv").read_bytes()
    with pytest.raises(UnicodeDecodeError):
        south.decode("utf-8")
    assert "Филиал" in south.decode("cp1251")
    assert (SAMPLES / "04_west.csv").read_bytes().startswith(b"\xef\xbb\xbf")
    book = load_workbook(SAMPLES / "05_center.xlsx", read_only=True)
    try:
        assert book.sheetnames == ["Notes", "Sales"]
    finally:
        book.close()
    assert batch_rows["06_port.csv", 6][0]["sku"] == "=DEMO()"


def test_regeneration_is_repeatable_and_does_not_replace_the_oracle(tmp_path):
    generator = runpy.run_path(str(ROOT / "scripts" / "generate_sample_data.py"))
    output, profiles = tmp_path / "inputs", tmp_path / "profiles"
    output.mkdir()
    oracle_path = output / "expected_results.json"
    oracle_path.write_text("independent oracle", encoding="utf-8")
    generate = generator["generate_sample_data"]
    paths = generate(output, profiles)
    first = {path.name: path.read_bytes() for path in paths}
    generate(output, profiles)
    assert first == {path.name: path.read_bytes() for path in paths}
    assert first == {name: (SAMPLES / name).read_bytes() for name in first}
    assert oracle_path.read_text(encoding="utf-8") == "independent oracle"
    assert len(list(profiles.glob("*.json"))) == 5
    for profile_path in profiles.glob("*.json"):
        checked_in = ROOT / "config" / "sample_profiles" / profile_path.name
        assert json.loads(profile_path.read_text(encoding="utf-8")) == json.loads(
            checked_in.read_text(encoding="utf-8")
        )
