"""Stable duplicate retention, configurable keys, and conflict reporting."""

from datetime import date
from decimal import Decimal

from excel_csv_automation.deduplication import deduplicate_rows
from excel_csv_automation.models import (
    CANONICAL_FIELDS,
    MappingProfile,
    NormalizedRow,
    SourceReference,
)


def make_profile(**overrides):
    values = {
        "profile_id": "test",
        "column_map": {field: field for field in CANONICAL_FIELDS},
    }
    values.update(overrides)
    return MappingProfile(**values)


def row(filename, row_number, input_index, **overrides):
    values = {
        "date": date(2026, 11, 1),
        "branch": "North",
        "sku": "0007",
        "quantity": Decimal("2"),
        "unit_price": Decimal("1.25"),
        "transaction_id": "0001",
    }
    values.update(overrides)
    return NormalizedRow(
        **values,
        source=SourceReference(filename, row_number, input_index=input_index),
        original_values={field: str(value) for field, value in values.items()},
    )


def test_first_valid_row_by_input_and_physical_order_is_retained():
    later_file = row("later.csv", 2, 1)
    later_row = row("first.csv", 4, 0)
    first_row = row("first.csv", 2, 0)
    result = deduplicate_rows([later_file, later_row, first_row], make_profile())
    assert result.retained == (first_row,)
    assert [item.source for item in result.duplicates] == [later_row.source, later_file.source]
    assert all(item.retained_source == first_row.source for item in result.duplicates)
    assert all(item.category == "duplicate" for item in result.duplicates)


def test_different_payload_with_same_key_is_flagged_as_conflict():
    first = row("first.csv", 2, 0)
    changed = row("second.csv", 3, 1, quantity=Decimal("9"))
    result = deduplicate_rows([first, changed], make_profile())
    assert result.retained == (first,)
    conflict = result.duplicates[0]
    assert conflict.category == "duplicate_conflict"
    assert conflict.reasons[0] == "duplicate_conflict"
    assert conflict.source == changed.source
    assert conflict.retained_source == first.source
    assert conflict.original_values["quantity"] == "9"


def test_configured_business_key_controls_duplicate_identity():
    profile = make_profile(duplicate_key=("branch", "sku"))
    first = row("first.csv", 2, 0, transaction_id="001")
    same_key = row("second.csv", 2, 1, transaction_id="002")
    other_branch = row("third.csv", 2, 2, branch="South", transaction_id="003")
    result = deduplicate_rows([first, same_key, other_branch], profile)
    assert result.retained == (first, other_branch)
    assert result.duplicates[0].retained_source == first.source
    assert result.duplicates[0].category == "duplicate_conflict"


def test_empty_input_produces_empty_result():
    result = deduplicate_rows([], make_profile())
    assert result.retained == ()
    assert result.duplicates == ()
