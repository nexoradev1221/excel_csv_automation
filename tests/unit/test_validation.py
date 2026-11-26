"""Required-field and exact numeric/date validation behavior."""

from datetime import date, datetime
from decimal import Decimal

import pytest

from excel_csv_automation.mapping import map_import
from excel_csv_automation.models import (
    CANONICAL_FIELDS,
    FileImport,
    MappedFile,
    MappingProfile,
    ProcessingException,
    SourceReference,
    SourceRow,
)
from excel_csv_automation.validation import validate_rows


def make_profile(**overrides):
    values = {
        "profile_id": "test",
        "column_map": {field: field for field in CANONICAL_FIELDS},
    }
    values.update(overrides)
    return MappingProfile(**values)


def make_row(**overrides):
    values = {
        "date": "2026-11-03",
        "branch": " North ",
        "sku": " 0007 ",
        "quantity": "2.125",
        "unit_price": "1.2500",
        "transaction_id": " 00001 ",
    }
    values.update(overrides)
    return SourceRow(SourceReference("sales.csv", 2, input_index=1), values)


def validate(row, profile=None):
    profile = profile or make_profile()
    mapped = MappedFile(source=row.source, rows=(row,))
    return validate_rows(mapped, profile)


def test_valid_row_is_normalized_exactly_and_original_values_are_retained():
    row = make_row(sku="=DEMO()")
    result = validate(row)
    assert not result.rejected and not result.file_errors
    normalized = result.accepted[0]
    assert normalized.date == date(2026, 11, 3)
    assert normalized.branch == "North"
    assert normalized.sku == "=DEMO()"
    assert normalized.quantity == Decimal("2.125")
    assert normalized.unit_price == Decimal("1.2500")
    assert normalized.transaction_id == "00001"
    assert normalized.source == row.source
    assert normalized.original_values == row.values


def test_native_excel_date_signed_quantity_zero_price_and_comma_decimal():
    row = make_row(date=datetime(2026, 11, 4), quantity="-1,25", unit_price=0)
    profile = make_profile(decimal_separator=",", quantity_decimal_places=2)
    result = validate(row, profile)
    assert not result.rejected
    assert result.accepted[0].date == date(2026, 11, 4)
    assert result.accepted[0].quantity == Decimal("-1.25")
    assert result.accepted[0].unit_price == Decimal("0")


@pytest.mark.parametrize(
    "field,value,code",
    [
        ("date", "31.02.2026", "invalid_date"),
        ("date", "=TODAY()", "formula_not_supported"),
        ("branch", "  ", "missing_required"),
        ("quantity", "NaN", "nonfinite_number"),
        ("quantity", "=SUM(A1:A2)", "formula_not_supported"),
        ("quantity", "1,000", "invalid_number"),
        ("quantity", "1.2345", "excess_precision"),
        ("unit_price", "-1", "negative_unit_price"),
        ("unit_price", "Infinity", "nonfinite_number"),
        ("sku", None, "missing_required"),
    ],
)
def test_invalid_value_is_rejected_with_reason_and_original_data(field, value, code):
    row = make_row(**{field: value})
    result = validate(row)
    assert not result.accepted
    exception = result.rejected[0]
    assert exception.category == "validation"
    assert exception.source == row.source
    assert code in exception.reasons
    assert exception.original_values[field] == value


def test_all_row_errors_are_reported_together_instead_of_stopping_at_first():
    row = make_row(date="not-a-date", branch="", quantity="nope", unit_price="-1")
    exception = validate(row).rejected[0]
    codes = exception.reasons[::2]
    assert codes == ("invalid_date", "missing_required", "invalid_number", "negative_unit_price")


def test_date_format_is_explicit_and_round_tripped_without_ambiguity():
    profile = make_profile(date_format="%d/%m/%Y")
    result = validate(make_row(date="01/02/2026"), profile)
    assert result.accepted[0].date == date(2026, 2, 1)
    strict = validate(make_row(date="1/2/2026"), profile)
    assert strict.rejected[0].reasons[0] == "invalid_date"


def test_file_errors_are_kept_without_row_processing():
    source = SourceReference("bad.csv", None)
    error = ProcessingException(
        category="file", source=source, reasons=("decode_error", "Invalid encoding.")
    )
    imported = FileImport(
        source=source,
        errors=(error,),
    )
    profile = make_profile()
    mapped = map_import(imported, profile)
    result = validate_rows(mapped, profile)
    assert not result.accepted and not result.rejected
    assert result.file_errors == (error,)
