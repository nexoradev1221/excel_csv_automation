"""Strict required-field, date, and decimal validation for mapped input rows."""

from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from excel_csv_automation.models import (
    MappedFile,
    MappingProfile,
    NormalizedRow,
    ProcessingException,
    SourceRow,
    ValidationResult,
)


def _required_text(value, field: str) -> tuple[str | None, tuple[str, str] | None]:
    if value is None or isinstance(value, bool):
        return None, ("missing_required", f"{field} is required.")
    text = str(value).strip()
    if not text:
        return None, ("missing_required", f"{field} is required.")
    return text, None


def _parse_date(value, profile: MappingProfile) -> tuple[date | None, tuple[str, str] | None]:
    if isinstance(value, datetime):
        return value.date(), None
    if isinstance(value, date):
        return value, None
    if value is None or isinstance(value, bool) or not str(value).strip():
        return None, ("missing_required", "date is required.")
    text = str(value).strip()
    if text.startswith("="):
        return None, ("formula_not_supported", "Formula cells cannot be used as dates.")
    try:
        parsed = datetime.strptime(text, profile.date_format).date()
    except ValueError:
        return None, (
            "invalid_date",
            f"Date does not match the configured format {profile.date_format!r}.",
        )
    if parsed.strftime(profile.date_format) != text:
        return None, (
            "invalid_date",
            f"Date does not match the configured format {profile.date_format!r} exactly.",
        )
    return parsed, None


def _parse_decimal(
    value, field: str, profile: MappingProfile, maximum_places: int
) -> tuple[Decimal | None, tuple[str, str] | None]:
    if value is None or isinstance(value, bool) or (isinstance(value, str) and not value.strip()):
        return None, ("missing_required", f"{field} is required.")
    if isinstance(value, str):
        text = value.strip()
        if text.startswith("="):
            return None, ("formula_not_supported", f"Formula cells cannot be used as {field}.")
        other_separator = "." if profile.decimal_separator == "," else ","
        if other_separator in text:
            return None, (
                "invalid_number",
                f"{field} contains an unexpected decimal or grouping separator.",
            )
        text = text.replace(profile.decimal_separator, ".")
    else:
        text = str(value)
    try:
        number = Decimal(text)
    except (InvalidOperation, ValueError):
        return None, ("invalid_number", f"{field} is not a valid decimal number.")
    if not number.is_finite():
        return None, ("nonfinite_number", f"{field} must be finite.")
    fractional_places = max(0, -number.as_tuple().exponent)
    if fractional_places > maximum_places:
        return (
            None,
            (
                "excess_precision",
                f"{field} has {fractional_places} decimal places; at most "
                f"{maximum_places} are allowed.",
            ),
        )
    return number, None


def _validate_row(
    row: SourceRow, profile: MappingProfile
) -> tuple[NormalizedRow | None, ProcessingException | None]:
    values = row.values
    reasons: list[str] = []

    parsed_date, error = _parse_date(values.get("date"), profile)
    if error:
        reasons.extend(error)
    branch, error = _required_text(values.get("branch"), "branch")
    if error:
        reasons.extend(error)
    sku, error = _required_text(values.get("sku"), "sku")
    if error:
        reasons.extend(error)
    transaction_id, error = _required_text(values.get("transaction_id"), "transaction_id")
    if error:
        reasons.extend(error)
    quantity, error = _parse_decimal(
        values.get("quantity"), "quantity", profile, profile.quantity_decimal_places
    )
    if error:
        reasons.extend(error)
    unit_price, error = _parse_decimal(
        values.get("unit_price"), "unit_price", profile, profile.price_decimal_places
    )
    if error:
        reasons.extend(error)
    elif unit_price is not None and unit_price < 0:
        reasons.extend(("negative_unit_price", "unit_price cannot be negative."))

    if reasons:
        return None, ProcessingException(
            category="validation",
            source=row.source,
            reasons=tuple(reasons),
            original_values=dict(values),
        )
    assert parsed_date is not None
    assert branch is not None and sku is not None and transaction_id is not None
    assert quantity is not None and unit_price is not None
    return (
        NormalizedRow(
            date=parsed_date,
            branch=branch,
            sku=sku,
            quantity=quantity,
            unit_price=unit_price,
            transaction_id=transaction_id,
            source=row.source,
            original_values=dict(values),
        ),
        None,
    )


def validate_rows(mapped: MappedFile, profile: MappingProfile) -> ValidationResult:
    """Validate each mapped row independently and retain every rejection reason."""
    if mapped.errors:
        return ValidationResult(source=mapped.source, file_errors=mapped.errors)

    accepted = []
    rejected = []
    for row in mapped.rows:
        normalized, error = _validate_row(row, profile)
        if error is not None:
            rejected.append(error)
        else:
            assert normalized is not None
            accepted.append(normalized)
    return ValidationResult(
        source=mapped.source,
        accepted=tuple(accepted),
        rejected=tuple(rejected),
    )
