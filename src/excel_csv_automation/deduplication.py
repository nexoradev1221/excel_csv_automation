"""Stable duplicate detection across validated transaction rows."""

from collections.abc import Iterable

from excel_csv_automation.models import (
    DeduplicationResult,
    MappingProfile,
    NormalizedRow,
    ProcessingException,
)


def _payload(row: NormalizedRow) -> tuple:
    """Business values used to identify conflicting duplicate payloads."""
    return (
        row.date,
        row.branch,
        row.sku,
        row.quantity,
        row.unit_price,
        row.transaction_id,
    )


def deduplicate_rows(rows: Iterable[NormalizedRow], profile: MappingProfile) -> DeduplicationResult:
    """Retain the first valid occurrence in file/row order and report later ones.

    Only validated rows are accepted. This lets a later valid occurrence survive
    when an earlier row with the same business key failed validation.
    """
    ordered = sorted(
        rows,
        key=lambda row: (
            row.source.input_index,
            row.source.row_number if row.source.row_number is not None else 0,
        ),
    )
    seen: dict[tuple, NormalizedRow] = {}
    retained = []
    duplicates = []
    for row in ordered:
        key = tuple(getattr(row, field) for field in profile.duplicate_key)
        previous = seen.get(key)
        if previous is None:
            seen[key] = row
            retained.append(row)
            continue

        conflict = _payload(previous) != _payload(row)
        category = "duplicate_conflict" if conflict else "duplicate"
        code = "duplicate_conflict" if conflict else "duplicate"
        message = (
            "Transaction key already exists with different values; review this row."
            if conflict
            else "Transaction key already exists; the first valid occurrence was retained."
        )
        duplicates.append(
            ProcessingException(
                category=category,
                source=row.source,
                reasons=(code, message),
                original_values=dict(row.original_values),
                retained_source=previous.source,
            )
        )
    return DeduplicationResult(retained=tuple(retained), duplicates=tuple(duplicates))
