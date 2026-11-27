"""Exact line totals, summaries, and run-level reconciliation."""

from collections import defaultdict
from collections.abc import Iterable, Mapping
from decimal import ROUND_HALF_UP, Decimal, localcontext

from excel_csv_automation.models import (
    AcceptedRow,
    DeduplicationResult,
    RunCounts,
    RunResult,
    SummaryRow,
    ValidationResult,
)


def _line_total(quantity: Decimal, unit_price: Decimal) -> Decimal:
    """Multiply without losing input significance, then round once to cents."""
    precision = max(
        28,
        len(quantity.as_tuple().digits) + len(unit_price.as_tuple().digits) + 4,
        quantity.adjusted() + unit_price.adjusted() + 3,
    )
    with localcontext() as context:
        context.prec = precision
        return (quantity * unit_price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _sum_money(amounts: list[Decimal]) -> Decimal:
    """Sum rounded line values using enough precision to avoid context rounding."""
    if not amounts:
        return Decimal("0.00")
    integer_digits = max(1, max(amount.adjusted() + 1 for amount in amounts))
    precision = max(28, integer_digits + len(str(len(amounts))) + 2)
    with localcontext() as context:
        context.prec = precision
        return sum(amounts, Decimal("0.00"))


def calculate_run(
    run_id: str,
    validations: Iterable[ValidationResult],
    deduplication: DeduplicationResult,
    *,
    run_metadata: Mapping[str, str] | None = None,
) -> RunResult:
    """Calculate accepted lines and summaries, retaining exceptions and audit counts."""
    if not run_id.strip():
        raise ValueError("Run ID cannot be blank")
    validations = tuple(validations)
    accepted_input = sum(len(result.accepted) for result in validations)
    rejected_rows = sum(len(result.rejected) for result in validations)
    file_errors = sum(len(result.file_errors) for result in validations)
    if accepted_input != len(deduplication.retained) + len(deduplication.duplicates):
        raise ValueError("Deduplication results do not account for every validated row")

    accepted = tuple(
        AcceptedRow(row=row, line_total=_line_total(row.quantity, row.unit_price))
        for row in deduplication.retained
    )
    branch_month: dict[tuple[str, str], list[Decimal]] = defaultdict(list)
    for item in accepted:
        month = item.row.date.strftime("%Y-%m")
        branch_month[item.row.branch, month].append(item.line_total)
    summaries = tuple(
        SummaryRow(branch=branch, month=month, line_total=_sum_money(amounts))
        for (branch, month), amounts in sorted(branch_month.items())
    )
    grand_total = _sum_money([item.line_total for item in accepted])

    file_exceptions = [error for result in validations for error in result.file_errors]
    row_exceptions = [error for result in validations for error in result.rejected]
    exceptions = sorted(
        [*file_exceptions, *row_exceptions, *deduplication.duplicates],
        key=lambda error: (
            error.source.input_index,
            error.source.row_number if error.source.row_number is not None else 0,
            error.category,
        ),
    )
    counts = RunCounts(
        parsed=accepted_input + rejected_rows,
        accepted=len(accepted),
        rejected=rejected_rows,
        duplicates=len(deduplication.duplicates),
        file_errors=file_errors,
    )
    if counts.parsed != counts.accepted + counts.rejected + counts.duplicates:
        raise ValueError(
            "Run rows do not reconcile across accepted, rejected, and duplicate groups"
        )
    return RunResult(
        run_id=run_id,
        counts=counts,
        accepted=accepted,
        exceptions=tuple(exceptions),
        summaries=summaries,
        grand_total=grand_total,
        run_metadata=dict(run_metadata or {}),
    )
