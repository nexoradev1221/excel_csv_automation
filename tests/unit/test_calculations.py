"""Monetary rounding, aggregation, and run-count reconciliation."""

from datetime import date
from decimal import Decimal

import pytest

from excel_csv_automation.calculations import calculate_run
from excel_csv_automation.models import (
    DeduplicationResult,
    NormalizedRow,
    ProcessingException,
    RunCounts,
    SourceReference,
    ValidationResult,
)


def normalized(filename, row, branch, month, quantity, price, transaction_id):
    return NormalizedRow(
        date=date.fromisoformat(f"{month}-01"),
        branch=branch,
        sku="SKU",
        quantity=Decimal(quantity),
        unit_price=Decimal(price),
        transaction_id=transaction_id,
        source=SourceReference(filename, row),
    )


def test_round_half_up_is_applied_per_line_before_summary():
    first = normalized("a.csv", 2, "North", "2026-10", "1", "1.005", "01")
    second = normalized("a.csv", 3, "North", "2026-10", "-1", "1.005", "02")
    third = normalized("a.csv", 4, "South", "2026-11", "0", "9.99", "03")
    validation = ValidationResult(source=first.source, accepted=(first, second, third))
    run = calculate_run("run-1", [validation], DeduplicationResult(retained=(first, second, third)))
    assert [item.line_total for item in run.accepted] == [
        Decimal("1.01"),
        Decimal("-1.01"),
        Decimal("0.00"),
    ]
    assert run.summaries[0].line_total == Decimal("0.00")
    assert run.summaries[1].line_total == Decimal("0.00")
    assert run.grand_total == Decimal("0.00")


def test_count_partition_includes_duplicates_and_rejected_rows():
    retained = normalized("a.csv", 2, "North", "2026-10", "2", "3", "01")
    duplicate = ProcessingException(
        category="duplicate",
        source=SourceReference("b.csv", 4, input_index=1),
        reasons=("duplicate", "Repeated key"),
        retained_source=retained.source,
    )
    rejection = ProcessingException(
        category="validation",
        source=SourceReference("a.csv", 3),
        reasons=("invalid_date", "Invalid date"),
    )
    validation = ValidationResult(
        source=retained.source,
        accepted=(retained, normalized("a.csv", 4, "North", "2026-10", "2", "3", "01")),
        rejected=(rejection,),
    )
    run = calculate_run(
        "run-2",
        [validation],
        DeduplicationResult(retained=(retained,), duplicates=(duplicate,)),
    )
    assert run.counts == RunCounts(parsed=3, accepted=1, rejected=1, duplicates=1, file_errors=0)
    assert [item.category for item in run.exceptions] == ["validation", "duplicate"]
    assert run.grand_total == Decimal("6.00")


def test_empty_run_has_zero_grand_total_and_no_summaries():
    source = SourceReference("empty.csv", None)
    run = calculate_run("empty", [ValidationResult(source=source)], DeduplicationResult())
    assert run.counts == RunCounts()
    assert run.grand_total == Decimal("0.00")
    assert run.summaries == ()
    assert run.accepted == ()


def test_file_errors_are_separate_from_row_count_partition():
    source = SourceReference("bad.xlsx", None)
    error = ProcessingException(
        category="file", source=source, reasons=("xlsx_read_error", "Cannot read workbook")
    )
    run = calculate_run(
        "file-error",
        [ValidationResult(source=source, file_errors=(error,))],
        DeduplicationResult(),
    )
    assert run.counts == RunCounts(file_errors=1)
    assert run.exceptions == (error,)


def test_row_count_mismatch_is_rejected():
    row = normalized("a.csv", 2, "North", "2026-10", "1", "1", "01")
    with pytest.raises(ValueError, match="do not account"):
        calculate_run(
            "mismatch",
            [ValidationResult(source=row.source, accepted=(row,))],
            DeduplicationResult(),
        )


def test_run_id_must_not_be_blank():
    with pytest.raises(ValueError, match="cannot be blank"):
        calculate_run(" ", [], DeduplicationResult())
