"""Framework-independent contracts shared by import, validation, and report services.

Values remain unparsed in SourceRow. Only validated dates and exact Decimal values
belong in NormalizedRow. Line totals and summaries are populated by later services.
"""

from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import date as Date
from datetime import datetime
from decimal import Decimal
from typing import Literal

type CanonicalField = Literal["date", "branch", "sku", "quantity", "unit_price", "transaction_id"]
CANONICAL_FIELDS: tuple[CanonicalField, ...] = (
    "date",
    "branch",
    "sku",
    "quantity",
    "unit_price",
    "transaction_id",
)
type CellValue = str | int | float | bool | Date | datetime | Decimal | None
type ExceptionCategory = Literal["validation", "duplicate", "duplicate_conflict", "file"]


@dataclass(frozen=True, slots=True)
class MappingProfile:
    """Version 1 profile; column_map maps canonical names to source header names."""

    profile_id: str
    column_map: Mapping[str, str]
    csv_delimiter: str = ","
    csv_encoding: str = "utf-8-sig"
    date_format: str = "%Y-%m-%d"
    decimal_separator: str = "."
    sheet_name: str | None = None
    duplicate_key: tuple[CanonicalField, ...] = ("branch", "transaction_id")
    quantity_decimal_places: int = 3
    price_decimal_places: int = 4
    schema_version: int = 1

    def __post_init__(self) -> None:
        if self.schema_version != 1:
            raise ValueError("Unsupported mapping profile schema version")
        if not self.profile_id.strip():
            raise ValueError("Profile ID cannot be blank")
        if set(self.column_map) != set(CANONICAL_FIELDS):
            raise ValueError("Map each of the six canonical fields exactly once")
        headers = list(self.column_map.values())
        if any(not header.strip() for header in headers) or len(set(headers)) != len(headers):
            raise ValueError("Source headers must be nonblank and unique")
        if len(self.csv_delimiter) != 1 or self.csv_delimiter in '\r\n"':
            raise ValueError("CSV delimiter must be one non-newline, non-quote character")
        if self.decimal_separator not in (".", ","):
            raise ValueError("Decimal separator must be a dot or comma")
        if not self.csv_encoding.strip() or not self.date_format.strip():
            raise ValueError("Encoding and date format cannot be blank")
        if not self.duplicate_key or len(set(self.duplicate_key)) != len(self.duplicate_key):
            raise ValueError("Duplicate key must contain distinct canonical fields")
        if not set(self.duplicate_key).issubset(CANONICAL_FIELDS):
            raise ValueError("Duplicate key contains an unknown canonical field")
        if self.quantity_decimal_places < 0 or self.price_decimal_places < 0:
            raise ValueError("Decimal precision cannot be negative")


@dataclass(frozen=True, slots=True)
class SourceReference:
    """Physical row in an input; row_number=None denotes a whole-file error.

    input_index disambiguates uploaded files with identical names and records order.
    Row numbers include the header: first data row is normally 2.
    """

    filename: str
    row_number: int | None
    sheet_name: str | None = None
    input_index: int = 0

    def __post_init__(self) -> None:
        if not self.filename.strip():
            raise ValueError("Source filename cannot be blank")
        if self.row_number is not None and self.row_number < 1:
            raise ValueError("Source row numbers are one-based")
        if self.input_index < 0:
            raise ValueError("Input index cannot be negative")


@dataclass(frozen=True, slots=True)
class SourceRow:
    source: SourceReference
    values: Mapping[str, CellValue]


@dataclass(frozen=True, slots=True)
class NormalizedRow:
    date: Date
    branch: str
    sku: str
    quantity: Decimal
    unit_price: Decimal
    transaction_id: str
    source: SourceReference
    original_values: Mapping[str, CellValue] = field(default_factory=dict)


@dataclass(frozen=True, slots=True)
class AcceptedRow:
    row: NormalizedRow
    line_total: Decimal


@dataclass(frozen=True, slots=True)
class ProcessingException:
    category: ExceptionCategory
    source: SourceReference
    reasons: tuple[str, ...]
    original_values: Mapping[str, CellValue] = field(default_factory=dict)
    retained_source: SourceReference | None = None


@dataclass(frozen=True, slots=True)
class SummaryRow:
    branch: str
    month: str
    line_total: Decimal


@dataclass(frozen=True, slots=True)
class RunCounts:
    parsed: int = 0
    accepted: int = 0
    rejected: int = 0
    duplicates: int = 0
    file_errors: int = 0


@dataclass(frozen=True, slots=True)
class RunResult:
    run_id: str
    counts: RunCounts
    accepted: tuple[AcceptedRow, ...] = ()
    exceptions: tuple[ProcessingException, ...] = ()
    summaries: tuple[SummaryRow, ...] = ()
    grand_total: Decimal = Decimal("0.00")


@dataclass(frozen=True, slots=True)
class FileImport:
    """One file's raw import. Errors exclude the file: rows must then be empty."""

    source: SourceReference
    headers: tuple[str, ...] = ()
    rows: tuple[SourceRow, ...] = ()
    errors: tuple[ProcessingException, ...] = ()

    def preview(self, max_rows: int = 5) -> tuple[SourceRow, ...]:
        if max_rows < 0:
            raise ValueError("Preview size cannot be negative")
        return self.rows[:max_rows]


@dataclass(frozen=True, slots=True)
class MappedFile:
    """One successfully mapped input, or a file-level mapping error."""

    source: SourceReference
    rows: tuple[SourceRow, ...] = ()
    errors: tuple[ProcessingException, ...] = ()


@dataclass(frozen=True, slots=True)
class ValidationResult:
    """Partition mapped rows into normalized rows and row/file exceptions."""

    source: SourceReference
    accepted: tuple[NormalizedRow, ...] = ()
    rejected: tuple[ProcessingException, ...] = ()
    file_errors: tuple[ProcessingException, ...] = ()
