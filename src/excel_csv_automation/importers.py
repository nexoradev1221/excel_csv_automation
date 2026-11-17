"""Strict, auditable CSV import without value coercion or partial-file acceptance.

Headers are trimmed; cells remain original text (including leading zeroes and
empty strings). Physical blank lines are ignored, but explicit empty records are
retained for later validation. Mapping and value validation belong to later stages.
"""

import csv
import io
from dataclasses import replace
from pathlib import Path

from excel_csv_automation.models import (
    FileImport,
    MappingProfile,
    ProcessingException,
    SourceReference,
    SourceRow,
)


def _failure(source, code, message, *, headers=(), row_number=None, values=None):
    return FileImport(
        source=source,
        headers=headers,
        errors=(
            ProcessingException(
                category="file",
                source=replace(source, row_number=row_number),
                reasons=(code, message),
                original_values={} if values is None else values,
            ),
        ),
    )


def import_csv(
    data: bytes, filename: str, profile: MappingProfile, *, input_index: int = 0
) -> FileImport:
    """Parse upload bytes using explicit settings; exclude the whole file on errors.

    Source references use the first physical line of each record, even for quoted
    multiline cells. A malformed record invalidates all parsed rows in this file;
    a future batch orchestrator must explicitly handle partial-batch acceptance.
    Header-only input is a successful empty import, distinguishable from no header.
    """
    source = SourceReference(filename=filename, row_number=None, input_index=input_index)
    if Path(filename).suffix.lower() != ".csv":
        return _failure(source, "unsupported_file_type", "Select a CSV file for CSV import.")
    try:
        text = data.decode(profile.csv_encoding, errors="strict").removeprefix("\ufeff")
    except LookupError:
        return _failure(
            source, "unknown_encoding", f"Unknown CSV encoding: {profile.csv_encoding}."
        )
    except UnicodeDecodeError:
        return _failure(
            source,
            "decode_error",
            f"Cannot decode CSV using {profile.csv_encoding}; select the matching encoding.",
        )
    if "\x00" in text:
        return _failure(
            source,
            "binary_or_wrong_encoding",
            "Decoded CSV contains NUL characters; check the file format and encoding.",
        )

    reader = csv.reader(io.StringIO(text, newline=""), delimiter=profile.csv_delimiter, strict=True)
    headers: tuple[str, ...] = ()
    rows = []
    record_start = 1
    try:
        raw_headers = next(reader, None)
        if not raw_headers:
            return _failure(source, "missing_header", "CSV must start with a non-empty header row.")
        headers = tuple(header.strip() for header in raw_headers)
        if any(not header for header in headers):
            return _failure(
                source, "blank_header", "Every CSV column needs a name.", headers=headers
            )
        if len(set(headers)) != len(headers):
            return _failure(
                source,
                "duplicate_header",
                "CSV header names must be unique after trimming whitespace.",
                headers=headers,
            )
        while True:
            record_start = reader.line_num + 1
            fields = next(reader, None)
            if fields is None:
                break
            if not fields:  # No data record; preserve physical line numbering for the next row.
                continue
            if len(fields) != len(headers):
                return _failure(
                    source,
                    "row_width_mismatch",
                    f"CSV record at line {record_start} has {len(fields)} cells; "
                    f"expected {len(headers)}. Check the delimiter and quoting.",
                    headers=headers,
                    row_number=record_start,
                    values={f"column_{index + 1}": value for index, value in enumerate(fields)},
                )
            rows.append(
                SourceRow(
                    source=replace(source, row_number=record_start),
                    values=dict(zip(headers, fields, strict=True)),
                )
            )
    except csv.Error as error:
        return _failure(
            source,
            "csv_parse_error",
            f"CSV syntax error near line {reader.line_num}: {error}.",
            headers=headers,
            row_number=record_start,
        )
    return FileImport(source=source, headers=headers, rows=tuple(rows))


def import_csv_file(
    path: str | Path, profile: MappingProfile, *, input_index: int = 0
) -> FileImport:
    """Read a local CSV without writing or changing its source file."""
    path = Path(path)
    source = SourceReference(filename=path.name, row_number=None, input_index=input_index)
    if path.suffix.lower() != ".csv":
        return _failure(source, "unsupported_file_type", "Select a CSV file for CSV import.")
    try:
        data = path.read_bytes()
    except OSError:
        return _failure(source, "file_read_error", "Cannot read this CSV file; check its access.")
    return import_csv(data, path.name, profile, input_index=input_index)
