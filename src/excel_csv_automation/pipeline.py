"""Batch orchestration from uploaded source bytes through report-ready results."""

import json
from dataclasses import dataclass, replace
from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4

from excel_csv_automation.calculations import calculate_run
from excel_csv_automation.deduplication import deduplicate_rows
from excel_csv_automation.importers import import_csv, import_xlsx
from excel_csv_automation.mapping import map_import
from excel_csv_automation.models import (
    CanonicalField,
    MappingProfile,
    RunResult,
    ValidationResult,
)
from excel_csv_automation.validation import validate_rows


@dataclass(frozen=True, slots=True)
class BatchFile:
    filename: str
    data: bytes
    profile: MappingProfile


@dataclass(frozen=True, slots=True)
class BatchOutcome:
    validations: tuple[ValidationResult, ...]
    result: RunResult | None

    @property
    def file_errors(self):
        return tuple(error for item in self.validations for error in item.file_errors)


def _new_run_id() -> str:
    return datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ-") + uuid4().hex[:8]


def process_batch(
    files: tuple[BatchFile, ...] | list[BatchFile],
    duplicate_key: tuple[CanonicalField, ...],
    *,
    allow_partial: bool = False,
    run_id: str | None = None,
) -> BatchOutcome:
    """Run every file in displayed order; require explicit consent for partial batches."""
    files = tuple(files)
    if not files:
        raise ValueError("Select at least one input file")
    # Use the profile contract to validate and normalize the run-level key selection.
    first_profile = files[0].profile
    run_profile = replace(first_profile, duplicate_key=duplicate_key)

    validations = []
    accepted_rows = []
    for input_index, item in enumerate(files):
        extension = Path(item.filename).suffix.lower()
        if extension == ".csv":
            imported = import_csv(item.data, item.filename, item.profile, input_index=input_index)
        elif extension == ".xlsx":
            imported = import_xlsx(item.data, item.filename, item.profile, input_index=input_index)
        else:
            # Reuse the import service's consistent file-level unsupported-type error.
            imported = import_csv(item.data, item.filename, item.profile, input_index=input_index)
        validated = validate_rows(map_import(imported, item.profile), item.profile)
        validations.append(validated)
        accepted_rows.extend(validated.accepted)

    validations = tuple(validations)
    if any(item.file_errors for item in validations) and not allow_partial:
        return BatchOutcome(validations=validations, result=None)

    duplicate_result = deduplicate_rows(accepted_rows, run_profile)
    profile_settings = [
        {
            "file": item.filename,
            "profile_id": item.profile.profile_id,
            "column_map": dict(item.profile.column_map),
            "csv_delimiter": item.profile.csv_delimiter,
            "csv_encoding": item.profile.csv_encoding,
            "date_format": item.profile.date_format,
            "decimal_separator": item.profile.decimal_separator,
            "sheet_name": item.profile.sheet_name,
            "quantity_decimal_places": item.profile.quantity_decimal_places,
            "price_decimal_places": item.profile.price_decimal_places,
        }
        for item in files
    ]
    excluded_files = [
        {"file": error.source.filename, "reasons": list(error.reasons)}
        for item in validations
        for error in item.file_errors
    ]
    metadata = {
        "duplicate_key": ", ".join(run_profile.duplicate_key),
        "duplicate_policy": (
            "First valid occurrence by uploaded file order and physical source row; "
            "conflicting payloads are flagged."
        ),
        "profile_settings": json.dumps(profile_settings, ensure_ascii=False, sort_keys=True),
        "excluded_files": json.dumps(excluded_files, ensure_ascii=False, sort_keys=True),
        "partial_processing": str(allow_partial).lower(),
    }
    result = calculate_run(
        run_id or _new_run_id(), validations, duplicate_result, run_metadata=metadata
    )
    return BatchOutcome(validations=validations, result=result)
