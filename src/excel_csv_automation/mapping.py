"""Explicit source-header mapping with source provenance preserved."""

from dataclasses import replace

from excel_csv_automation.models import (
    CANONICAL_FIELDS,
    FileImport,
    MappedFile,
    MappingProfile,
    ProcessingException,
)


def map_import(imported: FileImport, profile: MappingProfile) -> MappedFile:
    """Map one imported file or return a clear file-level missing-column error."""
    if imported.errors:
        return MappedFile(source=imported.source, errors=imported.errors)

    missing = tuple(
        header for header in profile.column_map.values() if header not in imported.headers
    )
    if missing:
        message = "Required mapped source columns are missing: " + ", ".join(missing) + "."
        return MappedFile(
            source=imported.source,
            errors=(
                ProcessingException(
                    category="file",
                    source=replace(imported.source, row_number=None),
                    reasons=("missing_mapped_columns", message),
                ),
            ),
        )

    return MappedFile(
        source=imported.source,
        rows=tuple(
            replace(
                row,
                values={
                    canonical: row.values[source_header]
                    for canonical in CANONICAL_FIELDS
                    if (source_header := profile.column_map[canonical]) in row.values
                },
            )
            for row in imported.rows
        ),
    )
