"""Versioned, local JSON mapping-profile persistence."""

import json
import os
import re
import tempfile
from pathlib import Path

from excel_csv_automation.models import MappingProfile

_PROFILE_ID = re.compile(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}\Z")


class ProfileError(ValueError):
    """A profile name or stored profile cannot be safely used."""


def _profile_path(profile_id: str, directory: str | Path) -> Path:
    if not isinstance(profile_id, str) or not _PROFILE_ID.fullmatch(profile_id):
        raise ProfileError("Profile ID must use 1-64 ASCII letters, digits, '_' or '-'.")
    return Path(directory) / f"{profile_id}.json"


def save_profile(profile: MappingProfile, directory: str | Path = "profiles") -> Path:
    """Atomically save a validated profile in a local JSON file."""
    target = _profile_path(profile.profile_id, directory)
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.is_symlink():
        raise ProfileError("Profile path cannot be a symbolic link.")

    content = {
        "schema_version": profile.schema_version,
        "profile_id": profile.profile_id,
        "column_map": dict(profile.column_map),
        "csv_delimiter": profile.csv_delimiter,
        "csv_encoding": profile.csv_encoding,
        "date_format": profile.date_format,
        "decimal_separator": profile.decimal_separator,
        "sheet_name": profile.sheet_name,
        "duplicate_key": list(profile.duplicate_key),
        "quantity_decimal_places": profile.quantity_decimal_places,
        "price_decimal_places": profile.price_decimal_places,
    }
    temporary_path: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w", encoding="utf-8", newline="\n", dir=target.parent, delete=False
        ) as stream:
            temporary_path = stream.name
            json.dump(content, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_path, target)
    except OSError as error:
        if temporary_path is not None:
            Path(temporary_path).unlink(missing_ok=True)
        raise ProfileError(f"Could not save profile '{profile.profile_id}'.") from error
    return target


def load_profile(profile_id: str, directory: str | Path = "profiles") -> MappingProfile:
    """Load and validate one profile by safe identifier, never by arbitrary filename."""
    path = _profile_path(profile_id, directory)
    if path.is_symlink():
        raise ProfileError("Profile path cannot be a symbolic link.")
    try:
        content = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(content, dict):
            raise ProfileError("Profile file must contain a JSON object.")
        if content.get("profile_id") != profile_id:
            raise ProfileError("Stored profile ID does not match its filename.")
        if not isinstance(content.get("duplicate_key"), list):
            raise ProfileError("Stored duplicate_key must be a list.")
        content["duplicate_key"] = tuple(content["duplicate_key"])
        return MappingProfile(**content)
    except OSError as error:
        raise ProfileError(f"Could not read profile '{profile_id}'.") from error
    except (json.JSONDecodeError, TypeError, ValueError) as error:
        if isinstance(error, ProfileError):
            raise
        raise ProfileError(f"Profile '{profile_id}' is invalid or unsupported.") from error


def list_profile_ids(directory: str | Path = "profiles") -> tuple[str, ...]:
    """List safely named local profiles in stable case-insensitive order."""
    root = Path(directory)
    if not root.exists():
        return ()
    if not root.is_dir():
        raise ProfileError("Profile storage path must be a directory.")
    return tuple(
        sorted(
            (path.stem for path in root.glob("*.json") if _PROFILE_ID.fullmatch(path.stem)),
            key=str.casefold,
        )
    )
