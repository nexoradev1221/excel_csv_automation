"""Rebuild only the named synthetic demo inputs and profiles (never expected_results.json).

Run with `uv run python scripts/generate_sample_data.py`. XLSX fixture authoring uses
openpyxl, the product's required standalone Excel dependency. No external service,
customer asset, or artifact-authoring runtime is required by this repository.
"""

import argparse
import csv
import io
import json
import re
import zipfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

FIELDS = ("date", "branch", "sku", "quantity", "unit_price", "transaction_id")
CANONICAL = dict(zip(FIELDS, FIELDS, strict=True))
ALIASES = {
    "date": "Date",
    "branch": "Office",
    "sku": "Product",
    "quantity": "Units",
    "unit_price": "Price",
    "transaction_id": "Txn",
}
RUSSIAN = {
    "date": "Дата",
    "branch": "Филиал",
    "sku": "Артикул",
    "quantity": "Количество",
    "unit_price": "Цена",
    "transaction_id": "Номер",
}
ALIAS_ORDER = ("transaction_id", "sku", "quantity", "unit_price", "date", "branch")
NORTH_FIRST = (datetime(2026, 10, 3), "Север", "0001", 2, "10.00", "0001")
SOUTH_FIRST = ("03.10.2026", "Юг", "0001", "3", "2,50", "0101")
CENTER_VALID = (datetime(2026, 11, 9), "Центр", "0003", 1, "8.00", "0403")

# Dates, identifiers, amounts, and intentional errors are invented demo content.
SOURCES = [
    {
        "filename": "01_north.xlsx",
        "profile_id": "canonical_xlsx",
        "branch": "Север",
        "rows": [
            NORTH_FIRST,
            (datetime(2026, 10, 4), "Север", "0002", 1.5, "3.335", "0002"),
            (datetime(2026, 11, 1), "Север", "0001", -1, "10.00", "0003"),
            ("31/02/2026", "Север", "0003", 1, "4.00", "0004"),
            NORTH_FIRST,
        ],
    },
    {
        "filename": "02_south.csv",
        "profile_id": "russian_csv",
        "branch": "Юг",
        "rows": [
            SOUTH_FIRST,
            ("04.11.2026", "Юг", "0002", "2", "1,005", "0102"),
            ("05.11.2026", "Юг", "0003", "1", "bad", "0103"),
            ("06.11.2026", "Юг", "", "1", "4,00", "0104"),
            SOUTH_FIRST,
        ],
    },
    {
        "filename": "03_east.xlsx",
        "profile_id": "alias_xlsx",
        "branch": "Восток",
        "rows": [
            (datetime(2026, 10, 5), "Восток", "0001", 4, "1.25", "0201"),
            (datetime(2026, 11, 6), "Восток", "0002", 0.5, "9.99", "0202"),
            (datetime(2026, 11, 7), "Восток", "0003", "NaN", "1.00", "0203"),
            (datetime(2026, 11, 8), "Восток", "0004", 0, "10.00", "0204"),
            (datetime(2026, 10, 5), "Восток", "0001", 5, "1.25", "0201"),
        ],
    },
    {
        "filename": "04_west.csv",
        "profile_id": "alias_csv",
        "branch": "Запад",
        "rows": [
            ("2026-10-07", "Запад", "0003", "2", "6.10", "0301"),
            ("2026-11-07", "Запад", "0004", "3", "0.10", "0302"),
            ("2026-11-08", "Запад", "0005", "1", "-2.00", "0303"),
            ("2026-11-09", "Запад", "0006", "1", "0.00", "0304"),
            ("2026-11-10", "Запад", "0007", "", "3.00", "0305"),
        ],
    },
    {
        "filename": "05_center.xlsx",
        "profile_id": "canonical_xlsx",
        "branch": "Центр",
        "rows": [
            (datetime(2026, 10, 8), "Центр", "0001", 1, "100.005", "0401"),
            (datetime(2026, 11, 8), "Центр", "0002", 2.25, "2.00", "0402"),
            ("2026-99-01", "Центр", "0003", 1, "8.00", "0403"),
            CENTER_VALID,
            CENTER_VALID,
        ],
    },
    {
        "filename": "06_port.csv",
        "profile_id": "canonical_csv",
        "branch": "Порт",
        "rows": [
            ("2026-10-10", "Порт", "0001", "10", "0.10", "0501"),
            ("2026-11-10", "Порт", "0002", "-2", "1.50", "0502"),
            ("2026-11-11", "Порт", "0003", "1", "0.005", "0503"),
            ("2026-10-03", "Север", "0001", "2", "10.00", "0001"),
            ("2026-11-12", "Порт", "=DEMO()", "1", "2.00", "0001"),
        ],
    },
]


def profile(profile_id, column_map, **settings):
    return {
        "schema_version": 1,
        "profile_id": profile_id,
        "column_map": column_map,
        "csv_delimiter": ",",
        "csv_encoding": "utf-8-sig",
        "date_format": "%Y-%m-%d",
        "decimal_separator": ".",
        "sheet_name": None,
        "duplicate_key": ["branch", "transaction_id"],
        "quantity_decimal_places": 3,
        "price_decimal_places": 4,
        **settings,
    }


PROFILES = {
    "canonical_xlsx": profile("canonical_xlsx", CANONICAL, sheet_name="Sales"),
    "russian_csv": profile(
        "russian_csv",
        RUSSIAN,
        csv_delimiter=";",
        csv_encoding="cp1251",
        date_format="%d.%m.%Y",
        decimal_separator=",",
    ),
    "alias_xlsx": profile("alias_xlsx", ALIASES, sheet_name="Orders"),
    "alias_csv": profile("alias_csv", ALIASES, csv_delimiter="|"),
    "canonical_csv": profile("canonical_csv", CANONICAL, csv_encoding="utf-8"),
}


def write_xlsx(path, settings, field_order, rows, *, add_notes=False):
    book = Workbook()
    sheet = book.active
    if add_notes:
        sheet.title = "Notes"
        sheet.append(["Synthetic demo only; select the Sales sheet using the sample profile."])
        sheet = book.create_sheet()
    sheet.title = settings["sheet_name"]
    sheet.append([settings["column_map"][name] for name in field_order])
    for row in rows:
        values = dict(zip(FIELDS, row, strict=True))
        sheet.append([values[name] for name in field_order])
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="284B63")
    for column in sheet.columns:
        sheet.column_dimensions[column[0].column_letter].width = 20
        for cell in column[1:]:
            if isinstance(cell.value, datetime):
                cell.number_format = "yyyy-mm-dd"
    book.properties.creator = "Synthetic demo generator"
    book.properties.created = datetime(2000, 1, 1)
    content = io.BytesIO()
    book.save(content)
    book.close()
    # Normalize metadata/ZIP dates so fixture regeneration has no timestamp churn.
    with zipfile.ZipFile(content) as source, zipfile.ZipFile(path, "w") as target:
        for name in sorted(source.namelist()):
            data = source.read(name)
            if name == "docProps/core.xml":
                data = re.sub(
                    rb"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)",
                    rb"\g<1>2000-01-01T00:00:00Z\g<2>",
                    data,
                )
            info = zipfile.ZipInfo(name, date_time=(2000, 1, 1, 0, 0, 0))
            target.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED)


def generate_sample_data(output_dir, profile_dir):
    output_dir, profile_dir = Path(output_dir), Path(profile_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    profile_dir.mkdir(parents=True, exist_ok=True)
    for name, settings in PROFILES.items():
        (profile_dir / f"{name}.json").write_text(
            json.dumps(settings, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
    batch = []
    for source in SOURCES:
        settings = PROFILES[source["profile_id"]]
        order = ALIAS_ORDER if source["profile_id"].startswith("alias") else FIELDS
        path = output_dir / source["filename"]
        if path.suffix == ".xlsx":
            write_xlsx(
                path,
                settings,
                order,
                source["rows"],
                add_notes=source["filename"] == "05_center.xlsx",
            )
        else:
            with path.open("w", encoding=settings["csv_encoding"], newline="") as stream:
                writer = csv.writer(
                    stream, delimiter=settings["csv_delimiter"], lineterminator="\n"
                )
                writer.writerow([settings["column_map"][name] for name in order])
                for row in source["rows"]:
                    values = dict(zip(FIELDS, row, strict=True))
                    writer.writerow([values[name] for name in order])
        batch.append({key: source[key] for key in ("filename", "profile_id", "branch")})
    (output_dir / "batch.json").write_text(
        json.dumps({"schema_version": 1, "files": batch}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return [output_dir / source["filename"] for source in SOURCES]


def main():
    root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=root / "sample_data")
    parser.add_argument("--profile-dir", type=Path, default=root / "config" / "sample_profiles")
    args = parser.parse_args()
    paths = generate_sample_data(args.output_dir, args.profile_dir)
    print(f"Generated {len(paths)} synthetic branch inputs and {len(PROFILES)} profiles")
    print("Expected results are maintained independently and were not regenerated.")


if __name__ == "__main__":
    main()
