# Excel and CSV Processing Automation

A local Streamlit application that consolidates synthetic branch sales files into an auditable
workbook. The application is intended for a portfolio demonstration. Do not use it for live
business decisions without validating the documented assumptions against real source reports.

## Features

- Import multiple CSV and XLSX files. Choose CSV delimiter and encoding, and select an XLSX sheet.
- Map source headers to date, branch, SKU, quantity, unit price, and transaction ID. Preview source
  and mapped rows before processing.
- Reuse or save local mapping profiles in `profiles/`; bundled examples are in
  `config/sample_profiles/`.
- Validate required fields, explicit date formats, numeric values, finite decimals, decimal
  precision, and nonnegative unit prices. Source rows and rejection reasons remain reviewable.
- Find duplicates across the batch using a configurable canonical key. Keep the first valid row by
  file and physical row order; report duplicates and payload conflicts.
- Calculate each accepted line using `Decimal` and round once to cents with `ROUND_HALF_UP`. Export
  branch/month summaries, accepted detail, audit counts, and exception references.
- Download `consolidated_sales.xlsx` and a UTF-8 BOM `exceptions.csv`. Formula-like text is kept
  literal in the workbook and escaped in the CSV.

## Setup and launch

Install Python 3.12 and [uv](https://docs.astral.sh/uv/getting-started/installation/), then run from
the repository root:

```powershell
uv sync --frozen
uv run excel-csv-automation --check
uv run excel-csv-automation
```

Open the localhost URL printed by Streamlit. The server binds to `127.0.0.1`; input files are
processed in memory and are not sent to an external service. Local profiles are written under
`profiles/`, which is ignored by Git. The application does not overwrite source files.

## Operator workflow

1. Upload the branch CSV/XLSX files together.
2. Select a sample or local profile for each file. For XLSX, select the correct worksheet; for CSV,
   choose delimiter and encoding.
3. Review the source preview and map the six canonical fields to the source headers. Save a profile
   if the mapping should be reused later.
4. Choose the duplicate key for the batch. Review the file-level errors before processing. If one or
   more files failed, the application requires the operator to opt in before processing only the
   readable files.
5. Review row counts, totals, summaries, and exceptions. Download both reports.

The supplied demo inputs are synthetic. Start with all six files and their corresponding sample
profiles to reproduce the independent acceptance oracle.

## Checks

```powershell
uv run pytest -q
uv run ruff check .
uv run ruff format --check .
uv run excel-csv-automation --check
```

The tests cover CSV/XLSX import, profile persistence, mapping, validation, duplicate retention,
financial calculations, and reopened workbook/CSV outputs. The CI workflow also checks dependency
imports, wheel packaging, and Streamlit rendering on Windows and Linux. See `progress.md` for the
actual results and PR/CI records; application acceptance is only claimed when supported by those
recorded results.

## Demo rules and limitations

- Duplicate identity defaults to `(branch, transaction_id)`. A different key may be chosen for the
  full batch. The first valid occurrence is retained, and a later conflicting payload is flagged.
- Dates must follow the selected format unless the XLSX cell contains a native date. Numbers use
  the selected decimal separator and configured precision. Signed quantities represent returns;
  unit prices cannot be negative. One currency per batch is assumed.
- XLSX formulas are not evaluated. Formula cells in date/numeric fields are rejected; text cells
  are exported literally.
- Profiles and reports are local. The application has no authentication, hosted service, ERP
  integration, inventory accounting, macro execution, or legacy `.xls` support.
- Fixtures contain synthetic data only. Keep operator files in ignored `local_data/`, profiles in
  ignored `profiles/`, and local outputs in ignored `outputs/`. Never commit credentials or genuine
  customer records.

## Synthetic demo data

See [sample_data/README.md](sample_data/README.md) and `sample_data/batch.json` for the six inputs,
profiles, deliberate errors, and hand-specified expected results. Rebuild the inputs with:

```powershell
uv run python scripts/generate_sample_data.py
```

The independent oracle specifies 30 input rows, 18 accepted rows, 7 rejected rows, 5 duplicates,
and an accepted grand total of `159.53`.

## Для оператора

Загрузите CSV- и XLSX-файлы филиалов, выберите профиль для каждого файла, проверьте лист Excel,
кодировку и разделитель CSV, затем сопоставьте исходные заголовки с шестью обязательными полями.
Перед обработкой проверьте предварительный просмотр и ключ дублей. Если есть ошибки файлов,
частичная обработка запускается только после явного выбора оператора. После обработки проверьте
количество принятых строк, итоги и список исключений, затем скачайте обе версии отчёта.

Профили сохраняются локально. Примеры находятся в `config/sample_profiles/`. Входные данные
демонстрации синтетические; для реальных отчётов сначала проверьте допущения о датах, денежных
значениях, дублях и возвратах.
