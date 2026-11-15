# Excel and CSV Processing Automation

A standalone local application for consolidating branch sales reports. **Development is in progress.**
Only the package setup and Russian application shell are currently available. File import,
validation, duplicate detection, consolidation, and exports are not implemented yet.

See [requirements.md](requirements.md), [plan.md](plan.md), and [progress.md](progress.md)
for scope, the November 15-28 implementation schedule, and actual verification results.

## Setup

Install Python 3.12 and [uv](https://docs.astral.sh/uv/getting-started/installation/), then run
from this repository:

```powershell
uv sync --frozen
uv run excel-csv-automation --check
uv run excel-csv-automation
```

Open the localhost URL printed by Streamlit. The launcher binds to `127.0.0.1`.
No platform account, credentials, paid services, or other portfolio project is required.
`uv.lock` pins application and development dependencies. Do not edit `requirements.md`
as a Python dependency file; it is the business specification.

## Checks

```powershell
uv run ruff check .
uv run ruff format --check .
uv run excel-csv-automation --check
```

Behavior tests and their `uv run pytest` command will be added with the processing features.
The initial CI checks cover lint, formatting, dependency imports, and wheel packaging.

## Demo defaults and limitations

The agreed demo defaults in `plan.md` assume one currency per batch, explicit input date
and numeric formats, signed quantities for returns, and nonnegative unit prices. Monetary
calculations will use `Decimal` with two-place `ROUND_HALF_UP` rounding per line. Duplicate
identity defaults to `(branch, transaction_id)` with configurable keys. These are planned
rules, not functionality delivered by this initial shell.

No production hosting, authentication, ERP integration, inventory accounting, macro execution,
or legacy `.xls` support is included. Sample fixtures will be synthetic. Keep real operator
files in ignored `local_data/`, local profiles in `profiles/`, and generated reports in
`outputs/`. Do not commit secrets or customer records.

## Для оператора

Запустите приложение командой `uv run excel-csv-automation` и откройте локальную ссылку
в браузере. Сейчас доступен только начальный экран. Обработка файлов и скачивание отчётов
появятся на следующих этапах разработки.
