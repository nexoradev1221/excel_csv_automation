# Synthetic six-branch acceptance dataset

Every value is invented for this demo. There are no customer records or external assets.
Use the order in `batch.json`; its profile IDs refer to `config/sample_profiles/*.json`.
Each input has a header followed by five physical source rows numbered 2-6.

| File | Profile | Format | Important cases |
| --- | --- | --- | --- |
| 01_north.xlsx | canonical_xlsx | Sales sheet, native dates | Return, invalid date, duplicate |
| 02_south.csv | russian_csv | Windows-1251, semicolon, comma decimals | Cyrillic headers, bad price, missing SKU, duplicate |
| 03_east.xlsx | alias_xlsx | Orders sheet, reordered English headers | Non-finite quantity, zero quantity, conflicting duplicate |
| 04_west.csv | alias_csv | UTF-8 BOM, pipe, English aliases | Negative price, zero price, missing quantity |
| 05_center.xlsx | canonical_xlsx | Sales sheet after Notes | Invalid occurrence before valid ID, duplicate, half-cent rounding |
| 06_port.csv | canonical_csv | UTF-8, comma | Cross-file copied North row, return, same ID in a different branch |

The Port file deliberately contains a copied North transaction to exercise cross-file
business keys. Row 6 also has a formula-like SKU (`=DEMO()`), which is literal input text
and must remain non-executable when reports are exported. Identifiers have leading zeroes.

## Independent expected results

`expected_results.json` is a hand-specified oracle, not output from the generator or the
future processing pipeline. It lists every accepted row and rounded amount, each rejection
reason, each duplicate's retained source, and every branch/month summary.

Expected classification: **30 parsed = 18 accepted + 7 rejected + 5 duplicates**, with no
file-level errors. Expected accepted total: **159.53** (October **150.71**, November **8.82**).
These are test targets; application acceptance tests have not yet passed because processing
services are still being implemented. Quantities permit three decimal places; prices permit
four; line totals round to two places with ROUND_HALF_UP before aggregation.

For example, 0.5 x 9.99 = 4.995 rounds to 5.00; 1 x 100.005 rounds to 100.01;
-2 x 1.50 = -3.00. The first valid Center ID 0403 is row 5, not its invalid row 4.

## Regeneration

```powershell
uv run python scripts/generate_sample_data.py
uv run pytest
```

The generator overwrites only the six named demo inputs, batch metadata, and five sample
profiles. It never regenerates expected_results.json. Use `--output-dir` and `--profile-dir`
with disposable folders when experimenting. XLSX document and ZIP timestamps are fixed to
avoid unnecessary binary diffs. No report calculations or external services are used.
