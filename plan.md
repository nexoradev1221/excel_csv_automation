# Implementation Plan: Excel and CSV Processing Automation

Status: planning only; application implementation has not started.  
Requirements: `requirements.md`, DEMO-SRS-01, version 1.0 (2026-11-14).  
Target repository: `nexoradev1221/excel_csv_automation`.

## 1. Requirements analysis

The product must let an office operator consolidate six branches' weekly files into an auditable monthly sales workbook without changing code. The main risks are incorrectly parsing source values, counting duplicate transactions, and producing summaries that disagree with accepted detail rows.

The baseline covers XLSX/CSV import, reusable column mappings, validation, duplicate detection, calculation, summaries, exports, and visible processing results. It also includes synthetic demo inputs, automated tests, actual application screenshots, and documentation. This plan describes future implementation; it does not claim those deliverables already exist.

The background mentions returns and inventory, but the specified canonical fields and acceptance tests define a transaction consolidation workflow. Inventory balance calculations, stock reconciliation, and a separate returns subsystem are outside the baseline. Signed quantities can represent returns under the provisional rules below.

### Requirement coverage

| Requirement | Planned behavior | Verification |
| --- | --- | --- |
| FR-01 | Upload several XLSX/CSV files; choose delimiter and encoding for CSV and sheet for XLSX per file/profile; preview parsed headers and rows. | Parser tests covering formats, encodings, sheets, and malformed files. |
| FR-02 | Map source headers to date, branch, SKU, quantity, unit price, and transaction ID; preview normalized rows. | Tests for multiple layouts, missing mappings, and conflicting mappings. |
| FR-03 | Check schema and values; retain rejected source rows and explicit reasons. | Invalid-date, invalid-number, blank-value, and missing-column tests; AT-03. |
| FR-04 | Apply configurable business keys across the entire run; retain one accepted occurrence and list duplicates with their retained source reference. | Within-file and cross-file duplicate tests; AT-02. |
| FR-05 | Calculate accepted line totals and branch/month summaries using exact decimal arithmetic. | Independently specified expected totals and reconciliation checks; AT-04. |
| FR-06 | Export a formatted workbook containing cleaned detail, summaries, and exceptions, plus a separate exception report. | Reopen workbook and verify values, sheets, formatting, and source references. |
| FR-07 | Save, reload, and edit local mapping profiles; show processing stages, counts, and download controls. | Profile round-trip tests and an operator workflow; AT-01. |

## 2. Proposed architecture

- Python application with a local Streamlit browser UI, pandas for tabular import/transform work, and openpyxl for XLSX handling and presentation.
- Framework-independent processing services so calculations and validation can be tested without the UI.
- Local versioned JSON profiles and configuration; no database, hosted service, paid API, or credentials required for the demo.
- Bind the application to localhost by default. Use per-run temporary storage and explicit download/save actions; exclude operator files and outputs from Git.
- Use `pyproject.toml` for Python dependencies and tool configuration, plus a reproducible dependency lock after selecting compatible versions. Keep the existing `requirements.md` as the business specification.
- Check Python/library compatibility during implementation before selecting exact versions. This plan does not prescribe unverified current version numbers.

### Planned repository structure

```text
requirements.md
plan.md
README.md
pyproject.toml
<dependency lock file>
.gitignore
src/excel_csv_automation/
  app.py
  models.py
  importers.py
  mapping.py
  validation.py
  deduplication.py
  calculations.py
  exports.py
  profiles.py
  pipeline.py
config/sample_profiles/
sample_data/                  # Six synthetic branch files and expected results
scripts/generate_sample_data.py
tests/
  unit/
  integration/
docs/screenshots/
.github/workflows/tests.yml
```

## 3. Data contract and processing rules

These are proposed demo defaults for unresolved business choices. Document them in the README and sample profiles; validate them with the client before commercial use.

1. **Canonical fields:** `date`, `branch`, `sku`, `quantity`, `unit_price`, and `transaction_id` are required. Preserve identifiers as text, including leading zeroes; trim surrounding whitespace without silently changing case.
2. **Provenance:** attach source filename, sheet (where applicable), physical source row number, and processing-run identifier before transformation. Preserve original cell values for rejected rows and duplicate review.
3. **Parsing:** explicit profile settings control date format, decimal separator, CSV delimiter, and encoding. Support UTF-8/UTF-8 BOM and a Windows-1251 sample. Reject ambiguous dates and lossy decoding rather than guessing silently. Accept native XLSX date values.
4. **Schema errors:** report unreadable files and missing/ambiguous columns as file-level exceptions. Exclude affected files from totals and require an explicit operator choice to proceed with a partial batch; display the omission in results and exports.
5. **Row validation:** reject missing required values, non-finite or malformed numbers, invalid dates, and negative unit prices. Proposed demo policy permits signed decimal quantities for returns. Zero quantity and zero price are allowed but must be documented. Require explicit profile precision rules rather than silently rounding excessive input precision.
6. **Money:** parse validated numeric text into `Decimal`, avoid binary-float aggregation, calculate each line as quantity times unit price, and round once per line to two decimal places using `ROUND_HALF_UP`. Sum those accepted rounded lines for every summary. Assume one currency per batch; do not combine currencies or implement exchange conversion.
7. **Duplicates:** default to `(branch, transaction_id)`, with configurable non-empty canonical key fields. This assumes one transaction ID identifies one line within a branch; datasets with repeated IDs for different SKUs need an explicitly adjusted key. Validate rows first so an invalid occurrence cannot suppress a later valid one. Retain the first valid occurrence in the displayed file order and source-row order. Flag repeated keys with differing payloads as conflicts requiring review, while excluding later occurrences from totals.
8. **Audit counts:** every parsed data row ends in exactly one accepted, rejected, or duplicate category; a row may have several validation reasons. File-level failures are counted separately. Show counts and reconciliation in both results and exported run metadata.
9. **Summaries:** derive calendar month from the validated date; aggregate accepted rounded totals by branch and month, with a grand total. Empty accepted sets produce a clear result and an exceptions download without suggesting a successful sales consolidation.
10. **Safe export:** write untrusted text as literal spreadsheet text to prevent formula execution; do not evaluate source formulas. Reject formulas in required numeric/date cells unless a reliable cached literal is explicitly supported and documented. Never overwrite the operator's source files.

### Output contract

- `consolidated_sales.xlsx`: `Cleaned_Data`, `Summary`, `Exceptions`, and `Run_Info` sheets. Include headers, filters, frozen header rows, readable widths, date formats, and monetary formats.
- `exceptions.csv`: UTF-8 BOM report with category, source references, original row values, validation reason codes/messages, and retained-row references for duplicates. Ensure exported CSV text cannot become executable spreadsheet formulas.
- Both exports are downloadable from the UI; a documented local save workflow is sufficient. Generated timestamps/run IDs may vary, but classification and financial results must be deterministic.

## 4. Implementation phases

### Phase 1: Project foundation and demo contract

- Create the standalone package, dependency configuration, ignores for secrets/local data, and basic test workflow.
- Define typed profile, normalized-row, exception, and run-result structures.
- Generate six synthetic branch files spanning XLSX/CSV, at least two layouts, UTF-8 and Windows-1251, multiple months, invalid rows, and cross-file duplicates.
- Commit a hand-reviewed expected-results manifest with accepted/rejected/duplicate counts and exact branch/month totals. Keep expected values independent of production calculation code.

Exit criterion: fixtures, schemas, and baseline business rules are documented and reproducible without any other portfolio project.

### Phase 2: Import, mapping, and profiles

- Implement per-file import settings, sheet enumeration/selection, header preview, and actionable parse failures.
- Build mapping validation and transformations while preserving source references.
- Implement versioned JSON profile save/load and sample profiles; validate profile contents and sanitize filenames.

Exit criterion: all six inputs normalize through configuration, with meaningful errors for unsupported layouts and settings.

### Phase 3: Validation, duplicates, and calculations

- Implement schema checks, explicit parsing, row validation, deterministic duplicate classification, and conflict reasons.
- Implement exact line totals and branch/month aggregation.
- Add run counts and enforce detail-to-summary and row-count reconciliation invariants.

Exit criterion: automated checks demonstrate AT-02, AT-03, and AT-04 against the fixed expected-results manifest.

### Phase 4: Report generation

- Build formatted XLSX exports and the separate exception CSV with source-row traceability and literal-text safeguards.
- Include selected profile settings, duplicate key/policy, excluded files, row counts, rounding policy, and reconciliation results in run metadata.
- Verify outputs by reopening them with openpyxl and a CSV reader; inspect formatting using a representative generated workbook.

Exit criterion: export content reconciles exactly and exceptions remain reviewable outside the application.

### Phase 5: Operator interface

- Provide Russian customer-facing wording for upload, preview, mapping, settings, progress, validation feedback, and downloads.
- Guide the operator through upload -> select/reuse profiles -> review -> process -> inspect counts/exceptions -> export.
- Prevent accidental reruns from changing retention order or creating inconsistent results; keep run state explicit.
- Handle malformed files, empty batches, all-rejected data, and output/save failures with clear recoverable messages.

Exit criterion: an operator completes AT-01 with the six supplied files without modifying source code.

### Phase 6: Verification and handover

- Run unit/integration tests and a clean-environment installation and startup smoke test.
- Walk through every acceptance test using the running UI and save actual application screenshots showing upload/profile selection, results, and exceptions.
- Write a concise English README covering setup, launch, tests, demo workflow, profiles, output rules, assumptions, and limitations; include Russian operating guidance where useful.
- Review tracked files for credentials, genuine personal/customer data, temporary files, and dependencies on unrelated projects.

Exit criterion: every deliverable in requirements sections 7-8 is present, runnable, and supported by recorded verification.

## 5. Acceptance and test strategy

| Test | Planned evidence |
| --- | --- |
| AT-01 | Fresh local setup; load six sample inputs and reusable profiles; complete UI workflow and download both reports without editing code. Record steps and real screenshots. |
| AT-02 | Known within-file/cross-file duplicates, custom keys, conflicting duplicates, and an invalid earlier occurrence; assert only the retained valid rows contribute to totals and all excluded duplicates reference their retained rows. |
| AT-03 | Invalid/ambiguous dates, malformed amounts, blanks, non-finite values, and missing columns; assert affected rows/files are excluded and exceptions identify their source and reasons. |
| AT-04 | Compare exported detail and each branch/month/grand total with independently specified expected decimal totals, including rounding and negative-quantity cases. |

Additional targeted coverage: selected XLSX sheet, leading-zero IDs, Windows-1251 Cyrillic text, profile persistence, empty/all-invalid batches, duplicate ordering, repeated processing, formula-like text, and workbook export structure. Tests should verify observable behavior and financial correctness rather than mirror internal implementation.

## 6. Limits and open decisions

- Confirm duplicate identity for real source reports, especially whether a transaction contains several line items.
- Confirm signed-quantity return rules, currency, decimal precision, date locales, and whether zero amounts should be permitted.
- Confirm expected file sizes/row counts and representative hardware before setting supported upload limits or performance targets; measure the supplied demo rather than promise unspecified scale.
- The baseline has operator/admin workflows but no authentication, role enforcement, multi-user server, or production security-hardening commitment. Profiles are managed locally.
- Exclude legacy `.xls`, macro execution, ERP integration, inventory accounting, production hosting, maintenance, advanced integrations, and unrelated portfolio dependencies.
- Scheduling and production commitments need separate agreement. Implement phases in dependency order; estimate effort after validating the open business rules.

## 7. Immediate repository milestone

Publish the reviewed `requirements.md` and this `plan.md` under `nexoradev1221/excel_csv_automation`, then begin application work only when requested. Repository creation and publication require an authenticated account with permission to create repositories for that owner.
