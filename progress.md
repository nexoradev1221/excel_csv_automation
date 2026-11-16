# Implementation progress

## November 15, 2026 — Day 1 (Asia/Vladivostok)

- Reviewed requirements and the full scheduled plan; accepted its documented demo defaults.
- Confirmed documentation PR #1 was already merged; synchronized local `main` before work.
- Started `chore/project-foundation` for package setup, dependency locking, local Russian UI
  shell, setup instructions, ignores, and initial checks. No transaction processing is claimed.
- The scheduled sandbox is read-only. Approved escalated commands allow the authorized project
  work and GitHub access for this session; future runs may require their own approvals.
- Added a standalone Python 3.12 package, localhost launcher, Russian shell, locked dependencies,
  setup/limitations README, and Windows/Linux CI (including Streamlit rendering).
- Local checks passed: frozen installation, Ruff lint/format, dependency import check, wheel build,
  and Streamlit AppTest render without exceptions. Fixed a launcher line-length lint failure
  before the initial commit. Actual processing behavior tests start with processing features.
- Commits: [f7d9c15](https://github.com/nexoradev1221/excel_csv_automation/commit/f7d9c15)
  (package/shell) and [c488535](https://github.com/nexoradev1221/excel_csv_automation/commit/c488535)
  (CI); both pushed. Review/CI: [PR #2](https://github.com/nexoradev1221/excel_csv_automation/pull/2).
- Windows and Linux CI both passed for commit `9a57000` in
  [run 43689248097](https://github.com/nexoradev1221/excel_csv_automation/actions/runs/43689248097).
- A real subprocess launch through `python -m excel_csv_automation` returned HTTP 200 from
  the localhost health endpoint; the test-owned server was stopped afterward.
- Day 1 foundation deliverables are complete. PR #2 is the canonical merge/check record;
  this final progress-only update will also pass CI before merge. No processing features
  or acceptance-test completion are claimed yet.

Day 1 PR #2 was merged as `3cb0b5a`; main was synchronized and verified before Day 2.

## November 16, 2026 — Day 2 (Asia/Vladivostok)

- Created `feat/demo-dataset` from clean, up-to-date main; no unrelated user changes found.
- Added framework-independent typed profiles, source references, rows, exceptions, summaries,
  counts, and run results. Profile keys, delimiter settings, and source coordinates have
  boundary checks; processing validation/calculation services are still future work.
- Added six synthetic branch files (three XLSX, three CSV), three header layouts, five sample
  profiles, native XLSX dates, Windows-1251, UTF-8/BOM, two months, returns, invalid values,
  duplicates/conflicts, leading-zero IDs, and a literal formula-like SKU for future export checks.
- The Center workbook requires selecting Sales after Notes. Invalid Center ID 0403 occurs
  before a valid occurrence; the Port file contains a copied North row for cross-file deduplication.
- Added a hand-specified expected-results oracle, independent of generator/production services:
  30 parsed = 18 accepted + 7 rejected + 5 duplicates; expected accepted total 159.53.
- Added deterministic fixture generation without regenerating expected results. Tests verify
  every source-row reference, all expected line/branch/month/grand totals, each rejection,
  retained duplicate references/conflicts, encodings/sheets, and repeatable fixture bytes.
- Local checks: Ruff lint/format passed; all 21 tests passed. Corrected Python 3.12 type-alias
  lint findings before the first commit. No application acceptance-test completion is claimed.
- CI now runs pytest on Windows and Linux; `.gitattributes` preserves fixture LF line endings
  and XLSX binary bytes across platforms. Existing runtime/build/render checks remain.
- Commits pushed: [dea9843](https://github.com/nexoradev1221/excel_csv_automation/commit/dea9843)
  (contracts), [18c8a3d](https://github.com/nexoradev1221/excel_csv_automation/commit/18c8a3d)
  (fixtures/oracle/tests), and `b3fe18d` (CI/progress). Review and final CI evidence:
  [PR #3](https://github.com/nexoradev1221/excel_csv_automation/pull/3). Merge only after
  both Windows and Linux checks pass; the PR provides the canonical merge/check record.

Next: Day 3 CSV import with explicit delimiter/encoding, physical source-row provenance,
header preview, and actionable file-level errors. No current blocker.
