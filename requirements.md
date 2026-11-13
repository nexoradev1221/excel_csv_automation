# Client Software Requirements — Excel and CSV Processing Automation

**Brief ID:** DEMO-SRS-01  
**Target Git repository:** `excel_csv_automation`
**Version:** 1.0 — 2026-11-14  

## 1. Client Background and Business Problem

A regional wholesale distributor receives sales, returns and inventory reports from six branches in differing Excel and CSV layouts. Its office manager spends several hours each week aligning columns, removing duplicates and assembling a monthly workbook. Manual errors have caused inconsistent totals.

## 2. Requested Business Outcome

An office operator uploads weekly branch files and produces a standardized, auditable consolidated sales workbook without editing scripts.

## 3. Expected Users

Office personnel and an administrator.

## 4. Functional Requirements

- **FR-01:** Import XLSX and CSV files with configurable delimiter, encoding and sheet selection.
- **FR-02:** Map differently named source columns to a standard template (date, branch, SKU, quantity, unit price, transaction ID).
- **FR-03:** Validate required columns, dates, numeric types and missing values; report rejected rows with reasons.
- **FR-04:** Identify duplicate transactions using configurable business keys; retain a reviewable exception list.
- **FR-05:** Combine accepted rows, calculate line totals and generate branch/month summaries.
- **FR-06:** Export a formatted XLSX report with raw-cleaned data, summary and exceptions sheets.
- **FR-07:** Save reusable mapping profiles and show processing progress and results.

## 5. Input Data and Expected Outputs

- Input: six synthetic branch files; at least two column layouts, one CSV encoding variation, duplicate IDs and invalid rows.
- Output: a consolidated workbook and an exception report, both downloadable or saved locally.

## 6. Acceptance Tests / Definition of Done

- **AT-01:** Given the supplied six test files, the operator produces a consolidated report without modifying source code.
- **AT-02:** Duplicate transactions appear only once in accepted totals and are listed in exceptions.
- **AT-03:** Invalid amounts or dates do not silently enter totals; reports identify affected source rows.
- **AT-04:** Summary totals reconcile exactly with the accepted detail rows.

## 7. Delivery Scope and Constraints

- **Suggested implementation:** Python with pandas/openpyxl; desktop GUI or local browser UI. No live ERP integration.
- Deliver a runnable application, complete source code, dependency files, sample configuration, automated tests, sample dataset, and real screenshots of the running application.
- Include a concise English README with setup, test, usage, and limitations; provide Russian end-user wording where the application is customer-facing.
- Operate as its own product with its own repository, data and configuration. It must never require any of the other nine portfolio projects.
- Avoid paid third-party services for the core demo; where a platform account or external credentials are required, provide a local simulator or clearly documented sample mode.
- Do not place secrets, confidential documents, genuine personal records, or copyrighted customer assets in the public repository.
- Deliver only the listed baseline features; production hosting, ongoing maintenance, advanced integrations, and additional scope require separate agreement.

## 8. Final Deliverables

1. Public-ready source repository named `excel_csv_automation`.
2. Runnable demonstration with synthetic or properly licensed sample data.
3. Installation and operating instructions, screenshots and automated tests.
4. Demonstration of all acceptance tests listed in Section 6.

## 9. Assumptions and Open Points

- The client scenarios and example volumes are illustrative and should be validated before real commercial development.
- UI layout, detailed database schema, security hardening, deployment targets and deadlines will be specified during technical design.
- Any expansion beyond the basic scope is subject to a new estimate and a revised requirements version.
