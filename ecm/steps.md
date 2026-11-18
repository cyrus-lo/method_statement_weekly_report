# ECM Weekly Record Workflow

## Purpose

This workflow has two independent steps and is independent of the Django/PostgreSQL reporting flow. It must use its own environment.

1. **Step 1:** read ECM document records and project readiness, update the master `ECM_Record.xlsx`, and create a dated backup for the reporting period.
2. **Step 1b Excel handoff:** after `openpyxl` saves and closes the workbook, Step 1b uses one hidden Excel session to reopen it and apply the Count=Y filter. Excel starts once, the workbook opens once, and both the workbook and Excel close once. The `_1b_after.xlsx` backup is created only after this Excel save/close/quit sequence succeeds. It does not perform J/K.
3. **Post-Step-1b manual handoff:** after Step 1b succeeds, complete J `Use SMM` and K `Use KMS` in a separate Excel session.
4. **Post-Step-1b handoff:** after Step 1b finalizes `ECM_Record.xlsx`, create a separate dated backup of that finalized workbook before the J/K handoff and Step 2.
5. **Step 2:** read the finalized `ECM_Record.xlsx`, update the persistent SMM working workbook, and create a dated backup before adding a new date-period Summary sheet.

## Source files

- ECM exports: `csv_from_ecm/properties.csv` or `csv_from_ecm/properties (x).csv`
- ECM review template: `ECM_Record_template.xlsx`, sheet `ECM_1-12 Jul (WIP)`
- Current SMM workbook: `SMM_xxx.xlsx`
  - Readiness sheet: `Project Realiness Chris-In Use`

The Step 1 processor selects the newest modified matching CSV by default. It must allow an explicit CSV path override and log the selected file.

## Environment

1. Create and use `ecm/.venv` only.
2. Do not use `django/.venv`.
3. Do not use Django code, PostgreSQL credentials, or the Django report generator.
4. Install ECM dependencies from `ecm/requirements.txt`.
5. Do not edit `ECM_Record_template.xlsx` or `SMM_KMS_PM_Master_Document_template.xlsx` in place. The persistent working workbooks in `output/` are updated only after a dated backup is created.

## Inputs and reporting period

1. Require inclusive `--start-date` and `--end-date` parameters in `YYYY-MM-DD` format.
2. Include a CSV row when its `Document date` is on or after the start date and on or before the end date.
3. Reject an end date earlier than the start date.
4. Use the reporting period in generated worksheet names. For example, the period 2026-07-01 to 2026-07-12 uses `(OK) 1 Jul - 12 Jul`.

## Step 1 — Update ECM_Record.xlsx

**Input:** `properties.csv` or `properties (x).csv`, plus the current `SMM_xxx.xlsx` for project-readiness checking.

**Input/output:** `output/ECM_Record.xlsx` is the persistent ECM workbook. It is updated in place.

**Backup:** Before updating the master workbook, create `output/ECM_Record_YYYY-MM-DD_to_YYYY-MM-DD.xlsx`. The backup uses one `.xlsx` extension and is replaced when the same reporting period is run again.

After the Step 1b hidden Excel filter session succeeds and closes, Step 1b automatically creates a second backup before the separate J/K handoff:

`output/ECM_Record_YYYY-MM-DD_to_YYYY-MM-DD_1b_after.xlsx`

This is a post-finalization handoff snapshot copied from the finalized `output/ECM_Record.xlsx`. Step 2 leaves this snapshot unchanged; it reads the reviewed data and writes Display values to the live `output/ECM_Record.xlsx` workbook.

### CSV processing

1. Read the CSV with `utf-8-sig` encoding and `csv.DictReader`.
2. Require the ECM headers:
   - `Document title`
   - `Document date`
   - `Name`
   - `Document revision`
   - `Modified By`
   - `Major Version`
   - `SMM ID`
   - `Modified On`
   - `Outgoing ref. no.`
   - `Path`
   - `Link`
3. Trim whitespace from headers and values.
4. Parse `Document date` and `Modified On` as dates/times.
5. Parse `Major Version` as an integer.
6. Normalize blank `SMM ID` values.
7. Report malformed required dates or versions with the affected source row; do not guess a value.

### ECM review workbook

1. On the first run, initialize `output/ECM_Record.xlsx` by copying `ECM_Record_template.xlsx`.
2. Before every update, copy the existing `output/ECM_Record.xlsx` to the dated backup path.
3. Validate that `ECM_1-12 Jul (WIP)` exists and that the target columns match the required layout.
4. Keep `ECM_1-12 Jul (WIP)` unchanged. Replace any existing worksheet for the requested reporting period with a fresh copy; for July 1–13, create `ECM_1-13 Jul (WIP)`.
5. Extract the initial project code from `Name`:
   - Split the normalized Name on `/`.
   - Use the numeric segment immediately after the first slash.
   - Example: `CWFWSJV/1046/M/1773-2026.pdf` becomes project code `1046`.
   - If that segment is missing or non-numeric, leave column A blank for manual completion.
6. Write the selected period's rows in the new period worksheet using the template layout:
   - A: assigned project code
   - B–H: required ECM source fields, including SMM ID in H
   - I: count/check value
   - J: Use SMM (Y/N), for manual update
   - K: Use KMS (Y/N), for manual update
   - L: preserve the template layout for manual use
   - N: Display, written as a static value by Step 2 from I/J/K as `via KMS`, `w/o KMS`, or blank
   - P: `Outgoing ref. no. (without revision)`, derived from the Step 0 CSV by removing a matching terminal `/{Document revision}` suffix; reference-only and does not affect Count decisions
7. Sort rows by:
   1. normalized Project Code in ascending order, with blank project codes last
   2. case-insensitive Name
   3. Major Version
   4. Document Date
8. Highlight the entire row in orange when SMM ID is missing. A missing SMM ID does not itself exclude a row.
9. Do not add an Audit worksheet.
10. Do not hide any rows. Apply an Excel filter to the `Count (Y/N)` column with Y selected. Step 1b performs this final filter in one hidden Excel session; J/K remain a later manual operation.

- The filter hides `N` rows in Excel but does not remove them from the worksheet. Clear the filter to review lower major versions and other `N` rows.

11. Make only the new period worksheet selected and active when the updated workbook opens.

### Project readiness and unresolved projects

If the project code cannot be extracted from `Name`:

1. Leave the project-code cell in column A blank for manual handling.
2. Write N in the count/check column I.
3. Do not treat the row as an automated count.

For resolved project codes:

1. Read project code and readiness status from `Project Realiness Chris-In Use` in the current `SMM_xxx.xlsx`.
2. Normalize whitespace, letter case, and numeric/text project-code representations before matching.
3. If a project's column E value is `No`:
   - write N in column I for every matching ECM row;
   - exclude the row from totals.
4. A project absent from the readiness sheet is also not ready:
   - write N in I;
   - exclude the row from totals.
5. Listed projects not marked `No` continue through the remaining eligibility rules.

### Counting and first-document-date rules

### Major-version rule

1. Normalize Name for grouping by trimming and case-folding, while retaining its original display value.
2. For each normalized Name, retain only the greatest numeric Major Version.
3. Exclude lower major versions from the automated count.
4. If multiple rows share the greatest Major Version, select one canonical row using:
   1. latest Modified On;
   2. source-row order as the final tie-breaker.
5. Exclude all other maximum-version duplicates from the automated count.
6. At most one row per normalized Name can receive automated Y.

### First-document-date rule

1. Use the ECM search procedure to be supplied later.
2. The lookup takes the approved substring derived from `Outgoing ref. no.` and returns matching ECM documents and their Document Dates.
3. Until that procedure is implemented, provide a temporary first-date CSV with `Name` and `First Document Date` (`YYYY-MM-DD`) columns when running Step 1.
4. Complete this lookup before passing `ECM_Record.xlsx` to Chris and Jo.
5. For each normalized Name, the earliest ECM Document Date must be on or after `2026-04-16`.
6. `2026-04-16` itself is eligible.
7. A row without a verified eligible first date must not receive automated Y.

### Column I decision

1. Write Y only if the row:
   - is inside the reporting period;
   - is the canonical maximum-major-version row;
   - has a resolved project code;
   - belongs to a project listed in readiness and not marked `No` in column E;
   - passes the verified first-document-date rule.
2. Write N when a completed rule excludes the row.
3. Before the ECM search procedure is available, do not run final Step 1 processing or issue an ECM record workbook with automated final decisions.
4. Chris and Jo may manually update J, K, L, and a blank project-code cell in A after receiving `ECM_Record.xlsx`.

## Step 2 — Update SMM_KMS_PM_Master_Document.xlsx

**Input:** manually updated `output/ECM_Record.xlsx`.

**Working output:** `output/SMM_KMS_PM_Master_Document.xlsx`, updated in place. On its first run, initialize it from `SMM_KMS_PM_Master_Document_template.xlsx`.

**Backup:** after successfully updating the working output, create `output/SMM_KMS_PM_Master_Document_YYYY-MM-DD_to_YYYY-MM-DD.xlsx`. This dated snapshot contains the updated period Summary and is replaced when the same reporting period is run again.

1. Read the manually updated ECM record as the source for Step 2.
2. Use `SMM_KMS_PM_Master_Document_template.xlsx` only to initialize the persistent SMM working workbook. Do not modify the template.
3. Copy an existing period Summary sheet as the layout for a new worksheet. Keep the source Summary sheet unchanged.
4. Name the new worksheet for the reporting period. For 2026-07-01 to 2026-07-13, name it `(WIP) 1 Jul - 13 Jul`.
5. Make only the new Summary sheet selected and active when the output workbook opens.
6. Read every physical ECM row, regardless of the ECM worksheet's saved filter or hidden-row state.
7. Match rows by normalized numeric Project Code from ECM column A and the readiness sheet.
8. Only ECM rows with `Count (Y/N)` in column I equal to Y are eligible for Summary values.
   - For rows where I is N, ignore columns J and K completely; they may be blank or contain any value.
   - For rows where I is Y and J is N, K may be N or blank. K must not be Y.
9. Generate Summary columns E–H as follows:
   - E, `A. Draft via KMS`: count eligible rows where J, `Use SMM (Y/N)`, is Y and K, `Use KMS (Y/N)`, is Y.
   - F, `B. Draft NOT via KMS`: count eligible rows where J is Y and K is N.
   - G, `Total No. (A + B) Approved`: write the Excel formula `=E[row]+F[row]`.
   - H, `Number of Method Statement Record (Outbound) in ECM`: write G plus the count of eligible rows where I is Y and J is N, then automatically deduct one for each Mattex-added row.
   - A Mattex-added row is identified from Remark (column M), case-insensitively, by `Added by Mattex` or `ADD By Chris's Team` (Chris's Team is Mattex). Such `Y/Y/Y` and `Y/Y/N` rows remain in E or F respectively, but are deducted from H only. A Mattex-added Remark on a `Y/N` row is rejected because it must remain an ECM-not-prepared-via-SMM exception.
10. An eligible row with J=N cannot have K=Y. Treat that combination as a validation error.
11. Fail when an I=Y row has a blank or invalid project code, invalid Y/N values, or a project absent from `Project Realiness Chris-In Use`.
12. Add only projects having at least one I=Y ECM row, sort them by project code, and take their Project and BU values from the readiness sheet.
13. Update the date-period text in the merged E–G heading and the H heading.
14. Preserve all unrelated formulas, values, formatting, worksheets, merged cells, and workbook settings. Set the generated Summary sheet's columns as follows:

- B: final Excel width `92.55` (internal `openpyxl` width `93.1875`)
- C: final Excel width `15.82` (internal `openpyxl` width `16.46375`)
- D: hidden
- E, F, and G: final Excel width `19.82` (internal `openpyxl` width `20.46375`)
- H: final Excel width `26.82` (internal `openpyxl` width `27.46375`)
- Set the generated Summary worksheet zoom to `85%`.
- Do not change column A or unspecified columns. The internal values are calibrated because `openpyxl` width values differ from Excel's displayed widths for this workbook/font combination.

15. Later Step 2 runs retain prior period Summary sheets in `output/SMM_KMS_PM_Master_Document.xlsx`; each successful run creates one dated snapshot of the updated working workbook.
16. Project 997 has no special automated processing.

### Step 2 parameters

| Parameter                         | Required | Purpose                                                                                            |
| --------------------------------- | -------- | -------------------------------------------------------------------------------------------------- |
| `--start-date YYYY-MM-DD`       | Yes      | Start of the Summary reporting period.                                                             |
| `--end-date YYYY-MM-DD`         | Yes      | End of the Summary reporting period.                                                               |
| `--ecm-record PATH`             | No       | Manually reviewed ECM workbook. Defaults to`output/ECM_Record.xlsx`.                             |
| `--smm-template PATH`           | No       | Immutable SMM template used only to initialize the working workbook.                               |
| `--smm-record PATH`             | No       | Persistent SMM working workbook to update. Defaults to`output/SMM_KMS_PM_Master_Document.xlsx`.  |
| `--ecm-sheet NAME`              | No       | ECM review sheet. Defaults to the sheet name derived from the reporting period.                    |
| `--summary-template-sheet NAME` | No       | Existing SMM period sheet to copy for the new Summary layout. Defaults to`(OK) 1 Jul - 12 Jul`. |
| `--backup PATH`                 | No       | Dated SMM backup path. Defaults to the reporting-period backup in`output`.                       |

Example:

```powershell
.\ecm\.venv\Scripts\python.exe .\ecm\Step_2_Create_SMM_Summary.py `
   --start-date 2026-07-01 `
   --end-date 2026-07-13 `
   --ecm-record ".\ecm\output\ECM_Record.xlsx" `
   --smm-template ".\ecm\SMM_KMS_PM_Master_Document_template.xlsx" `
   --smm-record ".\ecm\output\SMM_KMS_PM_Master_Document.xlsx" `
   --summary-template-sheet "(OK) 1 Jul - 12 Jul"
```

The user manually changes `(WIP)` to `(OK)` after reviewing the generated Summary sheet.

## Verification

1. Confirm the workflow runs using only `ecm/.venv`.
2. Confirm the newest matching CSV is selected and logged.
3. Confirm Chinese text is preserved.
4. Confirm both templates remain unchanged after a run.
5. Confirm the major-version duplicates retain only the canonical maximum-version row.
6. Confirm missing SMM IDs are highlighted but not automatically excluded.
7. Confirm a readiness column-E value of `No` produces N in ECM column I.
8. Confirm unresolved project codes leave A blank and produce N in I.
9. Confirm final Step 1 processing does not run until the ECM search logic is available.
10. Confirm Step 1 updates `output/ECM_Record.xlsx` in place and creates the dated backup without changing the CSV, readiness workbook, or template.
11. Confirm Step 2 reads all ECM rows even when the review sheet is filtered to I=Y.
12. Confirm E counts I=Y/J=Y/K=Y, F counts I=Y/J=Y/K=N, every G cell is an E+F formula, and the automated H value equals G plus the count of I=Y/J=N rows before any manual Chris-record deduction.
13. Confirm Step 2 updates `output/SMM_KMS_PM_Master_Document.xlsx`, creates its dated backup, and adds a new `(WIP)` period worksheet while retaining prior Summary sheets unchanged.
14. Confirm only the new Summary sheet is selected and active.
15. Confirm both Step 2 templates remain byte-for-byte unchanged.
16. Open outputs in Excel to verify formulas, styles, filters, freeze panes, the Summary widths B=`92.55`, C=`15.82`, E:G=`19.82`, H=`26.82`, hidden D, dynamic names, and manual columns H–L.
