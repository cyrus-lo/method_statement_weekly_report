# ECM Weekly Record Workflow — Step 1a and 1b

## Purpose

This workflow is independent of the Django/PostgreSQL reporting flow and uses only `ecm/.venv`.

1. **Step 1a:** read the ECM export and project readiness, create a dated backup, and create a period worksheet in the persistent ECM review workbook.
2. **Manual first-date review:** enter a verified first document date in column M, `Remark`, for every Step 1a row that requires first-date review.
3. **Step 1b:** validate the manual dates, create a dated backup, finalize `Count (Y/N)` in column I, and apply the final Count=Y filter through one hidden Excel session. That session starts Excel once, opens the workbook once, then closes the workbook and quits Excel once.
4. **Post-Step-1b manual handoff:** after Step 1b succeeds, open Excel separately to complete J `Use SMM` and K `Use KMS`; these decisions are not performed by Step 1b.
5. **Post-Step-1b backup:** after the hidden Excel session applies Count=`Y`, saves `output/ECM_Record.xlsx`, closes the workbook, and quits Excel, Step 1b creates a separate `_1b_after.xlsx` backup. It is created only after every Step 1b operation succeeds and before the separate J/K work and Step 2 edits.
6. **Step 2:** read the finalized ECM workbook and update the SMM working workbook. Its existing process remains unchanged.

Do not edit either immutable template in place. The persistent output workbooks in `output/` are changed only after their phase-specific backup is created.

## Environment

1. Use `ecm/.venv` only; do not use `django/.venv`.
2. Install dependencies from `ecm/requirements.txt`.
3. Do not use Django code, PostgreSQL credentials, or the Django report generator.
4. The existing `Step_1_Create_ECM_Record.py` remains unchanged as the legacy backup/reference script. Run the new phase scripts for this workflow.

## Common source rules

1. The ECM input is `csv_from_ecm/properties.csv` or `csv_from_ecm/properties (x).csv`. Step 1a selects the newest matching CSV by modified date unless `--csv` is provided.
2. Require inclusive `--start-date` and `--end-date` in `YYYY-MM-DD`; reject an end date earlier than the start date.
3. Include a CSV record only when its `Document date` falls within the inclusive period.
4. Read readiness from `Project Realiness Chris-In Use` in the current SMM workbook.
5. Preserve every selected CSV row. No row is deleted because it has a lower major version, duplicate SMM ID, missing SMM ID, unavailable project, or unresolved project code.
6. Sort rows by normalized Project Code ascending (blank last), case-insensitive Name, Major Version, then Document Date.

## Step 1a — Create ECM review worksheet

### Input and output files

| Phase | Input files | Output files | Purpose |
| --- | --- | --- | --- |
| **Step 1a — Create ECM review record** | ECM export CSV from `csv_from_ecm`; current SMM workbook for `Project Realiness Chris-In Use`; immutable `ECM_Record_template.xlsx` only when initializing the persistent workbook | Updated `output/ECM_Record.xlsx`; dated pre-update backup `output/ECM_Record_YYYY-MM-DD_to_YYYY-MM-DD_1a.xlsx` | Creates the reporting-period worksheet, writes permanent I=`N` exclusions, and leaves I blank for otherwise eligible rows. |
| **Manual handoff** | Updated `output/ECM_Record.xlsx` | The same `output/ECM_Record.xlsx` | Enter each verified first document date in M `Remark` only for a first-date review row. J, K, and L retain their existing manual purposes. |
| **Step 1b — Finalize count** | Manually updated `output/ECM_Record.xlsx`, specifically the selected reporting-period worksheet | Finalized `output/ECM_Record.xlsx`; dated pre-finalization backup `output/ECM_Record_YYYY-MM-DD_to_YYYY-MM-DD_1b.xlsx` | Finalizes I. Blank I and I=`Y` rows are handled identically; only I=`N` skips later rules. A first-date review date on or before `2026-04-16` produces I=`N`; after produces I=`Y`. A row with a blank Document revision becomes I=`Y` without an M date. No physical data row has blank I in the finalized output. |
| **Post-Step-1b handoff backup** | Successfully finalized and Excel-filtered `output/ECM_Record.xlsx` | `output/ECM_Record_YYYY-MM-DD_to_YYYY-MM-DD_1b_after.xlsx` | Create only after the hidden Excel session saves and closes the workbook and quits Excel, before the separate J/K handoff and Step 2. This preserves the exact finalized ECM workbook used as the Step 2 input. |

### Step 1a processing

1. Read the ECM CSV with `utf-8-sig` encoding and `csv.DictReader`.
2. Require these source headers: `Document title`, `Document date`, `Name`, `Document revision`, `Modified By`, `Major Version`, `SMM ID`, `Modified On`, `Outgoing ref. no.`, `Path`, and `Link`.
3. Trim headers and values. Parse `Document date` and `Modified On`; parse `Major Version` as an integer. Report malformed required values with the affected source row.
4. On the first run, initialize `output/ECM_Record.xlsx` from `ECM_Record_template.xlsx`. Before modification, create the `_1a.xlsx` backup. A backup for the same period may be replaced by a rerun.
5. Retain `ECM_1-12 Jul (WIP)` unchanged. Replace only the requested period worksheet, for example `ECM_1-13 Jul (WIP)`.
6. Write the template layout: A Project Code, B–H ECM fields, I Count, J Use SMM, K Use KMS, L Work Flow, M Remark, N Display, preserve column O, and write the derived outgoing reference in P, headed `Outgoing ref. no. (without revision)`.
7. Leave column N blank during Step 1a. Step 2 writes a static value after the post-Step-1b J/K decisions have been validated: I=`Y`, J=`Y`, K=`Y` writes `via KMS`; I=`Y`, J=`Y`, K=`N` writes `w/o KMS`; otherwise it writes blank.
8. Extract the numeric segment immediately after the first slash in Name as the initial project code. If it is absent or non-numeric, leave A blank.
9. Apply every non-date exclusion before manual review:
   - unresolved project code;
   - project absent from readiness or marked `No` in readiness column E;
   - lower major version for the same normalized Name;
   - noncanonical duplicate at the greatest major version. The canonical duplicate is latest `Modified On`, then lowest source-row number.
10. Write I=`N` for every excluded row. For a lower-major-version row, write `Duplicate with its version x` in M `Remark`, where `x` is the greatest Major Version for the same normalized Name. For an otherwise eligible canonical record, leave I blank. A blank I is not eligible for Step 2 until Step 1b finalizes it.
11. For P, remove one terminal `/{Document revision}` suffix from the source `Outgoing ref. no.` only when it matches exactly. Keep the value unchanged when the revision is blank or the suffix does not match. P is reference-only, does not affect Count decisions, and is preserved by Step 1b.
12. Clear the Count filter in 1a so users can review all rows, including blank-I and permanent N records. Do not hide rows. Select and activate only the new period worksheet.

### Run Step 1a

```powershell
.\ecm\.venv\Scripts\python.exe .\ecm\Step_1a_Create_ECM_Record.py `
  --start-date 2026-07-01 `
  --end-date 2026-07-13 `
  --smm-workbook ".\ecm\SMM_KMS_PM_Master_Document_template.xlsx"
```

Optional: specify `--csv`, `--ecm-record`, `--ecm-template`, or `--backup`.

## Manual first-document-date review

1. Open the new period worksheet in `output/ECM_Record.xlsx`.
2. A **first-date review row** is a physical data row where I `Count (Y/N)` is not `N` and E `Document revision` is non-empty after trimming spaces. Its M `Remark` requires first-date input and the revision rule.
3. For each first-date review row, find and verify the first document date using the approved ECM search procedure, then enter it only in M `Remark`.
4. M accepts a native Excel date or exact ISO text `YYYY-MM-DD`. After a successful Step 1b run, it is saved as `1st version is submitted on D Mon YYYY`; Step 1b accepts that saved form if it is run again.
5. Do not enter an M date solely because I is blank when E `Document revision` is blank. Step 1b finalizes that non-N row as I=`Y`.
6. Do not use Step 2 until Step 1b has successfully finalized the period worksheet and the separate post-Step-1b Excel session has completed J and K. J, K, L, other manual fields, and manually completed blank project codes in A are not replaced by Step 1b.

## Step 1b — Finalize Count (Y/N)

1. Open the requested period worksheet in the manually reviewed persistent `output/ECM_Record.xlsx`. Column P contains the outgoing reference with a matching terminal revision suffix removed and is preserved during finalization.
2. Validate headers A:K, require M to be named `Remark`, and require P to be named `Outgoing ref. no. (without revision)`.
3. Keep I=`N` records as `N`; they are terminal exclusions and all later rules skip them. Blank I and I=`Y` rows are handled identically. Every non-N physical data row with a non-empty trimmed E `Document revision` is a first-date review row and requires a valid M first date before any changes are saved.
4. Reject the entire run if a first-date review row has a blank or invalid M value. The workbook is not saved or modified in that case.
5. For each first-date review row:
  - M on or before `2026-04-16` → I=`N`.
  - M after `2026-04-16` → I=`Y`.
  - Replace M with `1st version is submitted on D Mon YYYY` after the date has been validated.
6. For a non-N row whose E `Document revision` is blank after trimming spaces, M is not required and Step 1b sets I=`Y`. Preserve M on this type of row and on existing I=`N` rows. Before the output is saved, convert any remaining blank I on a physical data row to I=`Y`.
7. After validating every first-date review row, create `output/ECM_Record_YYYY-MM-DD_to_YYYY-MM-DD_1b.xlsx`. A backup for the same period may be replaced by a rerun. Then update I and M, preserving J–L, other worksheets, and unrelated workbook settings.
8. Step 1b leaves N and P unchanged. Step 2 writes the static N `Display` value for every physical data row before generating the SMM summary. N remains presentation-only; P remains a reference field; Step 2 continues to use I/J/K as its source fields.
9. After I is finalized, highlight the full row orange when H `SMM ID` is blank. This final formatting action does not exclude or delete the record.
10. Apply the Count filter with I=`Y` selected, select and activate only the finalized period worksheet, then save it. The filter does not delete N rows; clear it in Excel to review them.
11. The Excel filter operation is the only Desktop Excel operation performed by Step 1b. It uses one hidden Excel session and does not wait for or enter J/K values.

### Run Step 1b

```powershell
.\ecm\.venv\Scripts\python.exe .\ecm\Step_1b_Finalize_ECM_Record.py `
  --start-date 2026-07-01 `
  --end-date 2026-07-13 `
  --ecm-record ".\ecm\output\ECM_Record.xlsx"
```

Optional: specify `--ecm-sheet` or `--backup`.

## Post-Step-1b backup before Step 2

Step 1b automatically copies the finalized and Excel-filtered `output/ECM_Record.xlsx` only after the hidden Excel session has saved and closed the workbook and quit Excel. Step 2 leaves this handoff snapshot unchanged while writing Display values to the live ECM workbook. The default snapshot is:

`output/ECM_Record_YYYY-MM-DD_to_YYYY-MM-DD_1b_after.xlsx`

This snapshot is distinct from the Step 1b pre-finalization backup and represents the finalized workbook before the separate manual Use SMM/Use KMS decisions and Step 2 updates.

## Step 2 — Update SMM_KMS_PM_Master_Document.xlsx

Step 2 reads every physical row from the manually completed live ECM worksheet, regardless of the saved Excel filter, but includes only finalized I=`Y` records in Summary totals. It updates the live ECM Display values and SMM workbook without changing the Step 1b snapshot. Refer to `steps.md` for the full Step 2 rules, parameters, output workbook handling, Summary formulas, and format requirements.

## Verification

1. Confirm both new scripts run with `ecm/.venv` and do not alter `Step_1_Create_ECM_Record.py`.
2. Confirm 1a retains all selected records, writes permanent N exclusions, and leaves only candidates requiring later finalization blank in I.
3. Confirm M on a first-date review row accepts a native Excel date, `YYYY-MM-DD` text, or the saved `1st version is submitted on D Mon YYYY` form, and is saved in that standard form.
4. Confirm M=`2026-04-15` and M=`2026-04-16` produce I=`N`; M=`2026-04-17` produces I=`Y` when all prior rules passed.
5. Confirm an invalid or missing M value on a first-date review row fails before the persistent workbook is saved, while a blank Document revision does not require M.
6. Confirm J–L values, and M values other than validated first-date review values, are retained by 1b; confirm no physical data row has blank I after a successful run.
7. Confirm Step 2 writes literal N values: `via KMS`, `w/o KMS`, or blank according to the Display mapping; confirm no N value begins with `=`.
8. Confirm the Step 1a backup, Step 1b pre-finalization backup, and post-Step-1b/pre-Step-2 backup exist, and immutable templates are unchanged.
9. Clear the final filter and confirm lower major versions and duplicate SMM IDs remain as physical N rows.
