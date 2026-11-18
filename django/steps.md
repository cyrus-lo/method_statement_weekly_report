# Method Statement Interaction Weekly Report

## Purpose

Generate a weekly Excel report from PostgreSQL data, replacing Django-admin navigation, webpage filtering, web scraping, and manual copy/paste.

## Inputs

- Excel template: `Method_Statement_Interaction_Weekly_Report_template.xlsx`
- Inclusive reporting period: start date and end date
- PostgreSQL credentials supplied in a local `.env` file
- Session query: `sessions.sql`
- Submission-count query: `submission_counts.sql`
- Report program: `Method_Statement_Interaction_Weekly_Report.py`

## Database connection details

Store the following values in a local `.env` file. Do not commit this file to source control.

- `PGHOST`
- `PGPORT` — normally `5432`
- `PGDATABASE`
- `PGUSER`
- `PGPASSWORD`
- `PGSSLMODE` — if required by the database, for example `require`
- `REPORTING_TIMEZONE` — named IANA timezone used for the report's calendar-day
	boundaries; defaults to `Asia/Hong_Kong`

Use a read-only PostgreSQL account for this report.

## Run the report

1. Use the project-local `.venv` in the `django` folder.
2. Copy `.env.example` to `.env` and provide the PostgreSQL read-only account details.
3. Run `Method_Statement_Interaction_Weekly_Report.py` with `--start-date` and `--end-date` in `YYYY-MM-DD` format.
4. The generated workbook is saved in the `output` folder. Its filename includes the reporting period.

Example PowerShell command:

```powershell
.\django\.venv\Scripts\python.exe .\django\Method_Statement_Interaction_Weekly_Report.py --start-date 2026-07-01 --end-date 2026-07-12
```

## Reporting period

1. Provide start and end dates in ISO format: `YYYY-MM-DD` to the report program.
2. The program passes identical date values to both SQL files using bound PostgreSQL parameters.
3. Before either query runs, the program sets its PostgreSQL session timezone to
	`REPORTING_TIMEZONE` (default `Asia/Hong_Kong`). Both queries therefore use
	that timezone when filtering `documents_session.created_at`.
4. The end date is inclusive because the queries select records before the following calendar day.

## Data extraction

### 1. Sessions

1. Run `sessions.sql` through the report program for the selected period.
2. Export the complete result to a worksheet named `Sessions`.
3. Sort the worksheet by `username`, then `project_code`.
4. Retain all columns returned by the query, including the session-level `count_submission_times` and `pending_issue` display value.

### 2. Submission Counts

1. Run `submission_counts.sql` through the report program using the same reporting period.
2. Export the complete result to a worksheet named `Submission Counts`.
3. Keep the rows ordered by `project_code`.

## Summary worksheet

1. Copy the Excel template and save the result as a new timestamped workbook. Do not alter the source template.
2. Preserve the existing consolidated worksheet (the first worksheet in the template) and its rows 1–3, then rename it dynamically from the reporting period. For example, `2026-07-01` to `2026-07-12` becomes `1 - 12 Jul - Consolidated`.
3. Clear existing Summary data below row 3.
4. Build one Summary row for every project code present in either extracted dataset.
5. Write the project code to column A, starting at row 4.
6. Sort Summary project codes in ascending numeric order.
7. Exclude project code `0` from the Summary / Consolidation sheet.
8. Use zero where a project has no value for a metric.
9. Set fixed Summary column widths so Excel displays approximately `A = 40.09`, `B = 16`, `C:H = 19.09`, `I = 19`. The script uses the workbook's internal width values rather than the UI display numbers. Do not auto-size the Summary sheet from content.

### Summary field mapping

| Summary heading | Calculation by `project_code` |
| --- | --- |
| Number of KMS User | Distinct non-null count of `username` from `sessions.sql` |
| Number of Question Asked by KMS User to Retrieve Historical Method Statement Record | Sum of `count_retrieve_mode` |
| Number of PDF Viewed by KMS User | Sum of `count_pdf_viewed` |
| Number of Method Statement Content Selected by User to Further Edit in WORD Plugin | Sum of `count_content_copied` |
| Number of Question Asked by KMS User to Gather Work Trade Information | Sum of `count_ask_mode` |
| Number of Method Statement Drafted & Submitted to SMM | `count_submission_times` from `submission_counts.sql`; this is the distinct DOCX submission-ID count |
| Number of Reported Issue | Sum of `total_issue_count` from `sessions.sql` |

The session-level `count_submission_times` field is retained in `Sessions` only. The Summary submission value is taken from `submission_counts.sql` to avoid counting multiple events belonging to the same submission ID.

## Validation

1. Confirm the raw `Sessions` and `Submission Counts` worksheets match their respective SQL results.
2. Confirm all records dated on the selected end date are included, and records from the following date are excluded.
3. Check representative projects manually for each Summary calculation.
4. Confirm a project returned by only one query still appears in Summary with zeroes for unavailable metrics.
5. Confirm the output includes the dynamically named consolidated worksheet, `Sessions`, and `Submission Counts`, and that the original template remains unchanged.
