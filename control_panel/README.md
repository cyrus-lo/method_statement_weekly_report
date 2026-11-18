# Local Workflow Control Panel

A localhost-only browser UI for the existing weekly report and ECM Python workflows. It runs only on the current Windows workstation and does not replace the report/ECM business logic.

## What it provides

- Four date-driven workflow actions: weekly Method Statement report, ECM Step 1a, ECM Step 1b, and ECM Step 2.
- One active job at a time, preventing overlapping edits to persistent Excel workbooks.
- Live process logs, a SQLite run history, and immutable copies of successful output workbooks/backups.
- Preflight checks for invalid dates, missing input CSVs, and Excel temporary lock files.
- Clear ECM handoff: **Step 1a → Excel manual Remark-date review → Step 1b → complete Use SMM/Use KMS values → Step 2**.

## Prerequisites

- Windows with the `py` launcher and Python 3.11 or later available.
- Existing working environments at `django/.venv` and `ecm/.venv`.
- Excel closed for any workbook that a workflow is about to modify.
- Connectivity and permissions required by the existing weekly-report PostgreSQL configuration in `django/.env`.

## Configure

1. Copy `.env.example` to `.env` in this folder.
2. Keep the defaults when this repository folder structure is unchanged. Otherwise set absolute paths for the workflow Python interpreters, templates, CSV directory, and persistent output workbooks.
3. Do not add PostgreSQL credentials to this file. The weekly report continues to read them from `django/.env`.
4. Keep `.env` private. It is ignored by source control.

`CONTROL_PANEL_DATA_DIR` stores the local SQLite history, logs, and immutable copied artifacts. Do not point it at `django/output` or `ecm/output`.

## Start

Run `start_control_panel.ps1` from PowerShell. The first run creates `control_panel/.venv`, installs only UI dependencies, validates the two existing workflow environments, starts Uvicorn on `127.0.0.1:8765`, and opens the local dashboard.

To restart the app, run `restart_control_panel.ps1`. It stops only the V1 Uvicorn process, waits for port `8765` to be released, and starts the app again. If workflow child processes are active, the restart is refused; use `restart_control_panel.ps1 -Force` only when those workers should also be stopped. Add `-NoBrowser` to avoid opening a browser window.

The service is deliberately bound to `127.0.0.1`; it is not available over the network. Close the terminal/process that hosts Uvicorn to stop it.

## Operating sequence

1. For any card, enter an inclusive start and end date, then choose **Run workflow**.
2. Open the run-detail page to monitor the live log and download copied artifacts after success.
3. For ECM Step 1a, open `ecm/output/ECM_Record.xlsx` after the run completes. Enter verified first-document dates in `Remark` (column M) for the required rows, save, and **close Excel**.
4. Run ECM Step 1b. Step 1b finalizes Count/Remark, creates the pre-1b backup, closes the `openpyxl` workbook, then uses one hidden Excel session to reopen the workbook and apply the Count=Y filter. After Excel saves and closes successfully, Step 1b creates the post-1b backup; it does not perform the later J/K manual decisions.
5. After Step 1b succeeds, open `ecm/output/ECM_Record.xlsx` in Excel, complete manual `Use SMM`/`Use KMS` values on Count=Y rows, save, and close Excel.
6. Run ECM Step 2 to update the SMM working workbook.

A non-zero workflow exit, missing expected artifact, process timeout, or copy failure marks the run as failed and preserves its log. Existing script validation messages remain authoritative for workbook-content errors.

## Recovery and history

- Each successful artifact is copied to `control_panel/data/runs/<run-id>/artifacts`; later reruns cannot change an earlier download.
- The application database is `control_panel/data/history.sqlite3`.
- Back up the `control_panel/data` folder if run records and downloads must be retained.
- Existing operational workbooks and their workflow-created backups remain in `django/output` and `ecm/output`; they are not replaced by the application data folder.
- Restarting the control panel marks any job that was in progress at the time of shutdown as failed. Do not rerun until the underlying Excel files have been checked.

## Verification

From the dedicated control-panel environment, run `pytest control_panel/tests` from the repository root. Existing ECM regression tests remain separate under `ecm/tests` and should continue to run with the ECM environment.
