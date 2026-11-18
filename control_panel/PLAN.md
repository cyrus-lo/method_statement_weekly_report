# Local Workflow Control Panel — Plan

Build a localhost-only browser application for one Windows workstation. It will keep the four existing scripts as the source of business logic, launch each with its existing virtual environment and command-line contract, and provide forms, one-at-a-time execution, streamed logs, immutable output copies, and SQLite-backed run history. This avoids rewriting the tested Excel/PostgreSQL logic while providing a maintainable UI that could later be hosted centrally.

## Confirmed decisions

- Deployment is one Windows workstation only.
- The UI is a local browser app, not a packaged desktop executable.
- No login is required.
- The application is in this new top-level `control_panel` folder.
- The UI must show live progress/logs, retain run history, and offer output downloads.
- No other users will edit the shared ECM workbook while this workstation runs it.
- All settings, including secret PostgreSQL values and shared-folder paths, are maintained in a local `.env`; no settings screen is included.

## Architecture

- Keep the existing [django](../django) and [ecm](../ecm) scripts in place. The folder named `django` is an existing report-script location, not a Django web project.
- Use FastAPI with server-rendered Jinja pages and small browser-side JavaScript for status polling/log refresh. This is intentionally lighter than a separate SPA and API, appropriate for a local control panel.
- Use a dedicated application virtual environment. Preserve the existing isolated ECM and report virtual environments; launch the report generator with the Python interpreter in `django/.venv` and ECM jobs with `ecm/.venv`. This prevents dependency conflicts and retains the scripts' expected import paths.
- Wrap runs in a single-worker job coordinator. It creates a database record, starts the applicable command with a known working directory, captures stdout/stderr to a run log, and updates the job state through queued, running, succeeded, or failed. The UI disables all Run buttons while a job is active.
- Store application metadata in a local SQLite database under `control_panel/data`. Store each successful output and backup as a copied, immutable artifact inside a run-specific data directory, so a later run cannot overwrite a historical download. Original persistent workbooks in their configured/shared location remain the operational source of truth.
- Bind the server only to `127.0.0.1`; do not expose it to the LAN. Run it under the Windows account that has PostgreSQL connectivity and read/write permissions to every configured workbook/output folder.

## Proposed layout

- `control_panel/main.py` — FastAPI application factory, startup validation, routes, and local-only server configuration.
- `control_panel/services/runner.py` — queued one-at-a-time subprocess executor, stdout/stderr capture, log persistence, process exit handling, and artifact collection.
- `control_panel/services/workflows.py` — declarative definitions for the four button actions: labels, form fields, interpreter, working directory, script command, success artifacts, and preflight rules.
- `control_panel/services/history.py` — SQLite schema, job/artifact persistence, and cleanup/retention functions.
- `control_panel/services/config.py` — strongly validated environment configuration, normalized absolute paths, and safe child-process environment construction.
- `control_panel/templates` — dashboard, job detail/log view, history list, and error/not-found pages.
- `control_panel/static` — small CSS and JavaScript assets for accessible forms, polling, logs, and download links.
- `control_panel/tests` — route, configuration, job-coordinator, artifact, and workflow-command tests using a fake child process rather than real Excel/PostgreSQL runs.
- `control_panel/requirements.txt`, `.env.example`, `README.md`, and Windows launcher scripts — repeatable installation, configuration, local start, and backup instructions.

## Steps

### Phase 1 — Establish the application boundary and configuration

1. Create the independent `control_panel` structure and dependency manifest with FastAPI, Uvicorn, Jinja2, multipart form support, and test dependencies. Keep generated data, logs, downloaded artifacts, and actual `.env` ignored by Git.
2. Define one local `.env` contract: absolute locations of the two script interpreters, the report template/output directory, the ECM CSV directory, ECM/SMM templates and persistent working workbooks, SQLite/data root, and PostgreSQL values. Include `.env.example` with placeholders only; never copy a password into logs, pages, artifact names, or source control.
3. Validate configuration at startup and before each run: paths exist, each interpreter is executable, scripts/templates/SQL files are present, required folders are writable, data root is separate from the operational output paths, and the report database variables are supplied. Display actionable missing-setting errors rather than starting a broken run.
4. Resolve all configured paths before subprocess launch and construct commands as argument arrays, never by shell string concatenation. This retains paths with spaces and eliminates command injection from UI input.

### Phase 2 — Adapt each existing CLI as an application workflow

5. Model the Method Statement weekly report action with required start/end dates and configured template/output directory. Invoke the existing `main()` behavior through its current CLI, preserving its inclusive-date validation, PostgreSQL `.env` handling, and generated workbook naming. Register the newly generated report as the run's download artifact.
6. Model ECM Step 1a with required start/end dates and configured CSV/SMM/ECM paths. Let the existing script select the newest configured properties CSV unless the design later adds an explicit source selector. Preserve the script's `_1a` backup and capture both the updated ECM working workbook and its backup as artifacts.
7. Provide a clear manual-review state after Step 1a succeeds: show the shared ECM workbook location, state that blank Count cells require M/Remark dates, and prevent the UI from presenting Step 1a as equivalent to finalization. The user opens and edits the configured workbook in Excel outside the web UI.
8. Model ECM Step 1b with required start/end dates and configured ECM workbook. Before start, preflight that the workbook and expected period sheet exist and no Excel temporary lock file indicates that the workbook is open. Rely on its existing all-rows validation as the authoritative review-date check; surface its exact failure message/log with no persistent mutation on failure. On success, retain the finalized workbook, `_1b` backup, and `_1b_after` handoff backup artifacts. Step 1b saves and closes the workbook through openpyxl, then applies the Count=Y filter in one hidden Excel session: start Excel once, reopen the workbook once, save the filter, close the workbook once, and quit Excel once. Create `_1b_after` only after that Excel session succeeds. J/K decisions remain a separate manual Excel handoff after Step 1b.
9. Model ECM Step 2 with required start/end dates and configured ECM/SMM template/persistent workbook paths. Preserve its existing workbook/header/flag validation and capture the updated SMM workbook and dated backup. Indicate that J/K values must already be completed for Count=Y rows.
10. Do not import scripts directly into request handlers. Continue running each as its existing CLI from the folder it expects, which protects CLI compatibility and avoids changing the existing script tests. Add a short process timeout and report a timed-out run as failed; exclude unsafe “cancel” support for file-mutating Excel jobs.

### Phase 3 — Run history, artifacts, and UI

11. Create SQLite tables for runs, logs, and artifacts. Persist workflow name, sanitized parameters, timestamps, state, exit code, concise error summary, full captured log path, and artifact metadata/checksum. Never persist database passwords or a fully expanded sensitive environment.
12. Implement a dashboard with four explicit cards/buttons: Weekly Method Statement Report, ECM Step 1a — Create Review, ECM Step 1b — Finalize Review, and ECM Step 2 — Create SMM Summary. Each form requires start/end dates and uses browser and server validation that end date is not before start date. Show the intended input/output paths and a confirmation notice for actions that overwrite persistent workbooks.
13. Implement job detail pages with state, timestamps, selected dates, a polling live log panel, errors, output/backup download links, and the next required manual action. Refresh only the job state/log portion rather than reloading the whole page.
14. Implement a filterable history page showing every prior run, outcome, dates, duration, and links to its immutable log/artifacts. Use a configurable artifact-retention policy and clearly surface any missing copied artifact rather than linking to a mutable original.
15. Add friendly preflight and failure presentation for invalid dates, unavailable database configuration, missing/locked workbooks, invalid Excel manual-review data, missing CSVs, output permission errors, and failed backups. Preserve raw technical logs behind the run detail page for troubleshooting.

### Phase 4 — Windows operation and verification

16. Supply a PowerShell installer/start script that creates the application virtual environment, installs the application dependencies, checks that the existing two script environments are available, starts Uvicorn on `127.0.0.1`, and opens the default browser at the local dashboard. Document optional Windows Task Scheduler startup only after manual operation is proven.
17. Document workstation prerequisites: Python versions compatible with all three environments, Excel for manual ECM review, network access to PostgreSQL and shared workbook locations, write permissions, required environment values, backup/recovery procedure, and how to stop/restart the local service.
18. Retain and run existing ECM regression tests unchanged. Add UI/service tests for configuration failure, command construction, the single-active-run rule, streamed stdout/stderr, nonzero exit handling, artifact copying, path traversal prevention on downloads, and correctly rendered history/job states. Use temporary files and fake subprocesses for these tests.
19. Complete an operator acceptance run on copies of representative workbooks: generate a Django report; complete 1a; edit the shared workbook's M values and J/K choices in Excel; finalize 1b; generate Step 2; verify original templates were untouched, expected backups/artifacts are downloadable, and the same run details remain accurate after a later rerun.

## Relevant existing files

- [../django/Method_Statement_Interaction_Weekly_Report.py](../django/Method_Statement_Interaction_Weekly_Report.py) — retains the weekly-report command and its `parse_arguments()`, `get_database_settings()`, and `main()` contracts; the application supplies the date/template/output arguments through a child process.
- [../django/requirements.txt](../django/requirements.txt) — existing isolated report dependencies; keep separate from UI dependencies.
- [../ecm/Step_1a_Create_ECM_Record.py](../ecm/Step_1a_Create_ECM_Record.py) — retain `main()` and `apply_preliminary_counting_rules()` as the button action's existing implementation.
- [../ecm/Step_1b_Finalize_ECM_Record.py](../ecm/Step_1b_Finalize_ECM_Record.py) — retain `finalize_pending_rows()` validation/backup behavior and its period-sheet selection.
- [../ecm/Step_2_Create_SMM_Summary.py](../ecm/Step_2_Create_SMM_Summary.py) — retain `aggregate_ecm_record()` and `create_smm_workbook()` behavior for the final action.
- [../ecm/Step_1_Create_ECM_Record.py](../ecm/Step_1_Create_ECM_Record.py) — remains a legacy dependency of Step 1a/1b and must not be moved or changed as part of the UI work.
- [../ecm/requirements.txt](../ecm/requirements.txt) — existing isolated ECM dependency set; keep separate from UI dependencies.
- [../ecm/step_new.md](../ecm/step_new.md) — source of the deliberate 1a → Excel manual review → 1b workflow contract and immutable-template rules.
- [../ecm/tests/test_step_1a_1b.py](../ecm/tests/test_step_1a_1b.py) — retain as a regression suite while adding UI-specific tests separately.

## Verification

1. Start the application on the workstation and confirm it is reachable only through localhost; attempt LAN access and verify it is unavailable.
2. Start each action with valid and invalid date ranges and verify the UI has one active job maximum, validation is shown without execution, and logs stream until a terminal status.
3. Simulate subprocess output, error output, nonzero exit, timeout, and missing executable in automated tests; verify history/log persistence and no false success artifacts.
4. Verify a successful run copies the generated persistent workbook and backup into that run's artifact directory; run again for the same reporting period and verify the first run's downloads are unchanged.
5. Verify Step 1b rejects an Excel-open temporary lock and a missing/incomplete M/Remark review without mutating the working workbook; complete valid review dates and verify the existing `_1b` and `_1b_after` backup/finalization rules remain intact. Mock the filter helper to verify one call after finalization and no call after validation failure.
6. Run the existing ECM test suite and the new application suite, then perform the complete four-action acceptance sequence with representative non-production inputs.

## Scope boundaries

- **Included:** one-workstation localhost UI, four button-triggered workflows, dates/forms, one-at-a-time execution, progress/logs, historical results, artifact downloads, local `.env` configuration, Windows launch documentation, and Excel manual handoff.
- **Excluded:** authentication/SSO, LAN/server hosting, multiple concurrent users/runners, a settings-management UI, browser-based Excel editing, automatic scheduling, a packaged executable, automatic Excel conflict resolution, and cancellation of in-flight file-mutating jobs.
- **Future central deployment:** can reuse the same workflow definitions and UI, but must add authentication, authorization, multi-user locking, secure secrets management, and a production job worker before exposure beyond localhost.
