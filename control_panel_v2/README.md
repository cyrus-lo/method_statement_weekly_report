# Workflow Control Panel V2

V2 is a separate, localhost-only Workflow Control Panel. It runs on `127.0.0.1:8766` and is **read only by default**. Set `CONTROL_PANEL_V2_PREVIEW_ONLY=false` only after the cutover checks below are complete. The header mode button can switch the current server process between preview-only and live execution mode; this choice is runtime-only and resets to the environment setting after a restart.

## Safety model

- V2 has its own virtual environment, SQLite database, logs, and copied artifacts under `control_panel_v2/data`.
- V2 does not share V1's database or data directory.
- In preview mode, V2 rejects workflow execution before it creates a run or accesses a workbook.
- Report History owns the period workflow: ECM Step 0, Step 1a, Step 1b, and Step 2. Step 0 is not exposed on Manual Run.
- Step 0 uses a separate Playwright/Edge worker and a dedicated persistent Edge profile. The user completes login/SSO/MFA manually with `--login`; export mode does not automate credentials.
- Dashboard groups status by reporting period. Each period has an independent Weekly Report and ECM workflow sequence.
- ECM prerequisites and date locking are enforced per period. The runner still executes one child process at a time, and the current ECM/SMM workbooks remain shared, so this change does not enable simultaneous workbook runs for different periods.
- Step 1b finalizes Count/Remark and creates its pre-1b backup, then uses one hidden Excel session to reopen the saved workbook and apply the Count=Y filter. Only after Excel saves and closes successfully is the post-1b backup created; Use SMM/Use KMS are completed manually in a separate Excel session after Step 1b succeeds.
- If Step 1b has no run recorded for a period, Step 2 can be started only after explicitly confirming the warning in Report History. This bypass skips Step 1b review/finalization and its backups; ensure all required ECM review, Count, Use SMM, and Use KMS values are already complete. A queued, running, or failed Step 1b run cannot be bypassed.
- Step 2 updates the live ECM workbook's Display values and SMM workbook; it does not overwrite the post-Step-1b ECM snapshot.

## First-time history import

Use an empty V2 data directory. The importer uses SQLite's backup API to obtain a consistent database snapshot and copies only immutable logs/artifacts into V2 paths:

```text
control_panel_v2/.venv/Scripts/python.exe -m control_panel_v2.services.bootstrap --source-data-dir control_panel/data
```

Do not run a second import into a V2 data directory that contains history.

## Start

Run `start_control_panel_v2.ps1`. It provisions the V2 environment, starts Uvicorn on port 8766, and opens the V2 Dashboard.

To restart the app, run `restart_control_panel_v2.ps1`. It stops only the V2 Uvicorn process, waits for port `8766` to be released, and starts the app again. If workflow child processes are active, the restart is refused; use `restart_control_panel_v2.ps1 -Force` only when those workers should also be stopped. Add `-NoBrowser` to avoid opening a browser window.

## Step 0 worker

Provision the separate worker environment, then install its browser runtime:

```text
ecm/step0_worker/.venv/Scripts/python.exe -m pip install -r ecm/step0_worker/requirements.txt
ecm/step0_worker/.venv/Scripts/python.exe -m playwright install msedge
```

The V2 Step 0 development default opens the dedicated Edge profile visibly so
ECM navigation can be inspected. Set `STEP0_HEADED=false` to use the normal
headless-first behavior; when `STEP0_LOGIN_IF_NEEDED=true`, authentication
still opens the dedicated profile visibly only when required. Development
defaults are 60 seconds per ECM operation and 120 seconds for interactive
authentication. On a headed failure, the browser is retained for 120 seconds
by default and then closed automatically. Override these values with
`STEP0_HEADED`, `STEP0_FAILURE_BROWSER_HOLD_SECONDS`,
`STEP0_TIMEOUT_SECONDS`, and `STEP0_AUTH_TIMEOUT_SECONDS` when needed.

The dedicated Edge profile is persistent and is reused between runs. To recover
from a suspected stale session, set `STEP0_RESET_PROFILE=true` for one V2 run.
Before starting the browser, Step 0 renames the existing profile to a timestamped
`.backup-*` directory and creates a fresh profile. Complete ECM login/SSO/MFA in
the visible Edge window, then set the option back to `false`; it is deliberately
not enabled by default and should not be used on every run.

Before the worker opens the ECM Search UI, it verifies that Navigator has
rendered a usable shell. A blank document, login page, wrong origin, or missing
Navigator shell fails with a specific readiness reason rather than a misleading
Search-selector error. Failures save a screenshot, DOM snapshot, and a compact
browser diagnostics JSON file directly in that failed run's artifacts. The
JSON records safe Console, page-error, and failed-request metadata; it does not
record credentials, cookies, tokens, or full page text. These diagnostics may
still contain ECM data and should be handled as sensitive files.

For manual session setup, authentication can still be completed explicitly:

```text
ecm/step0_worker/.venv/Scripts/python.exe ecm/step0_worker/export_worker.py --login --profile-dir control_panel_v2/.step0-edge-profile
```

The worker validates the 11-column `Export All` CSV, rejects empty results, and publishes timestamped files under `ecm/csv_from_ecm` without overwriting an existing export.

## Cutover

Before enabling execution in V2: stop V1, verify no V1 run is active, close Excel, back up V1 data, make a final import, install the Step 0 worker environment, authenticate the dedicated Edge profile, and change `CONTROL_PANEL_V2_PREVIEW_ONLY` to `false`. V2 runs one child process at a time and retains logs and copied artifacts in its own data directory.
