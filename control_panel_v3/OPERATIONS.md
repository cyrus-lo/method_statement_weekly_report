# V3 Operations and Cutover Runbook

This runbook applies only to `control_panel_v3`. V1 and V2 are separate services and must not be stopped, edited, or used as V3 data sources during routine operation. V3 defaults to preview-only and binds to localhost. No production workflow should be run until its owner approves the change window and the checks below are complete.

## 1. Preflight

1. Confirm V2 is healthy on port 8766 and leave it running. Confirm V3's configured port (default 8767) is not occupied by any other service.
2. Review `control_panel_v3/.env` if present. Keep `CONTROL_PANEL_V3_HOST` at `127.0.0.1` or `localhost`; do not use 8766. Keep every V3 writable path inside `control_panel_v3/` and distinct from other writable paths.
3. Confirm the ECM/SMM templates, readiness workbook, and browser profile are V3-owned copies. Confirm the Step 1a/Step 2 readiness workbook contains the intended operator decisions. Do not assume a new template copy contains current approvals.
4. Confirm V3 is still preview-only in the UI. Check the V3 period history, current active run, archive page, and available disk space. Ensure no queued/running V3 job exists before restart or cutover.
5. For weekly reporting, verify the read-only PostgreSQL settings in `django/.env`: required `PGHOST`, `PGDATABASE`, `PGUSER`, and `PGPASSWORD`; optional `PGPORT` (default `5432`) and `PGSSLMODE`; and `REPORTING_TIMEZONE` (default `Asia/Hong_Kong`). The V3 weekly worker reads this file as the sole settings source. Never put secrets in run logs or command-line arguments; V3 passes them only through the isolated worker environment.
6. Review unmatched first-date lookup rows before Step 1b. Review Step 2's Use SMM / Use KMS decisions and backups before execution. A bypass is an exception: verify skipped work and backups manually and retain the run's bypass audit marker.

## 2. Start and readiness

From the workspace root, run `control_panel_v3/start_control_panel_v3.ps1`. It installs the V3 requirements into V3's virtual environment, refuses a non-local bind or occupied port, starts only `control_panel_v3.main:app`, and waits for the V3 `/health` endpoint before opening the browser. The health endpoint is a process-readiness check, not proof that ECM, Excel, or PostgreSQL integrations are healthy.

Open `http://127.0.0.1:<CONTROL_PANEL_V3_PORT>/`. Verify the page title identifies V3, the mode says Preview only, and the period/report history loads. Do not switch to live mode merely to test startup.

## 3. Preview rehearsal (no live integrations)

1. Keep Preview only enabled. Create a disposable reporting period and inspect the available workflow states, gates, lookup scope selector, and Run Next Step control.
2. Submit a workflow in preview mode and confirm the page reports preview-only and no run record is created.
3. Inspect a prior V3 run detail and `/history` for status, artifact links, lookup scope/counts, and any Step 1b bypass audit marker.
4. Confirm the V2 Archive is read-only, clearly separated, and does not satisfy V3 workflow gates.
5. Remove only the disposable, no-run period when appropriate. Do not delete runs or evidence as part of a rehearsal.

## 4. Live cutover (operator approval required)

1. Use a scheduled change window and obtain approval from the ECM/workbook owner. Ensure operators are available for SSO/MFA and the manual Use SMM / Use KMS review.
2. Use only V3's dedicated Edge profile; never share, copy, or reset V2's profile. Be available to complete SSO/MFA manually when prompted. A prior sign-in does not guarantee that the session will remain authenticated; do not automate credentials or MFA.
3. Verify backups and output paths resolve under V3. Confirm source templates and V1/V2 data remain untouched.
4. Switch the V3 process to Live in the UI. This mode flag is process-local and resets to the configured preview default on restart; it is not an authorization system.
5. Execute one approved period at a time, follow the gated Step 0 → Step 1a → first-date lookup → Step 1b → manual review → Step 2 sequence, and inspect outputs before continuing. Weekly reporting is independent.
6. Validate CSV row counts, lookup hit/miss summary, workbook sheets/counts, summary dates/formulas, artifact downloads, and preserved unrelated workbook content. Stop on any unexplained mismatch; do not blindly rerun an upstream step because downstream output may be stale.
7. Record the operator, change ticket, period, run IDs, validation, and approval in the team's normal change record. V3 history records workflow parameters and execution evidence but does not replace change management.

### Step 0 parity: operator guidance and implementation status

By default Step 0 opens a visible V3 Edge window immediately. If prompted, complete SSO/MFA in that window within the authentication timeout and leave it open; after readiness is detected, the run continues on the same page. If configured headless, a visible sign-in fallback is used when login is needed and enabled. No credentials or MFA are automated, and a persistent profile does not guarantee successful authentication.

The V3 settings and defaults are:

- `CONTROL_PANEL_V3_STEP0_HEADED=true`: visible by default; `false` selects headless.
- `CONTROL_PANEL_V3_STEP0_LOGIN_IF_NEEDED=true`: permit manual sign-in and the visible fallback when needed.
- `CONTROL_PANEL_V3_STEP0_TIMEOUT_SECONDS=60`: browser operation timeout in seconds.
- `CONTROL_PANEL_V3_STEP0_AUTH_TIMEOUT_SECONDS=120`: manual sign-in wait in seconds.
- `CONTROL_PANEL_V3_STEP0_FAILURE_BROWSER_HOLD_SECONDS=120`: bounded visible failure-window hold in seconds; `0` disables it.

The two timeouts must be positive and the hold non-negative. Allow sufficient `CONTROL_PANEL_V3_JOB_TIMEOUT_SECONDS` for the waits and export; the minimum validation is not a guarantee of sufficient time for all stages. Restart V3 safely to load configuration changes.

On timeout, cancellation, or another failure, inspect the V3 run status and log, stop before retrying, and investigate the reported stage; do not assume every readiness failure is an SSO/MFA problem. Preserve run-local screenshots, DOM snapshots, and diagnostics as restricted evidence: they may contain sensitive ECM data. Limit access and retention according to local policy and do not share them casually. Reset the V3 profile only deliberately using the existing opt-in and its backup; never use V2's profile as a workaround.

An operator-approved live smoke test is still required: confirm the visible startup, manual sign-in if prompted, continuation, and CSV contents/handoff for the approved period. Preserve the run ID and restricted diagnostics and stop on failure. Neither passing mocked tests nor a health check proves ECM authentication is fixed.

### First-date lookup sign-in

The lookup opens a visible Edge window with the dedicated V3 profile. If prompted, complete sign-in and MFA in that window and leave it open; after the login screen clears, the worker opens the `Search ECM Document` saved search. The lookup uses `CONTROL_PANEL_V3_WORKER_TIMEOUT_SECONDS` for its authentication wait and browser operations, independently of the Step 0-specific settings. A cleared login screen does not by itself prove the saved search is ready; a saved-search failure is reported separately. Screenshots, page snapshots, and browser diagnostics may contain sensitive data and must be treated as restricted evidence.

Close the V3 working ECM workbook in Excel before running the lookup. After a successful run, review the lookup summary and any unmatched rows before Step 1b. Successful row matches are saved independently, so a failed run may have already updated some rows; inspect the working workbook and run log before choosing a rerun scope. Do not copy or reset another version's profile, and never enter credentials into logs or chat.

For each eligible workbook row, the lookup uses the normalized outgoing reference in Excel column P; the Step 0 CSV is used to validate the row identity, not as a fallback search value. Before each search the worker verifies `In/Out = OUT`, `Outgoing ref. no. = Starts With`, and ECM `Version = All versions` through the visible UI. It rejects a still-visible old result grid unless a fresh result-pane update is observed, waits for pagination to advance, and logs per-query counts and result dates. A blank column-P reference or an unverified fresh result stops the lookup rather than writing a date. These checks are mocked in tests; operator-approved live validation is still required.

## 5. Restart and recovery

Run `control_panel_v3/restart_control_panel_v3.ps1`. It resolves the same host/port settings, stops only a process whose command line identifies the V3 ASGI app, refuses unrelated port owners, and refuses to restart while V3 has child processes unless `-Force` is supplied. It waits for the port to be released, starts V3, and verifies `/health`.

Use `-Force` only with explicit operator approval after confirming the worker can be interrupted safely. Forced termination may leave an external browser/Excel action or a partially completed external workflow; inspect the run log, status, workbook backups, staged outputs, and ECM state before retrying. Never delete the run database or evidence to clear a stuck run. If a port release or HTTP health check times out, stop and inspect the exact V3 process and logs before another restart.

## 6. Rollback

1. Return V3 to Preview only if it is responsive; otherwise stop only the verified V3 server process. Do not stop a process based solely on its port number.
2. Preserve the V3 database, run logs, artifacts, workbook backups, and browser diagnostics for investigation. Treat screenshots/DOM snapshots as sensitive evidence.
3. Restore a workbook only from the corresponding verified V3 backup after comparing timestamps and obtaining workbook-owner approval. Do not restore from V1/V2 or overwrite later operator changes.
4. Keep V2 on its existing port, database, source tree, and browser profile. V3 rollback must not require a V2 migration or modification.
5. Document the failed run and recovery decision. Resume only after the cause is understood and the owner approves another attempt.

## Safety boundary

Preview rehearsals and mocked tests are not end-to-end validation of live ECM, desktop Excel COM, or PostgreSQL connectivity. Live production readiness requires an operator-controlled smoke test and reconciliation against the authoritative source and destination records. Do not claim production readiness based only on `/health` or passing unit tests.
