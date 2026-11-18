# V3 Operations and Cutover Runbook

This runbook applies only to `control_panel_v3`. V1 and V2 are separate services and must not be stopped, edited, or used as V3 data sources during routine operation. V3 defaults to preview-only and binds to localhost. No production workflow should be run until its owner approves the change window and the checks below are complete.

## 1. Preflight

1. Confirm V2 is healthy on port 8766 and leave it running. Confirm V3's configured port (default 8767) is not occupied by any other service.
2. Review `control_panel_v3/.env` if present. Keep `CONTROL_PANEL_V3_HOST` at `127.0.0.1` or `localhost`; do not use 8766. Keep every V3 writable path inside `control_panel_v3/` and distinct from other writable paths.
3. Confirm the ECM/SMM templates, readiness workbook, and browser profile are V3-owned copies. Confirm the Step 1a/Step 2 readiness workbook contains the intended operator decisions. Do not assume a new template copy contains current approvals.
4. Confirm V3 is still preview-only in the UI. Check the V3 period history, current active run, archive page, and available disk space. Ensure no queued/running V3 job exists before restart or cutover.
5. For weekly reporting, verify read-only PostgreSQL credentials are configured in the local environment and the selected period/timezone are correct. Never put secrets in run logs or command-line arguments.
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
2. Verify V3's dedicated Edge profile is authenticated manually. Do not automate credentials or MFA.
3. Verify backups and output paths resolve under V3. Confirm source templates and V1/V2 data remain untouched.
4. Switch the V3 process to Live in the UI. This mode flag is process-local and resets to the configured preview default on restart; it is not an authorization system.
5. Execute one approved period at a time, follow the gated Step 0 → Step 1a → first-date lookup → Step 1b → manual review → Step 2 sequence, and inspect outputs before continuing. Weekly reporting is independent.
6. Validate CSV row counts, lookup hit/miss summary, workbook sheets/counts, summary dates/formulas, artifact downloads, and preserved unrelated workbook content. Stop on any unexplained mismatch; do not blindly rerun an upstream step because downstream output may be stale.
7. Record the operator, change ticket, period, run IDs, validation, and approval in the team's normal change record. V3 history records workflow parameters and execution evidence but does not replace change management.

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
