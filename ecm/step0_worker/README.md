# ECM Step 0 browser worker

This worker opens the approved `SMM_ECM_Reporting` saved search in IBM Content
Navigator, changes only the inclusive `Document date` range, runs Search, and
downloads `Export All` as a validated properties CSV.

The worker uses a dedicated persistent Microsoft Edge profile and never
controls the desktop-wide cursor. Before visible locator actions, it may move
Playwright's page-scoped cursor through timed intermediate coordinates and add
optional short pauses. A small number of optional pauses can be longer, but
every such pause is capped at 1.5 seconds. Locator activation and its DOM
dispatch fallback remain the reliable action path.

## Setup

Create the separate environment and install the worker dependency:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

The V2 settings point to this environment by default:

`ecm/step0_worker/.venv/Scripts/python.exe`

## Manual login/session refresh

Run the worker in login mode from the repository root. Complete ECM login,
SSO, and MFA in the dedicated visible Edge window, then return to the terminal
and press Enter:

```powershell
ecm\step0_worker\.venv\Scripts\python.exe ecm\step0_worker\export_worker.py `
  --login `
  --profile-dir control_panel_v2\.step0-edge-profile
```

The profile directory is local-only and must not be committed or copied as an
artifact.

## V2 export mode

V2 supplies the period dates and timestamped output path. Export mode is
headless by default and fails with a clear authentication message if the
dedicated session has expired.

V2 can pass `--login-if-needed`. In that mode the worker first tries the export
headlessly. If the dedicated profile is on an ECM login or SSO page, it opens the
same profile in a visible Edge window so the user can complete password or MFA
entry. The worker detects the logged-in Search page, closes the visible window,
and resumes the export automatically. It never asks for a password on stdin,
because the control panel intentionally disconnects child-process stdin.

The output is validated against the 11-column contract consumed by Step 1a.
No-result searches fail and do not publish a final CSV.
