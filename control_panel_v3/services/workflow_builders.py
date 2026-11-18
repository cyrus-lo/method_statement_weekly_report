"""Build shell-free commands using only V3-owned inputs and outputs."""

from __future__ import annotations

import os
from datetime import date
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from .config import CONTROL_PANEL_ROOT, Settings
from .workflows import ArtifactSource, PreparedWorkflow, WORKFLOWS


class PreflightError(ValueError):
    """A workflow cannot be safely prepared with current V3 state."""


def period_sheet_name(start_date: date, end_date: date) -> str:
    if start_date.month == end_date.month:
        return f"ECM_{start_date.day}-{end_date.day} {end_date:%b} (WIP)"
    return f"ECM_{start_date.day} {start_date:%b}-{end_date.day} {end_date:%b} (WIP)"


def _timestamped_csv(csv_dir: Path, start: date, end: date, run_id: int) -> Path:
    timestamp = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S")
    return csv_dir / f"properties_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_{timestamp}_v3run{run_id}.csv"


def _period_csv(csv_dir: Path, start: date, end: date) -> Path:
    return csv_dir / f"properties_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.csv"


def _inside(path: Path) -> bool:
    try:
        path.resolve().relative_to(CONTROL_PANEL_ROOT.resolve())
        return True
    except ValueError:
        return False


def _require_v3_file(path: Path, label: str) -> None:
    if not _inside(path) or not path.is_file():
        raise PreflightError(f"Required V3 {label} is missing or outside the V3 root: {path}")


def _require_v3_runtime(path: Path, label: str) -> None:
    if not path.is_file():
        raise PreflightError(f"Required {label} was not found: {path}")


def prepare_workflow(workflow_id: str, start: date, end: date, run_id: int, settings: Settings, *, lookup_scope: str = "all", previous_lookup_summary: Path | None = None) -> PreparedWorkflow:
    if workflow_id not in WORKFLOWS:
        raise PreflightError(f"Unknown V3 workflow: {workflow_id}")
    if end < start:
        raise PreflightError("End date cannot be earlier than start date.")
    definition = WORKFLOWS[workflow_id]
    python = str(settings.python_executable)
    if workflow_id == "weekly-report":
        _require_v3_runtime(settings.python_executable, "Python interpreter")
        _require_v3_file(settings.weekly_worker_script, "weekly-report worker")
        _require_v3_file(settings.weekly_template_path, "weekly-report template")
        if not settings.weekly_reporting_timezone:
            raise PreflightError("CONTROL_PANEL_V3_WEEKLY_REPORTING_TIMEZONE must not be empty.")
        database_environment = {
            "CONTROL_PANEL_V3_WEEKLY_PGHOST": "PGHOST",
            "CONTROL_PANEL_V3_WEEKLY_PGDATABASE": "PGDATABASE",
            "CONTROL_PANEL_V3_WEEKLY_PGUSER": "PGUSER",
            "CONTROL_PANEL_V3_WEEKLY_PGPASSWORD": "PGPASSWORD",
        }
        missing = [name for name in database_environment if not os.getenv(name, "").strip()]
        if missing:
            raise PreflightError("Set the V3 weekly database settings before running this workflow: " + ", ".join(missing))
        try:
            ZoneInfo(settings.weekly_reporting_timezone)
        except (ValueError, ZoneInfoNotFoundError) as error:
            raise PreflightError(f"Invalid V3 weekly reporting timezone: {settings.weekly_reporting_timezone}") from error
        output_dir = settings.django_output_dir / f"weekly-run-{run_id}"
        if not _inside(output_dir):
            raise PreflightError("V3 weekly-report output must remain under control_panel_v3.")
        output_name = f"Method_Statement_Interaction_Weekly_Report_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx"
        output_path = output_dir / output_name
        command = [python, "-m", "control_panel_v3.workers.weekly_report", "--start-date", start.isoformat(), "--end-date", end.isoformat(), "--template", str(settings.weekly_template_path), "--output-dir", str(output_dir)]
        environment_overrides = {
            "PATH": os.getenv("PATH", ""),
            "SYSTEMROOT": os.getenv("SYSTEMROOT", ""),
            "WINDIR": os.getenv("WINDIR", ""),
            "TEMP": os.getenv("TEMP", ""),
            "TMP": os.getenv("TMP", ""),
            "PGHOST": "",
            "PGDATABASE": "",
            "PGUSER": "",
            "PGPASSWORD": "",
            "PGPORT": "",
            "PGSSLMODE": "",
            "REPORTING_TIMEZONE": "",
            "CONTROL_PANEL_V3_WEEKLY_PGHOST": os.environ["CONTROL_PANEL_V3_WEEKLY_PGHOST"],
            "CONTROL_PANEL_V3_WEEKLY_PGDATABASE": os.environ["CONTROL_PANEL_V3_WEEKLY_PGDATABASE"],
            "CONTROL_PANEL_V3_WEEKLY_PGUSER": os.environ["CONTROL_PANEL_V3_WEEKLY_PGUSER"],
            "CONTROL_PANEL_V3_WEEKLY_PGPASSWORD": os.environ["CONTROL_PANEL_V3_WEEKLY_PGPASSWORD"],
            "CONTROL_PANEL_V3_WEEKLY_PGPORT": os.getenv("CONTROL_PANEL_V3_WEEKLY_PGPORT", "5432").strip() or "5432",
            "CONTROL_PANEL_V3_WEEKLY_PGSSLMODE": os.getenv("CONTROL_PANEL_V3_WEEKLY_PGSSLMODE", "").strip(),
            "CONTROL_PANEL_V3_WEEKLY_REPORTING_TIMEZONE": settings.weekly_reporting_timezone,
        }
        return PreparedWorkflow(
            definition,
            command,
            settings.project_root,
            (ArtifactSource("Weekly Method Statement report", output_path),),
            environment_overrides=environment_overrides,
        )
    if workflow_id == "ecm-step-0":
        script = settings.step0_worker_script
        _require_v3_file(script, "Step 0 worker")
        _require_v3_runtime(settings.python_executable, "Python interpreter")
        output = _timestamped_csv(settings.csv_dir, start, end, run_id)
        period_input = _period_csv(settings.csv_dir, start, end)
        diagnostics_dir = settings.runs_dir / str(run_id) / "diagnostics"
        if not _inside(output) or not _inside(period_input) or not _inside(settings.edge_profile_dir) or not _inside(diagnostics_dir):
            raise PreflightError("V3 Step 0 output and Edge profile must remain under control_panel_v3.")
        command = [python, "-m", "control_panel_v3.workers.ecm_step0_export", "--start-date", start.isoformat(), "--end-date", end.isoformat(), "--output", str(output), "--period-input", str(period_input), "--ecm-url", settings.ecm_url, "--profile-dir", str(settings.edge_profile_dir), "--root-dir", str(CONTROL_PANEL_ROOT), "--timeout-seconds", str(settings.worker_timeout_seconds), "--login-if-needed"]
        command.extend(["--diagnostics-dir", str(diagnostics_dir)])
        if os.getenv("CONTROL_PANEL_V3_RESET_EDGE_PROFILE", "false").strip().casefold() in {"1", "true", "yes", "on"}:
            command.append("--reset-profile")
        diagnostic_artifacts = tuple(ArtifactSource(f"Sensitive browser diagnostic: {path}", diagnostics_dir / path, required=False) for path in ("failure-step0.png", "failure-step0.html", "failure-step0.diagnostics.json"))
        return PreparedWorkflow(definition, command, settings.project_root, (ArtifactSource("ECM properties CSV", output), ArtifactSource("ECM period CSV", period_input), *diagnostic_artifacts))

    if workflow_id == "ecm-step-1a":
        period_csv = _period_csv(settings.csv_dir, start, end)
        _require_v3_file(period_csv, "exact-period CSV")
        _require_v3_runtime(settings.python_executable, "Python interpreter")
        _require_v3_file(settings.step1a_worker_script, "Step 1a worker")
        workbook_path = settings.ecm_output_dir / "ECM_Record.xlsx"
        backup_path = settings.ecm_output_dir / "backups" / f"ECM_Record_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_1a_run{run_id}.xlsx"
        for label, path in (("readiness workbook", settings.readiness_workbook_path), ("ECM template", settings.ecm_template_path)):
            _require_v3_file(path, label)
        for path in (workbook_path, backup_path):
            if not _inside(path):
                raise PreflightError("V3 Step 1a workbook and backup paths must remain under control_panel_v3.")
        if not _inside(settings.ecm_template_path) or not _inside(settings.readiness_workbook_path):
            raise PreflightError("V3 Step 1a template and readiness workbook paths must remain under control_panel_v3.")
        if workbook_path.with_name(f"~${workbook_path.name}").exists():
            raise PreflightError("Close the V3 ECM workbook before running Step 1a.")
        command = [python, "-m", "control_panel_v3.workers.ecm_step1a", "--start-date", start.isoformat(), "--end-date", end.isoformat(), "--csv", str(period_csv), "--readiness-workbook", str(settings.readiness_workbook_path), "--workbook", str(workbook_path), "--template", str(settings.ecm_template_path), "--backup", str(backup_path)]
        return PreparedWorkflow(definition, command, settings.project_root, (ArtifactSource("V3 ECM review workbook", workbook_path), ArtifactSource("Pre-Step-1a ECM backup", backup_path)))

    if workflow_id == "ecm-first-date-lookup":
        period_csv = _period_csv(settings.csv_dir, start, end)
        _require_v3_file(period_csv, "period CSV")
        workbook_path = settings.ecm_output_dir / "ECM_Record.xlsx"
        worksheet = period_sheet_name(start, end)
        if not workbook_path.is_file():
            raise PreflightError("V3 ECM_Record.xlsx is not available; complete Step 1a first.")
        if not _inside(workbook_path) or not _inside(settings.edge_profile_dir):
            raise PreflightError("V3 workbook and Edge profile must remain under control_panel_v3.")
        if workbook_path.with_name(f"~${workbook_path.name}").exists():
            raise PreflightError("Close the V3 ECM workbook before running first-date lookup.")
        script = settings.lookup_worker_script
        _require_v3_file(script, "first-date lookup worker")
        _require_v3_runtime(settings.python_executable, "Python interpreter")
        summary_path = settings.runs_dir / str(run_id) / "lookup-summary.json"
        diagnostics_dir = settings.runs_dir / str(run_id) / "diagnostics"
        if not _inside(summary_path) or not _inside(diagnostics_dir):
            raise PreflightError("V3 lookup summary and diagnostics must remain under the V3 run root.")
        command = [python, "-m", "control_panel_v3.workers.ecm_outgoing_reference_lookup", "--workbook", str(workbook_path), "--worksheet", worksheet, "--csv", str(period_csv), "--profile-dir", str(settings.edge_profile_dir), "--ecm-url", settings.ecm_url, "--timeout-seconds", str(settings.worker_timeout_seconds), "--summary", str(summary_path), "--diagnostics-dir", str(diagnostics_dir), "--scope", lookup_scope]
        if lookup_scope == "previously-unmatched":
            if previous_lookup_summary is None or not _inside(previous_lookup_summary) or not previous_lookup_summary.is_file():
                raise PreflightError("Previously-unmatched rerun requires the prior V3 lookup summary artifact.")
            command.extend(["--previous-summary", str(previous_lookup_summary)])
        if os.getenv("CONTROL_PANEL_V3_RESET_EDGE_PROFILE", "false").strip().casefold() in {"1", "true", "yes", "on"}:
            command.append("--reset-profile")
        diagnostic_artifacts = tuple(ArtifactSource(f"Sensitive browser diagnostic: {path}", diagnostics_dir / path, required=False) for path in ("failure-lookup.png", "failure-lookup.html", "failure-lookup.diagnostics.json"))
        return PreparedWorkflow(definition, command, settings.project_root, (ArtifactSource("ECM review workbook", workbook_path), ArtifactSource("First-date lookup summary", summary_path), *diagnostic_artifacts))

    if workflow_id == "ecm-step-1b":
        _require_v3_runtime(settings.python_executable, "Python interpreter")
        _require_v3_file(settings.step1b_worker_script, "Step 1b worker")
        workbook_path = settings.ecm_output_dir / "ECM_Record.xlsx"
        _require_v3_file(workbook_path, "ECM review workbook")
        worksheet = period_sheet_name(start, end)
        backup_path = settings.ecm_output_dir / "backups" / f"ECM_Record_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_1b_run{run_id}.xlsx"
        post_backup_path = settings.ecm_output_dir / "snapshots" / f"ECM_Record_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}_1b_after_run{run_id}.xlsx"
        if workbook_path.with_name(f"~${workbook_path.name}").exists():
            raise PreflightError("Close the V3 ECM workbook before running Step 1b.")
        if any(not _inside(path) for path in (backup_path, post_backup_path)):
            raise PreflightError("V3 Step 1b workbook snapshots must remain under control_panel_v3.")
        command = [python, "-m", "control_panel_v3.workers.ecm_step1b", "--start-date", start.isoformat(), "--end-date", end.isoformat(), "--workbook", str(workbook_path), "--worksheet", worksheet, "--backup", str(backup_path), "--post-backup", str(post_backup_path), "--excel-filter-timeout-seconds", str(settings.excel_filter_timeout_seconds)]
        return PreparedWorkflow(definition, command, settings.project_root, (
            ArtifactSource("V3 finalized ECM workbook", workbook_path),
            ArtifactSource("Pre-Step-1b ECM backup", backup_path),
            ArtifactSource("Post-Step-1b ECM snapshot", post_backup_path),
        ))
    if workflow_id == "ecm-step-2":
        _require_v3_runtime(settings.python_executable, "Python interpreter")
        script = CONTROL_PANEL_ROOT / "workers" / "ecm_step2.py"
        _require_v3_file(script, "Step 2 worker")
        ecm_workbook = settings.ecm_output_dir / "ECM_Record.xlsx"
        smm_workbook = settings.smm_output_dir / "SMM_KMS_PM_Master_Document.xlsx"
        ecm_sheet = period_sheet_name(start, end)
        backup = settings.smm_backup_dir / f"SMM_KMS_PM_Master_Document_{start:%Y-%m-%d}_to_{end:%Y-%m-%d}.xlsx"
        for label, path in (("ECM review workbook", ecm_workbook), ("V3 SMM template", settings.smm_template_path), ("V3 readiness workbook", settings.readiness_workbook_path)):
            _require_v3_file(path, label)
        for label, path in (("SMM workbook", smm_workbook), ("SMM backup", backup)):
            if not _inside(path):
                raise PreflightError(f"V3 Step 2 {label} must remain under control_panel_v3.")
        if ecm_workbook.with_name(f"~${ecm_workbook.name}").exists() or smm_workbook.with_name(f"~${smm_workbook.name}").exists():
            raise PreflightError("Close the V3 ECM and SMM workbooks before running Step 2.")
        command = [python, "-m", "control_panel_v3.workers.ecm_step2", "--start-date", start.isoformat(), "--end-date", end.isoformat(), "--ecm-workbook", str(ecm_workbook), "--ecm-sheet", ecm_sheet, "--smm-workbook", str(smm_workbook), "--smm-template", str(settings.smm_template_path), "--readiness-workbook", str(settings.readiness_workbook_path), "--summary-layout-sheet", settings.summary_layout_sheet, "--backup", str(backup)]
        return PreparedWorkflow(definition, command, settings.project_root, (
            ArtifactSource("V3 updated ECM workbook", ecm_workbook),
            ArtifactSource("V3 SMM working workbook", smm_workbook),
            ArtifactSource("Post-Step-2 SMM backup", backup),
        ))
    raise PreflightError(f"V3 live command builder is not available for {workflow_id} yet.")
