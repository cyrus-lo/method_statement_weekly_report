"""FastAPI entry point for the isolated V3 workflow control panel."""

from __future__ import annotations

import json
from datetime import date, datetime, timedelta
from urllib.parse import urlencode, urlsplit
from zoneinfo import ZoneInfo

from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from .services.archive import ArchiveStore
from .services.bootstrap import import_v2_archive
from .services.config import ConfigurationError, load_settings
from .services.history import HistoryStore
from .services.runner import WorkflowBusyError, WorkflowRunner
from .services.workflows import (
    ECM_WORKFLOW_IDS,
    MANUAL_WORKFLOW_IDS,
    REPORT_WORKFLOW_IDS,
    WORKFLOWS,
    build_period_dashboards,
    ecm_workflow_gate_error,
    eligible_next_ecm_workflow,
)
from .services.workflow_builders import PreflightError, prepare_workflow

ROOT = load_settings()
TEMPLATES = Jinja2Templates(directory=str(ROOT.project_root / "control_panel_v3" / "templates"))
TEMPLATES.env.filters["hkt"] = lambda value: datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(ZoneInfo("Asia/Hong_Kong")).strftime("%Y-%m-%d %H:%M:%S HKT")


def _local_return_path(value: str) -> str:
    parsed = urlsplit(value)
    if not value.startswith("/") or value.startswith("//") or "\\" in value or parsed.scheme or parsed.netloc:
        raise HTTPException(status_code=400, detail="Return path must stay inside the V3 control panel.")
    return value


def _date_range(start_date: str, end_date: str) -> tuple[date, date]:
    try:
        start, end = date.fromisoformat(start_date), date.fromisoformat(end_date)
    except ValueError as error:
        raise ValueError("Dates must use YYYY-MM-DD format.") from error
    if end < start:
        raise ValueError("End date cannot be earlier than start date.")
    return start, end


def _default_dates() -> tuple[str, str]:
    today = date.today()
    previous_sunday = today - timedelta(days=today.weekday() + 1)
    return (previous_sunday - timedelta(days=6)).isoformat(), previous_sunday.isoformat()


def create_app() -> FastAPI:
    settings = load_settings()
    errors = settings.validate()
    if errors:
        raise ConfigurationError(" ".join(errors))
    settings.ensure_application_directories()
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    history.fail_interrupted_runs()
    archive = ArchiveStore(settings.archive_database_path, settings.archive_dir / "runs")
    runner = WorkflowRunner(settings, history)

    app = FastAPI(title="Workflow Control Panel V3", docs_url=None, redoc_url=None)
    app.state.settings = settings
    app.state.history = history
    app.state.archive = archive
    app.state.runner = runner
    app.state.preview_only = settings.preview_only
    app.mount("/static", StaticFiles(directory=str(settings.project_root / "control_panel_v3" / "static")), name="static")

    def context(error: str | None = None) -> dict[str, object]:
        summaries = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)
        lookup_previous_counts: dict[int, dict[str, int]] = {}
        for item in summaries:
            previous_run = item.runs.get("ecm-first-date-lookup")
            if previous_run is None:
                continue
            summary_artifact = next((artifact for artifact in history.list_artifacts(previous_run.id) if artifact.name == "First-date lookup summary"), None)
            if summary_artifact and summary_artifact.stored_path.is_file():
                try:
                    payload = json.loads(summary_artifact.stored_path.read_text(encoding="utf-8"))
                    lookup_previous_counts[item.period.id] = {key: int(payload.get(key, 0)) for key in ("hits", "miss_count", "searched")}
                except (OSError, json.JSONDecodeError, TypeError, ValueError):
                    pass
        start, end = _default_dates()
        return {
            "periods": build_period_dashboards(summaries),
            "period_summaries": summaries,
            "lookup_previous_counts": lookup_previous_counts,
            "deleted_period_summaries": history.list_report_period_summaries(REPORT_WORKFLOW_IDS, is_deleted=True),
            "active_run": history.active_run(),
            "configuration_errors": settings.validate(),
            "preview_only": app.state.preview_only,
            "workflows": [WORKFLOWS[key] for key in MANUAL_WORKFLOW_IDS],
            "report_workflows": [WORKFLOWS[key] for key in REPORT_WORKFLOW_IDS],
            "default_start": start,
            "default_end": end,
            "error": error,
        }

    @app.get("/")
    def dashboard(request: Request, error: str | None = None):
        return TEMPLATES.TemplateResponse(request=request, name="dashboard.html", context=context(error))

    @app.get("/health")
    def health():
        return JSONResponse({"status": "ok", "service": "control_panel_v3", "preview_only": app.state.preview_only}, headers={"Cache-Control": "no-store"})

    @app.get("/report-history")
    def report_history(request: Request, error: str | None = None):
        return TEMPLATES.TemplateResponse(request=request, name="report_history.html", context=context(error))

    @app.post("/mode/toggle")
    def toggle_mode(return_to: str = Form("/")):
        target = _local_return_path(return_to)
        app.state.preview_only = not app.state.preview_only
        return RedirectResponse(target, status_code=303)

    @app.post("/report-periods")
    def create_period(start_date: str = Form(...), end_date: str = Form(...)):
        try:
            start, end = _date_range(start_date, end_date)
            history.create_report_period(start.isoformat(), end.isoformat())
        except ValueError as error:
            return RedirectResponse("/report-history?" + urlencode({"error": str(error)}), status_code=303)
        return RedirectResponse("/report-history", status_code=303)

    @app.post("/report-periods/{period_id}")
    def update_period(period_id: int, start_date: str = Form(...), end_date: str = Form(...)):
        try:
            start, end = _date_range(start_date, end_date)
            history.update_report_period(period_id, start.isoformat(), end.isoformat())
        except (KeyError, ValueError) as error:
            return RedirectResponse("/report-history?" + urlencode({"error": str(error)}), status_code=303)
        return RedirectResponse("/report-history", status_code=303)

    @app.post("/report-periods/{period_id}/delete")
    def delete_period(period_id: int):
        try:
            history.delete_report_period(period_id)
        except (KeyError, ValueError) as error:
            return RedirectResponse("/report-history?" + urlencode({"error": str(error)}), status_code=303)
        return RedirectResponse("/report-history", status_code=303)

    @app.post("/report-periods/{period_id}/restore")
    def restore_period(period_id: int):
        try:
            history.restore_report_period(period_id)
        except (KeyError, ValueError) as error:
            return RedirectResponse("/report-history?" + urlencode({"error": str(error)}), status_code=303)
        return RedirectResponse("/report-history", status_code=303)

    def _submit_period_run(period_id: int, workflow_id: str, *, confirm_step_1b_bypass: bool = False, source_run_id: int | None = None, lookup_scope: str = "all"):
        if workflow_id not in REPORT_WORKFLOW_IDS:
            raise HTTPException(status_code=404, detail="Unknown V3 workflow.")
        if app.state.preview_only:
            return RedirectResponse("/report-history?" + urlencode({"error": "V3 is preview-only; no run was created."}), status_code=303)
        try:
            period = history.get_report_period(period_id)
            summary = next((item for item in history.list_report_period_summaries(REPORT_WORKFLOW_IDS) if item.period.id == period_id), None)
            if summary is None or period.is_deleted:
                raise ValueError("This period is unavailable for execution.")
            if source_run_id is not None:
                source_run = history.get_run(source_run_id)
                if source_run.period_id != period_id or source_run.workflow_id != "ecm-step-0" or source_run.state != "succeeded":
                    raise ValueError("The selected Step 0 source run is not a successful run for this period.")
                if summary.runs.get("ecm-step-0") is None or summary.runs["ecm-step-0"].id != source_run_id:
                    raise ValueError("The selected Step 0 source run is no longer the latest run for this period.")
                if eligible_next_ecm_workflow(source_run, [summary]) != "ecm-step-1a":
                    raise ValueError("Run Next Step is no longer available; refresh the period workflow state.")
                if workflow_id != "ecm-step-1a":
                    raise ValueError("A Step 0 source run can hand off only to Step 1a.")
            elif workflow_id == "ecm-step-1a":
                raise ValueError("Start Step 1a from the successful Step 0 run using Run Next Step.")
            if workflow_id == "ecm-first-date-lookup" and lookup_scope not in {"all", "blank", "previously-unmatched"}:
                raise ValueError("Select a supported first-date lookup rerun scope.")
            prior_lookup_summary = None
            if workflow_id == "ecm-first-date-lookup" and lookup_scope == "previously-unmatched":
                previous_run = summary.runs.get("ecm-first-date-lookup")
                if previous_run is None:
                    raise ValueError("There is no prior lookup run for previously-unmatched scope.")
                prior_artifact = next((item for item in history.list_artifacts(previous_run.id) if item.name == "First-date lookup summary"), None)
                if prior_artifact is None or not prior_artifact.stored_path.is_file():
                    raise ValueError("The prior lookup summary artifact is unavailable; choose another scope.")
                prior_lookup_summary = prior_artifact.stored_path
            gate_error = ecm_workflow_gate_error(summary, workflow_id, confirm_step_1b_bypass=confirm_step_1b_bypass)
            if gate_error:
                raise ValueError(gate_error)
            if workflow_id == "ecm-step-1a":
                latest_step0 = summary.runs.get("ecm-step-0")
                if latest_step0 is None:
                    raise ValueError("ECM Step 0 must succeed before ECM Step 1a can start.")
                if latest_step0.state != "succeeded":
                    raise ValueError("The latest ECM Step 0 run must succeed before ECM Step 1a can start.")
            workflow_configuration_errors = settings.validate_for_workflow(workflow_id)
            if workflow_configuration_errors:
                raise ValueError("Resolve workflow configuration errors before starting: " + " ".join(workflow_configuration_errors))
            run = runner.submit(
                workflow_id,
                WORKFLOWS[workflow_id].label,
                date.fromisoformat(period.start_date),
                date.fromisoformat(period.end_date),
                period_id=period_id,
                manual_step_1b_bypass=confirm_step_1b_bypass and workflow_id == "ecm-step-2",
                lookup_scope=lookup_scope if workflow_id == "ecm-first-date-lookup" else None,
                prepared_factory=lambda run_id: prepare_workflow(
                    workflow_id,
                    date.fromisoformat(period.start_date),
                    date.fromisoformat(period.end_date),
                    run_id,
                    settings,
                    lookup_scope=lookup_scope,
                    previous_lookup_summary=prior_lookup_summary,
                ),
            )
            if workflow_id == "ecm-step-1a":
                return RedirectResponse(f"/runs/{run.id}", status_code=303)
            return RedirectResponse("/report-history", status_code=303)
        except (KeyError, ValueError, PreflightError, WorkflowBusyError) as error:
            return RedirectResponse("/report-history?" + urlencode({"error": str(error)}), status_code=303)

    @app.post("/report-periods/{period_id}/runs/next")
    def run_next_step(period_id: int, source_run_id: int = Form(...)):
        return _submit_period_run(period_id, "ecm-step-1a", source_run_id=source_run_id)

    @app.post("/report-periods/{period_id}/runs/{workflow_id}")
    def start_period_run(period_id: int, workflow_id: str, confirm_step_1b_bypass: bool = Form(False), lookup_scope: str = Form("all")):
        return _submit_period_run(period_id, workflow_id, confirm_step_1b_bypass=confirm_step_1b_bypass, lookup_scope=lookup_scope)

    @app.get("/runs/{run_id}")
    def run_detail(request: Request, run_id: int):
        try:
            run = history.get_run(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail="Run not found.") from error
        summaries = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)
        artifacts = history.list_artifacts(run_id)
        lookup_summary = None
        if run.workflow_id == "ecm-first-date-lookup":
            summary_artifact = next((item for item in artifacts if item.name == "First-date lookup summary"), None)
            if summary_artifact and summary_artifact.stored_path.is_file():
                try:
                    lookup_summary = json.loads(summary_artifact.stored_path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError):
                    lookup_summary = None
        period_summary = next((item for item in summaries if item.period.id == run.period_id), None)
        prior_lookup_summary = None
        if run.workflow_id == "ecm-first-date-lookup" and period_summary:
            prior_lookup_run = next((item for item in history.list_runs(workflow_id="ecm-first-date-lookup", period_id=run.period_id) if item.id < run.id), None)
            if prior_lookup_run:
                artifact = next((item for item in history.list_artifacts(prior_lookup_run.id) if item.name == "First-date lookup summary"), None)
                if artifact and artifact.stored_path.is_file():
                    try:
                        prior_lookup_summary = json.loads(artifact.stored_path.read_text(encoding="utf-8"))
                    except (OSError, json.JSONDecodeError):
                        prior_lookup_summary = None
        return TEMPLATES.TemplateResponse(request=request, name="job_detail.html", context={
            "run": run, "workflow": WORKFLOWS.get(run.workflow_id), "artifacts": artifacts,
            "log": history.read_log(run_id), "next_workflow_id": eligible_next_ecm_workflow(run, [item for item in summaries if item.period.id == run.period_id]),
            "lookup_summary": lookup_summary,
            "prior_lookup_summary": prior_lookup_summary,
            "preview_only": app.state.preview_only, "active_run": history.active_run(),
            "next_step_disabled": bool(app.state.preview_only or history.active_run() or eligible_next_ecm_workflow(run, [item for item in summaries if item.period.id == run.period_id]) is None),
        })

    @app.get("/runs/{run_id}/status")
    def run_status(run_id: int):
        try:
            run = history.get_run(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail="Run not found.") from error
        if run.workflow_id not in REPORT_WORKFLOW_IDS:
            raise HTTPException(status_code=404, detail="Run not found.")
        artifacts = history.list_artifacts(run_id)
        return JSONResponse({
            "id": run.id,
            "state": run.state,
            "started_at": run.started_at,
            "finished_at": run.finished_at,
            "exit_code": run.exit_code,
            "error_summary": run.error_summary,
            "log": history.read_log(run_id),
            "artifacts": [{"id": item.id, "name": item.name, "filename": item.stored_path.name, "is_csv": item.stored_path.suffix.casefold() == ".csv"} for item in artifacts],
        }, headers={"Cache-Control": "no-store"})

    @app.get("/runs/{run_id}/artifacts/{artifact_id}")
    def download_artifact(run_id: int, artifact_id: int):
        try:
            artifact = history.get_artifact(run_id, artifact_id)
            stored = artifact.stored_path.resolve()
            stored.relative_to(settings.runs_dir.resolve())
        except (KeyError, ValueError) as error:
            raise HTTPException(status_code=404, detail="Artifact not found.") from error
        if not stored.is_file():
            raise HTTPException(status_code=410, detail="Artifact is no longer available.")
        return FileResponse(stored, filename=stored.name)

    @app.get("/runs/{run_id}/artifacts/{artifact_id}/view")
    def view_artifact(run_id: int, artifact_id: int):
        try:
            artifact = history.get_artifact(run_id, artifact_id)
            stored = artifact.stored_path.resolve()
            stored.relative_to(settings.runs_dir.resolve())
        except (KeyError, ValueError) as error:
            raise HTTPException(status_code=404, detail="Artifact not found.") from error
        if stored.suffix.casefold() != ".csv":
            raise HTTPException(status_code=404, detail="Only CSV artifacts can be viewed inline.")
        if not stored.is_file():
            raise HTTPException(status_code=410, detail="Artifact is no longer available.")
        return FileResponse(stored, media_type="text/plain; charset=utf-8", filename=stored.name, content_disposition_type="inline", headers={"Cache-Control": "no-store"})

    @app.get("/archive")
    def archived_history(request: Request, error: str | None = None):
        available = archive.is_imported()
        return TEMPLATES.TemplateResponse(request=request, name="archive_history.html", context={
            "preview_only": app.state.preview_only, "runs": archive.list_runs() if available else [],
            "periods": archive.list_periods() if available else [], "archive_available": available, "error": error,
        })

    @app.get("/archive/runs/{run_id}")
    def archived_run_detail(request: Request, run_id: int):
        try:
            run = archive.get_run(run_id)
            artifacts = archive.list_artifacts(run_id)
            log = archive.read_log(run_id)
        except (KeyError, FileNotFoundError, ValueError) as error:
            raise HTTPException(status_code=404, detail="Archived run not found.") from error
        return TEMPLATES.TemplateResponse(request=request, name="archive_job_detail.html", context={
            "preview_only": app.state.preview_only, "run": run, "artifacts": artifacts, "log": log,
        })

    @app.get("/archive/runs/{run_id}/artifacts/{artifact_id}")
    def archived_artifact(run_id: int, artifact_id: int):
        try:
            artifact = archive.get_artifact(run_id, artifact_id)
            stored = artifact.stored_path.resolve()
            stored.relative_to((settings.archive_dir / "runs").resolve())
        except (KeyError, ValueError, FileNotFoundError) as error:
            raise HTTPException(status_code=404, detail="Archived artifact not found.") from error
        if not stored.is_file():
            raise HTTPException(status_code=410, detail="Archived artifact is no longer available.")
        return FileResponse(stored, filename=stored.name)

    @app.post("/archive/import")
    def import_archive():
        try:
            import_v2_archive(settings.project_root / "control_panel_v2" / "data", settings.archive_dir)
        except (OSError, ValueError) as error:
            return RedirectResponse("/archive?" + urlencode({"error": str(error)}), status_code=303)
        return RedirectResponse("/archive", status_code=303)

    @app.get("/history")
    def run_history(request: Request):
        return TEMPLATES.TemplateResponse(request=request, name="history.html", context={"preview_only": app.state.preview_only, "runs": history.list_runs(), "workflows": [WORKFLOWS[key] for key in REPORT_WORKFLOW_IDS]})

    @app.get("/manual-run")
    def manual_run(request: Request):
        return TEMPLATES.TemplateResponse(request=request, name="manual_run.html", context={**context(), "workflows": [WORKFLOWS[key] for key in MANUAL_WORKFLOW_IDS]})

    return app


class LazyApplication:
    """Create the service only when an ASGI server begins serving requests."""

    def __init__(self) -> None:
        self._application: FastAPI | None = None

    async def __call__(self, scope, receive, send) -> None:
        if self._application is None:
            self._application = create_app()
        await self._application(scope, receive, send)


app = LazyApplication()
