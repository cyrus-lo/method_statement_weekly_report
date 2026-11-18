"""FastAPI entry point for the isolated, preview-first Workflow Control Panel V2."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from urllib.parse import urlencode, urlsplit
from zoneinfo import ZoneInfo

from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from control_panel.services.history import HistoryStore
from control_panel.services.runner import WorkflowBusyError, WorkflowRunner
from control_panel.services.workflows import WORKFLOWS, get_workflow

from .services.config import CONTROL_PANEL_ROOT, load_settings
from .services.dashboard import (
    build_period_dashboards,
    ecm_workflow_gate_error,
    eligible_next_ecm_workflow,
)


TEMPLATE_DIRECTORY = CONTROL_PANEL_ROOT / "templates"
STATIC_DIRECTORY = CONTROL_PANEL_ROOT / "static"
templates = Jinja2Templates(directory=str(TEMPLATE_DIRECTORY))
REPORT_WORKFLOW_IDS = (
    "weekly-report",
    "ecm-step-0",
    "ecm-step-1a",
    "ecm-step-1b",
    "ecm-step-2",
)
MANUAL_WORKFLOW_IDS = ("weekly-report", "ecm-step-1a", "ecm-step-1b", "ecm-step-2")


def _format_hong_kong_timestamp(value: str) -> str:
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZoneInfo("UTC"))
    return parsed.astimezone(ZoneInfo("Asia/Hong_Kong")).strftime("%Y-%m-%d %H:%M:%S HKT")


templates.env.filters["hkt"] = _format_hong_kong_timestamp


def _preview_error() -> str:
    return "V2 is in preview-only mode. Workflows continue to run only in V1 until cutover."


def _report_history_url(error: str | None = None) -> str:
    return "/report-history" if not error else "/report-history?" + urlencode({"error": error})


def _safe_return_path(value: str) -> str:
    parsed = urlsplit(value)
    if (
        not value.startswith("/")
        or value.startswith("//")
        or "\\" in value
        or parsed.scheme
        or parsed.netloc
    ):
        raise HTTPException(status_code=400, detail="Return path must stay inside the control panel.")
    return value


def _parse_date(value: str, label: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as error:
        raise ValueError(f"{label} must use YYYY-MM-DD.") from error


def _validate_period_dates(start_date: str, end_date: str) -> tuple[date, date]:
    start = _parse_date(start_date, "Start date")
    end = _parse_date(end_date, "End date")
    if end < start:
        raise ValueError("End date cannot be earlier than start date.")
    return start, end


def _default_period() -> tuple[str, str]:
    today = date.today()
    previous_sunday = today - timedelta(days=today.weekday() + 1)
    return (previous_sunday - timedelta(days=6)).isoformat(), previous_sunday.isoformat()


def create_app() -> FastAPI:
    settings = load_settings()
    settings.ensure_application_directories()
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    history.fail_interrupted_runs()
    runner = WorkflowRunner(settings, history)

    app = FastAPI(title="Workflow Control Panel V2 Preview", docs_url=None, redoc_url=None)
    app.state.settings = settings
    app.state.history = history
    app.state.runner = runner
    app.state.preview_only = settings.preview_only
    app.mount("/static", StaticFiles(directory=str(STATIC_DIRECTORY)), name="static")

    def page_context() -> dict[str, object]:
        summaries = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)
        dashboards = build_period_dashboards(summaries)
        unassigned_runs = [run for run in history.list_runs(limit=100) if run.period_id is None]
        return {
            "periods": dashboards,
            "active_run": runner.active_run(),
            "configuration_errors": settings.validate(),
            "preview_only": app.state.preview_only,
            "unassigned_run_count": len(unassigned_runs),
            "report_workflows": [WORKFLOWS[workflow_id] for workflow_id in REPORT_WORKFLOW_IDS],
        }

    @app.get("/")
    def dashboard(request: Request):
        return templates.TemplateResponse(request=request, name="dashboard.html", context=page_context())

    @app.get("/report-history")
    def report_history_page(request: Request, error: str | None = None):
        context = page_context()
        default_start, default_end = _default_period()
        context.update(
            {
                "error": error,
                "default_start": default_start,
                "default_end": default_end,
                "period_summaries": history.list_report_period_summaries(REPORT_WORKFLOW_IDS),
                "deleted_period_summaries": history.list_report_period_summaries(
                    REPORT_WORKFLOW_IDS, is_deleted=True
                ),
            }
        )
        return templates.TemplateResponse(request=request, name="report_history.html", context=context)

    @app.post("/mode/toggle")
    def toggle_mode(return_to: str = Form("/")) -> RedirectResponse:
        safe_return_path = _safe_return_path(return_to)
        app.state.preview_only = not app.state.preview_only
        return RedirectResponse(url=safe_return_path, status_code=303)

    @app.post("/report-periods")
    def create_report_period(start_date: str = Form(...), end_date: str = Form(...)) -> RedirectResponse:
        try:
            start, end = _validate_period_dates(start_date, end_date)
            history.create_report_period(start.isoformat(), end.isoformat())
        except ValueError as error:
            return RedirectResponse(url=_report_history_url(str(error)), status_code=303)
        return RedirectResponse(url="/report-history", status_code=303)

    @app.post("/report-periods/{period_id}")
    def update_report_period(
        period_id: int, start_date: str = Form(...), end_date: str = Form(...)
    ) -> RedirectResponse:
        try:
            start, end = _validate_period_dates(start_date, end_date)
            history.update_report_period(period_id, start.isoformat(), end.isoformat())
        except (KeyError, ValueError) as error:
            return RedirectResponse(url=_report_history_url(str(error)), status_code=303)
        return RedirectResponse(url="/report-history", status_code=303)

    @app.post("/report-periods/{period_id}/delete")
    def delete_report_period(period_id: int) -> RedirectResponse:
        try:
            history.delete_report_period(period_id)
        except (KeyError, ValueError) as error:
            return RedirectResponse(url=_report_history_url(str(error)), status_code=303)
        return RedirectResponse(url="/report-history", status_code=303)

    @app.post("/report-periods/{period_id}/restore")
    def restore_report_period(period_id: int) -> RedirectResponse:
        try:
            history.restore_report_period(period_id)
        except (KeyError, ValueError) as error:
            return RedirectResponse(url=_report_history_url(str(error)), status_code=303)
        return RedirectResponse(url="/report-history", status_code=303)

    @app.get("/manual-run")
    def manual_run_page(request: Request):
        return templates.TemplateResponse(
            request=request,
            name="manual_run.html",
            context={**page_context(), "workflows": [WORKFLOWS[workflow_id] for workflow_id in MANUAL_WORKFLOW_IDS]},
        )

    @app.post("/runs/{workflow_id}")
    def block_manual_execution(workflow_id: str) -> RedirectResponse:
        if workflow_id not in MANUAL_WORKFLOW_IDS:
            raise HTTPException(status_code=404, detail="Workflow is not available from Manual Run.")
        if app.state.preview_only:
            return RedirectResponse(url="/?error=" + _preview_error().replace(" ", "+"), status_code=303)
        raise HTTPException(status_code=501, detail="V2 Manual Run execution is not enabled.")

    @app.post("/report-periods/{period_id}/runs/{workflow_id}")
    def start_period_run(
        period_id: int,
        workflow_id: str,
        confirm_step_1b_bypass: bool = Form(False),
        source_run_id: int | None = Form(None),
    ) -> RedirectResponse:
        if workflow_id not in REPORT_WORKFLOW_IDS:
            raise HTTPException(status_code=404, detail="Workflow is not available from Report History.")
        if app.state.preview_only:
            return RedirectResponse(
                url=_report_history_url(_preview_error()), status_code=303
            )
        try:
            setup_problems = settings.validate()
            if workflow_id != "ecm-step-0":
                setup_problems = [
                    problem
                    for problem in setup_problems
                    if not problem.startswith("Step 0 ")
                    and not problem.startswith("ECM BASE_URL")
                    and not problem.startswith("STEP0_")
                ]
            if setup_problems:
                raise ValueError(
                    "Resolve the local configuration issues before starting a workflow."
                )
            period = history.get_report_period(period_id)
            if period.is_deleted:
                raise ValueError("This report period has been deleted and cannot start a workflow.")
            summaries = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)
            summary = next(
                (item for item in summaries if item.period.id == period.id), None
            )
            if summary is None:
                raise ValueError(f"Report period {period.id} does not exist.")
            if source_run_id is not None:
                source_run = history.get_run(source_run_id)
                if (
                    source_run.period_id != period.id
                    or eligible_next_ecm_workflow(source_run, summaries) != workflow_id
                ):
                    raise ValueError(
                        "This run is no longer eligible to start its next ECM step."
                    )
            gate_error = ecm_workflow_gate_error(
                summaries,
                summary,
                workflow_id,
                confirm_step_1b_bypass=confirm_step_1b_bypass,
            )
            if gate_error:
                raise ValueError(gate_error)
            run = runner.submit(
                workflow_id,
                _parse_date(period.start_date, "Period start date"),
                _parse_date(period.end_date, "Period end date"),
                period_id=period.id,
            )
        except (KeyError, ValueError, WorkflowBusyError) as error:
            return RedirectResponse(url=_report_history_url(str(error)), status_code=303)
        return RedirectResponse(url=f"/runs/{run.id}", status_code=303)

    @app.get("/runs/{run_id}")
    def run_detail(request: Request, run_id: int):
        try:
            run = history.get_run(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail="Run was not found.") from error
        if run.workflow_id not in REPORT_WORKFLOW_IDS:
            raise HTTPException(status_code=404, detail="Run was not found.")
        summaries = history.list_report_period_summaries(REPORT_WORKFLOW_IDS)
        next_workflow_id = eligible_next_ecm_workflow(run, summaries)
        active_run = runner.active_run()
        next_step_configuration_errors = settings.validate()
        if next_workflow_id == "ecm-step-1a":
            next_step_configuration_errors = [
                problem
                for problem in next_step_configuration_errors
                if not problem.startswith("Step 0 ")
                and not problem.startswith("ECM BASE_URL")
                and not problem.startswith("STEP0_")
            ]
        return templates.TemplateResponse(
            request=request,
            name="job_detail.html",
            context={
                "preview_only": app.state.preview_only,
                "active_run": active_run,
                "next_workflow_id": next_workflow_id,
                "next_step_disabled": (
                    app.state.preview_only
                    or bool(next_step_configuration_errors)
                    or active_run is not None
                ),
                "run": run,
                "workflow": get_workflow(run.workflow_id),
                "artifacts": history.list_artifacts(run_id),
                "log": history.read_log(run_id),
            },
        )

    @app.get("/runs/{run_id}/status")
    def run_status(run_id: int) -> JSONResponse:
        try:
            run = history.get_run(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail="Run was not found.") from error
        if run.workflow_id not in REPORT_WORKFLOW_IDS:
            raise HTTPException(status_code=404, detail="Run was not found.")
        return JSONResponse(
            {
                "id": run.id,
                "state": run.state,
                "started_at": run.started_at,
                "finished_at": run.finished_at,
                "exit_code": run.exit_code,
                "error_summary": run.error_summary,
                "log": history.read_log(run_id),
                "artifacts": [
                    {
                        "id": artifact.id,
                        "name": artifact.name,
                        "file_name": artifact.stored_path.name,
                        "download_url": f"/runs/{run_id}/artifacts/{artifact.id}",
                        "view_url": (
                            f"/runs/{run_id}/artifacts/{artifact.id}/view"
                            if artifact.stored_path.suffix.casefold() == ".csv"
                            else None
                        ),
                    }
                    for artifact in history.list_artifacts(run_id)
                ],
            }
        )

    def get_artifact_path(run_id: int, artifact_id: int):
        try:
            run = history.get_run(run_id)
            if run.workflow_id not in REPORT_WORKFLOW_IDS:
                raise KeyError(run_id)
            artifact = history.get_artifact(run_id, artifact_id)
            stored_path = artifact.stored_path.resolve()
            stored_path.relative_to(settings.runs_dir.resolve())
        except (KeyError, ValueError) as error:
            raise HTTPException(status_code=404, detail="Artifact was not found.") from error
        if not stored_path.is_file():
            raise HTTPException(status_code=410, detail="The saved artifact is no longer available.")
        return artifact, stored_path

    @app.get("/runs/{run_id}/artifacts/{artifact_id}")
    def download_artifact(run_id: int, artifact_id: int) -> FileResponse:
        artifact, stored_path = get_artifact_path(run_id, artifact_id)
        return FileResponse(path=stored_path, filename=stored_path.name)

    @app.get("/runs/{run_id}/artifacts/{artifact_id}/view")
    def view_artifact(run_id: int, artifact_id: int) -> FileResponse:
        artifact, stored_path = get_artifact_path(run_id, artifact_id)
        if stored_path.suffix.casefold() != ".csv":
            raise HTTPException(status_code=404, detail="Only CSV artifacts can be viewed inline.")
        return FileResponse(
            path=stored_path,
            media_type="text/plain; charset=utf-8",
            filename=stored_path.name,
            content_disposition_type="inline",
            headers={"Cache-Control": "no-store"},
        )

    @app.get("/history")
    def history_page(request: Request):
        return templates.TemplateResponse(
            request=request,
            name="history.html",
            context={
                "preview_only": app.state.preview_only,
                "runs": [
                    run
                    for run in history.list_runs()
                    if run.workflow_id in REPORT_WORKFLOW_IDS
                ],
                "workflows": [WORKFLOWS[workflow_id] for workflow_id in REPORT_WORKFLOW_IDS],
            },
        )

    return app


app = create_app()
