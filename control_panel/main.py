"""FastAPI entry point for the localhost-only workflow control panel."""

from __future__ import annotations

import re
from datetime import date, datetime, timedelta
from urllib.parse import urlencode
from zoneinfo import ZoneInfo

from fastapi import FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, JSONResponse, RedirectResponse
from openpyxl import load_workbook
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from control_panel.services.config import CONTROL_PANEL_ROOT, ConfigurationError, load_settings
from control_panel.services.history import HistoryStore
from control_panel.services.periods import HONG_KONG, latest_completed_report_period
from control_panel.services.runner import WorkflowBusyError, WorkflowRunner
from control_panel.services.workflows import (
    PreflightError,
    V1_WORKFLOW_IDS,
    STANDALONE_WORKFLOW_IDS,
    WORKFLOWS,
    get_workflow,
)

TEMPLATE_DIRECTORY = CONTROL_PANEL_ROOT / "templates"
STATIC_DIRECTORY = CONTROL_PANEL_ROOT / "static"
templates = Jinja2Templates(directory=str(TEMPLATE_DIRECTORY))
REPORT_HISTORY_WORKFLOW_IDS = V1_WORKFLOW_IDS


def _cell_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        try:
            return date.fromisoformat(value.strip()[:10])
        except ValueError:
            return None
    return None


def _source_period(sheet, name: str, reference_year: int) -> tuple[date, date] | None:
    token = name.removesuffix(" (OK)")
    weekly = re.search(r"ECM_(\d{1,2})-(\d{1,2})\s+([A-Za-z]{3})$", token)
    if weekly:
        start_day, end_day = int(weekly.group(1)), int(weekly.group(2))
        month_number = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
                        "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}.get(weekly.group(3).title())
        if month_number:
            start = date(reference_year, month_number, start_day)
            return start, date(reference_year, month_number, end_day)
    month = re.search(r"(20\d{2})-(0[1-9]|1[0-2])$", token)
    if month:
        year, month_number = int(month.group(1)), int(month.group(2))
        start = date(year, month_number, 1)
        next_month = date(year + (month_number == 12), month_number % 12 + 1, 1)
        return start, next_month - timedelta(days=1)
    quarter = re.search(r"(20\d{2})Q([1-4])$", token)
    if quarter:
        year, quarter_number = int(quarter.group(1)), int(quarter.group(2))
        month_number = (quarter_number - 1) * 3 + 1
        start = date(year, month_number, 1)
        next_month = date(year + 1, 1, 1) if month_number == 10 else date(year, month_number + 3, 1)
        return start, next_month - timedelta(days=1)
    dates = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if len(row) > 2:
            parsed = _cell_date(row[2])
            if parsed:
                dates.append(parsed)
    return (min(dates), max(dates)) if dates else None


def _preview_ecm_sources(workbook, start: date, end: date) -> list[dict[str, object]]:
    result = []
    for name in workbook.sheetnames:
        if not name.endswith("(OK)"):
            continue
        period = _source_period(workbook[name], name, start.year)
        result.append({
            "name": name,
            "period": f"{period[0]} to {period[1]}" if period else None,
            "selected": bool(period and start <= period[0] and period[1] <= end),
        })
    return result


def _format_hong_kong_timestamp(value: str) -> str:
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZoneInfo("UTC"))
    return parsed.astimezone(HONG_KONG).strftime("%Y-%m-%d %H:%M:%S HKT")


templates.env.filters["hkt"] = _format_hong_kong_timestamp
RUN_PAGE_LABELS = {
    "/": "Dashboard",
    "/manual-run": "Manual Run",
    "/ecm-period-report": "ECM Period Report",
}


def _run_page_path(value: str | None) -> str:
    """Return a safe local page path for workflow-run navigation."""
    return value if value in RUN_PAGE_LABELS else "/"


def _default_period() -> tuple[str, str]:
    today = date.today()
    previous_sunday = today - timedelta(days=today.weekday() + 1)
    return (previous_sunday - timedelta(days=6)).isoformat(), previous_sunday.isoformat()


def _parse_date(value: str, label: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as error:
        raise PreflightError(f"{label} must use YYYY-MM-DD.") from error


def _validate_period_dates(start_date: str, end_date: str) -> tuple[date, date]:
    start = _parse_date(start_date, "Start date")
    end = _parse_date(end_date, "End date")
    if end < start:
        raise PreflightError("End date cannot be earlier than start date.")
    return start, end


def _report_history_url(error: str | None = None) -> str:
    return "/report-history" if not error else "/report-history?" + urlencode({"error": error})


def create_app() -> FastAPI:
    settings = load_settings()
    history = HistoryStore(settings.database_path, settings.runs_dir)
    settings.ensure_application_directories()
    history.initialize()
    history.fail_interrupted_runs()
    runner = WorkflowRunner(settings, history)

    app = FastAPI(title="Local Workflow Control Panel", docs_url=None, redoc_url=None)
    app.state.settings = settings
    app.state.history = history
    app.state.runner = runner
    app.mount("/static", StaticFiles(directory=str(STATIC_DIRECTORY)), name="static")

    @app.get("/")
    def dashboard(request: Request, error: str | None = None):
        default_start, default_end = _default_period()
        return templates.TemplateResponse(
            request=request,
            name="dashboard.html",
            context={
                "workflows": [WORKFLOWS[workflow_id] for workflow_id in V1_WORKFLOW_IDS],
                "active_run": runner.active_run(),
                "configuration_errors": settings.validate(),
                "error": error,
                "default_start": default_start,
                "default_end": default_end,
                "settings": settings,
            },
        )

    @app.get("/manual-run")
    def manual_run_page(request: Request, error: str | None = None):
        default_start, default_end = _default_period()
        return templates.TemplateResponse(
            request=request,
            name="manual_run.html",
            context={
                "workflows": [WORKFLOWS[workflow_id] for workflow_id in V1_WORKFLOW_IDS],
                "active_run": runner.active_run(),
                "configuration_errors": settings.validate(),
                "error": error,
                "default_start": default_start,
                "default_end": default_end,
                "settings": settings,
            },
        )

    @app.get("/ecm-period-report")
    def ecm_period_report_page(
        request: Request, error: str | None = None,
        start_date: str | None = None, end_date: str | None = None,
    ):
        source_sheets: list[str] = []
        preview_error = None
        if start_date and end_date:
            try:
                _validate_period_dates(start_date, end_date)
                workbook = load_workbook(settings.ecm_record, read_only=True, data_only=True)
                try:
                    source_sheets = _preview_ecm_sources(
                        workbook, date.fromisoformat(start_date), date.fromisoformat(end_date)
                    )
                finally:
                    workbook.close()
            except Exception as exc:
                preview_error = str(exc)
        return templates.TemplateResponse(
            request=request,
            name="ecm_period_report.html",
            context={
                "active_run": runner.active_run(),
                "configuration_errors": settings.validate(),
                "error": error,
                "settings": settings,
                "start_date": start_date or "",
                "end_date": end_date or "",
                "source_sheets": source_sheets,
                "preview_error": preview_error,
            },
        )

    @app.post("/ecm-period-report/run")
    def start_ecm_period_report(
        start_date: str = Form(...), end_date: str = Form(...),
        source_sheet: list[str] = Form([]),
    ) -> RedirectResponse:
        try:
            setup_problems = settings.validate()
            if setup_problems:
                raise PreflightError("Resolve the local configuration issues before starting a workflow.")
            start, end = _validate_period_dates(start_date, end_date)
            if not source_sheet:
                raise PreflightError("Select at least one approved (OK) ECM source sheet.")
            extra_args = tuple(arg for name in source_sheet for arg in ("--source-sheet", name))
            run = runner.submit("ecm-period-report", start, end, extra_args=extra_args)
        except (PreflightError, WorkflowBusyError, ConfigurationError) as error:
            return RedirectResponse(
                url="/ecm-period-report?" + urlencode({"error": str(error)}), status_code=303
            )
        return RedirectResponse(
            url=f"/runs/{run.id}?" + urlencode({"origin": "/ecm-period-report"}), status_code=303
        )

    @app.post("/runs/{workflow_id}")
    def start_run(
        workflow_id: str,
        start_date: str = Form(...),
        end_date: str = Form(...),
        origin: str = Form("/"),
    ) -> RedirectResponse:
        origin_path = _run_page_path(origin)
        try:
            if workflow_id not in V1_WORKFLOW_IDS:
                raise PreflightError(f"Workflow is not available in V1: {workflow_id}")
            get_workflow(workflow_id)
            setup_problems = settings.validate()
            if setup_problems:
                raise PreflightError("Resolve the local configuration issues before starting a workflow.")
            start = _parse_date(start_date, "Start date")
            end = _parse_date(end_date, "End date")
            run = runner.submit(workflow_id, start, end)
        except (PreflightError, WorkflowBusyError, ConfigurationError) as error:
            return RedirectResponse(
                url=origin_path + "?" + urlencode({"error": str(error)}), status_code=303
            )
        return RedirectResponse(
            url=f"/runs/{run.id}?" + urlencode({"origin": origin_path}), status_code=303
        )

    @app.get("/report-history")
    def report_history_page(request: Request, error: str | None = None):
        default_start, default_end = latest_completed_report_period()
        report_workflows = [WORKFLOWS[workflow_id] for workflow_id in REPORT_HISTORY_WORKFLOW_IDS]
        return templates.TemplateResponse(
            request=request,
            name="report_history.html",
            context={
                "period_summaries": history.list_report_period_summaries(
                    REPORT_HISTORY_WORKFLOW_IDS
                ),
                "deleted_period_summaries": history.list_report_period_summaries(
                    REPORT_HISTORY_WORKFLOW_IDS, is_deleted=True
                ),
                "report_workflows": report_workflows,
                "active_run": runner.active_run(),
                "configuration_errors": settings.validate(),
                "default_start": default_start.isoformat(),
                "default_end": default_end.isoformat(),
                "error": error,
            },
        )

    @app.post("/report-periods")
    def create_report_period(
        start_date: str = Form(...), end_date: str = Form(...)
    ) -> RedirectResponse:
        try:
            start, end = _validate_period_dates(start_date, end_date)
            history.create_report_period(start.isoformat(), end.isoformat())
        except (PreflightError, ValueError) as error:
            return RedirectResponse(url=_report_history_url(str(error)), status_code=303)
        return RedirectResponse(url="/report-history", status_code=303)

    @app.post("/report-periods/{period_id}")
    def update_report_period(
        period_id: int,
        start_date: str = Form(...),
        end_date: str = Form(...),
    ) -> RedirectResponse:
        try:
            start, end = _validate_period_dates(start_date, end_date)
            history.update_report_period(period_id, start.isoformat(), end.isoformat())
        except (KeyError, PreflightError, ValueError) as error:
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

    @app.post("/report-periods/{period_id}/runs/{workflow_id}")
    def start_period_run(period_id: int, workflow_id: str) -> RedirectResponse:
        try:
            if workflow_id not in REPORT_HISTORY_WORKFLOW_IDS:
                raise PreflightError(f"Unknown report workflow: {workflow_id}")
            get_workflow(workflow_id)
            setup_problems = settings.validate()
            if setup_problems:
                raise PreflightError("Resolve the local configuration issues before starting a workflow.")
            period = history.get_report_period(period_id)
            if period.is_deleted:
                raise PreflightError("This report period has been deleted and cannot start a workflow.")
            run = runner.submit(
                workflow_id,
                _parse_date(period.start_date, "Period start date"),
                _parse_date(period.end_date, "Period end date"),
                period_id=period.id,
            )
        except (KeyError, PreflightError, WorkflowBusyError, ConfigurationError, ValueError) as error:
            return RedirectResponse(url=_report_history_url(str(error)), status_code=303)
        return RedirectResponse(url=f"/runs/{run.id}", status_code=303)

    @app.get("/runs/{run_id}")
    def run_detail(request: Request, run_id: int, origin: str | None = None):
        try:
            run = history.get_run(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail="Run was not found.") from error
        if run.workflow_id not in (*V1_WORKFLOW_IDS, *STANDALONE_WORKFLOW_IDS):
            raise HTTPException(status_code=404, detail="Run was not found.")
        return templates.TemplateResponse(
            request=request,
            name="job_detail.html",
            context={
                "run": run,
                "workflow": get_workflow(run.workflow_id),
                "artifacts": history.list_artifacts(run_id),
                "log": history.read_log(run_id),
                "run_origin": _run_page_path(origin),
                "run_origin_label": RUN_PAGE_LABELS[_run_page_path(origin)],
            },
        )

    @app.get("/runs/{run_id}/status")
    def run_status(run_id: int) -> JSONResponse:
        try:
            run = history.get_run(run_id)
        except KeyError as error:
            raise HTTPException(status_code=404, detail="Run was not found.") from error
        if run.workflow_id not in (*V1_WORKFLOW_IDS, *STANDALONE_WORKFLOW_IDS):
            raise HTTPException(status_code=404, detail="Run was not found.")
        artifacts = history.list_artifacts(run_id)
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
                        "size_bytes": artifact.size_bytes,
                        "download_url": f"/runs/{run_id}/artifacts/{artifact.id}",
                    }
                    for artifact in artifacts
                ],
            }
        )

    @app.get("/runs/{run_id}/artifacts/{artifact_id}")
    def download_artifact(run_id: int, artifact_id: int) -> FileResponse:
        try:
            run = history.get_run(run_id)
            if run.workflow_id not in (*V1_WORKFLOW_IDS, *STANDALONE_WORKFLOW_IDS):
                raise KeyError(run_id)
            artifact = history.get_artifact(run_id, artifact_id)
            stored_path = artifact.stored_path.resolve()
            stored_path.relative_to(settings.runs_dir.resolve())
        except (KeyError, ValueError) as error:
            raise HTTPException(status_code=404, detail="Artifact was not found.") from error
        if not stored_path.is_file():
            raise HTTPException(status_code=410, detail="The saved artifact is no longer available.")
        return FileResponse(path=stored_path, filename=stored_path.name)

    @app.get("/history")
    def history_page(request: Request, workflow_id: str | None = None, state: str | None = None):
        if workflow_id and workflow_id not in V1_WORKFLOW_IDS:
            raise HTTPException(status_code=404, detail="Workflow was not found.")
        if state and state not in {"queued", "running", "succeeded", "failed"}:
            raise HTTPException(status_code=400, detail="Invalid run state.")
        return templates.TemplateResponse(
            request=request,
            name="history.html",
            context={
                "runs": [
                    run
                    for run in history.list_runs(workflow_id=workflow_id, state=state)
                    if run.workflow_id in V1_WORKFLOW_IDS
                ],
                "workflows": [WORKFLOWS[workflow_id] for workflow_id in V1_WORKFLOW_IDS],
                "selected_workflow": workflow_id or "",
                "selected_state": state or "",
            },
        )

    return app


app = create_app()
