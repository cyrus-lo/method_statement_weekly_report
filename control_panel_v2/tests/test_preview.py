from __future__ import annotations

from datetime import datetime

from fastapi.testclient import TestClient

import control_panel_v2.main as v2_main
from control_panel.services.history import ReportPeriodRecord, ReportPeriodSummary, RunRecord
from control_panel_v2.services.dashboard import build_period_dashboards
from control_panel_v2.services.config import Settings


def test_preview_rejects_direct_workflow_post_without_creating_a_run(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    app = v2_main.create_app()
    client = TestClient(app)
    before = len(app.state.history.list_runs(limit=100))

    response = client.post("/report-periods/1/runs/ecm-step-1a", follow_redirects=False)

    assert response.status_code == 303
    assert "preview-only" in response.headers["location"]
    assert len(app.state.history.list_runs(limit=100)) == before
    assert "PREVIEW" in client.get("/").text


def test_mode_toggle_changes_label_and_redirects_to_originating_page(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    app = v2_main.create_app()
    client = TestClient(app)

    preview_page = client.get("/report-history")
    assert "PREVIEW · READ ONLY" in preview_page.text
    assert "LIVE · EXECUTION ENABLED" not in preview_page.text

    enabled = client.post(
        "/mode/toggle",
        data={"return_to": "/report-history"},
        follow_redirects=False,
    )

    assert enabled.status_code == 303
    assert enabled.headers["location"] == "/report-history"
    assert app.state.preview_only is False
    live_page = client.get("/report-history")
    assert "LIVE · EXECUTION ENABLED" in live_page.text
    assert "PREVIEW · READ ONLY" not in live_page.text

    disabled = client.post(
        "/mode/toggle",
        data={"return_to": "/history"},
        follow_redirects=False,
    )

    assert disabled.status_code == 303
    assert disabled.headers["location"] == "/history"
    assert app.state.preview_only is True
    assert "PREVIEW · READ ONLY" in client.get("/history").text


def test_mode_toggle_rejects_external_return_path(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    app = v2_main.create_app()
    client = TestClient(app)

    response = client.post(
        "/mode/toggle",
        data={"return_to": "https://example.com"},
        follow_redirects=False,
    )

    assert response.status_code == 400
    assert app.state.preview_only is True


def test_mode_toggle_changes_workflow_guard_without_persisting_settings(
    tmp_path, monkeypatch
) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    monkeypatch.setattr(Settings, "validate", lambda self: [])
    submitted: list[tuple[str, int | None]] = []

    def fake_submit(workflow_id, start_date, end_date, *, period_id=None):
        submitted.append((workflow_id, period_id))
        return app.state.history.create_run(
            workflow_id,
            "ECM Step 0",
            start_date.isoformat(),
            end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)

    blocked = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-0", follow_redirects=False
    )
    assert blocked.status_code == 303
    assert submitted == []

    client.post("/mode/toggle", data={"return_to": "/report-history"})
    allowed = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-0", follow_redirects=False
    )

    assert allowed.status_code == 303
    assert allowed.headers["location"].startswith("/runs/")
    assert submitted == [("ecm-step-0", period.id)]
    assert app.state.settings.preview_only is True
def test_step_0_is_visible_in_report_history_but_not_manual_run(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    client = TestClient(v2_main.create_app())

    report_history = client.get("/report-history")
    manual_run = client.get("/manual-run")

    assert "ECM Step 0 — Export Properties" in report_history.text
    assert "ECM Step 0 — Export Properties" not in manual_run.text


def test_report_history_uses_state_specific_terminal_action_labels(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    succeeded = app.state.history.create_run(
        "ecm-step-0",
        "ECM Step 0 — Export Properties",
        period.start_date,
        period.end_date,
        period_id=period.id,
    )
    app.state.history.mark_finished(succeeded.id, state="succeeded", exit_code=0)
    failed = app.state.history.create_run(
        "ecm-step-1a",
        "ECM Step 1a — Create Review",
        period.start_date,
        period.end_date,
        period_id=period.id,
    )
    app.state.history.mark_finished(failed.id, state="failed", exit_code=1)

    page = client.get("/report-history")

    assert page.status_code == 200
    assert page.text.count(">Rerun<") == 1
    assert page.text.count(">Retry<") == 1
    assert "disabled>Rerun</button>" in page.text
    assert "disabled>Retry</button>" in page.text


def test_period_run_uses_saved_dates_and_preview_blocks_before_submission(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-0", follow_redirects=False
    )

    assert response.status_code == 303
    assert "preview-only" in response.headers["location"]
    assert app.state.history.list_runs(limit=100) == []


def test_v2_history_includes_step_0_and_hides_unknown_workflow_records(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    app.state.history.create_run(
        "ecm-step-0", "ECM Step 0", period.start_date, period.end_date, period_id=period.id
    )
    app.state.history.create_run(
        "ecm-step-0-hidden", "Hidden", period.start_date, period.end_date
    )

    page = client.get("/history")

    assert page.status_code == 200
    assert "ECM Step 0" in page.text
    assert "ecm-step-0-hidden" not in page.text


def test_execution_rejects_out_of_order_ecm_step_before_submission(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "false")
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    monkeypatch.setattr(Settings, "validate", lambda self: [])

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-1a", follow_redirects=False
    )

    assert response.status_code == 303
    assert "Step+0+must+succeed" in response.headers["location"]
    assert app.state.history.list_runs(limit=100) == []


def test_execution_allows_ecm_flow_for_another_period(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "false")
    app = v2_main.create_app()
    client = TestClient(app)
    owner = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    independent = app.state.history.create_report_period("2026-07-27", "2026-08-02")
    app.state.history.create_run(
        "ecm-step-0", "ECM Step 0", owner.start_date, owner.end_date, period_id=owner.id
    )
    monkeypatch.setattr(Settings, "validate", lambda self: [])
    submitted: list[tuple[str, int | None]] = []

    def fake_submit(workflow_id, start_date, end_date, *, period_id=None):
        submitted.append((workflow_id, period_id))
        return app.state.history.create_run(
            workflow_id,
            "ECM Step 0",
            start_date.isoformat(),
            end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)

    response = client.post(
        f"/report-periods/{independent.id}/runs/ecm-step-0", follow_redirects=False
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/runs/")
    assert submitted == [("ecm-step-0", independent.id)]
    assert len(app.state.history.list_runs(limit=100)) == 2


def test_execution_rejects_invalid_configuration_before_submission(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "false")
    monkeypatch.setenv("STEP0_PYTHON", str(tmp_path / "missing-python.exe"))
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-0", follow_redirects=False
    )

    assert response.status_code == 303
    assert "configuration+issues" in response.headers["location"]
    assert app.state.history.list_runs(limit=100) == []


def test_step_1b_ignores_step_0_only_configuration_errors(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "false")
    monkeypatch.setenv(
        "ECM_RECORD", str(tmp_path / "test-workbooks" / "ECM_Record_test.xlsx")
    )
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    step_0 = app.state.history.create_run(
        "ecm-step-0", "ECM Step 0", period.start_date, period.end_date, period_id=period.id
    )
    app.state.history.mark_finished(step_0.id, state="succeeded", exit_code=0)
    step_1a = app.state.history.create_run(
        "ecm-step-1a", "ECM Step 1a", period.start_date, period.end_date, period_id=period.id
    )
    app.state.history.mark_finished(step_1a.id, state="succeeded", exit_code=0)
    app.state.settings.ecm_record.parent.mkdir(parents=True, exist_ok=True)
    app.state.settings.ecm_record.write_text("placeholder", encoding="utf-8")
    monkeypatch.setattr(
        Settings,
        "validate",
        lambda self: ["Step 0 Python environment is missing Playwright."],
    )
    submitted: list[tuple[str, int]] = []

    def fake_submit(workflow_id, start_date, end_date, *, period_id=None):
        submitted.append((workflow_id, period_id))
        return app.state.history.create_run(
            workflow_id,
            "ECM Step 1b",
            start_date.isoformat(),
            end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)
    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-1b", follow_redirects=False
    )

    assert response.status_code == 303
    assert submitted == [("ecm-step-1b", period.id)]


def test_execution_rejects_repeating_successful_step_0(tmp_path, monkeypatch) -> None:
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "false")
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    run = app.state.history.create_run(
        "ecm-step-0", "ECM Step 0", period.start_date, period.end_date, period_id=period.id
    )
    app.state.history.mark_finished(run.id, state="succeeded", exit_code=0)
    monkeypatch.setattr(Settings, "validate", lambda self: [])

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-0", follow_redirects=False
    )

    assert response.status_code == 303
    assert "already+succeeded" in response.headers["location"]
    assert len(app.state.history.list_runs(limit=100)) == 1


def _run(run_id: int, workflow_id: str, state: str) -> RunRecord:
    timestamp = datetime.now().astimezone().isoformat()
    return RunRecord(
        id=run_id,
        period_id=1,
        workflow_id=workflow_id,
        workflow_label=workflow_id,
        start_date="2026-07-20",
        end_date="2026-07-26",
        state=state,
        created_at=timestamp,
        started_at=timestamp,
        finished_at=timestamp if state in {"succeeded", "failed"} else None,
        exit_code=0 if state == "succeeded" else None,
        error_summary=None,
        log_path=tmp_path_placeholder(),
    )


def tmp_path_placeholder():
    from pathlib import Path

    return Path("run.log")


def test_successful_step_1b_requires_manual_decisions_before_step_2() -> None:
    period = ReportPeriodRecord(
        id=1,
        start_date="2026-07-20",
        end_date="2026-07-26",
        created_at="2026-07-26T00:00:00+08:00",
        updated_at="2026-07-26T00:00:00+08:00",
        is_deleted=False,
    )
    summary = ReportPeriodSummary(
        period=period,
        runs={
            "weekly-report": _run(1, "weekly-report", "succeeded"),
            "ecm-step-0": None,
            "ecm-step-1a": _run(2, "ecm-step-1a", "succeeded"),
            "ecm-step-1b": _run(3, "ecm-step-1b", "succeeded"),
            "ecm-step-2": None,
        },
        last_updated_at="2026-07-26T10:00:00+08:00",
    )

    dashboard = build_period_dashboards([summary])[0]

    assert dashboard.ecm_status == "Waiting for manual Use SMM / Use KMS decisions"
    assert "manual decisions" in dashboard.next_action


def test_legacy_period_with_step_1a_history_can_continue_without_step_0() -> None:
    period = ReportPeriodRecord(
        id=1,
        start_date="2026-07-20",
        end_date="2026-07-26",
        created_at="2026-07-26T00:00:00+08:00",
        updated_at="2026-07-26T00:00:00+08:00",
        is_deleted=False,
    )
    summary = ReportPeriodSummary(
        period=period,
        runs={
            "weekly-report": None,
            "ecm-step-1a": _run(1, "ecm-step-1a", "succeeded"),
            "ecm-step-1b": None,
            "ecm-step-2": None,
        },
        last_updated_at="2026-07-26T10:00:00+08:00",
    )

    dashboard = build_period_dashboards([summary])[0]

    assert dashboard.ecm_status == "Waiting for manual Remark review"
