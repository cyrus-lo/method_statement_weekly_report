from __future__ import annotations

from datetime import date
from pathlib import Path

from fastapi.testclient import TestClient

import control_panel.main as control_main

from .test_workflows import build_settings


def build_client(tmp_path: Path, monkeypatch) -> TestClient:
    settings = build_settings(tmp_path)
    monkeypatch.setattr(control_main, "load_settings", lambda: settings)
    return TestClient(control_main.create_app())


def test_report_history_renders_period_flow_and_creates_a_period(tmp_path: Path, monkeypatch) -> None:
    client = build_client(tmp_path, monkeypatch)

    page = client.get("/report-history")
    assert page.status_code == 200
    assert "Report History" in page.text
    assert "Standard view" in page.text
    assert "Detailed view" in page.text
    assert "Deleted view" in page.text
    assert 'id="detailed-view-tab"' in page.text
    assert 'aria-selected="true"' in page.text
    assert page.text.count("Last activity (HKT)") == 3
    assert "Weekly Method Statement Report" in page.text
    assert "ECM Step 0 — Export Properties" not in page.text
    assert "ECM Step 1a — Create Review" in page.text
    assert "ECM Step 1b — Finalize Review" in page.text
    assert "ECM Step 2 — Create SMM Summary" in page.text

    response = client.post(
        "/report-periods",
        data={"start_date": "2026-06-22", "end_date": "2026-06-30"},
        follow_redirects=False,
    )
    assert response.status_code == 303
    page = client.get("/report-history")
    assert "2026-06-22" in page.text
    assert "2026-06-30" in page.text
    assert page.text.count("Pending") == 8
    assert 'action="/report-periods/1/delete"' in page.text


def test_report_history_soft_deletes_and_restores_a_period(tmp_path: Path, monkeypatch) -> None:
    client = build_client(tmp_path, monkeypatch)
    period = client.app.state.history.create_report_period("2026-07-13", "2026-07-19")

    response = client.post(f"/report-periods/{period.id}/delete", follow_redirects=False)

    assert response.status_code == 303
    assert client.app.state.history.get_report_period(period.id).is_deleted
    page = client.get("/report-history")
    assert 'action="/report-periods/1/delete"' not in page.text
    assert 'action="/report-periods/1/restore"' in page.text
    assert "Deleted period. Workflow history is preserved." in page.text
    assert "Run workflow" not in page.text

    response = client.post(f"/report-periods/{period.id}/restore", follow_redirects=False)

    assert response.status_code == 303
    assert not client.app.state.history.get_report_period(period.id).is_deleted
    page = client.get("/report-history")
    assert 'action="/report-periods/1/delete"' in page.text
    assert 'action="/report-periods/1/restore"' not in page.text


def test_report_history_uses_state_specific_terminal_action_labels(
    tmp_path: Path, monkeypatch
) -> None:
    client = build_client(tmp_path, monkeypatch)
    app = client.app
    period = app.state.history.create_report_period("2026-07-13", "2026-07-19")
    succeeded = app.state.history.create_run(
        "weekly-report",
        "Weekly Method Statement Report",
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
    app.state.history.mark_finished(
        failed.id, state="failed", exit_code=1, error_summary="Test failure"
    )

    page = client.get("/report-history")

    assert page.status_code == 200
    assert page.text.count(">Rerun<") == 2
    assert page.text.count(">Retry<") == 2
    assert page.text.count(">Run workflow<") == 4


def test_report_history_rejects_deleting_a_period_with_an_active_run(tmp_path: Path, monkeypatch) -> None:
    client = build_client(tmp_path, monkeypatch)
    period = client.app.state.history.create_report_period("2026-07-13", "2026-07-19")
    client.app.state.history.create_run(
        "weekly-report", "Weekly Method Statement Report", period.start_date, period.end_date,
        period_id=period.id,
    )

    response = client.post(
        f"/report-periods/{period.id}/delete", follow_redirects=True
    )

    assert response.status_code == 200
    assert "cannot be deleted while a workflow is running" in response.text
    assert not client.app.state.history.get_report_period(period.id).is_deleted


def test_report_history_rejects_invalid_period_dates(tmp_path: Path, monkeypatch) -> None:
    client = build_client(tmp_path, monkeypatch)

    response = client.post(
        "/report-periods",
        data={"start_date": "2026-07-12", "end_date": "2026-07-01"},
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert "End date cannot be earlier than start date." in response.text


def test_period_workflow_run_uses_saved_period_dates(tmp_path: Path, monkeypatch) -> None:
    client = build_client(tmp_path, monkeypatch)
    app = client.app
    period = app.state.history.create_report_period("2026-07-01", "2026-07-12")
    submitted: dict[str, object] = {}

    def fake_submit(
        workflow_id: str, start_date: date, end_date: date, *, period_id: int | None = None
    ):
        submitted.update(
            workflow_id=workflow_id,
            start_date=start_date,
            end_date=end_date,
            period_id=period_id,
        )
        return app.state.history.create_run(
            workflow_id,
            "Weekly Method Statement Report",
            start_date.isoformat(),
            end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)
    response = client.post(
        f"/report-periods/{period.id}/runs/weekly-report", follow_redirects=False
    )

    assert response.status_code == 303
    assert submitted == {
        "workflow_id": "weekly-report",
        "start_date": date(2026, 7, 1),
        "end_date": date(2026, 7, 12),
        "period_id": period.id,
    }
    page = client.get("/report-history")
    assert "Period locked after first workflow run" in page.text
    assert "View run #1" in page.text


def test_v1_history_rejects_step_0_filter(tmp_path: Path, monkeypatch) -> None:
    client = build_client(tmp_path, monkeypatch)

    response = client.get("/history?workflow_id=ecm-step-0")

    assert response.status_code == 404