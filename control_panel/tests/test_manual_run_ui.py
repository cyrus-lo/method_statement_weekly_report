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


def test_dashboard_is_preserved_and_manual_run_is_a_separate_page(
    tmp_path: Path, monkeypatch
) -> None:
    client = build_client(tmp_path, monkeypatch)

    dashboard = client.get("/")
    manual_run = client.get("/manual-run")

    assert dashboard.status_code == 200
    assert "<title>Dashboard · Workflow Control Panel</title>" in dashboard.text
    assert "Run the weekly reporting workflow with confidence." in dashboard.text
    assert manual_run.status_code == 200
    assert "<title>Manual Run · Workflow Control Panel</title>" in manual_run.text
    assert "Run the weekly reporting workflow manually" in manual_run.text
    assert "Manual Run" in dashboard.text
    assert "Dashboard" in manual_run.text
    for page in (dashboard, manual_run):
        assert page.text.count("V1 · Stable") == 1
        assert page.text.count("Weekly Method Statement Report") == 1
        assert page.text.count("ECM Step 1a — Create Review") == 1
        assert page.text.count("ECM Step 1b — Finalize Review") == 1
        assert page.text.count("ECM Step 2 — Create SMM Summary") == 1


def test_manual_run_validation_error_returns_to_manual_run(
    tmp_path: Path, monkeypatch
) -> None:
    client = build_client(tmp_path, monkeypatch)

    response = client.post(
        "/runs/weekly-report",
        data={
            "start_date": "not-a-date",
            "end_date": "2026-07-26",
            "origin": "/manual-run",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/manual-run?")
    page = client.get(response.headers["location"])
    assert "Start date must use YYYY-MM-DD." in page.text
    assert "<title>Manual Run · Workflow Control Panel</title>" in page.text


def test_workflow_started_from_manual_run_keeps_manual_run_as_back_link(
    tmp_path: Path, monkeypatch
) -> None:
    client = build_client(tmp_path, monkeypatch)
    app = client.app

    def fake_submit(
        workflow_id: str, start_date: date, end_date: date, *, period_id: int | None = None
    ):
        return app.state.history.create_run(
            workflow_id,
            "Weekly Method Statement Report",
            start_date.isoformat(),
            end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)
    response = client.post(
        "/runs/weekly-report",
        data={
            "start_date": "2026-07-20",
            "end_date": "2026-07-26",
            "origin": "/manual-run",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    location = response.headers["location"]
    assert location.startswith("/runs/1?")
    detail = client.get(location)
    assert detail.status_code == 200
    assert '<a class="back-link" href="/manual-run">← Manual Run</a>' in detail.text


def test_invalid_run_origin_falls_back_to_dashboard(
    tmp_path: Path, monkeypatch
) -> None:
    client = build_client(tmp_path, monkeypatch)

    response = client.post(
        "/runs/weekly-report",
        data={
            "start_date": "not-a-date",
            "end_date": "2026-07-26",
            "origin": "https://example.com/",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/?")


def test_v1_rejects_step_0_from_direct_run_route(tmp_path: Path, monkeypatch) -> None:
    client = build_client(tmp_path, monkeypatch)

    response = client.post(
        "/runs/ecm-step-0",
        data={
            "start_date": "2026-07-20",
            "end_date": "2026-07-26",
            "origin": "/manual-run",
        },
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/manual-run?")
    assert len(client.app.state.history.list_runs(limit=100)) == 0
