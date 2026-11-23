from __future__ import annotations

import shutil
from collections.abc import Iterator
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from control_panel_v3 import main as app_module
from control_panel_v3.services.config import CONTROL_PANEL_ROOT, load_settings


@pytest.fixture
def client(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Iterator[TestClient]:
    settings = load_settings()
    test_root = CONTROL_PANEL_ROOT / "tests" / f"period-routes-{tmp_path.name}"
    shutil.rmtree(test_root, ignore_errors=True)
    roots = {
        "data_dir": test_root / "data",
        "archive_dir": test_root / "archive",
        "csv_dir": test_root / "csv",
        "ecm_output_dir": test_root / "ecm-output",
        "smm_output_dir": test_root / "smm-output",
        "django_output_dir": test_root / "django-output",
        "edge_profile_dir": test_root / "edge-profile",
        "ecm_template_path": test_root / "templates" / "ecm.xlsx",
        "smm_template_path": test_root / "templates" / "smm.xlsx",
        "weekly_template_path": test_root / "weekly-template" / "weekly.xlsx",
        "readiness_workbook_path": test_root / "readiness" / "readiness.xlsx",
    }
    test_settings = type(settings)(**{**settings.__dict__, **roots, "preview_only": True})
    monkeypatch.setattr(app_module, "load_settings", lambda: test_settings)
    app = app_module.create_app()
    try:
        yield TestClient(app)
    finally:
        shutil.rmtree(test_root, ignore_errors=True)


def test_create_period_post_accepts_valid_dates_and_reports_duplicates(client: TestClient) -> None:
    response = client.post(
        "/report-periods",
        data={"start_date": "2026-10-01", "end_date": "2026-10-07"},
        follow_redirects=False,
    )
    assert response.status_code == 303
    assert response.headers["location"] == "/report-history"
    assert len(client.app.state.history.list_report_period_summaries(())) == 1

    duplicate = client.post(
        "/report-periods",
        data={"start_date": "2026-10-01", "end_date": "2026-10-07"},
        follow_redirects=False,
    )
    assert duplicate.status_code == 303
    assert "already+exists" in duplicate.headers["location"]
    assert len(client.app.state.history.list_report_period_summaries(())) == 1


@pytest.mark.parametrize(
    ("start_date", "end_date", "message"),
    [
        ("not-a-date", "2026-10-07", "YYYY-MM-DD"),
        ("2026-10-08", "2026-10-07", "earlier than start date"),
    ],
)
def test_create_period_post_rejects_invalid_date_ranges(
    client: TestClient, start_date: str, end_date: str, message: str
) -> None:
    response = client.post(
        "/report-periods",
        data={"start_date": start_date, "end_date": end_date},
        follow_redirects=False,
    )
    assert response.status_code == 303
    assert message.replace(" ", "+") in response.headers["location"]
    assert client.app.state.history.list_report_period_summaries(()) == []


def test_manual_run_page_keeps_workflows_disabled_and_shared_navigation(client: TestClient) -> None:
    response = client.get("/manual-run")
    assert response.status_code == 200
    assert 'href="/manual-run"' in response.text
    assert 'href="/archive"' in response.text
    assert "/static/app.js?v=3" in response.text
    assert "ECM First-Date Lookup" in response.text
    assert response.text.count("Start from Report History</button>") == 5
    assert client.get("/static/app.js?v=3").status_code == 200


def test_dashboard_uses_period_first_layout_and_keeps_v3_lookup_step(client: TestClient) -> None:
    client.app.state.history.create_report_period("2026-10-01", "2026-10-07")
    response = client.get("/")
    assert response.status_code == 200
    assert "Workflow status by reporting period" in response.text
    assert "First-Date Lookup" in response.text
    assert "Weekly Method Statement Report" in response.text
    assert "Last activity:" in response.text
    assert 'href="/report-history"' in response.text


def test_report_history_shows_hkt_activity_and_confirms_period_deletion(client: TestClient) -> None:
    history = client.app.state.history
    period = history.create_report_period("2026-10-01", "2026-10-07")
    history.create_run("ecm-step-0", "ECM Step 0", period.start_date, period.end_date, period_id=period.id)

    response = client.get("/report-history")
    assert response.status_code == 200
    assert "Last activity (HKT)" in response.text
    assert "Queued:" in response.text
    assert "Delete this report period?" in response.text
    assert "Lookup scope" in response.text


def test_report_history_marks_active_run_for_status_polling(client: TestClient) -> None:
    history = client.app.state.history
    idle = client.get("/report-history")
    assert 'data-run-status-url=' not in idle.text

    period = history.create_report_period("2026-10-01", "2026-10-07")
    run = history.create_run("ecm-step-0", "ECM Step 0", period.start_date, period.end_date, period_id=period.id)
    active = client.get("/report-history")
    assert f'data-run-status-url="/runs/{run.id}/status"' in active.text
    assert 'data-active="true"' in active.text

    history.mark_finished(run.id, state="failed", exit_code=1, error_summary="test failure")
    finished = client.get("/report-history")
    assert 'data-run-status-url=' not in finished.text


def test_run_detail_uses_the_shared_script_only_once(client: TestClient) -> None:
    run = client.app.state.history.create_run("weekly-report", "Weekly report", "2026-10-01", "2026-10-07")
    response = client.get(f"/runs/{run.id}")
    assert response.status_code == 200
    assert response.text.count("app.js?v=3") == 1
    assert "Exit code" in response.text
    assert "Next required action" in response.text
    assert "Saved artifacts" in response.text


def test_period_update_delete_and_restore_conflict_are_reported_safely(client: TestClient) -> None:
    history = client.app.state.history
    first = history.create_report_period("2026-10-01", "2026-10-07")
    second = history.create_report_period("2026-10-08", "2026-10-14")

    invalid_update = client.post(
        f"/report-periods/{first.id}",
        data={"start_date": "2026-10-20", "end_date": "2026-10-19"},
        follow_redirects=False,
    )
    assert invalid_update.status_code == 303
    assert "earlier+than+start+date" in invalid_update.headers["location"]
    assert history.get_report_period(first.id).start_date == "2026-10-01"

    deleted = client.post(f"/report-periods/{second.id}/delete", follow_redirects=False)
    assert deleted.status_code == 303
    assert history.get_report_period(second.id).is_deleted

    updated = client.post(
        f"/report-periods/{first.id}",
        data={"start_date": "2026-10-08", "end_date": "2026-10-14"},
        follow_redirects=False,
    )
    assert updated.status_code == 303
    restored_conflict = client.post(f"/report-periods/{second.id}/restore", follow_redirects=False)
    assert restored_conflict.status_code == 303
    assert "already+uses+these+dates" in restored_conflict.headers["location"]
    assert history.get_report_period(second.id).is_deleted
