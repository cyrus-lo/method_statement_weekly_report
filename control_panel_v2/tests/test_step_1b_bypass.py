from __future__ import annotations

from fastapi.testclient import TestClient

import control_panel_v2.main as v2_main
from control_panel_v2.services.config import Settings


def _set_isolated_workbook_paths(monkeypatch, tmp_path, *, live: bool = False) -> None:
    """Keep every app instance in this module away from real ECM/SMM workbooks."""
    output_dir = tmp_path / "test-workbooks"
    output_dir.mkdir(parents=True, exist_ok=True)
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "false" if live else "true")
    monkeypatch.setenv("ECM_RECORD", str(output_dir / "ECM_Record_test.xlsx"))
    monkeypatch.setenv("SMM_RECORD", str(output_dir / "SMM_Record_test.xlsx"))


def test_step_2_without_step_1b_requires_explicit_confirmation(tmp_path, monkeypatch) -> None:
    _set_isolated_workbook_paths(monkeypatch, tmp_path, live=True)
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    monkeypatch.setattr(Settings, "validate", lambda self: [])
    submitted: list[str] = []
    monkeypatch.setattr(
        app.state.runner,
        "submit",
        lambda workflow_id, *args, **kwargs: submitted.append(workflow_id),
    )

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-2", follow_redirects=False
    )

    assert response.status_code == 303
    assert "Step+1b+must+succeed" in response.headers["location"]
    assert submitted == []
    assert app.state.history.list_runs(limit=100) == []


def test_step_2_runs_after_explicit_confirmation_when_step_1b_is_absent(
    tmp_path, monkeypatch
) -> None:
    _set_isolated_workbook_paths(monkeypatch, tmp_path, live=True)
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    monkeypatch.setattr(Settings, "validate", lambda self: [])
    submitted: list[tuple[str, int | None]] = []

    def fake_submit(workflow_id, start_date, end_date, *, period_id=None):
        submitted.append((workflow_id, period_id))
        return app.state.history.create_run(
            workflow_id,
            "ECM Step 2",
            start_date.isoformat(),
            end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-2",
        data={"confirm_step_1b_bypass": "true"},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/runs/")
    assert submitted == [("ecm-step-2", period.id)]


def test_step_2_runs_normally_after_step_1b_succeeds(tmp_path, monkeypatch) -> None:
    _set_isolated_workbook_paths(monkeypatch, tmp_path, live=True)
    app = v2_main.create_app()
    client = TestClient(app)
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
    step_1b = app.state.history.create_run(
        "ecm-step-1b",
        "ECM Step 1b",
        period.start_date,
        period.end_date,
        period_id=period.id,
    )
    app.state.history.mark_finished(step_1b.id, state="succeeded", exit_code=0)
    monkeypatch.setattr(Settings, "validate", lambda self: [])

    def fake_submit(workflow_id, start_date, end_date, *, period_id=None):
        return app.state.history.create_run(
            workflow_id,
            "ECM Step 2",
            start_date.isoformat(),
            end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-2", follow_redirects=False
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/runs/")


def test_step_2_bypass_cannot_override_existing_step_1b_runs(tmp_path, monkeypatch) -> None:
    for state in ("queued", "running", "failed"):
        case_dir = tmp_path / state
        _set_isolated_workbook_paths(monkeypatch, case_dir, live=True)
        app = v2_main.create_app()
        client = TestClient(app)
        period = app.state.history.create_report_period("2026-07-20", "2026-07-26")
        step_1b = app.state.history.create_run(
            "ecm-step-1b",
            "ECM Step 1b",
            period.start_date,
            period.end_date,
            period_id=period.id,
        )
        if state == "running":
            app.state.history.mark_running(step_1b.id)
        elif state == "failed":
            app.state.history.mark_finished(step_1b.id, state="failed", exit_code=1)
        monkeypatch.setattr(Settings, "validate", lambda self: [])
        submitted: list[str] = []
        monkeypatch.setattr(
            app.state.runner,
            "submit",
            lambda workflow_id, *args, **kwargs: submitted.append(workflow_id),
        )

        response = client.post(
            f"/report-periods/{period.id}/runs/ecm-step-2",
            data={"confirm_step_1b_bypass": "true"},
            follow_redirects=False,
        )

        assert response.status_code == 303
        assert "Step+1b+must+succeed" in response.headers["location"]
        assert submitted == []
        assert len(app.state.history.list_runs(limit=100)) == 1


def test_report_history_renders_step_2_bypass_confirmation_for_unstarted_step_1b(
    tmp_path, monkeypatch
) -> None:
    _set_isolated_workbook_paths(monkeypatch, tmp_path)
    app = v2_main.create_app()
    period = app.state.history.create_report_period("2026-07-20", "2026-07-26")

    page = TestClient(app).get("/report-history")

    assert page.status_code == 200
    assert 'data-step-one-b-state="not-started"' in page.text
    assert 'id="step-2-bypass-dialog"' in page.text
    assert "Process anyway" in page.text
    assert "skips Step 1b review/finalization and its pre- and post-Step-1b backups" in page.text
    assert f"/report-periods/{period.id}/runs/ecm-step-2" in page.text
