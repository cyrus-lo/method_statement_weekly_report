from __future__ import annotations

from fastapi.testclient import TestClient

import control_panel_v2.main as v2_main
from control_panel_v2.services.config import Settings


def _app(tmp_path, monkeypatch, *, live: bool = False):
    monkeypatch.setenv("CONTROL_PANEL_V2_DATA_DIR", str(tmp_path / "v2-data"))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "false" if live else "true")
    app = v2_main.create_app()
    return app, TestClient(app)


def _successful_step_0(app, period):
    run = app.state.history.create_run(
        "ecm-step-0",
        "ECM Step 0 — Export Properties",
        period.start_date,
        period.end_date,
        period_id=period.id,
    )
    app.state.history.mark_finished(run.id, state="succeeded", exit_code=0)
    return run


def test_successful_current_step_0_shows_next_step_action(tmp_path, monkeypatch) -> None:
    app, client = _app(tmp_path, monkeypatch)
    period = app.state.history.create_report_period("2026-10-01", "2026-10-11")
    run = _successful_step_0(app, period)

    page = client.get(f"/runs/{run.id}")

    assert page.status_code == 200
    assert f'action="/report-periods/{period.id}/runs/ecm-step-1a"' in page.text
    assert f'name="source_run_id" value="{run.id}"' in page.text
    assert ">Run Next Step</button>" in page.text
    assert "disabled>Run Next Step</button>" in page.text


def test_next_step_action_hidden_for_failed_stale_or_advanced_runs(
    tmp_path, monkeypatch
) -> None:
    app, client = _app(tmp_path, monkeypatch)
    failed_period = app.state.history.create_report_period("2026-10-01", "2026-10-07")
    failed = app.state.history.create_run(
        "ecm-step-0", "ECM Step 0", failed_period.start_date, failed_period.end_date,
        period_id=failed_period.id,
    )
    app.state.history.mark_finished(failed.id, state="failed", exit_code=1)

    stale_period = app.state.history.create_report_period("2026-10-08", "2026-10-14")
    stale = _successful_step_0(app, stale_period)
    newer = app.state.history.create_run(
        "ecm-step-0", "ECM Step 0 retry", stale_period.start_date,
        stale_period.end_date, period_id=stale_period.id,
    )
    app.state.history.mark_finished(newer.id, state="failed", exit_code=1)

    advanced_period = app.state.history.create_report_period("2026-10-15", "2026-10-21")
    advanced_step_0 = _successful_step_0(app, advanced_period)
    app.state.history.create_run(
        "ecm-step-1a", "ECM Step 1a", advanced_period.start_date,
        advanced_period.end_date, period_id=advanced_period.id,
    )

    for run in (failed, stale, advanced_step_0):
        page = client.get(f"/runs/{run.id}")
        assert page.status_code == 200
        assert "Run Next Step" not in page.text


def test_next_step_post_rechecks_source_run_and_uses_period_dates(
    tmp_path, monkeypatch
) -> None:
    app, client = _app(tmp_path, monkeypatch, live=True)
    period = app.state.history.create_report_period("2026-10-01", "2026-10-11")
    run = _successful_step_0(app, period)
    monkeypatch.setattr(Settings, "validate", lambda self: [])
    submitted = []

    def fake_submit(workflow_id, start_date, end_date, *, period_id=None):
        submitted.append((workflow_id, start_date.isoformat(), end_date.isoformat(), period_id))
        return app.state.history.create_run(
            workflow_id, "ECM Step 1a", start_date.isoformat(), end_date.isoformat(),
            period_id=period_id,
        )

    monkeypatch.setattr(app.state.runner, "submit", fake_submit)
    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-1a",
        data={"source_run_id": str(run.id)},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert response.headers["location"].startswith("/runs/")
    assert submitted == [("ecm-step-1a", "2026-10-01", "2026-10-11", period.id)]


def test_next_step_post_rejects_stale_source_run(tmp_path, monkeypatch) -> None:
    app, client = _app(tmp_path, monkeypatch, live=True)
    period = app.state.history.create_report_period("2026-10-01", "2026-10-11")
    source = _successful_step_0(app, period)
    app.state.history.create_run(
        "ecm-step-1a", "ECM Step 1a", period.start_date, period.end_date,
        period_id=period.id,
    )
    monkeypatch.setattr(Settings, "validate", lambda self: [])
    submitted = []
    monkeypatch.setattr(
        app.state.runner, "submit", lambda *args, **kwargs: submitted.append(args)
    )

    response = client.post(
        f"/report-periods/{period.id}/runs/ecm-step-1a",
        data={"source_run_id": str(source.id)},
        follow_redirects=False,
    )

    assert response.status_code == 303
    assert "no+longer+eligible" in response.headers["location"]
    assert submitted == []


def test_csv_artifact_can_be_viewed_inline_or_downloaded(tmp_path, monkeypatch) -> None:
    app, client = _app(tmp_path, monkeypatch)
    period = app.state.history.create_report_period("2026-10-01", "2026-10-11")
    run = _successful_step_0(app, period)
    artifact_dir = app.state.settings.runs_dir / str(run.id)
    artifact_dir.mkdir(parents=True, exist_ok=True)
    csv_path = artifact_dir / "properties.csv"
    csv_path.write_text("id,name\n1,Example\n", encoding="utf-8")
    artifact = app.state.history.add_artifact(run.id, "Properties", csv_path, csv_path)

    page = client.get(f"/runs/{run.id}")
    inline = client.get(f"/runs/{run.id}/artifacts/{artifact.id}/view")
    download = client.get(f"/runs/{run.id}/artifacts/{artifact.id}")
    status = client.get(f"/runs/{run.id}/status").json()

    assert "/view" in page.text and "Download" in page.text
    assert inline.status_code == 200
    assert inline.headers["content-type"].startswith("text/plain")
    assert inline.headers["content-disposition"].startswith("inline;")
    assert inline.headers["cache-control"] == "no-store"
    assert inline.text == "id,name\r\n1,Example\r\n"
    assert download.status_code == 200
    assert download.headers["content-disposition"].startswith("attachment;")
    assert status["artifacts"][0]["view_url"].endswith("/view")


def test_inline_view_rejects_non_csv_and_artifact_from_another_run(
    tmp_path, monkeypatch
) -> None:
    app, client = _app(tmp_path, monkeypatch)
    period = app.state.history.create_report_period("2026-10-01", "2026-10-11")
    run = _successful_step_0(app, period)
    artifact_dir = app.state.settings.runs_dir / str(run.id)
    artifact_dir.mkdir(parents=True, exist_ok=True)
    workbook = artifact_dir / "review.xlsx"
    workbook.write_bytes(b"workbook")
    artifact = app.state.history.add_artifact(run.id, "Workbook", workbook, workbook)

    assert client.get(f"/runs/{run.id}/artifacts/{artifact.id}/view").status_code == 404
    assert client.get(f"/runs/{run.id + 1}/artifacts/{artifact.id}/view").status_code == 404