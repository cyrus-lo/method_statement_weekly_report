from __future__ import annotations

import ast
import sqlite3
from pathlib import Path

import pytest

from control_panel_v3.services import bootstrap
from control_panel_v3.services.archive import ArchiveStore
from control_panel_v3.services.config import CONTROL_PANEL_ROOT, Settings, load_settings
from control_panel_v3.services.history import HistoryStore


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_default_settings_use_separate_v3_roots(monkeypatch: pytest.MonkeyPatch) -> None:
    for name in (
        "CONTROL_PANEL_V3_DATA_DIR", "CONTROL_PANEL_V3_ARCHIVE_DIR", "CONTROL_PANEL_V3_CSV_DIR",
        "CONTROL_PANEL_V3_ECM_OUTPUT_DIR", "CONTROL_PANEL_V3_SMM_OUTPUT_DIR",
        "CONTROL_PANEL_V3_DJANGO_OUTPUT_DIR", "CONTROL_PANEL_V3_EDGE_PROFILE_DIR",
        "CONTROL_PANEL_V3_READINESS_WORKBOOK",
        "CONTROL_PANEL_V3_WEEKLY_TEMPLATE",
    ):
        monkeypatch.delenv(name, raising=False)
    settings = load_settings()
    assert settings.port == 8767
    assert settings.preview_only is True
    for path in (settings.data_dir, settings.archive_dir, settings.csv_dir, settings.ecm_output_dir,
                 settings.smm_output_dir, settings.django_output_dir, settings.edge_profile_dir,
                 settings.readiness_workbook_path, settings.weekly_template_path):
        assert path.is_relative_to(CONTROL_PANEL_ROOT.resolve())
    assert settings.readiness_workbook_path == CONTROL_PANEL_ROOT / "working" / "readiness" / "SMM_KMS_PM_Master_Document.xlsx"
    assert settings.weekly_template_path == CONTROL_PANEL_ROOT / "working" / "templates" / "Method_Statement_Interaction_Weekly_Report_template.xlsx"


def test_v3_rejects_paths_outside_its_owned_root(tmp_path: Path) -> None:
    settings = Settings(
        project_root=PROJECT_ROOT,
        data_dir=tmp_path / "data", archive_dir=CONTROL_PANEL_ROOT / "archive-test",
        csv_dir=CONTROL_PANEL_ROOT / "csv-test", ecm_output_dir=CONTROL_PANEL_ROOT / "ecm-test",
        smm_output_dir=CONTROL_PANEL_ROOT / "smm-test", django_output_dir=CONTROL_PANEL_ROOT / "django-test",
        edge_profile_dir=CONTROL_PANEL_ROOT / "profile-test", preview_only=True, host="127.0.0.1",
        port=8767, job_timeout_seconds=60, ecm_url="https://example.test/", worker_timeout_seconds=60,
        weekly_template_path=CONTROL_PANEL_ROOT / "templates-test" / "weekly.xlsx",
    )
    assert any("must remain under" in error for error in settings.validate())


def test_v3_rejects_symlinked_writable_path_outside_its_root(tmp_path: Path) -> None:
    outside = tmp_path / "outside"
    outside.mkdir()
    link = CONTROL_PANEL_ROOT / "tests" / "path-link"
    try:
        link.symlink_to(outside, target_is_directory=True)
    except OSError as error:
        pytest.skip(f"Directory symlinks are not available: {error}")
    try:
        settings = load_settings()
        settings = Settings(**{**settings.__dict__, "data_dir": link})
        assert any("must remain under" in error for error in settings.validate())
    finally:
        link.unlink(missing_ok=True)


def test_v3_rejects_v2_port() -> None:
    settings = load_settings()
    settings = Settings(**{**settings.__dict__, "port": 8766})
    assert any("V2's port" in error for error in settings.validate())


def test_v3_rejects_overlapping_writable_roots() -> None:
    settings = load_settings()
    settings = Settings(**{**settings.__dict__, "archive_dir": settings.data_dir / "nested-archive"})
    assert any("must be distinct" in error for error in settings.validate())


def test_history_period_lifecycle_and_soft_delete(tmp_path: Path) -> None:
    store = HistoryStore(tmp_path / "history.sqlite3", tmp_path / "runs")
    store.initialize()
    period = store.create_report_period("2026-10-01", "2026-10-07")
    with pytest.raises(ValueError, match="already exists"):
        store.create_report_period("2026-10-01", "2026-10-07")
    deleted = store.delete_report_period(period.id)
    assert deleted.is_deleted
    restored = store.restore_report_period(period.id)
    assert not restored.is_deleted


def test_period_cannot_be_changed_after_a_run_or_deleted_while_active(tmp_path: Path) -> None:
    store = HistoryStore(tmp_path / "history.sqlite3", tmp_path / "runs")
    store.initialize()
    period = store.create_report_period("2026-10-01", "2026-10-07")
    run = store.create_run("ecm-step-0", "ECM Step 0", period.start_date, period.end_date, period_id=period.id)

    with pytest.raises(ValueError, match="running"):
        store.delete_report_period(period.id)
    store.mark_finished(run.id, state="succeeded", exit_code=0)
    with pytest.raises(ValueError, match="locked"):
        store.update_report_period(period.id, "2026-10-02", "2026-10-08")


def test_archive_import_copies_evidence_without_mutating_source(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    v2_root = tmp_path / "control_panel_v2"
    v3_root = tmp_path / "control_panel_v3"
    source = v2_root / "data"
    source_runs = source / "runs" / "1"
    source_runs.mkdir(parents=True)
    log_path = source_runs / "run.log"
    artifact_path = source_runs / "report.csv"
    log_path.write_text("V2 log\n", encoding="utf-8")
    artifact_path.write_text("value\n1\n", encoding="utf-8")
    source_db = source / "history.sqlite3"
    with sqlite3.connect(source_db) as connection:
        connection.executescript(
            "CREATE TABLE runs(id INTEGER PRIMARY KEY, log_path TEXT);"
            "CREATE TABLE artifacts(id INTEGER PRIMARY KEY, stored_path TEXT);"
        )
        connection.execute("INSERT INTO runs VALUES(1, ?)", (str(log_path),))
        connection.execute("INSERT INTO artifacts VALUES(1, ?)", (str(artifact_path),))

    monkeypatch.setattr(bootstrap, "V2_ROOT", v2_root)
    monkeypatch.setattr(bootstrap, "CONTROL_PANEL_ROOT", v3_root)
    destination = v3_root / "archive"
    bootstrap.import_v2_archive(source, destination)

    with sqlite3.connect(source_db) as connection:
        assert connection.execute("SELECT log_path FROM runs WHERE id=1").fetchone()[0] == str(log_path)
        assert connection.execute("SELECT stored_path FROM artifacts WHERE id=1").fetchone()[0] == str(artifact_path)
    with sqlite3.connect(destination / "history.sqlite3") as connection:
        archived_log = Path(connection.execute("SELECT log_path FROM runs WHERE id=1").fetchone()[0])
        archived_artifact = Path(connection.execute("SELECT stored_path FROM artifacts WHERE id=1").fetchone()[0])
    assert archived_log == destination / "runs" / "1" / "run.log"
    assert archived_artifact == destination / "runs" / "1" / "report.csv"
    assert archived_log.read_text(encoding="utf-8") == "V2 log\n"
    assert archived_artifact.read_text(encoding="utf-8") == "value\n1\n"


def test_archive_import_rejects_source_outside_v2_root(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="approved root"):
        bootstrap.import_v2_archive(tmp_path / "external-source", CONTROL_PANEL_ROOT / "archive-test")


def test_archive_store_reads_only_and_rejects_paths_outside_archive(tmp_path: Path) -> None:
    archive_dir = tmp_path / "archive"
    runs_dir = archive_dir / "runs"
    run_dir = runs_dir / "1"
    run_dir.mkdir(parents=True)
    log_path = run_dir / "run.log"
    artifact_path = run_dir / "report.csv"
    log_path.write_text("archived evidence", encoding="utf-8")
    artifact_path.write_text("value\n1\n", encoding="utf-8")
    database = archive_dir / "history.sqlite3"
    with sqlite3.connect(database) as connection:
        connection.executescript(
            "CREATE TABLE report_periods(id INTEGER PRIMARY KEY,start_date TEXT,end_date TEXT,created_at TEXT,updated_at TEXT,is_deleted INTEGER);"
            "CREATE TABLE runs(id INTEGER PRIMARY KEY,period_id INTEGER,workflow_id TEXT,workflow_label TEXT,start_date TEXT,end_date TEXT,state TEXT,created_at TEXT,started_at TEXT,finished_at TEXT,exit_code INTEGER,error_summary TEXT,log_path TEXT);"
            "CREATE TABLE artifacts(id INTEGER PRIMARY KEY,run_id INTEGER,name TEXT,source_path TEXT,stored_path TEXT,size_bytes INTEGER);"
        )
        connection.execute("INSERT INTO report_periods VALUES(1,'2026-10-01','2026-10-07','created','updated',0)")
        connection.execute("INSERT INTO runs VALUES(1,1,'ecm-step-0','ECM Step 0','2026-10-01','2026-10-07','succeeded','created',NULL,'finished',0,NULL,?)", (str(log_path),))
        connection.execute("INSERT INTO artifacts VALUES(1,1,'report.csv','source',?,?)", (str(artifact_path), artifact_path.stat().st_size))
    store = ArchiveStore(database, runs_dir)
    assert store.is_imported()
    assert [run.id for run in store.list_runs()] == [1]
    assert len(store.list_periods()) == 1
    assert store.read_log(1) == "archived evidence"
    with pytest.raises(sqlite3.OperationalError):
        with store._connect() as connection:
            connection.execute("DELETE FROM runs")

    outside = tmp_path / "outside.log"
    outside.write_text("not archive", encoding="utf-8")
    with sqlite3.connect(database) as connection:
        connection.execute("UPDATE runs SET log_path=? WHERE id=1", (str(outside),))
    with pytest.raises(ValueError, match="escapes"):
        store.read_log(1)


def test_archive_routes_are_separate_from_active_v3_history(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    from fastapi.testclient import TestClient

    from control_panel_v3 import main as app_module

    settings = load_settings()
    route_root = settings.project_root / "control_panel_v3" / "tests" / f"archive-route-{tmp_path.name}"
    roots = {name: route_root / name for name in (
        "data_dir", "archive_dir", "csv_dir", "ecm_output_dir", "smm_output_dir", "django_output_dir", "edge_profile_dir"
    )}
    import shutil
    shutil.rmtree(route_root, ignore_errors=True)
    test_settings = type(settings)(**{**settings.__dict__, **roots})
    test_settings.ensure_application_directories()
    monkeypatch.setattr(app_module, "import_v2_archive", lambda _source, destination: _make_route_test_archive(destination))
    monkeypatch.setattr(app_module, "load_settings", lambda: test_settings)
    app = app_module.create_app()
    active = app.state.history.create_run("ecm-step-0", "V3 Step 0", "2026-10-01", "2026-10-07")
    client = TestClient(app)
    assert client.get("/archive").status_code == 200
    assert client.get("/archive").text.find("No V2 archive") != -1
    assert client.get("/runs/1").status_code == 200
    assert client.post("/archive/import", follow_redirects=False).status_code == 303
    response = client.get("/archive")
    assert response.status_code == 200
    assert "Archived Run History" in response.text
    assert "Archived Step 0" in response.text
    assert "V3 Step 0" not in response.text
    assert client.get(f"/runs/{active.id}").status_code == 200
    archived_detail = client.get("/archive/runs/1")
    assert archived_detail.status_code == 200
    assert "test archive log" in archived_detail.text
    artifact_response = client.get("/archive/runs/1/artifacts/1")
    assert artifact_response.status_code == 200
    assert artifact_response.text.replace("\r\n", "\n") == "a\n1\n"
    repeat_import = client.post("/archive/import", follow_redirects=False)
    assert repeat_import.status_code == 303
    assert "not+empty" in repeat_import.headers["location"]

    shutil.rmtree(route_root, ignore_errors=True)


def _make_route_test_archive(destination: Path) -> None:
    if destination.exists() and any(destination.iterdir()):
        raise ValueError("V3 archive destination is not empty.")
    destination.mkdir(parents=True, exist_ok=True)
    runs = destination / "runs" / "1"
    runs.mkdir(parents=True)
    log = runs / "run.log"
    artifact = runs / "evidence.csv"
    log.write_text("test archive log", encoding="utf-8")
    artifact.write_text("a\n1\n", encoding="utf-8")
    with sqlite3.connect(destination / "history.sqlite3") as connection:
        connection.executescript(
            "CREATE TABLE report_periods(id INTEGER PRIMARY KEY,start_date TEXT,end_date TEXT,created_at TEXT,updated_at TEXT,is_deleted INTEGER);"
            "CREATE TABLE runs(id INTEGER PRIMARY KEY,period_id INTEGER,workflow_id TEXT,workflow_label TEXT,start_date TEXT,end_date TEXT,state TEXT,created_at TEXT,started_at TEXT,finished_at TEXT,exit_code INTEGER,error_summary TEXT,log_path TEXT);"
            "CREATE TABLE artifacts(id INTEGER PRIMARY KEY,run_id INTEGER,name TEXT,source_path TEXT,stored_path TEXT,size_bytes INTEGER);"
        )
        connection.execute("INSERT INTO report_periods VALUES(1,'2026-10-01','2026-10-07','created','updated',0)")
        connection.execute("INSERT INTO runs VALUES(1,1,'ecm-step-0','Archived Step 0','2026-10-01','2026-10-07','succeeded','2026-10-08',NULL,'2026-10-08',0,NULL,?)", (str(log),))
        connection.execute("INSERT INTO artifacts VALUES(1,1,'evidence.csv','source',?,?)", (str(artifact), artifact.stat().st_size))


def test_v3_modules_do_not_import_v1_or_v2_runtime(tmp_path: Path) -> None:
    root = CONTROL_PANEL_ROOT
    violations: list[str] = []
    for path in root.rglob("*.py"):
        if "tests" in path.parts:
            continue
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names = [alias.name for alias in node.names]
            elif isinstance(node, ast.ImportFrom):
                names = [node.module or ""]
            else:
                continue
            if any(name.startswith("control_panel.services") or name.startswith("control_panel_v2") for name in names):
                violations.append(f"{path}: {names}")
    assert violations == []
