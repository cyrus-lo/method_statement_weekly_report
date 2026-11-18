from __future__ import annotations

import sqlite3
from pathlib import Path

from control_panel.services.history import HistoryStore
from control_panel_v2.services.bootstrap import import_v1_history


def test_import_remaps_run_logs_into_v2_runs(tmp_path: Path) -> None:
    source = tmp_path / "v1"
    destination = tmp_path / "v2"
    source_runs = source / "runs" / "1"
    source_runs.mkdir(parents=True)
    source_log = source_runs / "run.log"
    source_log.write_text("imported log", encoding="utf-8")

    source_history = HistoryStore(source / "history.sqlite3", source / "runs")
    source_history.initialize()
    source_history.create_run(
        "weekly-report", "Weekly report", "2026-07-20", "2026-07-26"
    )
    with sqlite3.connect(source / "history.sqlite3") as connection:
        connection.execute("UPDATE runs SET log_path = ? WHERE id = 1", (str(source_log),))

    import_v1_history(source, destination)

    imported_history = HistoryStore(destination / "history.sqlite3", destination / "runs")
    imported_history.initialize()
    imported_run = imported_history.get_run(1)

    assert imported_run.log_path == (destination / "runs" / "1" / "run.log").resolve()
    assert imported_history.read_log(1) == "imported log"