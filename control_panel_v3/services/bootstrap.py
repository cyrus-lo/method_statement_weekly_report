"""One-time, read-only import of V2 run history into a separate V3 archive."""

from __future__ import annotations

import shutil
import sqlite3
from pathlib import Path

from .config import CONTROL_PANEL_ROOT, V2_ROOT


def _contained(path: Path, root: Path) -> Path:
    resolved = path.resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError as error:
        raise ValueError(f"Path escapes its approved root: {resolved}") from error
    return resolved


def import_v2_archive(source_data_dir: Path, destination_archive_dir: Path) -> None:
    """Copy V2 DB and run evidence; do not mutate the V2 source or active V3 DB."""
    source = _contained(source_data_dir, V2_ROOT)
    destination = _contained(destination_archive_dir, CONTROL_PANEL_ROOT)
    if destination == source or source == destination or destination.is_relative_to(source):
        raise ValueError("V3 archive destination must be separate from V2 source data.")
    source_db = source / "history.sqlite3"
    source_runs = source / "runs"
    if not source_db.is_file():
        raise ValueError(f"V2 history database was not found: {source_db}")
    if destination.exists() and any(destination.iterdir()):
        raise ValueError(f"V3 archive destination is not empty: {destination}")
    destination.mkdir(parents=True, exist_ok=True)
    target_db = destination / "history.sqlite3"
    target_runs = destination / "runs"
    try:
        with sqlite3.connect(f"file:{source_db.as_posix()}?mode=ro", uri=True) as source_connection:
            with sqlite3.connect(target_db) as target_connection:
                source_connection.backup(target_connection)
        target_runs.mkdir()
        if source_runs.is_dir():
            shutil.copytree(source_runs, target_runs, dirs_exist_ok=True)
        with sqlite3.connect(target_db) as connection:
            connection.row_factory = sqlite3.Row
            tables = {row["name"] for row in connection.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            if "runs" not in tables or "artifacts" not in tables:
                raise ValueError("V2 archive database does not contain expected runs and artifacts tables.")
            for row in connection.execute("SELECT id, log_path FROM runs").fetchall():
                old = Path(row["log_path"])
                if not old.is_absolute():
                    old = source / old
                old = _contained(old, source_runs)
                relative = old.relative_to(source_runs.resolve())
                new = _contained(target_runs / relative, target_runs)
                if not new.is_file():
                    raise ValueError(f"Copied V2 run log is missing: {new}")
                connection.execute("UPDATE runs SET log_path=? WHERE id=?", (str(new), row["id"]))
            for row in connection.execute("SELECT id, stored_path FROM artifacts").fetchall():
                old = _contained(Path(row["stored_path"]), source_runs)
                new = _contained(target_runs / old.relative_to(source_runs.resolve()), target_runs)
                if not new.is_file():
                    raise ValueError(f"Copied V2 artifact is missing: {new}")
                connection.execute("UPDATE artifacts SET stored_path=? WHERE id=?", (str(new), row["id"]))
    except Exception:
        shutil.rmtree(destination, ignore_errors=True)
        raise
