"""Safely seed V2 history from a V1 data snapshot."""

from __future__ import annotations

import argparse
import shutil
import sqlite3
from pathlib import Path

from .config import load_settings


def import_v1_history(source_data_dir: Path, destination_data_dir: Path) -> None:
    source = source_data_dir.resolve()
    destination = destination_data_dir.resolve()
    source_database = source / "history.sqlite3"
    if not source_database.is_file():
        raise ValueError(f"V1 history database was not found: {source_database}")
    if destination == source:
        raise ValueError("V2 data directory must differ from the V1 data directory.")
    if destination.exists() and any(destination.iterdir()):
        existing_database = destination / "history.sqlite3"
        existing_runs = destination / "runs"
        permitted_empty_runtime = (
            existing_database.is_file()
            and existing_runs.is_dir()
            and not any(existing_runs.iterdir())
        )
        if not permitted_empty_runtime:
            raise ValueError(f"V2 data directory is not empty: {destination}")
        with sqlite3.connect(existing_database) as connection:
            existing_run_count = connection.execute("SELECT COUNT(*) FROM runs").fetchone()[0]
            existing_period_count = connection.execute(
                "SELECT COUNT(*) FROM report_periods"
            ).fetchone()[0]
        if existing_run_count or existing_period_count:
            raise ValueError(f"V2 data directory is not empty: {destination}")
        shutil.rmtree(destination)
    destination.mkdir(parents=True, exist_ok=True)
    destination_database = destination / "history.sqlite3"
    try:
        with sqlite3.connect(source_database) as source_connection, sqlite3.connect(destination_database) as destination_connection:
            source_connection.backup(destination_connection)
        source_runs = source / "runs"
        destination_runs = destination / "runs"
        if source_runs.is_dir():
            shutil.copytree(source_runs, destination_runs)
        else:
            destination_runs.mkdir()
        with sqlite3.connect(destination_database) as connection:
            run_rows = connection.execute("SELECT id, log_path FROM runs").fetchall()
            for run_id, log_path in run_rows:
                old_log_path = Path(log_path)
                if not old_log_path.is_absolute():
                    old_log_path = source / old_log_path
                old_log_path = old_log_path.resolve()
                try:
                    relative = old_log_path.relative_to(source_runs.resolve())
                except ValueError as error:
                    raise ValueError(f"Run {run_id} log points outside V1 runs: {old_log_path}") from error
                new_log_path = (destination_runs / relative).resolve()
                if not new_log_path.is_file():
                    raise ValueError(f"Copied run {run_id} log was not found: {new_log_path}")
                connection.execute("UPDATE runs SET log_path = ? WHERE id = ?", (str(new_log_path), run_id))
            rows = connection.execute("SELECT id, stored_path FROM artifacts").fetchall()
            for artifact_id, stored_path in rows:
                old_path = Path(stored_path).resolve()
                try:
                    relative = old_path.relative_to(source_runs.resolve())
                except ValueError as error:
                    raise ValueError(f"Artifact {artifact_id} points outside V1 runs: {old_path}") from error
                new_path = (destination_runs / relative).resolve()
                if not new_path.is_file():
                    raise ValueError(f"Copied artifact {artifact_id} was not found: {new_path}")
                connection.execute("UPDATE artifacts SET stored_path = ? WHERE id = ?", (str(new_path), artifact_id))
    except Exception:
        shutil.rmtree(destination, ignore_errors=True)
        raise


def main() -> None:
    parser = argparse.ArgumentParser(description="Copy V1 control-panel history into an empty V2 data root.")
    parser.add_argument("--source-data-dir", type=Path, default=Path("control_panel/data"))
    args = parser.parse_args()
    settings = load_settings()
    import_v1_history(args.source_data_dir, settings.data_dir)
    print(f"Imported V1 history into {settings.data_dir}")


if __name__ == "__main__":
    main()
