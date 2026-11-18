"""Read-only access to the imported V2 history archive."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path

from .history import ArtifactRecord, RunRecord


@dataclass(frozen=True)
class ArchivedPeriodRecord:
    id: int
    start_date: str
    end_date: str
    created_at: str
    updated_at: str
    is_deleted: bool


class ArchiveStore:
    """Read an imported V2 archive without opening its SQLite DB for writes."""

    def __init__(self, database_path: Path, runs_dir: Path) -> None:
        self.database_path = database_path.resolve()
        self.runs_dir = runs_dir.resolve()

    def _connect(self) -> sqlite3.Connection:
        if not self.database_path.is_file():
            raise FileNotFoundError(f"V2 archive has not been imported: {self.database_path}")
        connection = sqlite3.connect(f"file:{self.database_path.as_posix()}?mode=ro", uri=True)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA query_only = ON")
        return connection

    @staticmethod
    def _run(row: sqlite3.Row) -> RunRecord:
        return RunRecord(
            id=row["id"], period_id=row["period_id"], workflow_id=row["workflow_id"],
            workflow_label=row["workflow_label"], start_date=row["start_date"], end_date=row["end_date"],
            state=row["state"], created_at=row["created_at"], started_at=row["started_at"],
            finished_at=row["finished_at"], exit_code=row["exit_code"], error_summary=row["error_summary"],
            log_path=Path(row["log_path"]),
            manual_step_1b_bypass=bool(row["manual_step_1b_bypass"]) if "manual_step_1b_bypass" in row.keys() else False,
        )

    def is_imported(self) -> bool:
        return self.database_path.is_file()

    def list_runs(self, *, limit: int = 250) -> list[RunRecord]:
        with self._connect() as connection:
            return [self._run(row) for row in connection.execute("SELECT * FROM runs ORDER BY id DESC LIMIT ?", (limit,))]

    def get_run(self, run_id: int) -> RunRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM runs WHERE id=?", (run_id,)).fetchone()
        if row is None:
            raise KeyError(f"Archived run {run_id} does not exist.")
        return self._run(row)

    def list_artifacts(self, run_id: int) -> list[ArtifactRecord]:
        self.get_run(run_id)
        with self._connect() as connection:
            rows = connection.execute("SELECT * FROM artifacts WHERE run_id=? ORDER BY id", (run_id,)).fetchall()
        return [ArtifactRecord(row["id"], row["run_id"], row["name"], row["source_path"], Path(row["stored_path"]), row["size_bytes"]) for row in rows]

    def get_artifact(self, run_id: int, artifact_id: int) -> ArtifactRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM artifacts WHERE id=? AND run_id=?", (artifact_id, run_id)).fetchone()
        if row is None:
            raise KeyError(f"Archived artifact {artifact_id} does not exist for run {run_id}.")
        return ArtifactRecord(row["id"], row["run_id"], row["name"], row["source_path"], Path(row["stored_path"]), row["size_bytes"])

    def read_log(self, run_id: int, max_characters: int = 60_000) -> str:
        path = self.get_run(run_id).log_path.resolve()
        try:
            path.relative_to(self.runs_dir)
        except ValueError as error:
            raise ValueError("Archived log path escapes the V3 archive run root.") from error
        if not path.is_file():
            return ""
        with path.open("r", encoding="utf-8", errors="replace") as stream:
            stream.seek(max(0, path.stat().st_size - max_characters))
            return stream.read(max_characters)

    def list_periods(self) -> list[ArchivedPeriodRecord]:
        with self._connect() as connection:
            rows = connection.execute("SELECT * FROM report_periods ORDER BY start_date DESC,id DESC").fetchall()
        return [ArchivedPeriodRecord(row["id"], row["start_date"], row["end_date"], row["created_at"], row["updated_at"], bool(row["is_deleted"])) for row in rows]
