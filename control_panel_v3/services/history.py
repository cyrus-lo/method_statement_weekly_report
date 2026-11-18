"""V3-owned SQLite period, run, artifact, and archive history."""

from __future__ import annotations

import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


def _timestamp() -> str:
    return datetime.now(ZoneInfo("Asia/Hong_Kong")).isoformat(timespec="seconds")


@dataclass(frozen=True)
class RunRecord:
    id: int
    period_id: int | None
    workflow_id: str
    workflow_label: str
    start_date: str
    end_date: str
    state: str
    created_at: str
    started_at: str | None
    finished_at: str | None
    exit_code: int | None
    error_summary: str | None
    log_path: Path
    manual_step_1b_bypass: bool = False
    lookup_scope: str | None = None

    @property
    def is_active(self) -> bool:
        return self.state in {"queued", "running"}


@dataclass(frozen=True)
class ReportPeriodRecord:
    id: int
    start_date: str
    end_date: str
    created_at: str
    updated_at: str
    is_deleted: bool


@dataclass(frozen=True)
class ReportPeriodSummary:
    period: ReportPeriodRecord
    runs: dict[str, RunRecord | None]
    last_updated_at: str

    @property
    def is_locked(self) -> bool:
        return any(record is not None for record in self.runs.values())


@dataclass(frozen=True)
class ArtifactRecord:
    id: int
    run_id: int
    name: str
    source_path: str
    stored_path: Path
    size_bytes: int


class HistoryStore:
    def __init__(self, database_path: Path, runs_dir: Path) -> None:
        self.database_path = database_path
        self.runs_dir = runs_dir

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        connection = sqlite3.connect(self.database_path, timeout=30)
        try:
            connection.row_factory = sqlite3.Row
            connection.execute("PRAGMA foreign_keys = ON")
            with connection:
                yield connection
        finally:
            connection.close()

    def initialize(self) -> None:
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self.runs_dir.mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS report_periods (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    start_date TEXT NOT NULL,
                    end_date TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    is_deleted INTEGER NOT NULL DEFAULT 0
                );
                CREATE UNIQUE INDEX IF NOT EXISTS idx_period_active_dates
                    ON report_periods(start_date, end_date) WHERE is_deleted = 0;
                CREATE TABLE IF NOT EXISTS runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    period_id INTEGER REFERENCES report_periods(id),
                    workflow_id TEXT NOT NULL,
                    workflow_label TEXT NOT NULL,
                    start_date TEXT NOT NULL,
                    end_date TEXT NOT NULL,
                    state TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    started_at TEXT,
                    finished_at TEXT,
                    exit_code INTEGER,
                    error_summary TEXT,
                    log_path TEXT NOT NULL,
                    manual_step_1b_bypass INTEGER NOT NULL DEFAULT 0
                    ,lookup_scope TEXT
                );
                CREATE TABLE IF NOT EXISTS artifacts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                    name TEXT NOT NULL,
                    source_path TEXT NOT NULL,
                    stored_path TEXT NOT NULL,
                    size_bytes INTEGER NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_runs_created_at ON runs(created_at DESC);
                CREATE INDEX IF NOT EXISTS idx_runs_period_workflow ON runs(period_id, workflow_id, id DESC);
                CREATE INDEX IF NOT EXISTS idx_artifacts_run ON artifacts(run_id);
                """
            )
            columns = {row["name"] for row in connection.execute("PRAGMA table_info(runs)")}
            if "manual_step_1b_bypass" not in columns:
                connection.execute("ALTER TABLE runs ADD COLUMN manual_step_1b_bypass INTEGER NOT NULL DEFAULT 0")
            if "lookup_scope" not in columns:
                connection.execute("ALTER TABLE runs ADD COLUMN lookup_scope TEXT")

    @staticmethod
    def _run(row: sqlite3.Row) -> RunRecord:
        return RunRecord(
            id=row["id"], period_id=row["period_id"], workflow_id=row["workflow_id"],
            workflow_label=row["workflow_label"], start_date=row["start_date"], end_date=row["end_date"],
            state=row["state"], created_at=row["created_at"], started_at=row["started_at"],
            finished_at=row["finished_at"], exit_code=row["exit_code"], error_summary=row["error_summary"],
            log_path=Path(row["log_path"]), manual_step_1b_bypass=bool(row["manual_step_1b_bypass"]), lookup_scope=row["lookup_scope"] if "lookup_scope" in row.keys() else None,
        )

    @staticmethod
    def _period(row: sqlite3.Row) -> ReportPeriodRecord:
        return ReportPeriodRecord(
            row["id"], row["start_date"], row["end_date"], row["created_at"], row["updated_at"], bool(row["is_deleted"])
        )

    def create_report_period(self, start_date: str, end_date: str) -> ReportPeriodRecord:
        now = _timestamp()
        try:
            with self._connect() as connection:
                cursor = connection.execute(
                    "INSERT INTO report_periods(start_date,end_date,created_at,updated_at) VALUES(?,?,?,?)",
                    (start_date, end_date, now, now),
                )
                period_id = int(cursor.lastrowid)
        except sqlite3.IntegrityError as error:
            raise ValueError("This report period already exists.") from error
        return self.get_report_period(period_id)

    def get_report_period(self, period_id: int) -> ReportPeriodRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM report_periods WHERE id=?", (period_id,)).fetchone()
        if row is None:
            raise KeyError(f"Report period {period_id} does not exist.")
        return self._period(row)

    def update_report_period(self, period_id: int, start_date: str, end_date: str) -> ReportPeriodRecord:
        now = _timestamp()
        try:
            with self._connect() as connection:
                period = connection.execute("SELECT * FROM report_periods WHERE id=?", (period_id,)).fetchone()
                if period is None:
                    raise KeyError(f"Report period {period_id} does not exist.")
                if period["is_deleted"]:
                    raise ValueError("This report period has been deleted and cannot be updated.")
                if connection.execute("SELECT 1 FROM runs WHERE period_id=? LIMIT 1", (period_id,)).fetchone():
                    raise ValueError("This period is locked because a workflow has already started.")
                connection.execute(
                    "UPDATE report_periods SET start_date=?, end_date=?, updated_at=? WHERE id=?",
                    (start_date, end_date, now, period_id),
                )
        except sqlite3.IntegrityError as error:
            raise ValueError("This report period already exists.") from error
        return self.get_report_period(period_id)

    def delete_report_period(self, period_id: int) -> ReportPeriodRecord:
        with self._connect() as connection:
            period = connection.execute("SELECT * FROM report_periods WHERE id=?", (period_id,)).fetchone()
            if period is None:
                raise KeyError(f"Report period {period_id} does not exist.")
            if period["is_deleted"]:
                raise ValueError("This report period has already been deleted.")
            if connection.execute(
                "SELECT 1 FROM runs WHERE period_id=? AND state IN ('queued','running') LIMIT 1", (period_id,)
            ).fetchone():
                raise ValueError("This period cannot be deleted while a workflow is running.")
            connection.execute("UPDATE report_periods SET is_deleted=1,updated_at=? WHERE id=?", (_timestamp(), period_id))
        return self.get_report_period(period_id)

    def restore_report_period(self, period_id: int) -> ReportPeriodRecord:
        try:
            with self._connect() as connection:
                period = connection.execute("SELECT * FROM report_periods WHERE id=?", (period_id,)).fetchone()
                if period is None:
                    raise KeyError(f"Report period {period_id} does not exist.")
                if not period["is_deleted"]:
                    raise ValueError("This report period is not deleted.")
                connection.execute("UPDATE report_periods SET is_deleted=0,updated_at=? WHERE id=?", (_timestamp(), period_id))
        except sqlite3.IntegrityError as error:
            raise ValueError("An active report period already uses these dates.") from error
        return self.get_report_period(period_id)

    def list_report_period_summaries(self, workflow_ids: tuple[str, ...], *, is_deleted: bool = False) -> list[ReportPeriodSummary]:
        with self._connect() as connection:
            periods = connection.execute("SELECT * FROM report_periods WHERE is_deleted=? ORDER BY start_date DESC,id DESC", (int(is_deleted),)).fetchall()
            runs = connection.execute("SELECT * FROM runs ORDER BY period_id,workflow_id,id DESC").fetchall()
        newest: dict[int, dict[str, RunRecord]] = {}
        for row in runs:
            run = self._run(row)
            if run.period_id is not None and run.workflow_id in workflow_ids:
                existing = newest.setdefault(run.period_id, {}).get(run.workflow_id)
                if existing is None or run.id > existing.id:
                    newest[run.period_id][run.workflow_id] = run
        summaries = []
        for row in periods:
            period = self._period(row)
            by_workflow = {workflow: newest.get(period.id, {}).get(workflow) for workflow in workflow_ids}
            timestamps = [period.updated_at]
            for run in by_workflow.values():
                if run:
                    timestamps.extend(item for item in (run.created_at, run.started_at, run.finished_at) if item)
            summaries.append(ReportPeriodSummary(period, by_workflow, max(timestamps)))
        return summaries

    def create_run(
        self,
        workflow_id: str,
        workflow_label: str,
        start_date: str,
        end_date: str,
        *,
        period_id: int | None = None,
        manual_step_1b_bypass: bool = False,
        lookup_scope: str | None = None,
    ) -> RunRecord:
        if manual_step_1b_bypass and workflow_id != "ecm-step-2":
            raise ValueError("Manual Step 1b bypass metadata is valid only for ECM Step 2 runs.")
        if lookup_scope is not None and (workflow_id != "ecm-first-date-lookup" or lookup_scope not in {"all", "blank", "previously-unmatched"}):
            raise ValueError("Lookup scope metadata is valid only for a supported first-date lookup run.")
        if period_id is not None:
            period = self.get_report_period(period_id)
            if period.is_deleted:
                raise ValueError("This report period has been deleted and cannot start a workflow.")
            if (start_date, end_date) != (period.start_date, period.end_date):
                raise ValueError("Workflow dates must match the selected report period.")
        now = _timestamp()
        with self._connect() as connection:
            cursor = connection.execute(
                "INSERT INTO runs(period_id,workflow_id,workflow_label,start_date,end_date,state,created_at,log_path,manual_step_1b_bypass,lookup_scope) VALUES(?,?,?,?,?,'queued',?,'',?,?)",
                (period_id, workflow_id, workflow_label, start_date, end_date, now, int(manual_step_1b_bypass), lookup_scope),
            )
            run_id = int(cursor.lastrowid)
            log_path = (self.runs_dir / str(run_id) / "run.log").resolve()
            log_path.parent.mkdir(parents=True, exist_ok=True)
            log_path.touch(exist_ok=True)
            connection.execute("UPDATE runs SET log_path=? WHERE id=?", (str(log_path), run_id))
        return self.get_run(run_id)

    def get_run(self, run_id: int) -> RunRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM runs WHERE id=?", (run_id,)).fetchone()
        if row is None:
            raise KeyError(f"Run {run_id} does not exist.")
        return self._run(row)

    def list_runs(self, *, workflow_id: str | None = None, state: str | None = None, limit: int = 100, period_id: int | None = None) -> list[RunRecord]:
        clauses, values = [], []
        if workflow_id:
            clauses.append("workflow_id=?"); values.append(workflow_id)
        if state:
            clauses.append("state=?"); values.append(state)
        if period_id is not None:
            clauses.append("period_id=?"); values.append(period_id)
        where = " WHERE " + " AND ".join(clauses) if clauses else ""
        with self._connect() as connection:
            rows = connection.execute(f"SELECT * FROM runs{where} ORDER BY id DESC LIMIT ?", (*values, limit)).fetchall()
        return [self._run(row) for row in rows]

    def active_run(self) -> RunRecord | None:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM runs WHERE state IN ('queued','running') ORDER BY id DESC LIMIT 1").fetchone()
        return self._run(row) if row else None

    def fail_interrupted_runs(self) -> None:
        with self._connect() as connection:
            connection.execute(
                "UPDATE runs SET state='failed',finished_at=?,error_summary='The V3 control panel stopped before this run completed.' WHERE state IN ('queued','running')",
                (_timestamp(),),
            )

    def mark_running(self, run_id: int) -> None:
        with self._connect() as connection:
            connection.execute("UPDATE runs SET state='running',started_at=? WHERE id=?", (_timestamp(), run_id))

    def mark_finished(self, run_id: int, *, state: str, exit_code: int | None, error_summary: str | None = None) -> None:
        with self._connect() as connection:
            connection.execute("UPDATE runs SET state=?,finished_at=?,exit_code=?,error_summary=? WHERE id=?", (state, _timestamp(), exit_code, error_summary, run_id))

    def append_log(self, run_id: int, text: str) -> None:
        with self.get_run(run_id).log_path.open("a", encoding="utf-8", newline="") as stream:
            stream.write(text)
            if text and not text.endswith("\n"):
                stream.write("\n")

    def read_log(self, run_id: int, max_characters: int = 60_000) -> str:
        path = self.get_run(run_id).log_path
        if not path.is_file():
            return ""
        with path.open("r", encoding="utf-8", errors="replace") as stream:
            stream.seek(max(0, path.stat().st_size - max_characters))
            return stream.read(max_characters)

    def add_artifact(self, run_id: int, name: str, source_path: Path, stored_path: Path) -> ArtifactRecord:
        if not stored_path.is_file():
            raise FileNotFoundError(f"Artifact was not found: {stored_path}")
        with self._connect() as connection:
            cursor = connection.execute(
                "INSERT INTO artifacts(run_id,name,source_path,stored_path,size_bytes) VALUES(?,?,?,?,?)",
                (run_id, name, str(source_path), str(stored_path.resolve()), stored_path.stat().st_size),
            )
            artifact_id = int(cursor.lastrowid)
        return self.get_artifact(run_id, artifact_id)

    def get_artifact(self, run_id: int, artifact_id: int) -> ArtifactRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM artifacts WHERE id=? AND run_id=?", (artifact_id, run_id)).fetchone()
        if row is None:
            raise KeyError(f"Artifact {artifact_id} does not exist for run {run_id}.")
        return ArtifactRecord(row["id"], row["run_id"], row["name"], row["source_path"], Path(row["stored_path"]), row["size_bytes"])

    def list_artifacts(self, run_id: int) -> list[ArtifactRecord]:
        with self._connect() as connection:
            rows = connection.execute("SELECT * FROM artifacts WHERE run_id=? ORDER BY id", (run_id,)).fetchall()
        return [ArtifactRecord(row["id"], row["run_id"], row["name"], row["source_path"], Path(row["stored_path"]), row["size_bytes"]) for row in rows]
