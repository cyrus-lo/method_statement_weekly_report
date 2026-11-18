"""SQLite-backed run history and immutable artifact metadata."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from zoneinfo import ZoneInfo


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
        return any(run is not None for run in self.runs.values())


@dataclass(frozen=True)
class ArtifactRecord:
    id: int
    run_id: int
    name: str
    source_path: str
    stored_path: Path
    size_bytes: int


def _timestamp() -> str:
    return datetime.now(ZoneInfo("Asia/Hong_Kong")).isoformat(timespec="seconds")


def _hong_kong_timestamp(value: str) -> str:
    """Return a canonical timezone-aware timestamp in Hong Kong time."""
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZoneInfo("UTC"))
    return parsed.astimezone(ZoneInfo("Asia/Hong_Kong")).isoformat(timespec="seconds")


class HistoryStore:
    """A small repository that opens a short-lived SQLite connection per operation."""

    def __init__(self, database_path: Path, runs_dir: Path) -> None:
        self.database_path = database_path
        self.runs_dir = runs_dir

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database_path)
        connection.row_factory = sqlite3.Row
        return connection

    def initialize(self) -> None:
        self.database_path.parent.mkdir(parents=True, exist_ok=True)
        self.runs_dir.mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            connection.executescript(
                """
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
                    log_path TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS report_periods (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    start_date TEXT NOT NULL,
                    end_date TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    is_deleted INTEGER NOT NULL DEFAULT 0
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
                CREATE INDEX IF NOT EXISTS idx_artifacts_run_id ON artifacts(run_id);
                """
            )
            self._migrate_period_id(connection)
            self._migrate_report_periods(connection)
            connection.execute(
                """CREATE INDEX IF NOT EXISTS idx_runs_period_workflow_id
                ON runs(period_id, workflow_id, id DESC)"""
            )
            self._migrate_timestamps(connection)

    @staticmethod
    def _migrate_period_id(connection: sqlite3.Connection) -> None:
        columns = {
            row["name"] for row in connection.execute("PRAGMA table_info(runs)").fetchall()
        }
        if "period_id" not in columns:
            connection.execute("ALTER TABLE runs ADD COLUMN period_id INTEGER")

    @staticmethod
    def _migrate_report_periods(connection: sqlite3.Connection) -> None:
        """Remove the legacy date-range uniqueness constraint and add soft deletion.

        The original table enforced one row per date range. Rebuilding all three
        related tables preserves their IDs and foreign-key definitions while
        allowing a replacement period to be created after a soft deletion.
        """
        columns = {
            row["name"]
            for row in connection.execute("PRAGMA table_info(report_periods)").fetchall()
        }
        if "is_deleted" in columns:
            return

        connection.executescript(
            """
            ALTER TABLE artifacts RENAME TO artifacts_legacy;
            ALTER TABLE runs RENAME TO runs_legacy;
            ALTER TABLE report_periods RENAME TO report_periods_legacy;

            CREATE TABLE report_periods (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                is_deleted INTEGER NOT NULL DEFAULT 0
            );
            CREATE TABLE runs (
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
                log_path TEXT NOT NULL
            );
            CREATE TABLE artifacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL REFERENCES runs(id) ON DELETE CASCADE,
                name TEXT NOT NULL,
                source_path TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                size_bytes INTEGER NOT NULL
            );

            INSERT INTO report_periods (
                id, start_date, end_date, created_at, updated_at, is_deleted
            ) SELECT id, start_date, end_date, created_at, updated_at, 0
            FROM report_periods_legacy;
            INSERT INTO runs (
                id, period_id, workflow_id, workflow_label, start_date, end_date, state,
                created_at, started_at, finished_at, exit_code, error_summary, log_path
            ) SELECT
                id, period_id, workflow_id, workflow_label, start_date, end_date, state,
                created_at, started_at, finished_at, exit_code, error_summary, log_path
            FROM runs_legacy;
            INSERT INTO artifacts (id, run_id, name, source_path, stored_path, size_bytes)
            SELECT id, run_id, name, source_path, stored_path, size_bytes
            FROM artifacts_legacy;

            DROP TABLE artifacts_legacy;
            DROP TABLE runs_legacy;
            DROP TABLE report_periods_legacy;
            """
        )

    @staticmethod
    def _migrate_timestamps(connection: sqlite3.Connection) -> None:
        """Normalize existing run timestamps to Hong Kong time.

        The conversion is idempotent and keeps the stored values timezone-aware,
        so existing UTC records can be migrated without changing their instant.
        """
        rows = connection.execute(
            "SELECT id, created_at, started_at, finished_at FROM runs"
        ).fetchall()
        for row in rows:
            converted = {
                "created_at": _hong_kong_timestamp(row["created_at"]),
                "started_at": (
                    _hong_kong_timestamp(row["started_at"]) if row["started_at"] else None
                ),
                "finished_at": (
                    _hong_kong_timestamp(row["finished_at"]) if row["finished_at"] else None
                ),
            }
            connection.execute(
                """UPDATE runs
                SET created_at = ?, started_at = ?, finished_at = ?
                WHERE id = ?""",
                (
                    converted["created_at"],
                    converted["started_at"],
                    converted["finished_at"],
                    row["id"],
                ),
            )

    @staticmethod
    def _run_from_row(row: sqlite3.Row) -> RunRecord:
        return RunRecord(
            id=row["id"],
            period_id=row["period_id"],
            workflow_id=row["workflow_id"],
            workflow_label=row["workflow_label"],
            start_date=row["start_date"],
            end_date=row["end_date"],
            state=row["state"],
            created_at=row["created_at"],
            started_at=row["started_at"],
            finished_at=row["finished_at"],
            exit_code=row["exit_code"],
            error_summary=row["error_summary"],
            log_path=Path(row["log_path"]),
        )

    @staticmethod
    def _period_from_row(row: sqlite3.Row) -> ReportPeriodRecord:
        return ReportPeriodRecord(
            id=row["id"],
            start_date=row["start_date"],
            end_date=row["end_date"],
            created_at=row["created_at"],
            updated_at=row["updated_at"],
            is_deleted=bool(row["is_deleted"]),
        )

    def create_report_period(self, start_date: str, end_date: str) -> ReportPeriodRecord:
        timestamp = _timestamp()
        with self._connect() as connection:
            existing = connection.execute(
                """SELECT 1 FROM report_periods
                WHERE start_date = ? AND end_date = ? AND is_deleted = 0""",
                (start_date, end_date),
            ).fetchone()
            if existing is not None:
                raise ValueError("This report period already exists.")
            cursor = connection.execute(
                """INSERT INTO report_periods (
                    start_date, end_date, created_at, updated_at
                ) VALUES (?, ?, ?, ?)""",
                (start_date, end_date, timestamp, timestamp),
            )
            period_id = int(cursor.lastrowid)
        return self.get_report_period(period_id)

    def get_report_period(self, period_id: int) -> ReportPeriodRecord:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM report_periods WHERE id = ?", (period_id,)
            ).fetchone()
        if row is None:
            raise KeyError(f"Report period {period_id} does not exist.")
        return self._period_from_row(row)

    def update_report_period(
        self, period_id: int, start_date: str, end_date: str
    ) -> ReportPeriodRecord:
        with self._connect() as connection:
            existing = connection.execute(
                "SELECT id, is_deleted FROM report_periods WHERE id = ?", (period_id,)
            ).fetchone()
            if existing is None:
                raise KeyError(f"Report period {period_id} does not exist.")
            if existing["is_deleted"]:
                raise ValueError("This report period has been deleted and cannot be updated.")
            has_runs = connection.execute(
                "SELECT 1 FROM runs WHERE period_id = ? LIMIT 1", (period_id,)
            ).fetchone()
            if has_runs is not None:
                raise ValueError("This period is locked because a workflow has already started.")
            duplicate = connection.execute(
                """SELECT 1 FROM report_periods
                WHERE start_date = ? AND end_date = ? AND is_deleted = 0 AND id != ?""",
                (start_date, end_date, period_id),
            ).fetchone()
            if duplicate is not None:
                raise ValueError("This report period already exists.")
            connection.execute(
                """UPDATE report_periods
                SET start_date = ?, end_date = ?, updated_at = ?
                WHERE id = ?""",
                (start_date, end_date, _timestamp(), period_id),
            )
        return self.get_report_period(period_id)

    def delete_report_period(self, period_id: int) -> ReportPeriodRecord:
        with self._connect() as connection:
            existing = connection.execute(
                "SELECT id, is_deleted FROM report_periods WHERE id = ?", (period_id,)
            ).fetchone()
            if existing is None:
                raise KeyError(f"Report period {period_id} does not exist.")
            if existing["is_deleted"]:
                raise ValueError("This report period has already been deleted.")
            active_run = connection.execute(
                """SELECT 1 FROM runs
                WHERE period_id = ? AND state IN ('queued', 'running') LIMIT 1""",
                (period_id,),
            ).fetchone()
            if active_run is not None:
                raise ValueError("This period cannot be deleted while a workflow is running.")
            connection.execute(
                """UPDATE report_periods
                SET is_deleted = 1, updated_at = ?
                WHERE id = ?""",
                (_timestamp(), period_id),
            )
        return self.get_report_period(period_id)

    def restore_report_period(self, period_id: int) -> ReportPeriodRecord:
        with self._connect() as connection:
            existing = connection.execute(
                "SELECT id, start_date, end_date, is_deleted FROM report_periods WHERE id = ?", (period_id,)
            ).fetchone()
            if existing is None:
                raise KeyError(f"Report period {period_id} does not exist.")
            if not existing["is_deleted"]:
                raise ValueError("This report period is not deleted.")
            duplicate = connection.execute(
                """SELECT 1 FROM report_periods
                WHERE start_date = ? AND end_date = ? AND is_deleted = 0 AND id != ?""",
                (existing["start_date"], existing["end_date"], period_id),
            ).fetchone()
            if duplicate is not None:
                raise ValueError("An active report period already uses these dates.")
            connection.execute(
                """UPDATE report_periods
                SET is_deleted = 0, updated_at = ?
                WHERE id = ?""",
                (_timestamp(), period_id),
            )
        return self.get_report_period(period_id)

    def list_report_period_summaries(
        self, workflow_ids: tuple[str, ...], *, is_deleted: bool = False
    ) -> list[ReportPeriodSummary]:
        with self._connect() as connection:
            period_rows = connection.execute(
                """SELECT * FROM report_periods
                WHERE is_deleted = ?
                ORDER BY start_date DESC, id DESC""",
                (int(is_deleted),),
            ).fetchall()
            run_rows = connection.execute(
                """SELECT * FROM runs
                WHERE period_id IS NOT NULL
                ORDER BY period_id, workflow_id, id DESC"""
            ).fetchall()

        newest_runs: dict[int, dict[str, RunRecord]] = {}
        for row in run_rows:
            run = self._run_from_row(row)
            if run.workflow_id not in workflow_ids or run.period_id is None:
                continue
            newest_runs.setdefault(run.period_id, {}).setdefault(run.workflow_id, run)

        summaries: list[ReportPeriodSummary] = []
        for row in period_rows:
            period = self._period_from_row(row)
            runs = {
                workflow_id: newest_runs.get(period.id, {}).get(workflow_id)
                for workflow_id in workflow_ids
            }
            activity = [period.updated_at]
            for run in runs.values():
                if run is not None:
                    activity.extend(
                        timestamp
                        for timestamp in (run.created_at, run.started_at, run.finished_at)
                        if timestamp
                    )
            summaries.append(
                ReportPeriodSummary(
                    period=period,
                    runs=runs,
                    last_updated_at=max(activity),
                )
            )
        return summaries

    def create_run(
        self,
        workflow_id: str,
        workflow_label: str,
        start_date: str,
        end_date: str,
        *,
        period_id: int | None = None,
    ) -> RunRecord:
        if period_id is not None:
            period = self.get_report_period(period_id)
            if period.is_deleted:
                raise ValueError("This report period has been deleted and cannot start a workflow.")
            if (start_date, end_date) != (period.start_date, period.end_date):
                raise ValueError("Workflow dates must match the selected report period.")
        created_at = _timestamp()
        with self._connect() as connection:
            cursor = connection.execute(
                """INSERT INTO runs (
                    period_id, workflow_id, workflow_label, start_date, end_date, state,
                    created_at, log_path
                ) VALUES (?, ?, ?, ?, ?, 'queued', ?, '')""",
                (period_id, workflow_id, workflow_label, start_date, end_date, created_at),
            )
            run_id = int(cursor.lastrowid)
            log_path = (self.runs_dir / str(run_id) / "run.log").resolve()
            log_path.parent.mkdir(parents=True, exist_ok=True)
            log_path.touch(exist_ok=True)
            connection.execute("UPDATE runs SET log_path = ? WHERE id = ?", (str(log_path), run_id))
        return self.get_run(run_id)

    def get_run(self, run_id: int) -> RunRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
        if row is None:
            raise KeyError(f"Run {run_id} does not exist.")
        return self._run_from_row(row)

    def list_runs(
        self, *, workflow_id: str | None = None, state: str | None = None, limit: int = 100
    ) -> list[RunRecord]:
        clauses: list[str] = []
        values: list[str | int] = []
        if workflow_id:
            clauses.append("workflow_id = ?")
            values.append(workflow_id)
        if state:
            clauses.append("state = ?")
            values.append(state)
        values.append(limit)
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        with self._connect() as connection:
            rows = connection.execute(
                f"SELECT * FROM runs {where} ORDER BY id DESC LIMIT ?", values
            ).fetchall()
        return [self._run_from_row(row) for row in rows]

    def active_run(self) -> RunRecord | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM runs WHERE state IN ('queued', 'running') ORDER BY id DESC LIMIT 1"
            ).fetchone()
        return self._run_from_row(row) if row else None

    def fail_interrupted_runs(self) -> None:
        """Mark jobs from a previous process instance as failed after a restart."""
        with self._connect() as connection:
            connection.execute(
                """UPDATE runs
                SET state = 'failed', finished_at = ?,
                    error_summary = 'The local control panel stopped before this run completed.'
                WHERE state IN ('queued', 'running')""",
                (_timestamp(),),
            )

    def mark_running(self, run_id: int) -> None:
        with self._connect() as connection:
            connection.execute(
                "UPDATE runs SET state = 'running', started_at = ? WHERE id = ?", (_timestamp(), run_id)
            )

    def mark_finished(
        self,
        run_id: int,
        *,
        state: str,
        exit_code: int | None,
        error_summary: str | None = None,
    ) -> None:
        with self._connect() as connection:
            connection.execute(
                """UPDATE runs
                SET state = ?, finished_at = ?, exit_code = ?, error_summary = ?
                WHERE id = ?""",
                (state, _timestamp(), exit_code, error_summary, run_id),
            )

    def append_log(self, run_id: int, text: str) -> None:
        run = self.get_run(run_id)
        with run.log_path.open("a", encoding="utf-8", newline="") as log_file:
            log_file.write(text)
            if text and not text.endswith("\n"):
                log_file.write("\n")

    def read_log(self, run_id: int, max_characters: int = 60_000) -> str:
        run = self.get_run(run_id)
        if not run.log_path.is_file():
            return ""
        content = run.log_path.read_text(encoding="utf-8", errors="replace")
        return content[-max_characters:]

    def add_artifact(
        self, run_id: int, name: str, source_path: Path, stored_path: Path
    ) -> ArtifactRecord:
        with self._connect() as connection:
            cursor = connection.execute(
                """INSERT INTO artifacts (run_id, name, source_path, stored_path, size_bytes)
                VALUES (?, ?, ?, ?, ?)""",
                (run_id, name, str(source_path), str(stored_path), stored_path.stat().st_size),
            )
            artifact_id = int(cursor.lastrowid)
        return ArtifactRecord(
            id=artifact_id,
            run_id=run_id,
            name=name,
            source_path=str(source_path),
            stored_path=stored_path,
            size_bytes=stored_path.stat().st_size,
        )

    def list_artifacts(self, run_id: int) -> list[ArtifactRecord]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT * FROM artifacts WHERE run_id = ? ORDER BY id", (run_id,)
            ).fetchall()
        return [
            ArtifactRecord(
                id=row["id"],
                run_id=row["run_id"],
                name=row["name"],
                source_path=row["source_path"],
                stored_path=Path(row["stored_path"]),
                size_bytes=row["size_bytes"],
            )
            for row in rows
        ]

    def get_artifact(self, run_id: int, artifact_id: int) -> ArtifactRecord:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT * FROM artifacts WHERE id = ? AND run_id = ?", (artifact_id, run_id)
            ).fetchone()
        if row is None:
            raise KeyError(f"Artifact {artifact_id} does not belong to run {run_id}.")
        return ArtifactRecord(
            id=row["id"],
            run_id=row["run_id"],
            name=row["name"],
            source_path=row["source_path"],
            stored_path=Path(row["stored_path"]),
            size_bytes=row["size_bytes"],
        )

    def artifact_paths(self, run_id: int) -> Iterable[Path]:
        return (artifact.stored_path for artifact in self.list_artifacts(run_id))
