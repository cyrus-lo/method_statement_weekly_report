from __future__ import annotations

import time
import sqlite3
from datetime import date
from pathlib import Path

from control_panel.services.history import HistoryStore
from control_panel.services.runner import WorkflowRunner
from control_panel.services.workflows import ArtifactSource, PreparedWorkflow, WORKFLOWS

from .test_workflows import build_settings


def wait_for_terminal_state(history: HistoryStore, run_id: int) -> str:
    for _ in range(50):
        state = history.get_run(run_id).state
        if state in {"succeeded", "failed"}:
            return state
        time.sleep(0.05)
    raise AssertionError("Workflow did not finish in time.")


def test_runner_captures_log_and_copies_artifact(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    settings.django_script.write_text(
        """from pathlib import Path
import sys
output_dir = Path(sys.argv[sys.argv.index('--output-dir') + 1])
start = sys.argv[sys.argv.index('--start-date') + 1]
end = sys.argv[sys.argv.index('--end-date') + 1]
output = output_dir / f'Method_Statement_Interaction_Weekly_Report_{start}_to_{end}.xlsx'
output.write_text('report', encoding='utf-8')
print(f'Created: {output}')
""",
        encoding="utf-8",
    )
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    runner = WorkflowRunner(settings, history)

    run = runner.submit("weekly-report", date(2026, 7, 1), date(2026, 7, 7))

    assert run.created_at.endswith("+08:00")

    assert wait_for_terminal_state(history, run.id) == "succeeded"
    assert "Created:" in history.read_log(run.id)
    artifacts = history.list_artifacts(run.id)
    assert len(artifacts) == 1
    assert artifacts[0].stored_path.read_text(encoding="utf-8") == "report"


def test_history_migrates_existing_utc_timestamps_to_hong_kong(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()

    with sqlite3.connect(settings.database_path) as connection:
        connection.execute(
            """INSERT INTO runs (
                workflow_id, workflow_label, start_date, end_date, state,
                created_at, started_at, finished_at, log_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                "weekly-report",
                "Weekly report",
                "2026-07-01",
                "2026-07-07",
                "succeeded",
                "2026-07-27T00:00:00+00:00",
                "2026-07-27T00:01:00+00:00",
                "2026-07-27T00:02:00+00:00",
                str(settings.runs_dir / "1" / "run.log"),
            ),
        )

    history.initialize()
    migrated = history.get_run(1)

    assert migrated.created_at == "2026-07-27T08:00:00+08:00"
    assert migrated.started_at == "2026-07-27T08:01:00+08:00"
    assert migrated.finished_at == "2026-07-27T08:02:00+08:00"

    history.initialize()
    assert history.get_run(1).created_at == migrated.created_at


def test_runner_marks_missing_artifact_as_failed(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    runner = WorkflowRunner(settings, history)

    run = runner.submit("weekly-report", date(2026, 7, 1), date(2026, 7, 7))

    assert wait_for_terminal_state(history, run.id) == "failed"
    assert "Expected output" in (history.get_run(run.id).error_summary or "")


def test_runner_copies_all_step_1b_artifacts(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    settings.ecm_record.write_text("finalized", encoding="utf-8")
    backup = tmp_path / "ecm-output" / "before.xlsx"
    post_backup = tmp_path / "ecm-output" / "after.xlsx"
    backup.write_text("before", encoding="utf-8")
    post_backup.write_text("after", encoding="utf-8")
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    runner = WorkflowRunner(settings, history)
    prepared = PreparedWorkflow(
        definition=WORKFLOWS["ecm-step-1b"],
        command=[str(settings.ecm_python), "-c", "print('step 1b')"],
        working_directory=tmp_path,
        artifacts=(
            ArtifactSource("Finalized ECM review workbook", settings.ecm_record),
            ArtifactSource("Step 1b backup", backup),
            ArtifactSource("Post-Step 1b backup", post_backup),
        ),
    )

    runner._execute(
        history.create_run(
            "ecm-step-1b",
            "ECM Step 1b — Finalize Review",
            "2026-07-01",
            "2026-07-07",
        ).id,
        prepared,
    )

    run = history.list_runs(limit=1)[0]
    assert run.state == "succeeded"
    assert [artifact.name for artifact in history.list_artifacts(run.id)] == [
        "Finalized ECM review workbook",
        "Step 1b backup",
        "Post-Step 1b backup",
    ]


def test_runner_fails_step_1b_when_post_backup_is_missing(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    settings.ecm_record.write_text("finalized", encoding="utf-8")
    backup = tmp_path / "ecm-output" / "before.xlsx"
    backup.write_text("before", encoding="utf-8")
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    runner = WorkflowRunner(settings, history)
    prepared = PreparedWorkflow(
        definition=WORKFLOWS["ecm-step-1b"],
        command=[str(settings.ecm_python), "-c", "print('step 1b')"],
        working_directory=tmp_path,
        artifacts=(
            ArtifactSource("Finalized ECM review workbook", settings.ecm_record),
            ArtifactSource("Step 1b backup", backup),
            ArtifactSource("Post-Step 1b backup", tmp_path / "missing-after.xlsx"),
        ),
    )

    run_id = history.create_run(
        "ecm-step-1b",
        "ECM Step 1b — Finalize Review",
        "2026-07-01",
        "2026-07-07",
    ).id
    runner._execute(run_id, prepared)

    run = history.get_run(run_id)
    assert run.state == "failed"
    assert "Expected output was not created" in (run.error_summary or "")


def test_runner_registers_step0_run_scoped_failure_diagnostics(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    history = HistoryStore(settings.database_path, settings.runs_dir)
    history.initialize()
    runner = WorkflowRunner(settings, history)
    prepared = PreparedWorkflow(
        definition=WORKFLOWS["ecm-step-0"],
        command=[
            str(settings.ecm_python),
            "-c",
            (
                "import os; from pathlib import Path; "
                "diagnostics = Path(os.environ['STEP0_DIAGNOSTICS_DIR']); "
                "diagnostics.mkdir(parents=True, exist_ok=True); "
                "(diagnostics / 'failure-test.diagnostics.json').write_text('{}'); "
                "raise SystemExit(1)"
            ),
        ],
        working_directory=tmp_path,
        artifacts=(),
    )
    run_id = history.create_run(
        "ecm-step-0", "ECM Step 0 — Export Properties", "2026-07-01", "2026-07-07"
    ).id

    runner._execute(run_id, prepared)

    artifacts = history.list_artifacts(run_id)
    assert history.get_run(run_id).state == "failed"
    assert [artifact.name for artifact in artifacts] == [
        "Step 0 diagnostic: failure-test.diagnostics.json"
    ]
    assert artifacts[0].stored_path == (
        settings.runs_dir / str(run_id) / "artifacts" / "failure-test.diagnostics.json"
    )
