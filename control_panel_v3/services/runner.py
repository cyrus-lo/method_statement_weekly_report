"""Single-process V3 subprocess runner with run-local evidence."""

from __future__ import annotations

import os
import shutil
import subprocess
import threading
from datetime import date

from .config import Settings
from .history import HistoryStore, RunRecord
from .workflows import PreparedWorkflow


class WorkflowBusyError(RuntimeError):
    """A different V3 child process is already active."""


class WorkflowRunner:
    _ecm_workbook_lock = threading.Lock()

    def __init__(self, settings: Settings, history: HistoryStore) -> None:
        self.settings = settings
        self.history = history
        self._lock = threading.Lock()
        self._active_run_id: int | None = None

    def active_run(self) -> RunRecord | None:
        return self.history.active_run()

    def submit(self, workflow_id: str, label: str, start_date: date, end_date: date, *, period_id: int | None = None, prepared: PreparedWorkflow | None = None, prepared_factory=None, manual_step_1b_bypass: bool = False, lookup_scope: str | None = None) -> RunRecord:
        with self._lock:
            active = self.history.active_run()
            if self._active_run_id is not None or active:
                raise WorkflowBusyError(f"Run {self._active_run_id or active.id} is still active.")
            run = self.history.create_run(workflow_id, label, start_date.isoformat(), end_date.isoformat(), period_id=period_id, manual_step_1b_bypass=manual_step_1b_bypass, lookup_scope=lookup_scope)
            if prepared_factory is not None:
                try:
                    prepared = prepared_factory(run.id)
                except Exception as error:
                    message = str(error) or type(error).__name__
                    self.history.append_log(run.id, f"PREFLIGHT ERROR: {message}")
                    self.history.mark_finished(run.id, state="failed", exit_code=None, error_summary=message[:1000])
                    raise
            self._active_run_id = run.id
            thread = threading.Thread(target=self._execute, args=(run.id, prepared), name=f"v3-workflow-{run.id}", daemon=True)
            thread.start()
            return run

    def _execute(self, run_id: int, prepared: PreparedWorkflow | None) -> None:
        exit_code: int | None = None
        workbook_locked = False
        try:
            self.history.mark_running(run_id)
            if prepared is None:
                self.history.append_log(run_id, "Preview plan accepted; no child command configured.")
                self.history.mark_finished(run_id, state="succeeded", exit_code=0)
                return
            self.history.append_log(run_id, f"Starting: {prepared.definition.label}")
            self.history.append_log(run_id, "Command: " + subprocess.list2cmdline(prepared.command))
            if prepared.definition.modifies_workbooks or any(flag in prepared.command for flag in ("--workbook", "--ecm-workbook", "--smm-workbook")):
                if not self._ecm_workbook_lock.acquire(blocking=False):
                    raise RuntimeError("Another V3 workbook workflow is currently accessing ECM workbooks.")
                workbook_locked = True
            environment = os.environ.copy()
            if prepared.environment_overrides:
                environment = {}
                environment.update(prepared.environment_overrides)
            process = subprocess.Popen(prepared.command, cwd=prepared.working_directory, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL, text=True, encoding="utf-8", errors="replace", env=environment)
            assert process.stdout is not None
            output_thread = threading.Thread(
                target=self._capture_output,
                args=(run_id, process.stdout),
                name=f"v3-output-{run_id}",
                daemon=True,
            )
            output_thread.start()
            try:
                exit_code = process.wait(timeout=self.settings.job_timeout_seconds)
            except subprocess.TimeoutExpired as error:
                process.terminate()
                try:
                    process.wait(timeout=20)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait()
                raise RuntimeError("V3 workflow exceeded its configured timeout and was stopped.") from error
            output_thread.join(timeout=5)
            process.stdout.close()
            if exit_code != 0:
                raise RuntimeError(f"Workflow exited with code {exit_code}.")
            self._copy_artifacts(run_id, prepared)
            self.history.mark_finished(run_id, state="succeeded", exit_code=exit_code)
        except Exception as error:
            if prepared is not None:
                try:
                    self._copy_artifacts(run_id, prepared, optional_only=True)
                except Exception as artifact_error:
                    self.history.append_log(run_id, f"OPTIONAL ARTIFACT COPY ERROR: {artifact_error}")
            message = str(error) or type(error).__name__
            self.history.append_log(run_id, f"ERROR: {message}")
            self.history.mark_finished(run_id, state="failed", exit_code=exit_code, error_summary=message[:1000])
        finally:
            if workbook_locked:
                self._ecm_workbook_lock.release()
            with self._lock:
                if self._active_run_id == run_id:
                    self._active_run_id = None

    def _copy_artifacts(self, run_id: int, prepared: PreparedWorkflow, *, optional_only: bool = False) -> None:
        artifact_dir = self.settings.runs_dir / str(run_id) / "artifacts"
        artifact_dir.mkdir(parents=True, exist_ok=True)
        existing_names = {artifact.name for artifact in self.history.list_artifacts(run_id)}
        for source in prepared.artifacts:
            if optional_only and source.required:
                continue
            if source.name in existing_names:
                continue
            if not source.path.is_file():
                if source.required:
                    raise RuntimeError(f"Expected V3 output was not created: {source.path}")
                continue
            destination = artifact_dir / source.path.name
            shutil.copy2(source.path, destination)
            self.history.add_artifact(run_id, source.name, source.path, destination)

    def _capture_output(self, run_id: int, stream) -> None:
        for line in iter(stream.readline, ""):
            self.history.append_log(run_id, line)
