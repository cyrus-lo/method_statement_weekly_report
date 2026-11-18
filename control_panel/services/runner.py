"""Single-worker subprocess coordinator for the existing workflow CLIs."""

from __future__ import annotations

import os
import shutil
import subprocess
import threading
from datetime import date

from .config import Settings
from .history import HistoryStore, RunRecord
from .workflows import ArtifactSource, PreparedWorkflow, prepare_workflow


class WorkflowBusyError(RuntimeError):
    """Raised when a workbook-mutating job is already running."""


class WorkflowRunner:
    """Run exactly one existing workflow CLI at a time and retain its evidence."""

    def __init__(self, settings: Settings, history: HistoryStore) -> None:
        self.settings = settings
        self.history = history
        self._lock = threading.Lock()
        self._active_run_id: int | None = None

    def active_run(self) -> RunRecord | None:
        active = self.history.active_run()
        return active

    def submit(
        self,
        workflow_id: str,
        start_date: date,
        end_date: date,
        *,
        period_id: int | None = None,
        extra_args: tuple[str, ...] = (),
    ) -> RunRecord:
        prepared = prepare_workflow(workflow_id, start_date, end_date, self.settings, extra_args)
        with self._lock:
            active = self.history.active_run()
            if self._active_run_id is not None or active is not None:
                active_id = self._active_run_id or active.id
                raise WorkflowBusyError(f"Run {active_id} is still active.")
            run = self.history.create_run(
                prepared.definition.id,
                prepared.definition.label,
                start_date.isoformat(),
                end_date.isoformat(),
                period_id=period_id,
            )
            self._active_run_id = run.id
            thread = threading.Thread(
                target=self._execute,
                args=(run.id, prepared),
                name=f"workflow-run-{run.id}",
                daemon=True,
            )
            thread.start()
        return run

    def _execute(self, run_id: int, prepared: PreparedWorkflow) -> None:
        exit_code: int | None = None
        try:
            self.history.mark_running(run_id)
            self.history.append_log(run_id, f"Starting: {prepared.definition.label}")
            self.history.append_log(run_id, f"Working directory: {prepared.working_directory}")
            self.history.append_log(run_id, "Command: " + subprocess.list2cmdline(prepared.command))
            exit_code = self._run_process(run_id, prepared)
            if exit_code != 0:
                raise RuntimeError(f"Workflow process exited with code {exit_code}.")
            self._copy_artifacts(run_id, prepared.artifacts)
        except Exception as error:  # The raw exception is retained in the per-run log.
            message = str(error) or error.__class__.__name__
            self.history.append_log(run_id, f"\nERROR: {message}")
            self._copy_failure_diagnostics(run_id, prepared)
            self.history.mark_finished(
                run_id,
                state="failed",
                exit_code=exit_code,
                error_summary=message[:1_000],
            )
        else:
            self.history.append_log(run_id, "\nWorkflow completed successfully.")
            self.history.mark_finished(run_id, state="succeeded", exit_code=exit_code)
        finally:
            with self._lock:
                if self._active_run_id == run_id:
                    self._active_run_id = None

    def _copy_failure_diagnostics(self, run_id: int, prepared: PreparedWorkflow) -> None:
        """Register run-local worker diagnostics without masking the original failure."""
        if prepared.definition.id != "ecm-step-0":
            return
        diagnostics_dir = self.settings.runs_dir / str(run_id) / "artifacts"
        if not diagnostics_dir.is_dir():
            return
        for source in diagnostics_dir.iterdir():
            if not source.is_file():
                continue
            try:
                self.history.add_artifact(run_id, f"Step 0 diagnostic: {source.name}", source, source)
                self.history.append_log(run_id, f"Saved diagnostic artifact: {source.name}")
            except OSError as error:
                self.history.append_log(run_id, f"Diagnostic artifact registration failed: {error}")

    def _run_process(self, run_id: int, prepared: PreparedWorkflow) -> int:
        environment = os.environ.copy()
        if prepared.definition.id == "ecm-step-0":
            diagnostics_dir = self.settings.runs_dir / str(run_id) / "artifacts"
            diagnostics_dir.mkdir(parents=True, exist_ok=True)
            environment["STEP0_DIAGNOSTICS_DIR"] = str(diagnostics_dir)
        process = subprocess.Popen(
            prepared.command,
            cwd=prepared.working_directory,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            stdin=subprocess.DEVNULL,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=environment,
        )

        def capture_output() -> None:
            assert process.stdout is not None
            for line in iter(process.stdout.readline, ""):
                self.history.append_log(run_id, line)
            process.stdout.close()

        output_thread = threading.Thread(target=capture_output, name=f"log-capture-{run_id}", daemon=True)
        output_thread.start()
        try:
            return process.wait(timeout=self.settings.job_timeout_seconds)
        except subprocess.TimeoutExpired as error:
            process.terminate()
            try:
                process.wait(timeout=20)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait()
            raise RuntimeError(
                f"Workflow exceeded the {self.settings.job_timeout_seconds}-second timeout and was stopped."
            ) from error
        finally:
            output_thread.join(timeout=30)

    def _copy_artifacts(self, run_id: int, artifacts: tuple[ArtifactSource, ...]) -> None:
        artifact_dir = self.settings.runs_dir / str(run_id) / "artifacts"
        artifact_dir.mkdir(parents=True, exist_ok=True)
        for artifact in artifacts:
            if not artifact.path.is_file():
                raise RuntimeError(f"Expected output was not created: {artifact.path}")
            destination = artifact_dir / artifact.path.name
            shutil.copy2(artifact.path, destination)
            self.history.add_artifact(run_id, artifact.name, artifact.path, destination)
            self.history.append_log(run_id, f"Saved artifact: {destination.name}")
