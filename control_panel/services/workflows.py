"""Workflow definitions, command construction, and safe preflight checks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from .config import Settings


class PreflightError(ValueError):
    """Raised when a workflow should not start on the current workstation."""


@dataclass(frozen=True)
class ArtifactSource:
    name: str
    path: Path


@dataclass(frozen=True)
class WorkflowDefinition:
    id: str
    label: str
    short_description: str
    prerequisite: str
    next_action: str
    modifies_workbooks: bool


@dataclass(frozen=True)
class PreparedWorkflow:
    definition: WorkflowDefinition
    command: list[str]
    working_directory: Path
    artifacts: tuple[ArtifactSource, ...]


WORKFLOWS = {
    "weekly-report": WorkflowDefinition(
        id="weekly-report",
        label="Weekly Method Statement Report",
        short_description="Create the PostgreSQL-backed weekly Excel report.",
        prerequisite="The Django .env database settings must be valid.",
        next_action="Review or download the generated weekly report.",
        modifies_workbooks=False,
    ),
    "ecm-step-0": WorkflowDefinition(
        id="ecm-step-0",
        label="ECM Step 0 — Export Properties",
        short_description="Search the approved ECM report period and export a validated properties CSV.",
        prerequisite=(
            "The dedicated Edge session must already be authenticated. If the session has expired, "
            "run the Step 0 login/refresh procedure before retrying."
        ),
        next_action="Run ECM Step 1a with this period's exported properties CSV.",
        modifies_workbooks=False,
    ),
    "ecm-step-1a": WorkflowDefinition(
        id="ecm-step-1a",
        label="ECM Step 1a — Create Review",
        short_description="Create the period worksheet and prepare it for manual first-date review.",
        prerequisite="Close ECM_Record.xlsx and ensure a current properties CSV is available.",
        next_action="Open ECM_Record.xlsx in Excel and enter verified first-document dates in Remark (column M).",
        modifies_workbooks=True,
    ),
    "ecm-step-1b": WorkflowDefinition(
        id="ecm-step-1b",
        label="ECM Step 1b — Finalize Review",
        short_description="Validate the manual review and finalize Count (Y/N).",
        prerequisite="Complete the Step 1a manual Remark review, then close ECM_Record.xlsx in Excel.",
        next_action="After Step 1b succeeds, complete Use SMM / Use KMS decisions in Excel, then run ECM Step 2.",
        modifies_workbooks=True,
    ),
    "ecm-step-2": WorkflowDefinition(
        id="ecm-step-2",
        label="ECM Step 2 — Create SMM Summary",
        short_description="Update the SMM working workbook from the finalized ECM review.",
        prerequisite="Step 1b must have finalized the period; required Count, Use SMM, and Use KMS values must be complete.",
        next_action="Review or download the updated SMM workbook and its period backup.",
        modifies_workbooks=True,
    ),
    "ecm-period-report": WorkflowDefinition(
        id="ecm-period-report",
        label="ECM Monthly / Quarterly Report",
        short_description="Append approved ECM sheets and build a new period SMM summary.",
        prerequisite="Restore any previous generated sheets, close both workbooks, and confirm the selected source sheets.",
        next_action="Review the new ECM (WIP) sheet and change it to (OK) when approved.",
        modifies_workbooks=True,
    ),
}

V1_WORKFLOW_IDS = ("weekly-report", "ecm-step-1a", "ecm-step-1b", "ecm-step-2")
STANDALONE_WORKFLOW_IDS = ("ecm-period-report",)


def get_workflow(workflow_id: str) -> WorkflowDefinition:
    try:
        return WORKFLOWS[workflow_id]
    except KeyError as error:
        raise PreflightError(f"Unknown workflow: {workflow_id}") from error


def period_sheet_name(start_date: date, end_date: date) -> str:
    if start_date.month == end_date.month:
        return f"ECM_{start_date.day}-{end_date.day} {end_date:%b} (WIP)"
    return f"ECM_{start_date.day} {start_date:%b}-{end_date.day} {end_date:%b} (WIP)"


def _backup_path(prefix: str, start_date: date, end_date: date, suffix: str = "") -> str:
    return f"{prefix}_{start_date:%Y-%m-%d}_to_{end_date:%Y-%m-%d}{suffix}.xlsx"


def _ensure_unlocked(path: Path) -> None:
    if not path.is_file():
        raise PreflightError(f"Required workbook was not found: {path}")
    lock_file = path.with_name(f"~${path.name}")
    if lock_file.exists():
        raise PreflightError(f"Close the workbook in Excel before running this step: {path.name}")


def _newest_csv(directory: Path) -> Path:
    candidates = [
        path
        for path in directory.glob("properties*.csv")
        if path.is_file() and path.name.casefold() != "properties*.csv"
    ]
    if not candidates:
        raise PreflightError(f"No properties*.csv input was found in {directory}")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def _validate_dates(start_date: date, end_date: date) -> None:
    if end_date < start_date:
        raise PreflightError("End date cannot be earlier than start date.")


def _required_setting(settings: Settings, name: str) -> object:
    value = getattr(settings, name, None)
    if value is None:
        raise PreflightError(f"Step 0 setting is not configured: {name}")
    return value


def _step0_output_path(directory: Path, start_date: date, end_date: date) -> Path:
    timestamp = datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%f")
    return directory / (
        f"properties_{start_date:%Y-%m-%d}_to_{end_date:%Y-%m-%d}_{timestamp}.csv"
    )


def _step0_period_input_path(directory: Path, start_date: date, end_date: date) -> Path:
    return directory / (
        f"properties_{start_date:%Y-%m-%d}_to_{end_date:%Y-%m-%d}.csv"
    )


def prepare_workflow(
    workflow_id: str, start_date: date, end_date: date, settings: Settings,
    extra_args: tuple[str, ...] = (),
) -> PreparedWorkflow:
    """Build a shell-free child-process command only after local preflight passes."""
    definition = get_workflow(workflow_id)
    _validate_dates(start_date, end_date)
    start = start_date.isoformat()
    end = end_date.isoformat()

    if workflow_id == "weekly-report":
        output = settings.django_output_dir / (
            f"Method_Statement_Interaction_Weekly_Report_{start}_to_{end}.xlsx"
        )
        return PreparedWorkflow(
            definition=definition,
            command=[
                str(settings.django_python), str(settings.django_script),
                "--start-date", start, "--end-date", end,
                "--template", str(settings.django_template),
                "--output-dir", str(settings.django_output_dir),
            ],
            working_directory=settings.django_script.parent,
            artifacts=(ArtifactSource("Weekly report", output),),
        )

    if workflow_id == "ecm-step-0":
        step0_python = _required_setting(settings, "step0_python")
        step0_script = _required_setting(settings, "step0_script")
        step0_profile_dir = _required_setting(settings, "step0_profile_dir")
        ecm_url = _required_setting(settings, "ecm_url")
        step0_timeout_seconds = _required_setting(settings, "step0_timeout_seconds")
        output = _step0_output_path(settings.ecm_csv_dir, start_date, end_date)
        period_input = _step0_period_input_path(settings.ecm_csv_dir, start_date, end_date)
        command = [
            str(step0_python), str(step0_script),
            "--start-date", start, "--end-date", end,
            "--output", str(output),
            "--period-input", str(period_input),
            "--ecm-url", str(ecm_url),
            "--profile-dir", str(step0_profile_dir),
            "--timeout-seconds", str(step0_timeout_seconds),
        ]
        if getattr(settings, "step0_headed", False):
            command.append("--headed")
            hold_seconds = getattr(settings, "step0_failure_browser_hold_seconds", 0)
            command.extend(["--failure-browser-hold-seconds", str(hold_seconds)])
        if getattr(settings, "step0_login_if_needed", False):
            command.extend(
                [
                    "--login-if-needed",
                    "--auth-timeout-seconds",
                    str(getattr(settings, "step0_auth_timeout_seconds", 600)),
                ]
            )
        if getattr(settings, "step0_reset_profile", False):
            command.append("--reset-profile")
        return PreparedWorkflow(
            definition=definition,
            command=command,
            working_directory=Path(step0_script).parent,
            artifacts=(
                ArtifactSource("ECM properties CSV", output),
                ArtifactSource("ECM period CSV", period_input),
            ),
        )

    ecm_output_dir = settings.ecm_record.parent
    if workflow_id == "ecm-step-1a":
        _ensure_unlocked(settings.ecm_record) if settings.ecm_record.exists() else None
        if getattr(settings, "uses_period_csv_handoff", False):
            csv_path = _step0_period_input_path(settings.ecm_csv_dir, start_date, end_date)
            if not csv_path.is_file():
                raise PreflightError(
                    f"No Step 0 properties CSV was found for {start_date} to {end_date}: {csv_path}"
                )
        else:
            csv_path = _newest_csv(settings.ecm_csv_dir)
        backup = ecm_output_dir / _backup_path("ECM_Record", start_date, end_date, "_1a")
        readiness_workbook = settings.smm_record if settings.smm_record.is_file() else settings.smm_template
        return PreparedWorkflow(
            definition=definition,
            command=[
                str(settings.ecm_python), str(settings.step_1a_script),
                "--start-date", start, "--end-date", end,
                "--csv", str(csv_path),
                "--smm-workbook", str(readiness_workbook),
                "--ecm-record", str(settings.ecm_record),
                "--ecm-template", str(settings.ecm_template),
                "--backup", str(backup),
            ],
            working_directory=settings.step_1a_script.parent,
            artifacts=(
                ArtifactSource("ECM review workbook", settings.ecm_record),
                ArtifactSource("Step 1a backup", backup),
            ),
        )

    if workflow_id == "ecm-step-1b":
        _ensure_unlocked(settings.ecm_record)
        backup = ecm_output_dir / _backup_path("ECM_Record", start_date, end_date, "_1b")
        post_1b_backup = ecm_output_dir / _backup_path(
            "ECM_Record", start_date, end_date, "_1b_after"
        )
        return PreparedWorkflow(
            definition=definition,
            command=[
                str(settings.ecm_python), str(settings.step_1b_script),
                "--start-date", start, "--end-date", end,
                "--ecm-record", str(settings.ecm_record),
                "--ecm-sheet", period_sheet_name(start_date, end_date),
                "--backup", str(backup),
                "--post-1b-backup", str(post_1b_backup),
            ],
            working_directory=settings.step_1b_script.parent,
            artifacts=(
                ArtifactSource("Finalized ECM review workbook", settings.ecm_record),
                ArtifactSource("Step 1b backup", backup),
                ArtifactSource("Post-Step 1b backup", post_1b_backup),
            ),
        )

    if workflow_id == "ecm-step-2":
        _ensure_unlocked(settings.ecm_record)
        if settings.smm_record.exists():
            _ensure_unlocked(settings.smm_record)
        backup = ecm_output_dir / _backup_path("SMM_KMS_PM_Master_Document", start_date, end_date)
        return PreparedWorkflow(
            definition=definition,
            command=[
                str(settings.ecm_python), str(settings.step_2_script),
                "--start-date", start, "--end-date", end,
                "--ecm-record", str(settings.ecm_record),
                "--smm-template", str(settings.smm_template),
                "--smm-record", str(settings.smm_record),
                "--summary-template-sheet", settings.smm_summary_template_sheet,
                "--backup", str(backup),
            ],
            working_directory=settings.step_2_script.parent,
            artifacts=(
                ArtifactSource("SMM working workbook", settings.smm_record),
                ArtifactSource("SMM period backup", backup),
            ),
        )

    if workflow_id == "ecm-period-report":
        _ensure_unlocked(settings.ecm_record)
        _ensure_unlocked(settings.smm_record)
        ecm_backup = ecm_output_dir / _backup_path("ECM_Record", start_date, end_date, "_period")
        smm_backup = ecm_output_dir / _backup_path("SMM_KMS_PM_Master_Document", start_date, end_date, "_period")
        return PreparedWorkflow(
            definition=definition,
            command=[
                str(settings.ecm_python), str(settings.ecm_period_script),
                "--start-date", start, "--end-date", end,
                "--ecm-record", str(settings.ecm_record),
                "--smm-record", str(settings.smm_record),
                "--smm-template", str(settings.smm_template),
                "--summary-template-sheet", settings.smm_summary_template_sheet,
                "--ecm-backup", str(ecm_backup),
                "--smm-backup", str(smm_backup),
                *extra_args,
            ],
            working_directory=settings.ecm_period_script.parent,
            artifacts=(
                ArtifactSource("ECM working workbook", settings.ecm_record),
                ArtifactSource("SMM working workbook", settings.smm_record),
                ArtifactSource("ECM period backup", ecm_backup),
                ArtifactSource("SMM period backup", smm_backup),
            ),
        )

    raise PreflightError(f"Workflow is not implemented: {workflow_id}")
