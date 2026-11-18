"""V3 workflow inventory and period transition policy."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .history import ReportPeriodSummary, RunRecord


@dataclass(frozen=True)
class WorkflowDefinition:
    id: str
    label: str
    description: str
    next_action: str
    modifies_workbooks: bool


WORKFLOWS = {
    "weekly-report": WorkflowDefinition("weekly-report", "Weekly Method Statement Report", "Generate the independent weekly report.", "Review the generated report.", False),
    "ecm-step-0": WorkflowDefinition("ecm-step-0", "ECM Step 0 — Export Properties", "Export a validated, period-specific ECM CSV.", "Create the V3 ECM review workbook from this period's CSV.", False),
    "ecm-step-1a": WorkflowDefinition("ecm-step-1a", "ECM Step 1a — Create Review", "Create a review sheet from the period CSV.", "Run the ECM first-date lookup.", True),
    "ecm-first-date-lookup": WorkflowDefinition("ecm-first-date-lookup", "ECM First-Date Lookup", "Find the oldest ECM document date by the exact CSV outgoing reference.", "Review unmatched rows, then run Step 1b.", True),
    "ecm-step-1b": WorkflowDefinition("ecm-step-1b", "ECM Step 1b — Finalize Review", "Validate Remark dates and finalize Count (Y/N).", "Complete Use SMM / Use KMS decisions, then run Step 2.", True),
    "ecm-step-2": WorkflowDefinition("ecm-step-2", "ECM Step 2 — Create SMM Summary", "Validate Use SMM / Use KMS decisions and transactionally update the V3 SMM workbook.", "Review the generated period summary, workbook, and post-save backup.", True),
}
REPORT_WORKFLOW_IDS = tuple(WORKFLOWS)
MANUAL_WORKFLOW_IDS = ("weekly-report", "ecm-step-1a", "ecm-first-date-lookup", "ecm-step-1b", "ecm-step-2")
ECM_WORKFLOW_IDS = ("ecm-step-0", "ecm-step-1a", "ecm-first-date-lookup", "ecm-step-1b", "ecm-step-2")
ECM_SEQUENCE = ("ecm-step-0", "ecm-step-1a", "ecm-first-date-lookup", "ecm-step-1b", "ecm-step-2")


@dataclass(frozen=True)
class ArtifactSource:
    name: str
    path: Path
    required: bool = True


@dataclass(frozen=True)
class PreparedWorkflow:
    definition: WorkflowDefinition
    command: list[str]
    working_directory: Path
    artifacts: tuple[ArtifactSource, ...] = ()
    environment_overrides: dict[str, str] | None = None


def eligible_next_ecm_workflow(run: RunRecord, summaries: list[ReportPeriodSummary]) -> str | None:
    if run.workflow_id != "ecm-step-0" or run.state != "succeeded" or run.period_id is None:
        return None
    summary = next((item for item in summaries if item.period.id == run.period_id), None)
    if summary is None or summary.runs.get("ecm-step-0") is None or summary.runs["ecm-step-0"].id != run.id:
        return None
    return "ecm-step-1a"


def ecm_workflow_gate_error(summary: ReportPeriodSummary, workflow_id: str, *, confirm_step_1b_bypass: bool = False) -> str | None:
    if workflow_id not in ECM_WORKFLOW_IDS:
        return None
    latest = summary.runs
    step0, one_a = latest.get("ecm-step-0"), latest.get("ecm-step-1a")
    lookup, one_b = latest.get("ecm-first-date-lookup"), latest.get("ecm-step-1b")
    if workflow_id == "ecm-step-0":
        if step0 and step0.state == "succeeded":
            return "ECM Step 0 already succeeded for this period; reuse its saved CSV."
        return None
    if workflow_id == "ecm-step-1a":
        if not step0 or step0.state != "succeeded":
            return "ECM Step 0 must succeed before ECM Step 1a can start."
        return None
    if workflow_id == "ecm-first-date-lookup":
        if not one_a or one_a.state != "succeeded":
            return "ECM Step 1a must succeed before first-date lookup can start."
        return None
    if workflow_id == "ecm-step-1b":
        if not one_a or one_a.state != "succeeded":
            return "ECM Step 1a must succeed before ECM Step 1b can start."
        if not lookup or lookup.state != "succeeded":
            return "The first-date lookup must succeed before ECM Step 1b can start."
        return None
    if one_b is None and confirm_step_1b_bypass:
        return None
    if one_b is None or one_b.state != "succeeded":
        return "ECM Step 1b must succeed before ECM Step 2 can start."
    return None


def build_period_dashboards(summaries: list[ReportPeriodSummary]) -> list[dict[str, object]]:
    result: list[dict[str, object]] = []
    for summary in summaries:
        runs = summary.runs
        status = "Not started"
        next_action = "Run ECM Step 0 to export the period properties CSV."
        attention = None
        for workflow in ECM_SEQUENCE:
            run = runs.get(workflow)
            if run is None:
                next_action = f"Next: {WORKFLOWS[workflow].label}. {WORKFLOWS[workflow].description}"
                break
            if run.state in {"queued", "running"}:
                status, attention, next_action = run.state.capitalize(), f"Wait for {WORKFLOWS[workflow].label} to finish.", f"View run #{run.id}."
                break
            if run.state == "failed":
                status = "Failed — action required"
                attention = run.error_summary or f"Resolve {WORKFLOWS[workflow].label} before continuing."
                next_action = f"Review run #{run.id} and retry after resolving the failure."
                break
        else:
            status, next_action = "Succeeded", "Review the finalized SMM workbook and backups."
        result.append({"summary": summary, "period": summary.period, "runs": runs, "ecm_status": status, "next_action": next_action, "attention": attention, "weekly_status": (runs["weekly-report"].state.capitalize() if runs.get("weekly-report") else "Not started")})
    return result
