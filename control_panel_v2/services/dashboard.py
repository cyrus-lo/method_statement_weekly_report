"""Derived, period-first operational status for V2 Dashboard."""

from __future__ import annotations

from dataclasses import dataclass

from control_panel.services.history import ReportPeriodSummary, RunRecord


ECM_WORKFLOW_IDS = ("ecm-step-0", "ecm-step-1a", "ecm-step-1b", "ecm-step-2")


@dataclass(frozen=True)
class PeriodDashboard:
    summary: ReportPeriodSummary
    weekly_status: str
    ecm_status: str
    next_action: str
    attention: str | None

    @property
    def period(self):
        return self.summary.period

    @property
    def runs(self) -> dict[str, RunRecord | None]:
        return self.summary.runs


def _run_status(run: RunRecord | None) -> str:
    if run is None:
        return "Not started"
    if run.state == "failed":
        return "Failed — action required"
    return run.state.capitalize()


def _is_active(run: RunRecord | None) -> bool:
    return run is not None and run.state in {"queued", "running"}


def eligible_next_ecm_workflow(
    run: RunRecord, summaries: list[ReportPeriodSummary]
) -> str | None:
    """Return the safe next-step action for a current successful run, if any."""
    if run.workflow_id != "ecm-step-0" or run.state != "succeeded" or run.period_id is None:
        return None

    summary = next(
        (item for item in summaries if item.period.id == run.period_id), None
    )
    if summary is None or summary.runs.get("ecm-step-0") is None:
        return None
    if summary.runs["ecm-step-0"].id != run.id:
        return None
    if any(
        summary.runs.get(workflow_id) is not None
        for workflow_id in ("ecm-step-1a", "ecm-step-1b", "ecm-step-2")
    ):
        return None
    return "ecm-step-1a"


def ecm_workflow_gate_error(
    summaries: list[ReportPeriodSummary],
    summary: ReportPeriodSummary,
    workflow_id: str,
    *,
    confirm_step_1b_bypass: bool = False,
) -> str | None:
    """Return an execution error when an ECM transition is unsafe or unconfirmed."""
    if workflow_id not in ECM_WORKFLOW_IDS:
        return None

    step_0 = summary.runs.get("ecm-step-0")
    one_a = summary.runs.get("ecm-step-1a")
    one_b = summary.runs.get("ecm-step-1b")

    if workflow_id == "ecm-step-0":
        if step_0 is not None and step_0.state == "succeeded":
            return "ECM Step 0 already succeeded for this period; use its saved CSV artifact."
        if step_0 is None and one_a is not None:
            return "This legacy period already passed Step 0 and cannot start a new Step 0 run."
        return None

    if workflow_id == "ecm-step-1a":
        if step_0 is None:
            if one_a is not None:
                return None
            return "ECM Step 0 must succeed before ECM Step 1a can start."
        if step_0.state != "succeeded":
            return "ECM Step 0 must succeed before ECM Step 1a can start."
        return None

    if workflow_id == "ecm-step-1b":
        if one_a is None or one_a.state != "succeeded":
            return "ECM Step 1a must succeed before ECM Step 1b can start."
        return None

    if one_b is None and confirm_step_1b_bypass:
        return None
    if one_b is None or one_b.state != "succeeded":
        return "ECM Step 1b must succeed before ECM Step 2 can start."
    return None


def build_period_dashboards(summaries: list[ReportPeriodSummary]) -> list[PeriodDashboard]:
    dashboards: list[PeriodDashboard] = []
    for summary in summaries:
        weekly = summary.runs.get("weekly-report")
        step_0 = summary.runs.get("ecm-step-0")
        one_a = summary.runs.get("ecm-step-1a")
        one_b = summary.runs.get("ecm-step-1b")
        two = summary.runs.get("ecm-step-2")
        if step_0 is None and one_a is None:
            ecm_status = "Not started"
            attention = None
            next_action = "Run ECM Step 0 to export the period properties CSV from ECM."
        elif _is_active(step_0):
            ecm_status = step_0.state.capitalize()
            attention = "Wait for the ECM properties export to finish."
            next_action = "View the active Step 0 run log."
        elif step_0 is not None and step_0.state == "failed":
            ecm_status = "Failed — action required"
            attention = step_0.error_summary or "Resolve the Step 0 export failure before continuing."
            next_action = "Review the failed Step 0 run and retry after fixing authentication or ECM access."
        elif one_a is None:
            ecm_status = "Not started"
            attention = None
            next_action = "Run ECM Step 1a using the exported properties CSV."
        elif _is_active(one_a):
            ecm_status = one_a.state.capitalize()
            attention = "Wait for ECM Step 1a to finish before opening the workbook."
            next_action = "View the active run log."
        elif one_a.state == "failed":
            ecm_status = "Failed — action required"
            attention = one_a.error_summary or "Resolve the Step 1a failure before continuing."
            next_action = "Review the failed Step 1a run and use the recovery procedure."
        elif one_b is None:
            ecm_status = "Waiting for manual Remark review"
            attention = "Open ECM_Record.xlsx, complete required Remark dates, save, and close Excel."
            next_action = "Complete the manual Remark review before ECM Step 1b."
        elif _is_active(one_b):
            ecm_status = one_b.state.capitalize()
            attention = "Wait for ECM Step 1b to finish."
            next_action = "View the active run log."
        elif one_b.state == "failed":
            ecm_status = "Failed — action required"
            attention = one_b.error_summary or "Resolve the Step 1b failure before continuing."
            next_action = "Review the failed Step 1b run and use the recovery procedure."
        elif two is None:
            ecm_status = "Waiting for manual Use SMM / Use KMS decisions"
            attention = "Step 1b is complete. Complete required Use SMM and Use KMS values in Excel, then save and close it."
            next_action = "Complete the post-Step-1b manual decisions before ECM Step 2."
        elif _is_active(two):
            ecm_status = two.state.capitalize()
            attention = "Wait for ECM Step 2 to finish."
            next_action = "View the active run log."
        elif two.state == "failed":
            ecm_status = "Failed — action required"
            attention = two.error_summary or "Resolve the Step 2 failure before continuing."
            next_action = "Review the failed Step 2 run and use the recovery procedure."
        else:
            ecm_status = "Succeeded"
            attention = None
            next_action = "Review or download the finalized SMM workbook and backup."
        dashboards.append(
            PeriodDashboard(
                summary=summary,
                weekly_status=_run_status(weekly),
                ecm_status=ecm_status,
                next_action=next_action,
                attention=attention,
            )
        )
    return dashboards
