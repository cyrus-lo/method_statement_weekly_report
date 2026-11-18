"""Calendar rules for report periods that remain aligned with month ends."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

HONG_KONG = ZoneInfo("Asia/Hong_Kong")


def _week_start(value: date) -> date:
    return value - timedelta(days=value.weekday())


def _first_day_of_month(value: date) -> date:
    return value.replace(day=1)


def report_period_for(value: date) -> tuple[date, date]:
    """Return the month-end-aware report period containing ``value``.

    Periods normally run Monday through Sunday. A week that would cross a
    month boundary is split between the adjacent periods: the preceding
    period extends to month end, and the new month's opening days join the
    following full Monday-to-Sunday week. For example, this produces
    2026-06-22 to 2026-06-30 and 2026-07-01 to 2026-07-12.
    """
    week_start = _week_start(value)
    week_end = week_start + timedelta(days=6)
    next_week_start = week_start + timedelta(days=7)
    next_week_end = next_week_start + timedelta(days=6)
    previous_week_start = week_start - timedelta(days=7)
    previous_week_end = week_start - timedelta(days=1)

    if week_start.month != week_end.month:
        month_start = _first_day_of_month(week_end)
        if value < month_start:
            return previous_week_start, month_start - timedelta(days=1)
        return month_start, week_end + timedelta(days=7)

    if previous_week_start.month != previous_week_end.month:
        return _first_day_of_month(value), week_end

    if next_week_start.month != next_week_end.month:
        return week_start, _first_day_of_month(next_week_end) - timedelta(days=1)

    return week_start, week_end


def latest_completed_report_period(today: date | None = None) -> tuple[date, date]:
    """Return the latest period whose inclusive end date has passed in HKT."""
    reference_date = today or datetime.now(HONG_KONG).date()
    candidate = report_period_for(reference_date)
    if candidate[1] < reference_date:
        return candidate
    return report_period_for(candidate[0] - timedelta(days=1))
