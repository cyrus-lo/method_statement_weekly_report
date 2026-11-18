"""Regression tests for weekly-report PostgreSQL timezone handling."""

from __future__ import annotations

import importlib.util
import os
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPT_PATH = Path(__file__).parents[1] / "Method_Statement_Interaction_Weekly_Report.py"
SPEC = importlib.util.spec_from_file_location("weekly_report", SCRIPT_PATH)
assert SPEC is not None and SPEC.loader is not None
weekly_report = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(weekly_report)


class FakeCursor:
    def __init__(self) -> None:
        self.calls: list[tuple[str, tuple[str, ...]]] = []

    def __enter__(self) -> FakeCursor:
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> bool:
        return False

    def execute(self, statement: str, parameters: tuple[str, ...]) -> None:
        self.calls.append((statement, parameters))


class FakeConnection:
    def __init__(self) -> None:
        self.cursor_instance = FakeCursor()

    def cursor(self) -> FakeCursor:
        return self.cursor_instance


class ReportingTimezoneTests(unittest.TestCase):
    def test_default_reporting_timezone_is_hong_kong(self) -> None:
        with patch.dict(os.environ, {}, clear=True):
            self.assertEqual(weekly_report.get_reporting_timezone(), "Asia/Hong_Kong")

    def test_set_connection_timezone_uses_parameterized_session_setting(self) -> None:
        connection = FakeConnection()

        weekly_report.set_connection_timezone(connection, "Asia/Hong_Kong")

        self.assertEqual(
            connection.cursor_instance.calls,
            [
                (
                    "SELECT set_config('TimeZone', %s, false) AS timezone_name",
                    ("Asia/Hong_Kong",),
                )
            ],
        )

    def test_project_1016_and_1007_are_combined(self) -> None:
        session_rows = [
            {
                "project_code": 1016,
                "username": "alice",
                "count_retrieve_mode": 2,
                "count_pdf_viewed": 3,
                "count_content_copied": 4,
                "count_ask_mode": 5,
                "total_issue_count": 6,
            },
            {
                "project_code": 1007,
                "username": "bob",
                "count_retrieve_mode": 7,
                "count_pdf_viewed": 8,
                "count_content_copied": 9,
                "count_ask_mode": 10,
                "total_issue_count": 11,
            },
        ]
        submission_rows = [
            {"project_code": 1016, "count_submission_times": 12},
            {"project_code": 1007, "count_submission_times": 13},
        ]

        summary_rows = weekly_report.build_summary_rows(session_rows, submission_rows)

        self.assertEqual(len(summary_rows), 1)
        self.assertEqual(summary_rows[0]["project_code"], 1016)
        self.assertEqual(summary_rows[0]["kms_user_count"], 2)
        self.assertEqual(summary_rows[0]["count_retrieve_mode"], 9)
        self.assertEqual(summary_rows[0]["count_pdf_viewed"], 11)
        self.assertEqual(summary_rows[0]["count_content_copied"], 13)
        self.assertEqual(summary_rows[0]["count_ask_mode"], 15)
        self.assertEqual(summary_rows[0]["total_issue_count"], 17)
        self.assertEqual(summary_rows[0]["count_submission_times"], 25)
        self.assertEqual(
            weekly_report.display_project_code(summary_rows[0]["project_code"]),
            "1016 (aka 1007)",
        )


if __name__ == "__main__":
    unittest.main()