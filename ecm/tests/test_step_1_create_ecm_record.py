"""Regression tests for Step 1 ECM review-record generation."""

from __future__ import annotations

import csv
import io
import sys
import tempfile
import unittest
from datetime import date, datetime
from pathlib import Path
from contextlib import redirect_stderr
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook

ECM_DIRECTORY = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ECM_DIRECTORY))

import Step_1_Create_ECM_Record as step_1  # noqa: E402
from Step_1_Create_ECM_Record import (  # noqa: E402
    EcmRecord,
    InputValidationError,
    REVIEW_HEADERS,
    apply_count_filter_in_excel,
    apply_counting_rules,
    project_code_from_fields,
    read_ecm_records,
    write_review_workbook,
)


class Step1MajorVersionTests(unittest.TestCase):
    def test_missing_required_csv_header_stops_before_processing_rows(self) -> None:
        headers = [header for header in (
            "Document title",
            "Document date",
            "Name",
            "Document revision",
            "Modified By",
            "Major Version",
            "Modified On",
            "Outgoing ref. no.",
            "Path",
            "Link",
        )]
        with tempfile.TemporaryDirectory() as temporary_directory:
            csv_path = Path(temporary_directory) / "properties.csv"
            with csv_path.open("w", encoding="utf-8", newline="") as source:
                writer = csv.DictWriter(source, fieldnames=headers)
                writer.writeheader()
                writer.writerow({header: "unread row value" for header in headers})

            with self.assertRaisesRegex(InputValidationError, r"missing header\(s\): SMM ID"):
                read_ecm_records(csv_path, date(2026, 7, 1), date(2026, 7, 13))

    def create_record(self, source_row: int, major_version: int, modified_on: datetime) -> EcmRecord:
        return EcmRecord(
            source_row=source_row,
            document_title="Method Statement of Tree Felling Works at RIW2",
            document_date=datetime(2026, 7, 3),
            name="CWCMGCJV/952/CSF/8730-2025/A.pdf",
            document_revision="A",
            modified_by="dm_admin",
            major_version=major_version,
            smm_id="15168",
            modified_on=modified_on,
            outgoing_reference="CWCMGCJV/952/CSF/8730-2025/A",
            path="/CWCMGCJV/952/CSF/8730-2025/A.pdf",
            link="https://example.test/document",
            project_code=952,
        )

    def create_review_workbook(self, workbook_path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "ECM_1-12 Jul (WIP)"
        for column, header in enumerate(REVIEW_HEADERS, start=1):
            worksheet.cell(1, column).value = header
        workbook.save(workbook_path)
        workbook.close()

    def create_review_workbook_with_sheet(self, workbook_path: Path, sheet_name: str) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = sheet_name
        for column, header in enumerate(REVIEW_HEADERS, start=1):
            worksheet.cell(1, column).value = header
        workbook.save(workbook_path)
        workbook.close()

    def test_ok_template_is_used_when_wip_template_is_missing(self) -> None:
        record = self.create_record(2, 1, datetime(2026, 7, 3, 15, 22))

        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "ECM_Record.xlsx"
            self.create_review_workbook_with_sheet(workbook_path, "ECM_1-12 Jul (OK)")
            write_review_workbook(workbook_path, date(2026, 7, 1), date(2026, 7, 12), [record])

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                self.assertIn("ECM_1-12 Jul (OK)", workbook.sheetnames)
                self.assertIn("ECM_1-12 Jul (WIP)", workbook.sheetnames)
                self.assertEqual(
                    "CWCMGCJV/952/CSF/8730-2025/A.pdf",
                    workbook["ECM_1-12 Jul (WIP)"].cell(2, 4).value,
                )
            finally:
                workbook.close()

    def test_ecm_template_is_preferred_and_row_height_is_preserved(self) -> None:
        record = self.create_record(2, 1, datetime(2026, 7, 3, 15, 22))

        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "ECM_Record.xlsx"
            self.create_review_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                legacy = workbook["ECM_1-12 Jul (WIP)"]
                legacy.row_dimensions[13].height = 180
                template = workbook.create_sheet("ECM_template")
                for column, header in enumerate(REVIEW_HEADERS, start=1):
                    template.cell(1, column).value = header
                template.row_dimensions[13].height = 15
                workbook.save(workbook_path)
            finally:
                workbook.close()

            write_review_workbook(
                workbook_path,
                date(2026, 7, 1),
                date(2026, 7, 12),
                [record],
            )

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                self.assertEqual(15, worksheet.row_dimensions[13].height)
            finally:
                workbook.close()

    def test_project_code_falls_back_to_outgoing_reference(self) -> None:
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            project_code = project_code_from_fields(
                "SCWJV-1089-MTH-BSGE-0002A.pdf",
                "SCWJV/1089/MTH/BSGE/0002A",
                40,
            )

        self.assertEqual(1089, project_code)
        self.assertIn("CSV row 40", stderr.getvalue())
        self.assertIn("Outgoing ref. no.", stderr.getvalue())

    def test_project_code_can_be_recovered_by_normalized_identifier_match(self) -> None:
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            project_code = project_code_from_fields(
                "SCWJV-1089-MTH-BSGE-0002A.pdf",
                "SCWJV.1089.MTH.BSGE.0002A",
                41,
            )

        self.assertEqual(1089, project_code)
        self.assertIn("matching Name and Outgoing ref. no.", stderr.getvalue())

    def test_project_code_remains_unresolved_when_identifiers_do_not_match(self) -> None:
        self.assertIsNone(
            project_code_from_fields(
                "SCWJV-1089-MTH-BSGE-0002A.pdf",
                "SCWJV.1090.MTH.BSGE.0002B",
                42,
            )
        )

    def test_lower_major_version_is_retained_and_marked_n(self) -> None:
        older = self.create_record(2, 1, datetime(2026, 7, 3, 15, 22))
        newer = self.create_record(3, 3, datetime(2026, 7, 8, 9, 28))
        records = [older, newer]
        first_dates = {older.normalized_name: date(2026, 7, 3)}

        apply_counting_rules(records, readiness={952: True}, first_dates=first_dates)

        self.assertEqual("N", older.count)
        self.assertIn("superseded_major_version", older.reasons)
        self.assertEqual("Y", newer.count)

        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "ECM_Record.xlsx"
            self.create_review_workbook(workbook_path)
            write_review_workbook(
                workbook_path,
                date(2026, 7, 1),
                date(2026, 7, 12),
                records,
            )

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                matching_rows = [
                    (worksheet.cell(row, 7).value, worksheet.cell(row, 8).value, worksheet.cell(row, 9).value)
                    for row in range(2, worksheet.max_row + 1)
                    if worksheet.cell(row, 8).value == "15168"
                ]
            finally:
                workbook.close()

        self.assertEqual([(1, "15168", "N"), (3, "15168", "Y")], matching_rows)

    def test_sort_order_is_project_code_name_major_version_then_document_date(self) -> None:
        later_name_earlier_date = self.create_record(2, 1, datetime(2026, 7, 3, 15, 22))
        later_name_earlier_date.name = "CWCMGCJV/952/CSF/A.pdf"
        later_name_earlier_date.document_date = datetime(2026, 7, 1)

        earlier_name_later_date = self.create_record(3, 1, datetime(2026, 7, 8, 9, 28))
        earlier_name_later_date.name = "CWCMGCJV/952/CSF/B.pdf"
        earlier_name_later_date.document_date = datetime(2026, 7, 12)

        same_name_higher_version = self.create_record(4, 3, datetime(2026, 7, 8, 9, 28))
        same_name_higher_version.name = "CWCMGCJV/952/CSF/B.pdf"
        same_name_higher_version.document_date = datetime(2026, 7, 2)

        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "ECM_Record.xlsx"
            self.create_review_workbook(workbook_path)
            write_review_workbook(
                workbook_path,
                date(2026, 7, 1),
                date(2026, 7, 12),
                [earlier_name_later_date, same_name_higher_version, later_name_earlier_date],
            )

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                ordered_values = [
                    (worksheet.cell(row, 4).value, worksheet.cell(row, 7).value, worksheet.cell(row, 3).value)
                    for row in range(2, worksheet.max_row + 1)
                ]
            finally:
                workbook.close()

        self.assertEqual(
            [
                ("CWCMGCJV/952/CSF/A.pdf", 1, datetime(2026, 7, 1)),
                ("CWCMGCJV/952/CSF/B.pdf", 1, datetime(2026, 7, 12)),
                ("CWCMGCJV/952/CSF/B.pdf", 3, datetime(2026, 7, 2)),
            ],
            ordered_values,
        )

    def test_count_filter_uses_one_hidden_excel_session(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            workbook_path = Path(temporary_directory) / "ECM_Record.xlsx"
            workbook_path.write_bytes(b"placeholder")

            process = Mock(returncode=0)
            process.communicate.return_value = ("", "")
            with patch.object(step_1.sys, "platform", "win32"), patch.object(
                step_1.subprocess, "Popen", return_value=process
            ) as popen:
                apply_count_filter_in_excel(workbook_path, "ECM_1-12 Jul (WIP)")

        popen.assert_called_once()
        command = popen.call_args.args[0]
        script = command[-1]
        self.assertIn("-STA", command)
        process.communicate.assert_called_once_with(timeout=90)
        self.assertEqual(1, script.count("New-Object -ComObject Excel.Application"))
        self.assertEqual(1, script.count("$excel.Workbooks.Open($path)"))
        self.assertEqual(1, script.count("$workbook.Close($false)"))
        self.assertEqual(1, script.count("$excel.Quit()"))
        self.assertIn("$excel.AskToUpdateLinks = $false", script)
        self.assertIn("$excel.EnableEvents = $false", script)
        self.assertIn("Start-Sleep -Milliseconds 500", script)
        self.assertIn("Available sheets:", script)


if __name__ == "__main__":
    unittest.main()
