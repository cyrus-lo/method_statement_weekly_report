"""Regression tests for the two-phase ECM Step 1 workflow."""

from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Side

ECM_DIRECTORY = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ECM_DIRECTORY))

from Step_1_Create_ECM_Record import (  # noqa: E402
    MISSING_SMM_ID_FILL,
    EcmRecord,
    InputValidationError,
    REVIEW_HEADERS,
    period_sheet_name,
)
from Step_1a_Create_ECM_Record import (  # noqa: E402
    apply_preliminary_counting_rules,
    outgoing_reference_without_revision,
    write_preliminary_review_workbook,
)
import Step_1b_Finalize_ECM_Record as step_1b  # noqa: E402
from Step_1b_Finalize_ECM_Record import finalize_pending_rows  # noqa: E402


class Step1a1bTests(unittest.TestCase):
    def test_outgoing_reference_removes_only_matching_terminal_revision_suffix(self) -> None:
        cases = (
            ("CWCMGCJV/952/CSF/8881-2026/A", "A", "CWCMGCJV/952/CSF/8881-2026"),
            ("CWCMGCJV/952/CSF/8881-2026", "A", "CWCMGCJV/952/CSF/8881-2026"),
            ("CWCMGCJV/952/CSF/8881-2026/AA", "A", "CWCMGCJV/952/CSF/8881-2026/AA"),
            ("CWCMGCJV/952/A/8881-2026", "A", "CWCMGCJV/952/A/8881-2026"),
            ("CWCMGCJV/952/CSF/8881-2026/A", "", "CWCMGCJV/952/CSF/8881-2026/A"),
            ("", "A", ""),
            ("  CWCMGCJV/952/CSF/8881-2026/A  ", " A ", "CWCMGCJV/952/CSF/8881-2026"),
        )
        for outgoing_reference, revision, expected in cases:
            with self.subTest(outgoing_reference=outgoing_reference, revision=revision):
                self.assertEqual(
                    expected,
                    outgoing_reference_without_revision(outgoing_reference, revision),
                )

    def create_record(self, source_row: int, name: str, version: int, project_code: int | None = 952) -> EcmRecord:
        return EcmRecord(
            source_row=source_row,
            document_title="Method Statement",
            document_date=datetime(2026, 7, 3),
            name=name,
            document_revision="A",
            modified_by="dm_admin",
            major_version=version,
            smm_id="15168",
            modified_on=datetime(2026, 7, 3, 15, source_row),
            outgoing_reference="ref",
            path="/path",
            link="https://example.test",
            project_code=project_code,
        )

    def create_workbook(self, workbook_path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "ECM_1-12 Jul (WIP)"
        for column, header in enumerate(REVIEW_HEADERS, start=1):
            worksheet.cell(1, column).value = header
        worksheet.cell(1, 12).value = "Work Flow"
        worksheet.cell(1, 13).value = "Remark"
        worksheet.cell(1, 14).value = "Display"
        worksheet.cell(1, 16).value = step_1b.OUTGOING_REFERENCE_HEADER
        workbook.save(workbook_path)
        workbook.close()

    def create_workbook_with_sheet(self, workbook_path: Path, sheet_name: str) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = sheet_name
        for column, header in enumerate(REVIEW_HEADERS, start=1):
            worksheet.cell(1, column).value = header
        worksheet.cell(1, 12).value = "Work Flow"
        worksheet.cell(1, 13).value = "Remark"
        worksheet.cell(1, 14).value = "Display"
        worksheet.cell(1, 16).value = step_1b.OUTGOING_REFERENCE_HEADER
        workbook.save(workbook_path)
        workbook.close()

    def test_1a_uses_ok_template_when_wip_template_is_missing(self) -> None:
        record = self.create_record(2, "CWCMGCJV/952/CSF/A.pdf", 1)
        apply_preliminary_counting_rules([record], {952: True})

        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "ECM_Record.xlsx"
            self.create_workbook_with_sheet(workbook_path, "ECM_1-12 Jul (OK)")
            write_preliminary_review_workbook(
                workbook_path,
                datetime(2026, 7, 1).date(),
                datetime(2026, 7, 13).date(),
                [record],
            )
            workbook = load_workbook(workbook_path, data_only=False)
            try:
                self.assertIn("ECM_1-12 Jul (OK)", workbook.sheetnames)
                self.assertIn("ECM_1-13 Jul (WIP)", workbook.sheetnames)
            finally:
                workbook.close()

    def test_1a_leaves_display_empty_until_step_2(self) -> None:
        record = self.create_record(2, "CWCMGCJV/952/CSF/A.pdf", 1)
        record.outgoing_reference = "CWCMGCJV/952/CSF/8881-2026/A"
        record.document_revision = "A"
        apply_preliminary_counting_rules([record], {952: True})

        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "ECM_Record.xlsx"
            self.create_workbook(workbook_path)
            write_preliminary_review_workbook(
                workbook_path,
                datetime(2026, 7, 1).date(),
                datetime(2026, 7, 13).date(),
                [record],
            )

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook["ECM_1-13 Jul (WIP)"]
                self.assertEqual(step_1b.OUTGOING_REFERENCE_HEADER, worksheet.cell(1, 16).value)
                self.assertIsNone(worksheet.cell(2, 14).value)
                self.assertEqual("CWCMGCJV/952/CSF/8881-2026", worksheet.cell(2, 16).value)
                self.assertEqual("CWCMGCJV/952/CSF/8881-2026/A", record.outgoing_reference)
                self.assertEqual("A1:P2", worksheet.auto_filter.ref)
            finally:
                workbook.close()

    def test_1a_preserves_records_and_leaves_only_candidates_pending(self) -> None:
        older = self.create_record(2, "CWCMGCJV/952/CSF/A.pdf", 1)
        canonical = self.create_record(3, "CWCMGCJV/952/CSF/A.pdf", 3)
        unresolved = self.create_record(4, "Unknown.pdf", 1, None)
        unresolved.smm_id = ""
        records = [canonical, unresolved, older]
        apply_preliminary_counting_rules(records, {952: True})

        self.assertEqual("N", older.count)
        self.assertEqual("", canonical.count)
        self.assertEqual("N", unresolved.count)

        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "ECM_Record.xlsx"
            self.create_workbook(workbook_path)
            write_preliminary_review_workbook(
                workbook_path,
                datetime(2026, 7, 1).date(),
                datetime(2026, 7, 13).date(),
                records,
            )
            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook[period_sheet_name(datetime(2026, 7, 1).date(), datetime(2026, 7, 13).date())]
                rows = [
                    (
                        worksheet.cell(row, 4).value,
                        worksheet.cell(row, 7).value,
                        worksheet.cell(row, 9).value,
                        worksheet.cell(row, 13).value,
                    )
                    for row in range(2, worksheet.max_row + 1)
                    if worksheet.cell(row, 4).value is not None
                ]
                self.assertEqual([], worksheet.auto_filter.filterColumn)
                self.assertNotEqual(
                    MISSING_SMM_ID_FILL.fgColor.rgb,
                    worksheet.cell(4, 1).fill.fgColor.rgb,
                )
            finally:
                workbook.close()

        self.assertEqual(
            [
                ("CWCMGCJV/952/CSF/A.pdf", 1, "N", "Duplicate with its version 3"),
                ("CWCMGCJV/952/CSF/A.pdf", 3, None, None),
                ("Unknown.pdf", 1, "N", None),
            ],
            rows,
        )

    def test_1a_treats_project_1007_as_ready_project_1016(self) -> None:
        record = self.create_record(
            2,
            "CWCMGCJV/1007/CSF/5260-2026.pdf",
            1,
            project_code=1007,
        )

        apply_preliminary_counting_rules([record], {1016: True})

        self.assertEqual("", record.count)
        self.assertNotIn("project_not_ready", record.reasons)

    def test_1a_removes_borders_from_unused_template_rows(self) -> None:
        records = [
            self.create_record(2, "CWCMGCJV/952/CSF/First.pdf", 1),
            self.create_record(3, "CWCMGCJV/952/CSF/Second.pdf", 1),
        ]
        apply_preliminary_counting_rules(records, {952: True})
        black_side = Side(style="thin", color="000000")
        black_border = Border(
            left=black_side,
            right=black_side,
            top=black_side,
            bottom=black_side,
        )

        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "ECM_Record.xlsx"
            self.create_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                for row_number in range(2, 6):
                    for column in range(1, 16):
                        worksheet.cell(row_number, column).border = black_border
                workbook.save(workbook_path)
            finally:
                workbook.close()

            write_preliminary_review_workbook(
                workbook_path,
                datetime(2026, 7, 1).date(),
                datetime(2026, 7, 13).date(),
                records,
            )

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook[period_sheet_name(datetime(2026, 7, 1).date(), datetime(2026, 7, 13).date())]
                self.assertEqual("thin", worksheet.cell(2, 1).border.left.style)
                self.assertEqual("thin", worksheet.cell(3, 1).border.left.style)
                for row_number in range(4, 6):
                    for column in range(1, 16):
                        self.assertIsNone(worksheet.cell(row_number, column).border.left)
                        self.assertIsNone(worksheet.cell(row_number, column).border.right)
                        self.assertIsNone(worksheet.cell(row_number, column).border.top)
                        self.assertIsNone(worksheet.cell(row_number, column).border.bottom)
            finally:
                workbook.close()

    def test_1b_finalizes_dates_and_preserves_manual_fields(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            workbook_path = directory_path / "ECM_Record.xlsx"
            backup_path = directory_path / "backup_1b.xlsx"
            post_1b_backup_path = directory_path / "backup_1b_after.xlsx"
            self.create_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                worksheet.cell(2, 1).value = 952
                worksheet.cell(2, 2).value = "First"
                worksheet.cell(2, 5).value = "A"
                worksheet.cell(2, 8).value = "15168"
                worksheet.cell(2, 9).value = None
                worksheet.cell(2, 10).value = "Y"
                worksheet.cell(2, 11).value = "N"
                worksheet.cell(2, 12).value = "keep L"
                worksheet.cell(2, 13).value = "2026-04-17"
                worksheet.cell(2, 16).value = "CWCMGCJV/952/CSF/8881-2026"
                worksheet.cell(3, 1).value = 952
                worksheet.cell(3, 2).value = "Second"
                worksheet.cell(3, 5).value = "B"
                worksheet.cell(3, 8).value = None
                worksheet.cell(3, 9).value = None
                worksheet.cell(3, 13).value = datetime(2026, 4, 16)
                workbook.save(workbook_path)
            finally:
                workbook.close()

            count_y, count_n = finalize_pending_rows(
                workbook_path,
                "ECM_1-12 Jul (WIP)",
                backup_path,
            )
            self.assertEqual((1, 1), (count_y, count_n))
            self.assertTrue(backup_path.is_file())

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                self.assertEqual("Y", worksheet.cell(2, 9).value)
                self.assertEqual("N", worksheet.cell(3, 9).value)
                self.assertEqual("Y", worksheet.cell(2, 10).value)
                self.assertEqual("N", worksheet.cell(2, 11).value)
                self.assertIsNone(worksheet.cell(2, 14).value)
                self.assertIsNone(worksheet.cell(3, 14).value)
                self.assertEqual("CWCMGCJV/952/CSF/8881-2026", worksheet.cell(2, 16).value)
                self.assertEqual("A1:P3", worksheet.auto_filter.ref)
                self.assertEqual("keep L", worksheet.cell(2, 12).value)
                self.assertEqual("1st version is submitted on 17 Apr 2026", worksheet.cell(2, 13).value)
                self.assertEqual("1st version is submitted on 16 Apr 2026", worksheet.cell(3, 13).value)
                self.assertEqual(["Y"], worksheet.auto_filter.filterColumn[0].filters.filter)
                self.assertNotEqual(
                    MISSING_SMM_ID_FILL.fgColor.rgb,
                    worksheet.cell(2, 1).fill.fgColor.rgb,
                )
                self.assertEqual(
                    MISSING_SMM_ID_FILL.fgColor.rgb,
                    worksheet.cell(3, 1).fill.fgColor.rgb,
                )
                self.assertEqual(
                    MISSING_SMM_ID_FILL.fgColor.rgb,
                    worksheet.cell(3, 13).fill.fgColor.rgb,
                )
            finally:
                workbook.close()

            self.assertFalse(post_1b_backup_path.exists())

    def test_1b_rejects_missing_pending_date_without_backup_or_update(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            workbook_path = directory_path / "ECM_Record.xlsx"
            backup_path = directory_path / "backup_1b.xlsx"
            self.create_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                worksheet.cell(2, 1).value = 952
                worksheet.cell(2, 2).value = "Pending"
                worksheet.cell(2, 5).value = "A"
                workbook.save(workbook_path)
            finally:
                workbook.close()

            with self.assertRaisesRegex(ValueError, "requires a verified"):
                finalize_pending_rows(workbook_path, "ECM_1-12 Jul (WIP)", backup_path)
            self.assertFalse(backup_path.exists())

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                self.assertIsNone(workbook["ECM_1-12 Jul (WIP)"].cell(2, 9).value)
            finally:
                workbook.close()

    def test_1b_rejects_missing_outgoing_reference_header(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            workbook_path = directory_path / "ECM_Record.xlsx"
            backup_path = directory_path / "backup_1b.xlsx"
            self.create_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                workbook["ECM_1-12 Jul (WIP)"].cell(1, 16).value = None
                workbook.save(workbook_path)
            finally:
                workbook.close()

            with self.assertRaisesRegex(ValueError, "column P"):
                finalize_pending_rows(workbook_path, "ECM_1-12 Jul (WIP)", backup_path)
            self.assertFalse(backup_path.exists())

    def test_1b_finalizes_blank_document_revision_without_m(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            workbook_path = directory_path / "ECM_Record.xlsx"
            backup_path = directory_path / "backup_1b.xlsx"
            self.create_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                worksheet.cell(2, 1).value = 952
                worksheet.cell(2, 2).value = "No revision"
                worksheet.cell(2, 5).value = "   "
                worksheet.cell(2, 8).value = "15168"
                worksheet.cell(2, 9).value = None
                worksheet.cell(2, 13).value = "Keep this manual remark"
                workbook.save(workbook_path)
            finally:
                workbook.close()

            self.assertEqual(
                (1, 0),
                finalize_pending_rows(workbook_path, "ECM_1-12 Jul (WIP)", backup_path),
            )

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                self.assertEqual("Y", worksheet.cell(2, 9).value)
                self.assertEqual("Keep this manual remark", worksheet.cell(2, 13).value)
            finally:
                workbook.close()

    def test_1b_reprocesses_y_rows_and_accepts_saved_first_version_remark(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            workbook_path = directory_path / "ECM_Record.xlsx"
            backup_path = directory_path / "backup_1b.xlsx"
            self.create_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                worksheet.cell(2, 1).value = 952
                worksheet.cell(2, 2).value = "Stays Y"
                worksheet.cell(2, 5).value = "A"
                worksheet.cell(2, 8).value = "15168"
                worksheet.cell(2, 9).value = "Y"
                worksheet.cell(2, 13).value = "1st version is submitted on 17 Apr 2026"
                worksheet.cell(3, 1).value = 952
                worksheet.cell(3, 2).value = "Changes to N"
                worksheet.cell(3, 5).value = "B"
                worksheet.cell(3, 8).value = "15169"
                worksheet.cell(3, 9).value = "Y"
                worksheet.cell(3, 13).value = "1st version is submitted on 10 Apr 2026"
                workbook.save(workbook_path)
            finally:
                workbook.close()

            self.assertEqual(
                (1, 1),
                finalize_pending_rows(workbook_path, "ECM_1-12 Jul (WIP)", backup_path),
            )

            workbook = load_workbook(workbook_path, data_only=False)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                self.assertEqual("Y", worksheet.cell(2, 9).value)
                self.assertEqual("N", worksheet.cell(3, 9).value)
                self.assertEqual("1st version is submitted on 17 Apr 2026", worksheet.cell(2, 13).value)
                self.assertEqual("1st version is submitted on 10 Apr 2026", worksheet.cell(3, 13).value)
                self.assertTrue(all(
                    worksheet.cell(row_number, 9).value is not None
                    for row_number in (2, 3)
                ))
            finally:
                workbook.close()

    def test_1b_starts_excel_filter_once_after_openpyxl_finalization(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            workbook_path = directory_path / "ECM_Record.xlsx"
            backup_path = directory_path / "backup_1b.xlsx"
            post_1b_backup_path = directory_path / "backup_1b_after.xlsx"
            args = SimpleNamespace(
                start_date=datetime(2026, 7, 1).date(),
                end_date=datetime(2026, 7, 13).date(),
                ecm_record=workbook_path,
                ecm_sheet=None,
                backup=backup_path,
                post_1b_backup=post_1b_backup_path,
            )
            events: list[str] = []

            def finalize(*arguments):
                events.append("finalize")
                return (1, 0)

            def apply_filter(*arguments):
                events.append("filter")

            def create_post_backup(*arguments):
                events.append("post-backup")

            with patch.object(step_1b, "parse_arguments", return_value=args), patch.object(
                step_1b, "finalize_pending_rows", side_effect=finalize
            ) as finalize_mock, patch.object(
                step_1b, "apply_count_filter_in_excel", side_effect=apply_filter
            ) as filter_mock, patch.object(
                step_1b.shutil, "copy2", side_effect=create_post_backup
            ) as copy_mock:
                self.assertEqual(0, step_1b.main())

        worksheet_name = period_sheet_name(args.start_date, args.end_date)
        finalize_mock.assert_called_once_with(
            workbook_path, worksheet_name, backup_path
        )
        filter_mock.assert_called_once_with(workbook_path, worksheet_name)
        copy_mock.assert_called_once_with(workbook_path, post_1b_backup_path)
        self.assertEqual(["finalize", "filter", "post-backup"], events)

    def test_1b_does_not_create_post_backup_when_excel_filter_fails(self) -> None:
        args = SimpleNamespace(
            start_date=datetime(2026, 7, 1).date(),
            end_date=datetime(2026, 7, 13).date(),
            ecm_record=Path("ECM_Record.xlsx"),
            ecm_sheet=None,
            backup=Path("backup_1b.xlsx"),
            post_1b_backup=Path("backup_1b_after.xlsx"),
        )

        with patch.object(step_1b, "parse_arguments", return_value=args), patch.object(
            step_1b, "finalize_pending_rows", return_value=(1, 0)
        ), patch.object(
            step_1b, "apply_count_filter_in_excel", side_effect=RuntimeError("Excel busy")
        ), patch.object(step_1b.shutil, "copy2") as copy_mock:
            with self.assertRaises(RuntimeError):
                step_1b.main()

        copy_mock.assert_not_called()

    def test_1b_does_not_start_excel_when_validation_fails(self) -> None:
        args = SimpleNamespace(
            start_date=datetime(2026, 7, 1).date(),
            end_date=datetime(2026, 7, 13).date(),
            ecm_record=Path("ECM_Record.xlsx"),
            ecm_sheet=None,
            backup=Path("backup_1b.xlsx"),
            post_1b_backup=Path("backup_1b_after.xlsx"),
        )

        with patch.object(step_1b, "parse_arguments", return_value=args), patch.object(
            step_1b,
            "finalize_pending_rows",
            side_effect=InputValidationError("missing Remark date"),
        ), patch.object(step_1b, "apply_count_filter_in_excel") as filter_mock:
            with self.assertRaises(InputValidationError):
                step_1b.main()

        filter_mock.assert_not_called()

    def test_1b_does_not_replace_workbook_when_save_fails(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            directory_path = Path(directory)
            workbook_path = directory_path / "ECM_Record.xlsx"
            backup_path = directory_path / "backup_1b.xlsx"
            self.create_workbook(workbook_path)
            workbook = load_workbook(workbook_path)
            try:
                worksheet = workbook["ECM_1-12 Jul (WIP)"]
                worksheet.cell(2, 1).value = 952
                worksheet.cell(2, 2).value = "Save failure"
                worksheet.cell(2, 5).value = "A"
                worksheet.cell(2, 13).value = "2026-04-17"
                workbook.save(workbook_path)
            finally:
                workbook.close()
            original_bytes = workbook_path.read_bytes()

            with patch.object(step_1b.os, "replace", side_effect=OSError("disk full")):
                with self.assertRaisesRegex(OSError, "disk full"):
                    finalize_pending_rows(workbook_path, "ECM_1-12 Jul (WIP)", backup_path)

            self.assertEqual(original_bytes, workbook_path.read_bytes())
            self.assertTrue(backup_path.is_file())


if __name__ == "__main__":
    unittest.main()
