from __future__ import annotations

import sys
import unittest
from copy import copy
from datetime import date
from hashlib import sha256
from pathlib import Path
from uuid import uuid4

from openpyxl import Workbook, load_workbook

ECM_ROOT = Path(__file__).resolve().parents[1]
REAL_ECM_RECORD = ECM_ROOT / "output" / "ECM_Record_2026-07-01_to_2026-07-13.xlsx"
TEST_WORKBOOK_DIRECTORY = Path(__file__).resolve().parent / "generated_workbooks"
sys.path.insert(0, str(ECM_ROOT))

from Step_2_Create_SMM_Summary import (  # noqa: E402
    InputValidationError,
    aggregate_ecm_record,
    create_smm_workbook,
)
from display import display_value  # noqa: E402


class AggregateEcmRecordTests(unittest.TestCase):
    def test_display_value_mapping(self) -> None:
        self.assertEqual("via KMS", display_value("Y", "Y", "Y"))
        self.assertEqual("w/o KMS", display_value("Y", "Y", "N"))
        self.assertEqual("", display_value("Y", "Y", ""))
        self.assertEqual("", display_value("N", "Y", "Y"))
        self.assertEqual("", display_value("Y", "N", "N"))

    def create_workbook(self, rows: list[tuple[object, ...]]) -> Path:
        TEST_WORKBOOK_DIRECTORY.mkdir(parents=True, exist_ok=True)
        path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "ECM_1-13 Jul (WIP)"
        worksheet.append(
            [
                "Project Code",
                "Document title",
                "Document date",
                "Name",
                "Document revision",
                "Modified By",
                "Major Version",
                "SMM ID",
                "Count (Y/N)",
                "Use SMM (Y/N)",
                "Use KMS (Y/N)",
                "Work Flow",
                "Remark",
            ]
        )
        for row in rows:
            worksheet.append(row)
        workbook.save(path)
        workbook.close()
        return path

    def test_aggregates_summary_metrics_and_adds_rows_not_using_smm(self) -> None:
        rows = [
            (100, None, None, None, None, None, None, "S1", "Y", "Y", "Y"),
            (100, None, None, None, None, None, None, "S2", "Y", "Y", "N"),
            (100, None, None, None, None, None, None, "S3", "Y", "N", "N"),
            (100, None, None, None, None, None, None, "S4", "N", "Y", "Y"),
            (101, None, None, None, None, None, None, None, "Y", "N", "N"),
        ]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].draft_via_kms, 1)
        self.assertEqual(metrics[100].draft_not_via_kms, 1)
        self.assertEqual(metrics[100].not_using_smm, 1)
        self.assertEqual(metrics[100].outbound_ecm, 3)
        self.assertEqual(metrics[101].draft_via_kms, 0)
        self.assertEqual(metrics[101].draft_not_via_kms, 0)
        self.assertEqual(metrics[101].not_using_smm, 1)
        self.assertEqual(metrics[101].outbound_ecm, 1)

    def test_reads_rows_even_when_hidden_by_an_excel_filter(self) -> None:
        rows = [
            (100, None, None, None, None, None, None, "S1", "Y", "Y", "Y"),
            (100, None, None, None, None, None, None, "S2", "Y", "Y", "N"),
        ]
        path = self.create_workbook(rows)
        from openpyxl import load_workbook

        workbook = load_workbook(path)
        worksheet = workbook["ECM_1-13 Jul (WIP)"]
        worksheet.row_dimensions[3].hidden = True
        workbook.save(path)
        workbook.close()

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].outbound_ecm, 2)
        self.assertEqual(metrics[100].draft_not_via_kms, 1)

    def test_handles_each_valid_count_smm_kms_combination(self) -> None:
        rows = [
            (100, None, None, None, None, None, None, "S1", "Y", "Y", "Y"),
            (100, None, None, None, None, None, None, "S2", "Y", "Y", "N"),
            (100, None, None, None, None, None, None, "S3", "Y", "N", "N"),
            (100, None, None, None, None, None, None, "S4", "Y", "N", "N"),
            (100, None, None, None, None, None, None, "S5", "N", "Y", "Y"),
            (100, None, None, None, None, None, None, "S6", "N", "N", "N"),
            (100, None, None, None, None, None, None, "S7", "N", "", ""),
        ]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].draft_via_kms, 1)
        self.assertEqual(metrics[100].draft_not_via_kms, 1)
        self.assertEqual(metrics[100].not_using_smm, 2)
        self.assertEqual(metrics[100].outbound_ecm, 4)

    def test_all_i_y_j_n_rows_add_to_h(self) -> None:
        rows = [
            (100, None, None, None, None, None, None, "S1", "Y", "N", "N"),
            (100, None, None, None, None, None, None, "S2", "Y", "N", ""),
        ]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].not_using_smm, 2)
        self.assertEqual(metrics[100].outbound_ecm, 2)

    def test_allows_blank_kms_when_count_is_y_and_smm_is_n(self) -> None:
        rows = [(100, None, None, None, None, None, None, "S1", "Y", "N", "")]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].draft_via_kms, 0)
        self.assertEqual(metrics[100].draft_not_via_kms, 0)
        self.assertEqual(metrics[100].not_using_smm, 1)
        self.assertEqual(metrics[100].outbound_ecm, 1)

    def test_ignores_blank_or_invalid_flags_when_count_is_n(self) -> None:
        rows = [
            (None, None, None, None, None, None, None, "S1", "N", "", ""),
            (None, None, None, None, None, None, None, "S2", "N", "Not reviewed", "Invalid"),
            (100, None, None, None, None, None, None, "S3", "Y", "Y", "Y"),
        ]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(set(metrics), {100})
        self.assertEqual(metrics[100].outbound_ecm, 1)

    def test_rejects_blank_flags_when_count_is_y(self) -> None:
        rows = [(100, None, None, None, None, None, None, "S1", "Y", "Y", "")]
        path = self.create_workbook(rows)

        with self.assertRaisesRegex(InputValidationError, r"Use KMS \(Y/N\) must be Y or N"):
            aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

    def test_real_july_ecm_record_rejects_unreviewed_blank_flags(self) -> None:
        if not REAL_ECM_RECORD.is_file():
            self.skipTest(f"Regression fixture not found: {REAL_ECM_RECORD}")

        with self.assertRaisesRegex(InputValidationError, r"Use SMM \(Y/N\) must be Y or N"):
            aggregate_ecm_record(REAL_ECM_RECORD, "ECM_1-13 Jul (WIP)")

    def test_rejects_eligible_row_without_project_code(self) -> None:
        rows = [(None, None, None, None, None, None, None, "S1", "Y", "Y", "Y")]
        path = self.create_workbook(rows)

        with self.assertRaisesRegex(InputValidationError, "Project Code is required"):
            aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

    def test_rejects_use_kms_without_use_smm(self) -> None:
        rows = [(100, None, None, None, None, None, None, "S1", "Y", "N", "Y")]
        path = self.create_workbook(rows)

        with self.assertRaisesRegex(InputValidationError, "Use KMS cannot be Y"):
            aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

    def test_maps_project_1007_to_canonical_1016(self) -> None:
        path = self.create_workbook(
            [(1007, None, None, None, None, None, None, "S1", "Y", "Y", "Y")]
        )

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(set(metrics), {1016})
        self.assertEqual(metrics[1016].draft_via_kms, 1)

    def test_deducts_mattex_added_records_from_outbound_only(self) -> None:
        rows = [
            (100, None, None, None, None, None, None, "S1", "Y", "Y", "Y", None, "Added by Mattex"),
            (100, None, None, None, None, None, None, "S2", "Y", "Y", "N", None, None),
            (100, None, None, None, None, None, None, "S3", "Y", "Y", "N", None, "ADD By Chris's Team"),
        ]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].draft_via_kms, 1)
        self.assertEqual(metrics[100].draft_not_via_kms, 2)
        self.assertEqual(metrics[100].raw_outbound_ecm, 3)
        self.assertEqual(metrics[100].mattex_added, 2)
        self.assertEqual(metrics[100].outbound_ecm, 1)

    def test_normalises_curly_apostrophe_and_whitespace_in_mattex_remark(self) -> None:
        rows = [
            (100, None, None, None, None, None, None, "S1", "Y", "Y", "Y", None, "  add  by Chris’s  Team  "),
        ]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].draft_via_kms, 1)
        self.assertEqual(metrics[100].mattex_added, 1)
        self.assertEqual(metrics[100].outbound_ecm, 0)

    def test_does_not_deduct_not_prepared_via_smm_row(self) -> None:
        rows = [(100, None, None, None, None, None, None, "S1", "Y", "N", "", None, None)]
        path = self.create_workbook(rows)

        metrics = aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

        self.assertEqual(metrics[100].not_using_smm, 1)
        self.assertEqual(metrics[100].mattex_added, 0)
        self.assertEqual(metrics[100].outbound_ecm, 1)

    def test_rejects_mattex_remark_on_not_using_smm_row(self) -> None:
        rows = [
            (100, None, None, None, None, None, None, "S1", "Y", "N", "N", None, "Added by Mattex"),
        ]
        path = self.create_workbook(rows)

        with self.assertRaisesRegex(InputValidationError, "Mattex-added Remark cannot be used"):
            aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")

    def test_rejects_missing_or_wrong_remark_header(self) -> None:
        path = self.create_workbook(
            [(100, None, None, None, None, None, None, "S1", "Y", "Y", "Y")]
        )
        workbook = load_workbook(path)
        worksheet = workbook["ECM_1-13 Jul (WIP)"]
        worksheet.cell(1, 13).value = "Notes"
        workbook.save(path)
        workbook.close()

        with self.assertRaisesRegex(InputValidationError, "M1: expected 'Remark'"):
            aggregate_ecm_record(path, "ECM_1-13 Jul (WIP)")


class CreateSmmWorkbookTests(unittest.TestCase):
    def create_ecm_workbook(self) -> Path:
        TEST_WORKBOOK_DIRECTORY.mkdir(parents=True, exist_ok=True)
        path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_ECM_Record.xlsx"
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "ECM_1-13 Jul (WIP)"
        worksheet.append(
            [
                "Project Code", "Document title", "Document date", "Name", "Document revision",
                "Modified By", "Major Version", "SMM ID", "Count (Y/N)", "Use SMM (Y/N)",
                "Use KMS (Y/N)", "Work Flow", "Remark",
            ]
        )
        worksheet.append([100, None, None, None, None, None, None, "S1", "Y", "Y", "Y"])
        worksheet.append([100, None, None, None, None, None, None, "S2", "Y", "Y", "N"])
        worksheet.append([101, None, None, None, None, None, None, "S3", "Y", "N", "N"])
        workbook.save(path)
        workbook.close()
        return path

    def create_smm_template(self) -> Path:
        path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_template.xlsx"
        workbook = Workbook()
        readiness = workbook.active
        readiness.title = "Project Realiness Chris-In Use"
        readiness.append([None, None, "Project", "BU"])
        readiness.append([None, None, 100, "BU 100"])
        readiness.append([None, None, 101, "BU 101"])
        summary = workbook.create_sheet("(OK) 1 Jul - 12 Jul")
        summary.cell(2, 2).value = "Project"
        summary.cell(2, 3).value = "BU"
        summary.cell(2, 4).value = "SMM (Y/N)"
        summary.cell(2, 8).value = "Number of Method Statement Record (Outbound) in ECM"
        summary.cell(3, 5).value = "A. Draft via KMS"
        summary.cell(3, 6).value = "B. Draft NOT via KMS"
        summary.cell(3, 7).value = "Total No. (A + B) Approved"
        summary.append([None, None, None, None, None, None, None, None])
        workbook.save(path)
        workbook.close()
        return path

    def test_updates_working_smm_workbook_and_creates_period_backup(self) -> None:
        ecm_path = self.create_ecm_workbook()
        template_path = self.create_smm_template()
        run_id = uuid4().hex
        working_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_SMM_KMS_PM_Master_Document.xlsx"
        backup_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_SMM_backup.xlsx"
        ecm_snapshot_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_ECM_1b_after.xlsx"
        ecm_snapshot_path.write_bytes(ecm_path.read_bytes())
        snapshot_digest = sha256(ecm_snapshot_path.read_bytes()).digest()
        template_digest = sha256(template_path.read_bytes()).digest()

        sheet_name, project_count = create_smm_workbook(
            ecm_record_path=ecm_path,
            smm_template_path=template_path,
            smm_record_path=working_path,
            backup_path=backup_path,
            start_date=date(2026, 7, 1),
            end_date=date(2026, 7, 13),
            ecm_worksheet_name="ECM_1-13 Jul (WIP)",
            summary_template_sheet="(OK) 1 Jul - 12 Jul",
        )

        self.assertEqual(sheet_name, "(WIP) 1 Jul - 13 Jul")
        self.assertEqual(project_count, 2)
        self.assertTrue(working_path.is_file())
        self.assertTrue(backup_path.is_file())
        self.assertEqual(sha256(ecm_snapshot_path.read_bytes()).digest(), snapshot_digest)
        self.assertNotEqual(sha256(ecm_path.read_bytes()).digest(), snapshot_digest)
        self.assertEqual(sha256(template_path.read_bytes()).digest(), template_digest)

        ecm_workbook = load_workbook(ecm_path, data_only=False)
        ecm_worksheet = ecm_workbook["ECM_1-13 Jul (WIP)"]
        self.assertEqual("via KMS", ecm_worksheet.cell(2, 14).value)
        self.assertEqual("w/o KMS", ecm_worksheet.cell(3, 14).value)
        self.assertIn(ecm_worksheet.cell(4, 14).value, (None, ""))
        ecm_workbook.close()

        backup = load_workbook(backup_path, data_only=False)
        self.assertIn(sheet_name, backup.sheetnames)
        backup.close()
        workbook = load_workbook(working_path, data_only=False)
        worksheet = workbook[sheet_name]
        self.assertEqual(workbook.active.title, sheet_name)
        self.assertEqual(worksheet.cell(5, 5).value, 1)
        self.assertEqual(worksheet.cell(5, 6).value, 1)
        self.assertEqual(worksheet.cell(5, 7).value, "=E5+F5")
        self.assertEqual(worksheet.cell(5, 8).value, 2)
        self.assertEqual(worksheet.cell(6, 8).value, 1)
        self.assertEqual(worksheet.column_dimensions["B"].width, 93.1875)
        self.assertEqual(worksheet.column_dimensions["C"].width, 16.46375)
        self.assertTrue(worksheet.column_dimensions["D"].hidden)
        for column in ("E", "F", "G"):
            self.assertEqual(worksheet.column_dimensions[column].width, 20.46375)
        self.assertEqual(worksheet.column_dimensions["H"].width, 27.46375)
        self.assertEqual(worksheet.sheet_view.zoomScale, 85)
        workbook.close()

    def test_only_nonzero_not_via_kms_figures_are_red(self) -> None:
        ecm_path = self.create_ecm_workbook()
        template_path = self.create_smm_template()
        template_workbook = load_workbook(template_path)
        template_sheet = template_workbook["(OK) 1 Jul - 12 Jul"]
        for column in (5, 6, 7, 8):
            font = copy(template_sheet.cell(5, column).font)
            font.color = "FFFF0000"
            template_sheet.cell(5, column).font = font
        template_workbook.save(template_path)
        template_workbook.close()

        run_id = uuid4().hex
        working_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_SMM.xlsx"
        backup_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_backup.xlsx"

        sheet_name, _ = create_smm_workbook(
            ecm_record_path=ecm_path,
            smm_template_path=template_path,
            smm_record_path=working_path,
            backup_path=backup_path,
            start_date=date(2026, 7, 1),
            end_date=date(2026, 7, 13),
            ecm_worksheet_name="ECM_1-13 Jul (WIP)",
            summary_template_sheet="(OK) 1 Jul - 12 Jul",
        )

        workbook = load_workbook(working_path, data_only=False)
        worksheet = workbook[sheet_name]
        self.assertEqual(worksheet.cell(5, 6).value, 1)
        self.assertEqual(worksheet.cell(5, 6).font.color.rgb, "FFFF0000")
        self.assertEqual(worksheet.cell(6, 6).value, 0)
        self.assertEqual(worksheet.cell(6, 6).font.color.rgb, "FF000000")
        for row_number in (5, 6):
            for column in (5, 7, 8):
                self.assertEqual(
                    worksheet.cell(row_number, column).font.color.rgb,
                    "FF000000",
                )
        workbook.close()

    def test_replaces_existing_period_summary_on_rerun(self) -> None:
        ecm_path = self.create_ecm_workbook()
        template_path = self.create_smm_template()
        working_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_SMM.xlsx"
        backup_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_backup.xlsx"
        kwargs = {
            "ecm_record_path": ecm_path,
            "smm_template_path": template_path,
            "smm_record_path": working_path,
            "backup_path": backup_path,
            "start_date": date(2026, 7, 1),
            "end_date": date(2026, 7, 13),
            "ecm_worksheet_name": "ECM_1-13 Jul (WIP)",
            "summary_template_sheet": "(OK) 1 Jul - 12 Jul",
        }
        create_smm_workbook(**kwargs)
        create_smm_workbook(**kwargs)

        workbook = load_workbook(working_path, data_only=False)
        self.assertEqual(workbook.sheetnames.count("(WIP) 1 Jul - 13 Jul"), 1)
        self.assertEqual(workbook.active.title, "(WIP) 1 Jul - 13 Jul")
        workbook.close()

    def test_writes_final_outbound_after_mattex_deduction(self) -> None:
        ecm_path = self.create_ecm_workbook()
        ecm_workbook = load_workbook(ecm_path)
        ecm_worksheet = ecm_workbook["ECM_1-13 Jul (WIP)"]
        ecm_worksheet.cell(2, 13).value = "Added by Mattex"
        ecm_worksheet.cell(3, 13).value = "ADD By Chris's Team"
        ecm_workbook.save(ecm_path)
        ecm_workbook.close()

        template_path = self.create_smm_template()
        run_id = uuid4().hex
        working_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_SMM.xlsx"
        backup_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_backup.xlsx"

        sheet_name, _ = create_smm_workbook(
            ecm_record_path=ecm_path,
            smm_template_path=template_path,
            smm_record_path=working_path,
            backup_path=backup_path,
            start_date=date(2026, 7, 1),
            end_date=date(2026, 7, 13),
            ecm_worksheet_name="ECM_1-13 Jul (WIP)",
            summary_template_sheet="(OK) 1 Jul - 12 Jul",
        )

        workbook = load_workbook(working_path, data_only=False)
        worksheet = workbook[sheet_name]
        self.assertEqual(worksheet.cell(5, 5).value, 1)
        self.assertEqual(worksheet.cell(5, 6).value, 1)
        self.assertEqual(worksheet.cell(5, 7).value, "=E5+F5")
        self.assertEqual(worksheet.cell(5, 8).value, 0)
        self.assertEqual(worksheet.cell(6, 8).value, 1)
        workbook.close()

    def test_maps_1007_to_1016_and_preserves_readiness_display_value(self) -> None:
        ecm_path = self.create_ecm_workbook()
        ecm_workbook = load_workbook(ecm_path)
        ecm_worksheet = ecm_workbook["ECM_1-13 Jul (WIP)"]
        ecm_worksheet.cell(2, 1).value = 1007
        ecm_worksheet.cell(3, 9).value = "N"
        ecm_worksheet.cell(4, 9).value = "N"
        ecm_workbook.save(ecm_path)
        ecm_workbook.close()

        template_path = self.create_smm_template()
        template_workbook = load_workbook(template_path)
        readiness = template_workbook["Project Realiness Chris-In Use"]
        readiness.cell(2, 3).value = "1016 - 1/WSD/19 (=1007's record in ECM)"
        template_workbook.save(template_path)
        template_workbook.close()

        run_id = uuid4().hex
        working_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_SMM.xlsx"
        backup_path = TEST_WORKBOOK_DIRECTORY / f"{self._testMethodName}_{run_id}_backup.xlsx"

        sheet_name, project_count = create_smm_workbook(
            ecm_record_path=ecm_path,
            smm_template_path=template_path,
            smm_record_path=working_path,
            backup_path=backup_path,
            start_date=date(2026, 7, 1),
            end_date=date(2026, 7, 13),
            ecm_worksheet_name="ECM_1-13 Jul (WIP)",
            summary_template_sheet="(OK) 1 Jul - 12 Jul",
        )

        self.assertEqual(project_count, 1)
        workbook = load_workbook(working_path, data_only=False)
        worksheet = workbook[sheet_name]
        self.assertEqual(worksheet.cell(5, 2).value, "1016 - 1/WSD/19 (1007)")
        workbook.close()


if __name__ == "__main__":
    unittest.main()
