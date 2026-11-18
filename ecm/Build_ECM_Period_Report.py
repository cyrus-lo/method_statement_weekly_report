"""Append approved ECM sheets and create a new period SMM summary."""
from __future__ import annotations

import argparse
import shutil
import sys
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from Step_2_Create_SMM_Summary import (  # noqa: E402
    aggregate_ecm_record,
    print_metrics_audit,
    read_project_metadata,
    rebuild_summary_sheet,
    summary_sheet_name,
    validate_summary_layout,
)
from display import write_display_values  # noqa: E402


def parse_date(value: str) -> date:
    return date.fromisoformat(value)


def period_token(start: date, end: date) -> str:
    if start.month == end.month and start.year == end.year:
        return f"{start:%Y-%m}"
    if start.month in (1, 4, 7, 10) and end.month == start.month + 2:
        return f"{start:%Y}Q{((start.month - 1) // 3) + 1}"
    return f"{start:%Y%m%d}_{end:%Y%m%d}"


def source_sheets(workbook, target: str, selected: list[str]) -> list[str]:
    # Only human-approved aggregation sheets are accepted as sources.
    sheets = selected or [name for name in workbook.sheetnames if name.endswith("(OK)") and name != target]
    if any(name not in workbook.sheetnames or not name.endswith("(OK)") for name in sheets):
        raise ValueError("Only existing ECM sheets marked (OK) may be selected as sources.")
    if not sheets:
        raise ValueError("No approved (OK) ECM sheets were found; WIP sheets are not valid sources.")
    return sheets


def append_sources(workbook, names: list[str], target: str) -> int:
    if target in workbook.sheetnames:
        raise ValueError(f"Target ECM sheet already exists: {target}. Restore the workbook before rerun.")
    output = workbook.copy_worksheet(workbook[names[0]])
    output.title = target
    for row in output.iter_rows(min_row=2):
        for cell in row:
            cell.value = None
    total = 0
    target_row = 2
    for name in names:
        source = workbook[name]
        if source.max_column < 11:
            raise ValueError(f"ECM source sheet has too few columns: {name}")
        for row in source.iter_rows(min_row=2, values_only=True):
            for column, value in enumerate(row, start=1):
                output.cell(target_row, column).value = value
            target_row += 1
            total += 1
    if total < 2:
        raise ValueError("Selected ECM sheets contain no data rows.")
    return total


def add_audit_sheet(workbook, name: str, headers: list[str], rows: list[list[object]]) -> None:
    if name in workbook.sheetnames:
        raise ValueError(f"Audit sheet already exists: {name}. Restore the workbook before rerun.")
    sheet = workbook.create_sheet(name)
    sheet.append(headers)
    for row in rows:
        sheet.append(row)


def source_register_rows(names: list[str], workbook) -> list[list[object]]:
    return [[name, workbook[name].max_row - 1, workbook[name].max_column, "OK"] for name in names]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--start-date", required=True, type=parse_date)
    parser.add_argument("--end-date", required=True, type=parse_date)
    parser.add_argument("--ecm-record", type=Path, required=True)
    parser.add_argument("--smm-record", type=Path, required=True)
    parser.add_argument("--smm-template", type=Path, required=True)
    parser.add_argument("--summary-template-sheet", required=True)
    parser.add_argument("--ecm-backup", type=Path, required=True)
    parser.add_argument("--smm-backup", type=Path, required=True)
    parser.add_argument("--source-sheet", action="append", default=[])
    args = parser.parse_args()
    if args.end_date < args.start_date:
        parser.error("--end-date cannot be earlier than --start-date")
    for path in (args.ecm_record, args.smm_record):
        lock = path.with_name(f"~${path.name}")
        if lock.exists():
            raise ValueError(f"Close the workbook in Excel before running: {path.name}")
        if not path.is_file():
            raise FileNotFoundError(path)
    ecm_target = f"ECM_{period_token(args.start_date, args.end_date)} (WIP)"
    smm_target = summary_sheet_name(args.start_date, args.end_date)
    ecm = load_workbook(args.ecm_record)
    try:
        sources = source_sheets(ecm, ecm_target, args.source_sheet)
        if ecm_target in ecm.sheetnames:
            raise ValueError(f"Target ECM sheet already exists: {ecm_target}. Restore before rerun.")
        smm = load_workbook(args.smm_record)
        try:
            if smm_target in smm.sheetnames:
                raise ValueError(f"Target SMM sheet already exists: {smm_target}. Restore before rerun.")
        finally:
            smm.close()
    finally:
        ecm.close()
    args.ecm_backup.parent.mkdir(parents=True, exist_ok=True)
    args.smm_backup.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(args.ecm_record, args.ecm_backup)
    shutil.copy2(args.smm_record, args.smm_backup)

    ecm = load_workbook(args.ecm_record)
    try:
        rows = append_sources(ecm, sources, ecm_target)
        add_audit_sheet(
            ecm,
            f"{ecm_target} Sources",
            ["Source worksheet", "Data rows", "Columns", "Status"],
            source_register_rows(sources, ecm),
        )
        ecm.save(args.ecm_record)
    finally:
        ecm.close()

    # Build metrics from the newly appended sheet using the existing Step 2 rules.
    temp = args.ecm_record.with_name(f".{args.ecm_record.stem}.{ecm_target.replace(' ', '_')}.tmp.xlsx")
    shutil.copy2(args.ecm_record, temp)
    try:
        metrics = aggregate_ecm_record(temp, ecm_target)
        if not metrics:
            raise ValueError("The appended ECM rows contain no Count=Y records.")
        write_display_values(args.ecm_record, ecm_target)
        print_metrics_audit(metrics)
        smm = load_workbook(args.smm_record)
        try:
            if args.summary_template_sheet not in smm.sheetnames:
                raise ValueError(f"Summary layout sheet not found: {args.summary_template_sheet}")
            metadata = read_project_metadata(smm)
            template = smm[args.summary_template_sheet]
            validate_summary_layout(template)
            if smm_target in smm.sheetnames:
                raise ValueError(f"Target SMM sheet already exists: {smm_target}. Restore before rerun.")
            output = smm.copy_worksheet(template)
            output.title = smm_target
            rebuild_summary_sheet(output, args.start_date, args.end_date, metrics, metadata)
            add_audit_sheet(
                smm,
                f"{smm_target} Alerts",
                ["Item", "Value"],
                [
                    ["Period", f"{args.start_date} to {args.end_date}"],
                    ["Source sheets", ", ".join(sources)],
                    ["Appended ECM rows", rows],
                    ["Classification projects", len(metrics)],
                    ["Overlap/duplicate review", "Human review required"],
                    ["Aggregation status", "WIP - change ECM sheet status to (OK) after review"],
                ],
            )
            smm.save(args.smm_record)
        finally:
            smm.close()
    finally:
        temp.unlink(missing_ok=True)
    print(f"Created {ecm_target} with {rows} appended rows and {smm_target}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
