from __future__ import annotations

import sys
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from typing import cast

import pytest

from control_panel.services.config import Settings
from control_panel.services.workflows import PreflightError, prepare_workflow


def build_settings(tmp_path: Path) -> Settings:
    scripts = tmp_path / "scripts"
    scripts.mkdir()
    for script in ("weekly.py", "step1a.py", "step1b.py", "step2.py"):
        (scripts / script).write_text("print('ok')\n", encoding="utf-8")
    django_output = tmp_path / "django-output"
    ecm_output = tmp_path / "ecm-output"
    csv_dir = tmp_path / "csv"
    data_dir = tmp_path / "data"
    for directory in (django_output, ecm_output, csv_dir):
        directory.mkdir()
    (csv_dir / "properties 1.csv").write_text("header\n", encoding="utf-8")
    for workbook in ("django-template.xlsx", "ecm-template.xlsx", "smm-template.xlsx"):
        (tmp_path / workbook).write_text("placeholder", encoding="utf-8")
    return Settings(
        project_root=tmp_path,
        data_dir=data_dir,
        django_python=Path(sys.executable),
        ecm_python=Path(sys.executable),
        django_script=scripts / "weekly.py",
        step_1a_script=scripts / "step1a.py",
        step_1b_script=scripts / "step1b.py",
        step_2_script=scripts / "step2.py",
        django_template=tmp_path / "django-template.xlsx",
        django_output_dir=django_output,
        ecm_csv_dir=csv_dir,
        ecm_template=tmp_path / "ecm-template.xlsx",
        ecm_record=ecm_output / "ECM_Record.xlsx",
        smm_template=tmp_path / "smm-template.xlsx",
        smm_record=ecm_output / "SMM.xlsx",
        smm_summary_template_sheet="Template",
        job_timeout_seconds=30,
    )


def test_step_0_command_uses_period_dates_and_timestamped_csv(tmp_path: Path) -> None:
    base_settings = build_settings(tmp_path)
    step0_script = tmp_path / "scripts" / "step0.py"
    step0_script.write_text("print('ok')\n", encoding="utf-8")
    settings = cast(
        Settings,
        SimpleNamespace(
            ecm_csv_dir=base_settings.ecm_csv_dir,
            step0_python=Path(sys.executable),
            step0_script=step0_script,
            step0_profile_dir=tmp_path / "edge-profile",
            ecm_url="https://ecm.example.test/navigator/",
            step0_timeout_seconds=900,
            step0_headed=True,
            step0_failure_browser_hold_seconds=300,
            step0_reset_profile=True,
            uses_period_csv_handoff=True,
        ),
    )

    prepared = prepare_workflow("ecm-step-0", date(2026, 7, 20), date(2026, 7, 26), settings)

    assert prepared.command[:2] == [str(Path(sys.executable)), str(settings.step0_script)]
    assert prepared.command[prepared.command.index("--start-date") + 1] == "2026-07-20"
    assert prepared.command[prepared.command.index("--end-date") + 1] == "2026-07-26"
    period_input = Path(prepared.command[prepared.command.index("--period-input") + 1])
    assert period_input.name == "properties_2026-07-20_to_2026-07-26.csv"
    output = prepared.artifacts[0].path
    assert output.parent == settings.ecm_csv_dir
    assert output.name.startswith("properties_2026-07-20_to_2026-07-26_")
    assert prepared.artifacts[1].path == period_input
    assert "--headed" in prepared.command
    hold_index = prepared.command.index("--failure-browser-hold-seconds") + 1
    assert prepared.command[hold_index] == "300"
    assert "--reset-profile" in prepared.command


def test_v2_step_0_command_can_request_login_if_needed(tmp_path: Path) -> None:
    base_settings = build_settings(tmp_path)
    step0_script = tmp_path / "scripts" / "step0.py"
    step0_script.write_text("print('ok')\n", encoding="utf-8")
    settings = cast(
        Settings,
        SimpleNamespace(
            ecm_csv_dir=base_settings.ecm_csv_dir,
            step0_python=Path(sys.executable),
            step0_script=step0_script,
            step0_profile_dir=tmp_path / "edge-profile",
            ecm_url="https://ecm.example.test/navigator/",
            step0_timeout_seconds=900,
            step0_auth_timeout_seconds=300,
            step0_login_if_needed=True,
            uses_period_csv_handoff=True,
        ),
    )

    prepared = prepare_workflow("ecm-step-0", date(2026, 7, 20), date(2026, 7, 26), settings)

    assert "--login-if-needed" in prepared.command
    auth_timeout_index = prepared.command.index("--auth-timeout-seconds") + 1
    assert prepared.command[auth_timeout_index] == "300"


def test_weekly_report_command_uses_argument_list(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    prepared = prepare_workflow("weekly-report", date(2026, 7, 1), date(2026, 7, 7), settings)

    assert prepared.command[:2] == [str(Path(sys.executable)), str(settings.django_script)]
    assert "--start-date" in prepared.command
    assert prepared.artifacts[0].path.name.endswith("2026-07-01_to_2026-07-07.xlsx")


def test_step_1a_selects_newest_properties_csv(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    newest = settings.ecm_csv_dir / "properties 2.csv"
    newest.write_text("header\n", encoding="utf-8")
    newest.touch()

    prepared = prepare_workflow("ecm-step-1a", date(2026, 7, 1), date(2026, 7, 7), settings)

    csv_index = prepared.command.index("--csv") + 1
    assert Path(prepared.command[csv_index]) == newest
    assert prepared.artifacts[1].path.name.endswith("_1a.xlsx")


def test_step_1b_passes_post_backup_and_collects_all_handoff_artifacts(tmp_path: Path) -> None:
    settings = build_settings(tmp_path)
    settings.ecm_record.write_text("placeholder", encoding="utf-8")

    prepared = prepare_workflow("ecm-step-1b", date(2026, 7, 1), date(2026, 7, 7), settings)

    post_backup_index = prepared.command.index("--post-1b-backup") + 1
    assert Path(prepared.command[post_backup_index]).name.endswith("_1b_after.xlsx")
    assert [artifact.name for artifact in prepared.artifacts] == [
        "Finalized ECM review workbook",
        "Step 1b backup",
        "Post-Step 1b backup",
    ]
    assert prepared.artifacts[2].path.name.endswith("_1b_after.xlsx")


def test_rejects_invalid_period(tmp_path: Path) -> None:
    with pytest.raises(PreflightError, match="End date"):
        prepare_workflow("weekly-report", date(2026, 7, 7), date(2026, 7, 1), build_settings(tmp_path))
