from __future__ import annotations

from dataclasses import replace
from datetime import date
from pathlib import Path

import pytest

from control_panel_v3.services import config
from control_panel_v3.services.config import ConfigurationError
from control_panel_v3.services.workflow_builders import PreflightError, prepare_workflow


@pytest.fixture
def settings(monkeypatch: pytest.MonkeyPatch):
    # Do not load an operator's local .env or inherit process-specific settings.
    monkeypatch.setattr(config, "load_dotenv", lambda *args, **kwargs: None)
    for name in tuple(config.os.environ):
        if name.startswith("CONTROL_PANEL_V3_"):
            monkeypatch.delenv(name)
    return config.load_settings()


def build(settings, workflow_id="ecm-step-0"):
    return prepare_workflow(workflow_id, date(2026, 11, 9), date(2026, 11, 15), 123, settings)


def value(command: list[str], flag: str) -> str:
    return command[command.index(flag) + 1]


def test_defaults_and_step0_command(settings) -> None:
    assert (settings.step0_headed, settings.step0_login_if_needed) == (True, True)
    assert (settings.step0_timeout_seconds, settings.step0_auth_timeout_seconds, settings.step0_failure_browser_hold_seconds) == (60, 120, 120)
    assert settings.worker_timeout_seconds == 180
    assert settings.validate_step0() == []
    prepared = build(settings)
    assert "--headed" in prepared.command
    assert "--login-if-needed" in prepared.command
    assert value(prepared.command, "--timeout-seconds") == "60"
    assert value(prepared.command, "--auth-timeout-seconds") == "120"
    assert value(prepared.command, "--failure-browser-hold-seconds") == "120"
    assert "--reset-profile" not in prepared.command
    assert value(prepared.command, "--root-dir") == str(config.CONTROL_PANEL_ROOT)
    assert all(artifact.path.is_relative_to(config.CONTROL_PANEL_ROOT) for artifact in prepared.artifacts)


def test_environment_overrides_and_headless_command(settings, monkeypatch: pytest.MonkeyPatch) -> None:
    for name, val in {
        "CONTROL_PANEL_V3_STEP0_HEADED": "false",
        "CONTROL_PANEL_V3_STEP0_LOGIN_IF_NEEDED": "no",
        "CONTROL_PANEL_V3_STEP0_TIMEOUT_SECONDS": "75",
        "CONTROL_PANEL_V3_STEP0_AUTH_TIMEOUT_SECONDS": "240",
        "CONTROL_PANEL_V3_STEP0_FAILURE_BROWSER_HOLD_SECONDS": "0",
        "CONTROL_PANEL_V3_RESET_EDGE_PROFILE": "true",
    }.items():
        monkeypatch.setenv(name, val)
    configured = config.load_settings()
    command = build(configured).command
    assert "--no-headed" in command and "--no-login-if-needed" in command
    from control_panel_v3.workers.ecm_step0_export import build_parser
    args = build_parser().parse_args(command[3:])
    assert not args.headed and not args.login_if_needed
    assert "--reset-profile" in command
    assert value(command, "--timeout-seconds") == "75"
    assert value(command, "--auth-timeout-seconds") == "240"
    assert value(command, "--failure-browser-hold-seconds") == "0"


@pytest.mark.parametrize("name,bad", [
    ("CONTROL_PANEL_V3_STEP0_HEADED", "maybe"),
    ("CONTROL_PANEL_V3_STEP0_LOGIN_IF_NEEDED", "perhaps"),
    ("CONTROL_PANEL_V3_STEP0_TIMEOUT_SECONDS", "0"),
    ("CONTROL_PANEL_V3_STEP0_TIMEOUT_SECONDS", "1.5"),
    ("CONTROL_PANEL_V3_STEP0_AUTH_TIMEOUT_SECONDS", "-1"),
    ("CONTROL_PANEL_V3_STEP0_FAILURE_BROWSER_HOLD_SECONDS", "-1"),
    ("CONTROL_PANEL_V3_STEP0_FAILURE_BROWSER_HOLD_SECONDS", "bad"),
])
def test_invalid_environment_is_rejected(settings, monkeypatch: pytest.MonkeyPatch, name: str, bad: str) -> None:
    monkeypatch.setenv(name, bad)
    with pytest.raises(ConfigurationError, match=name):
        config.load_settings()


@pytest.mark.parametrize("field,bad", [
    ("step0_headed", "false"),
    ("step0_login_if_needed", 0),
    ("step0_timeout_seconds", 0),
    ("step0_timeout_seconds", True),
    ("step0_auth_timeout_seconds", 1.5),
    ("step0_failure_browser_hold_seconds", -1),
])
def test_direct_settings_validation_and_builder(settings, field: str, bad: object) -> None:
    invalid = replace(settings, **{field: bad})
    assert invalid.validate_step0()
    assert invalid.validate_for_workflow("ecm-step-0")
    with pytest.raises(PreflightError):
        build(invalid)


def test_job_timeout_minimum_is_step0_specific(settings) -> None:
    too_short = replace(settings, job_timeout_seconds=300)
    assert any("JOB_TIMEOUT" in error for error in too_short.validate_step0())
    assert any("JOB_TIMEOUT" in error for error in too_short.validate_for_workflow("ecm-step-0"))
    assert not any("JOB_TIMEOUT" in error for error in too_short.validate_for_workflow("ecm-first-date-lookup"))
    with pytest.raises(PreflightError, match="JOB_TIMEOUT"):
        build(too_short)
    assert replace(settings, job_timeout_seconds=301).validate_step0() == []


def test_lookup_retains_shared_timeout(settings, monkeypatch: pytest.MonkeyPatch) -> None:
    # Simulate existing prerequisites without creating workbooks or launching a worker.
    monkeypatch.setattr(Path, "is_file", lambda self: True)
    configured = replace(settings, worker_timeout_seconds=211, step0_timeout_seconds=70)
    command = build(configured, "ecm-first-date-lookup").command
    assert value(command, "--timeout-seconds") == "211"
    assert "--headed" not in command
    assert "--failure-browser-hold-seconds" not in command


def test_step0_builder_preserves_root_guard(settings, tmp_path: Path) -> None:
    with pytest.raises(PreflightError, match="remain under"):
        build(replace(settings, edge_profile_dir=tmp_path / "outside"))