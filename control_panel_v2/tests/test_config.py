from __future__ import annotations

from pathlib import Path

from control_panel_v2.services.config import load_settings


def test_v2_defaults_keep_preview_mode_and_dedicated_step0_worker(monkeypatch) -> None:
    monkeypatch.delenv("CONTROL_PANEL_V2_PREVIEW_ONLY", raising=False)
    monkeypatch.delenv("STEP0_PYTHON", raising=False)
    monkeypatch.delenv("STEP0_EDGE_PROFILE_DIR", raising=False)

    settings = load_settings()

    assert settings.preview_only is True
    assert settings.step0_python == (
        settings.project_root / "ecm" / "step0_worker" / ".venv" / "Scripts" / "python.exe"
    ).resolve()
    assert settings.step0_profile_dir == (
        settings.project_root / "control_panel_v2" / ".step0-edge-profile"
    ).resolve()
    assert settings.step0_headed is True
    assert settings.step0_failure_browser_hold_seconds == 120
    assert settings.step0_reset_profile is False


def test_v2_can_override_step0_worker_paths(monkeypatch, tmp_path: Path) -> None:
    worker_python = tmp_path / "python.exe"
    profile_dir = tmp_path / "edge-profile"
    monkeypatch.setenv("STEP0_PYTHON", str(worker_python))
    monkeypatch.setenv("STEP0_EDGE_PROFILE_DIR", str(profile_dir))
    monkeypatch.setenv("CONTROL_PANEL_V2_PREVIEW_ONLY", "true")

    settings = load_settings()

    assert settings.step0_python == worker_python.resolve()
    assert settings.step0_profile_dir == profile_dir.resolve()