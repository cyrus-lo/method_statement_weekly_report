"""Configuration for the isolated Workflow Control Panel V2 preview."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


CONTROL_PANEL_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = CONTROL_PANEL_ROOT.parent


class ConfigurationError(ValueError):
    """Raised when V2 configuration is invalid."""


def _path_setting(name: str, default: str) -> Path:
    value = Path(os.getenv(name, default)).expanduser()
    if not value.is_absolute():
        value = PROJECT_ROOT / value
    return value.resolve()


def _positive_int(name: str, default: int) -> int:
    value = os.getenv(name, str(default))
    try:
        parsed = int(value)
    except ValueError as error:
        raise ConfigurationError(f"{name} must be a whole number; got {value!r}.") from error
    if parsed <= 0:
        raise ConfigurationError(f"{name} must be greater than zero.")
    return parsed


def _nonnegative_int(name: str, default: int) -> int:
    value = os.getenv(name, str(default))
    try:
        parsed = int(value)
    except ValueError as error:
        raise ConfigurationError(f"{name} must be a whole number; got {value!r}.") from error
    if parsed < 0:
        raise ConfigurationError(f"{name} must not be negative.")
    return parsed


def _boolean(name: str, default: bool) -> bool:
    value = os.getenv(name, str(default)).strip().casefold()
    if value in {"1", "true", "yes", "on"}:
        return True
    if value in {"0", "false", "no", "off"}:
        return False
    raise ConfigurationError(f"{name} must be true or false; got {value!r}.")


@dataclass(frozen=True)
class Settings:
    project_root: Path
    data_dir: Path
    preview_only: bool
    django_python: Path
    ecm_python: Path
    step0_python: Path
    django_script: Path
    step0_script: Path
    step_1a_script: Path
    step_1b_script: Path
    step_2_script: Path
    django_template: Path
    django_output_dir: Path
    ecm_csv_dir: Path
    ecm_template: Path
    ecm_record: Path
    smm_template: Path
    smm_record: Path
    smm_summary_template_sheet: str
    job_timeout_seconds: int
    step0_timeout_seconds: int
    step0_auth_timeout_seconds: int
    step0_login_if_needed: bool
    step0_headed: bool
    step0_failure_browser_hold_seconds: int
    step0_reset_profile: bool
    ecm_url: str
    step0_profile_dir: Path
    uses_period_csv_handoff: bool

    @property
    def database_path(self) -> Path:
        return self.data_dir / "history.sqlite3"

    @property
    def runs_dir(self) -> Path:
        return self.data_dir / "runs"

    def ensure_application_directories(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.runs_dir.mkdir(parents=True, exist_ok=True)

    def validate(self) -> list[str]:
        errors: list[str] = []
        if self.data_dir == (self.project_root / "control_panel" / "data").resolve():
            errors.append("CONTROL_PANEL_V2_DATA_DIR must not point to V1 control_panel/data.")
        required_files = {
            "Django Python interpreter": self.django_python,
            "ECM Python interpreter": self.ecm_python,
            "Step 0 Python interpreter": self.step0_python,
            "Weekly report script": self.django_script,
            "Step 0 worker script": self.step0_script,
            "ECM Step 1a script": self.step_1a_script,
            "ECM Step 1b script": self.step_1b_script,
            "ECM Step 2 script": self.step_2_script,
            "Weekly report template": self.django_template,
            "ECM template": self.ecm_template,
            "SMM template": self.smm_template,
        }
        for label, path in required_files.items():
            if not path.is_file():
                errors.append(f"{label} was not found: {path}")
        if self.step0_python.is_file():
            try:
                import subprocess

                probe = subprocess.run(
                    [str(self.step0_python), "-c", "import playwright"],
                    capture_output=True,
                    text=True,
                    timeout=10,
                )
            except (OSError, subprocess.SubprocessError) as error:
                errors.append(f"Step 0 Python environment could not be checked: {error}")
            else:
                if probe.returncode != 0:
                    errors.append("Step 0 Python environment is missing Playwright.")
        for label, path in {
            "Weekly report output folder": self.django_output_dir,
            "ECM CSV folder": self.ecm_csv_dir,
            "ECM output folder": self.ecm_record.parent,
            "SMM output folder": self.smm_record.parent,
        }.items():
            if not path.is_dir():
                errors.append(f"{label} was not found: {path}")
        if self.step0_timeout_seconds <= 0:
            errors.append("STEP0_TIMEOUT_SECONDS must be greater than zero.")
        if self.step0_auth_timeout_seconds <= 0:
            errors.append("STEP0_AUTH_TIMEOUT_SECONDS must be greater than zero.")
        if self.step0_failure_browser_hold_seconds < 0:
            errors.append("STEP0_FAILURE_BROWSER_HOLD_SECONDS must not be negative.")
        if not self.ecm_url:
            errors.append("ECM_BASE_URL must not be empty.")
        if self.step0_profile_dir == (self.project_root / "control_panel" / "data").resolve():
            errors.append("STEP0_EDGE_PROFILE_DIR must not point to V1 control_panel/data.")
        try:
            self.ensure_application_directories()
            self.step0_profile_dir.parent.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            errors.append(f"V2 data directory is not writable: {error}")
        return errors


def load_settings() -> Settings:
    load_dotenv(CONTROL_PANEL_ROOT / ".env")
    return Settings(
        project_root=PROJECT_ROOT,
        data_dir=_path_setting("CONTROL_PANEL_V2_DATA_DIR", "control_panel_v2/data"),
        preview_only=_boolean("CONTROL_PANEL_V2_PREVIEW_ONLY", True),
        django_python=_path_setting("DJANGO_PYTHON", "django/.venv/Scripts/python.exe"),
        ecm_python=_path_setting("ECM_PYTHON", "ecm/.venv/Scripts/python.exe"),
        step0_python=_path_setting(
            "STEP0_PYTHON", "ecm/step0_worker/.venv/Scripts/python.exe"
        ),
        django_script=PROJECT_ROOT / "django" / "Method_Statement_Interaction_Weekly_Report.py",
        step0_script=PROJECT_ROOT / "ecm" / "step0_worker" / "export_worker.py",
        step_1a_script=PROJECT_ROOT / "ecm" / "Step_1a_Create_ECM_Record.py",
        step_1b_script=PROJECT_ROOT / "ecm" / "Step_1b_Finalize_ECM_Record.py",
        step_2_script=PROJECT_ROOT / "ecm" / "Step_2_Create_SMM_Summary.py",
        django_template=_path_setting("DJANGO_TEMPLATE", "django/Method_Statement_Interaction_Weekly_Report_template.xlsx"),
        django_output_dir=_path_setting("DJANGO_OUTPUT_DIR", "django/output"),
        ecm_csv_dir=_path_setting("ECM_CSV_DIR", "ecm/csv_from_ecm"),
        ecm_template=_path_setting("ECM_TEMPLATE", "ecm/ECM_Record_template.xlsx"),
        ecm_record=_path_setting("ECM_RECORD", "ecm/output/ECM_Record.xlsx"),
        smm_template=_path_setting("SMM_TEMPLATE", "ecm/SMM_KMS_PM_Master_Document_template.xlsx"),
        smm_record=_path_setting("SMM_RECORD", "ecm/output/SMM_KMS_PM_Master_Document.xlsx"),
        smm_summary_template_sheet=os.getenv("SMM_SUMMARY_TEMPLATE_SHEET", "(OK) 1 Jul - 12 Jul").strip(),
        job_timeout_seconds=_positive_int("JOB_TIMEOUT_SECONDS", 7200),
        step0_timeout_seconds=_positive_int("STEP0_TIMEOUT_SECONDS", 60),
        step0_auth_timeout_seconds=_positive_int("STEP0_AUTH_TIMEOUT_SECONDS", 120),
        step0_login_if_needed=_boolean("STEP0_LOGIN_IF_NEEDED", True),
        step0_headed=_boolean("STEP0_HEADED", True),
        step0_failure_browser_hold_seconds=_nonnegative_int(
            "STEP0_FAILURE_BROWSER_HOLD_SECONDS", 120
        ),
        step0_reset_profile=_boolean("STEP0_RESET_PROFILE", False),
        ecm_url=os.getenv("ECM_BASE_URL", "https://ecm.chunwo.com/navigator/").strip(),
        step0_profile_dir=_path_setting(
            "STEP0_EDGE_PROFILE_DIR", "control_panel_v2/.step0-edge-profile"
        ),
        uses_period_csv_handoff=True,
    )
