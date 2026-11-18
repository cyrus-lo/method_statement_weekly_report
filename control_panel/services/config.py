"""Configuration for the localhost workflow control panel."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

CONTROL_PANEL_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = CONTROL_PANEL_ROOT.parent


class ConfigurationError(ValueError):
    """Raised when local application settings are incomplete or unsafe."""


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


@dataclass(frozen=True)
class Settings:
    project_root: Path
    data_dir: Path
    django_python: Path
    ecm_python: Path
    django_script: Path
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
    ecm_period_script: Path | None = None

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
        """Return setup problems without preventing the dashboard from explaining them."""
        errors: list[str] = []
        required_files = {
            "Django Python interpreter": self.django_python,
            "ECM Python interpreter": self.ecm_python,
            "Weekly report script": self.django_script,
            "ECM Step 1a script": self.step_1a_script,
            "ECM Step 1b script": self.step_1b_script,
            "ECM Step 2 script": self.step_2_script,
            "Weekly report template": self.django_template,
            "ECM template": self.ecm_template,
            "SMM template": self.smm_template,
        }
        if self.ecm_period_script is not None:
            required_files["ECM period report script"] = self.ecm_period_script
        for label, path in required_files.items():
            if path is None or not path.is_file():
                errors.append(f"{label} was not found: {path}")
        for label, path in {
            "Weekly report output folder": self.django_output_dir,
            "ECM CSV folder": self.ecm_csv_dir,
            "ECM output folder": self.ecm_record.parent,
            "SMM output folder": self.smm_record.parent,
        }.items():
            if not path.is_dir():
                errors.append(f"{label} was not found: {path}")
        if self.data_dir == self.django_output_dir or self.data_dir == self.ecm_record.parent:
            errors.append("CONTROL_PANEL_DATA_DIR must not be an operational workflow output folder.")
        try:
            self.ensure_application_directories()
        except OSError as error:
            errors.append(f"Application data directory is not writable: {error}")
        return errors


def load_settings() -> Settings:
    """Load non-secret UI settings from the local control-panel .env file."""
    load_dotenv(CONTROL_PANEL_ROOT / ".env")
    return Settings(
        project_root=PROJECT_ROOT,
        data_dir=_path_setting("CONTROL_PANEL_DATA_DIR", "control_panel/data"),
        django_python=_path_setting("DJANGO_PYTHON", "django/.venv/Scripts/python.exe"),
        ecm_python=_path_setting("ECM_PYTHON", "ecm/.venv/Scripts/python.exe"),
        django_script=PROJECT_ROOT / "django" / "Method_Statement_Interaction_Weekly_Report.py",
        step_1a_script=PROJECT_ROOT / "ecm" / "Step_1a_Create_ECM_Record.py",
        step_1b_script=PROJECT_ROOT / "ecm" / "Step_1b_Finalize_ECM_Record.py",
        step_2_script=PROJECT_ROOT / "ecm" / "Step_2_Create_SMM_Summary.py",
        ecm_period_script=PROJECT_ROOT / "ecm" / "Build_ECM_Period_Report.py",
        django_template=_path_setting(
            "DJANGO_TEMPLATE", "django/Method_Statement_Interaction_Weekly_Report_template.xlsx"
        ),
        django_output_dir=_path_setting("DJANGO_OUTPUT_DIR", "django/output"),
        ecm_csv_dir=_path_setting("ECM_CSV_DIR", "ecm/csv_from_ecm"),
        ecm_template=_path_setting("ECM_TEMPLATE", "ecm/ECM_Record_template.xlsx"),
        ecm_record=_path_setting("ECM_RECORD", "ecm/output/ECM_Record.xlsx"),
        smm_template=_path_setting("SMM_TEMPLATE", "ecm/SMM_KMS_PM_Master_Document_template.xlsx"),
        smm_record=_path_setting("SMM_RECORD", "ecm/output/SMM_KMS_PM_Master_Document.xlsx"),
        smm_summary_template_sheet=os.getenv(
            "SMM_SUMMARY_TEMPLATE_SHEET", "(OK) 1 Jul - 12 Jul"
        ).strip(),
        job_timeout_seconds=_positive_int("JOB_TIMEOUT_SECONDS", 7200),
    )
