"""V3 configuration with strict writable-path isolation from V1 and V2."""

from __future__ import annotations

import os
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from dotenv import dotenv_values, load_dotenv

CONTROL_PANEL_ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = CONTROL_PANEL_ROOT.parent
V1_ROOT = PROJECT_ROOT / "control_panel"
V2_ROOT = PROJECT_ROOT / "control_panel_v2"
DJANGO_ENV_PATH = PROJECT_ROOT / "django" / ".env"
WEEKLY_REQUIRED_ENV_KEYS = ("PGHOST", "PGDATABASE", "PGUSER", "PGPASSWORD")


class ConfigurationError(ValueError):
    """Raised when V3 settings violate the local configuration contract."""


def load_weekly_report_environment(env_path: Path | None = None) -> dict[str, str]:
    """Read and validate weekly-report settings without changing process env."""
    path = env_path or DJANGO_ENV_PATH
    try:
        parsed = dotenv_values(path, interpolate=False)
    except (OSError, UnicodeError) as error:
        raise ConfigurationError("Could not read weekly-report settings from django/.env.") from error

    missing = [
        key for key in WEEKLY_REQUIRED_ENV_KEYS
        if not isinstance(parsed.get(key), str) or not parsed[key].strip()
    ]
    if missing:
        raise ConfigurationError("Missing or blank django/.env setting(s): " + ", ".join(missing))

    port_value = parsed.get("PGPORT")
    port_text = port_value.strip() if isinstance(port_value, str) else ""
    if not port_text:
        port_text = "5432"
    try:
        port = int(port_text)
    except ValueError as error:
        raise ConfigurationError("django/.env PGPORT must be a whole number from 1 to 65535.") from error
    if not 1 <= port <= 65535:
        raise ConfigurationError("django/.env PGPORT must be a whole number from 1 to 65535.")

    timezone_value = parsed.get("REPORTING_TIMEZONE")
    timezone_name = "Asia/Hong_Kong" if timezone_value is None else timezone_value.strip()
    if not timezone_name:
        raise ConfigurationError("django/.env REPORTING_TIMEZONE must not be blank.")
    try:
        ZoneInfo(timezone_name)
    except (ValueError, ZoneInfoNotFoundError) as error:
        raise ConfigurationError("django/.env REPORTING_TIMEZONE is not a recognized timezone.") from error

    result = {key: parsed[key] for key in WEEKLY_REQUIRED_ENV_KEYS}
    result["PGPORT"] = str(port)
    sslmode = parsed.get("PGSSLMODE")
    if isinstance(sslmode, str) and sslmode.strip():
        result["PGSSLMODE"] = sslmode.strip()
    result["REPORTING_TIMEZONE"] = timezone_name
    return result


def _path_setting(name: str, default: str) -> Path:
    value = Path(os.getenv(name, default)).expanduser()
    if not value.is_absolute():
        value = PROJECT_ROOT / value
    return value.resolve()


def _boolean(name: str, default: bool) -> bool:
    value = os.getenv(name, str(default)).strip().casefold()
    if value in {"1", "true", "yes", "on"}:
        return True
    if value in {"0", "false", "no", "off"}:
        return False
    raise ConfigurationError(f"{name} must be true or false; got {value!r}.")


def _positive_int(name: str, default: int) -> int:
    value = os.getenv(name, str(default))
    try:
        result = int(value)
    except ValueError as error:
        raise ConfigurationError(f"{name} must be a whole number; got {value!r}.") from error
    if result <= 0:
        raise ConfigurationError(f"{name} must be greater than zero.")
    return result


def _nonnegative_int(name: str, default: int) -> int:
    value = os.getenv(name, str(default))
    try:
        result = int(value)
    except ValueError as error:
        raise ConfigurationError(f"{name} must be a whole number; got {value!r}.") from error
    if result < 0:
        raise ConfigurationError(f"{name} must be zero or greater.")
    return result


def _is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


@dataclass(frozen=True)
class Settings:
    project_root: Path
    data_dir: Path
    archive_dir: Path
    csv_dir: Path
    ecm_output_dir: Path
    smm_output_dir: Path
    django_output_dir: Path
    edge_profile_dir: Path
    preview_only: bool
    host: str
    port: int
    job_timeout_seconds: int
    ecm_url: str
    worker_timeout_seconds: int
    excel_filter_timeout_seconds: int = 90
    ecm_template_path: Path = CONTROL_PANEL_ROOT / "working" / "templates" / "ECM_Record_template.xlsx"
    smm_template_path: Path = CONTROL_PANEL_ROOT / "working" / "templates" / "SMM_KMS_PM_Master_Document_template.xlsx"
    readiness_workbook_path: Path = CONTROL_PANEL_ROOT / "working" / "readiness" / "SMM_KMS_PM_Master_Document.xlsx"
    weekly_template_path: Path = CONTROL_PANEL_ROOT / "working" / "templates" / "Method_Statement_Interaction_Weekly_Report_template.xlsx"
    summary_layout_sheet: str = "(WIP) 1 Jul - 12 Jul"
    step0_headed: bool = True
    step0_login_if_needed: bool = True
    step0_timeout_seconds: int = 60
    step0_auth_timeout_seconds: int = 120
    step0_failure_browser_hold_seconds: int = 120

    @property
    def python_executable(self) -> Path:
        return Path(sys.executable).resolve()

    @property
    def step0_worker_script(self) -> Path:
        return CONTROL_PANEL_ROOT / "workers" / "ecm_step0_export.py"

    @property
    def lookup_worker_script(self) -> Path:
        return CONTROL_PANEL_ROOT / "workers" / "ecm_outgoing_reference_lookup.py"

    @property
    def step1a_worker_script(self) -> Path:
        return CONTROL_PANEL_ROOT / "workers" / "ecm_step1a.py"

    @property
    def step1b_worker_script(self) -> Path:
        return CONTROL_PANEL_ROOT / "workers" / "ecm_step1b.py"

    @property
    def weekly_worker_script(self) -> Path:
        return CONTROL_PANEL_ROOT / "workers" / "weekly_report.py"

    @property
    def smm_backup_dir(self) -> Path:
        return CONTROL_PANEL_ROOT / "working" / "smm-backups"

    @property
    def database_path(self) -> Path:
        return self.data_dir / "history.sqlite3"

    @property
    def runs_dir(self) -> Path:
        return self.data_dir / "runs"

    @property
    def archive_database_path(self) -> Path:
        return self.archive_dir / "history.sqlite3"

    def validate_step0(self) -> list[str]:
        errors: list[str] = []
        for name, value in (("CONTROL_PANEL_V3_STEP0_HEADED", self.step0_headed),
                            ("CONTROL_PANEL_V3_STEP0_LOGIN_IF_NEEDED", self.step0_login_if_needed)):
            if not isinstance(value, bool):
                errors.append(f"{name} must be true or false.")
        durations = (
            ("CONTROL_PANEL_V3_STEP0_TIMEOUT_SECONDS", self.step0_timeout_seconds, 1),
            ("CONTROL_PANEL_V3_STEP0_AUTH_TIMEOUT_SECONDS", self.step0_auth_timeout_seconds, 1),
            ("CONTROL_PANEL_V3_STEP0_FAILURE_BROWSER_HOLD_SECONDS", self.step0_failure_browser_hold_seconds, 0),
        )
        for name, value, minimum in durations:
            if type(value) is not int or value < minimum:
                errors.append(f"{name} must be a whole number {'greater than zero' if minimum else 'zero or greater'}.")
        # A minimum budget, not an upper bound: a run can perform many operations.
        if not errors and type(self.job_timeout_seconds) is int:
            minimum_budget = self.step0_timeout_seconds + self.step0_auth_timeout_seconds + self.step0_failure_browser_hold_seconds
            if self.job_timeout_seconds <= minimum_budget:
                errors.append("CONTROL_PANEL_V3_JOB_TIMEOUT_SECONDS must exceed the sum of Step 0 operation, authentication and failure-hold timeouts.")
        return errors

    def validate(self) -> list[str]:
        errors: list[str] = []
        if self.host not in {"127.0.0.1", "localhost"}:
            errors.append("CONTROL_PANEL_V3_HOST must bind to localhost only.")
        if self.port == 8766:
            errors.append("CONTROL_PANEL_V3_PORT must not use V2's port 8766.")
        protected = (V1_ROOT.resolve(), V2_ROOT.resolve())
        writable_roots = (
            self.data_dir,
            self.archive_dir,
            self.smm_backup_dir,
            self.csv_dir,
            self.ecm_output_dir,
            self.smm_output_dir,
            self.readiness_workbook_path,
            self.django_output_dir,
            self.edge_profile_dir,
        )
        template_roots = (self.ecm_template_path.parent, self.smm_template_path.parent, self.weekly_template_path.parent)
        for path in writable_roots:
            resolved = path.resolve()
            if not _is_within(resolved, CONTROL_PANEL_ROOT.resolve()):
                errors.append(f"V3 writable path must remain under {CONTROL_PANEL_ROOT}: {resolved}")
            if any(_is_within(resolved, root) or _is_within(root, resolved) for root in protected):
                errors.append(f"V3 writable path overlaps a protected V1/V2 root: {resolved}")
        for index, root in enumerate(writable_roots):
            for other in writable_roots[index + 1 :]:
                if _is_within(root, other) or _is_within(other, root):
                    errors.append(f"V3 writable paths must be distinct: {root} and {other}")
        for template_root in template_roots:
            if not _is_within(template_root, CONTROL_PANEL_ROOT.resolve()):
                errors.append(f"V3 template copies must remain under {CONTROL_PANEL_ROOT}: {template_root}")
            if any(_is_within(template_root, protected_root) or _is_within(protected_root, template_root) for protected_root in protected):
                errors.append(f"V3 template copies must not overlap protected V1/V2 roots: {template_root}")
            if any(_is_within(template_root, writable) or _is_within(writable, template_root) for writable in writable_roots):
                errors.append(f"V3 template copies must not overlap mutable output roots: {template_root}")
        if any(_is_within(self.ecm_template_path, protected_root) for protected_root in protected):
            errors.append(f"V3 template copies must not point into protected V1/V2 roots: {self.ecm_template_path}")
        if any(_is_within(self.smm_template_path, protected_root) for protected_root in protected):
            errors.append(f"V3 template copies must not point into protected V1/V2 roots: {self.smm_template_path}")
        if any(_is_within(self.weekly_template_path, protected_root) for protected_root in protected):
            errors.append(f"V3 template copies must not point into protected V1/V2 roots: {self.weekly_template_path}")
        if any(_is_within(self.readiness_workbook_path, protected_root) for protected_root in protected):
            errors.append(f"V3 readiness workbook must not point into protected V1/V2 roots: {self.readiness_workbook_path}")
        if self.job_timeout_seconds <= 0:
            errors.append("CONTROL_PANEL_V3_JOB_TIMEOUT_SECONDS must be greater than zero.")
        if not self.ecm_url:
            errors.append("CONTROL_PANEL_V3_ECM_URL must not be empty.")
        if not _is_within(self.ecm_template_path, CONTROL_PANEL_ROOT.resolve()):
            errors.append("CONTROL_PANEL_V3_ECM_TEMPLATE must remain under control_panel_v3.")
        if not _is_within(self.smm_template_path, CONTROL_PANEL_ROOT.resolve()):
            errors.append("CONTROL_PANEL_V3_SMM_TEMPLATE must remain under control_panel_v3.")
        if not _is_within(self.weekly_template_path, CONTROL_PANEL_ROOT.resolve()):
            errors.append("CONTROL_PANEL_V3_WEEKLY_TEMPLATE must remain under control_panel_v3.")
        if not _is_within(self.readiness_workbook_path, CONTROL_PANEL_ROOT.resolve()):
            errors.append("CONTROL_PANEL_V3_READINESS_WORKBOOK must remain under control_panel_v3.")
        if not self.summary_layout_sheet.strip():
            errors.append("CONTROL_PANEL_V3_SUMMARY_LAYOUT_SHEET must not be empty.")
        if self.worker_timeout_seconds <= 0:
            errors.append("CONTROL_PANEL_V3_WORKER_TIMEOUT_SECONDS must be greater than zero.")
        if self.excel_filter_timeout_seconds <= 0:
            errors.append("CONTROL_PANEL_V3_EXCEL_FILTER_TIMEOUT_SECONDS must be greater than zero.")
        errors.extend(self.validate_step0())
        return errors

    def validate_for_workflow(self, workflow_id: str) -> list[str]:
        """Validate global isolation plus only settings required by this workflow."""
        errors: list[str] = []
        if self.host not in {"127.0.0.1", "localhost"}:
            errors.append("CONTROL_PANEL_V3_HOST must bind to localhost only.")
        if self.port == 8766:
            errors.append("CONTROL_PANEL_V3_PORT must not use V2's port 8766.")
        if self.job_timeout_seconds <= 0:
            errors.append("CONTROL_PANEL_V3_JOB_TIMEOUT_SECONDS must be greater than zero.")
        if workflow_id == "ecm-step-0":
            errors.extend(self.validate_step0())
        protected = (V1_ROOT.resolve(), V2_ROOT.resolve())
        writable_roots = (self.data_dir, self.archive_dir, self.smm_backup_dir, self.csv_dir, self.ecm_output_dir,
                          self.smm_output_dir, self.readiness_workbook_path, self.django_output_dir, self.edge_profile_dir)
        for path in writable_roots:
            if not _is_within(path, CONTROL_PANEL_ROOT.resolve()):
                errors.append(f"V3 writable path must remain under {CONTROL_PANEL_ROOT}: {path.resolve()}")
            if any(_is_within(path, root) or _is_within(root, path) for root in protected):
                errors.append(f"V3 writable path overlaps a protected V1/V2 root: {path.resolve()}")
        for index, path in enumerate(writable_roots):
            for other in writable_roots[index + 1:]:
                if _is_within(path, other) or _is_within(other, path):
                    errors.append(f"V3 writable paths must be distinct: {path} and {other}")

        if workflow_id == "weekly-report":
            if not _is_within(self.weekly_template_path, CONTROL_PANEL_ROOT.resolve()):
                errors.append("V3 weekly template must remain under control_panel_v3.")
            if self.worker_timeout_seconds <= 0:
                errors.append("CONTROL_PANEL_V3_WORKER_TIMEOUT_SECONDS must be greater than zero.")
            return errors

        if workflow_id in {"ecm-step-0", "ecm-step-1a", "ecm-first-date-lookup", "ecm-step-1b", "ecm-step-2"}:
            if not self.ecm_url:
                errors.append("CONTROL_PANEL_V3_ECM_URL must not be empty.")
            if self.worker_timeout_seconds <= 0:
                errors.append("CONTROL_PANEL_V3_WORKER_TIMEOUT_SECONDS must be greater than zero.")
            if not _is_within(self.ecm_template_path, CONTROL_PANEL_ROOT.resolve()):
                errors.append("V3 ECM template must remain under control_panel_v3.")
            if not _is_within(self.smm_template_path, CONTROL_PANEL_ROOT.resolve()):
                errors.append("V3 SMM template must remain under control_panel_v3.")
            if not _is_within(self.readiness_workbook_path, CONTROL_PANEL_ROOT.resolve()):
                errors.append("V3 readiness workbook must remain under control_panel_v3.")
            if self.excel_filter_timeout_seconds <= 0:
                errors.append("CONTROL_PANEL_V3_EXCEL_FILTER_TIMEOUT_SECONDS must be greater than zero.")
            if not self.summary_layout_sheet.strip():
                errors.append("CONTROL_PANEL_V3_SUMMARY_LAYOUT_SHEET must not be empty.")
        return errors

    def ensure_application_directories(self) -> None:
        errors = self.validate()
        if errors:
            raise ConfigurationError(" ".join(errors))
        for directory in (
            self.data_dir,
            self.runs_dir,
            self.archive_dir,
            self.smm_backup_dir,
            self.csv_dir,
            self.ecm_output_dir,
            self.smm_output_dir,
            self.django_output_dir,
            self.edge_profile_dir,
            self.ecm_template_path.parent,
            self.smm_template_path.parent,
            self.weekly_template_path.parent,
            self.readiness_workbook_path.parent,
        ):
            directory.mkdir(parents=True, exist_ok=True)
        template_sources = (
            (self.ecm_template_path, self.project_root / "ecm" / "ECM_Record_template.xlsx"),
            (self.smm_template_path, self.project_root / "ecm" / "SMM_KMS_PM_Master_Document_template.xlsx"),
        )
        for destination, source in template_sources:
            if not destination.exists() and source.is_file():
                shutil.copy2(source, destination)
        self.weekly_template_path.parent.mkdir(parents=True, exist_ok=True)
        weekly_source = self.project_root / "django" / "Method_Statement_Interaction_Weekly_Report_template.xlsx"
        if not self.weekly_template_path.exists() and weekly_source.is_file():
            shutil.copy2(weekly_source, self.weekly_template_path)
        if not self.readiness_workbook_path.exists() and self.smm_template_path.is_file():
            shutil.copy2(self.smm_template_path, self.readiness_workbook_path)


def load_settings() -> Settings:
    load_dotenv(CONTROL_PANEL_ROOT / ".env")
    return Settings(
        project_root=PROJECT_ROOT.resolve(),
        data_dir=_path_setting("CONTROL_PANEL_V3_DATA_DIR", "control_panel_v3/data"),
        archive_dir=_path_setting("CONTROL_PANEL_V3_ARCHIVE_DIR", "control_panel_v3/archive"),
        csv_dir=_path_setting("CONTROL_PANEL_V3_CSV_DIR", "control_panel_v3/working/csv"),
        ecm_output_dir=_path_setting("CONTROL_PANEL_V3_ECM_OUTPUT_DIR", "control_panel_v3/working/ecm"),
        smm_output_dir=_path_setting("CONTROL_PANEL_V3_SMM_OUTPUT_DIR", "control_panel_v3/working/smm"),
        django_output_dir=_path_setting("CONTROL_PANEL_V3_DJANGO_OUTPUT_DIR", "control_panel_v3/working/django"),
        edge_profile_dir=_path_setting("CONTROL_PANEL_V3_EDGE_PROFILE_DIR", "control_panel_v3/edge-profile"),
        ecm_template_path=_path_setting("CONTROL_PANEL_V3_ECM_TEMPLATE", "control_panel_v3/working/templates/ECM_Record_template.xlsx"),
        smm_template_path=_path_setting("CONTROL_PANEL_V3_SMM_TEMPLATE", "control_panel_v3/working/templates/SMM_KMS_PM_Master_Document_template.xlsx"),
        readiness_workbook_path=_path_setting("CONTROL_PANEL_V3_READINESS_WORKBOOK", "control_panel_v3/working/readiness/SMM_KMS_PM_Master_Document.xlsx"),
        weekly_template_path=_path_setting("CONTROL_PANEL_V3_WEEKLY_TEMPLATE", "control_panel_v3/working/templates/Method_Statement_Interaction_Weekly_Report_template.xlsx"),
        summary_layout_sheet=os.getenv("CONTROL_PANEL_V3_SUMMARY_LAYOUT_SHEET", "(WIP) 1 Jul - 12 Jul").strip(),
        preview_only=_boolean("CONTROL_PANEL_V3_PREVIEW_ONLY", True),
        host=os.getenv("CONTROL_PANEL_V3_HOST", "127.0.0.1").strip(),
        port=_positive_int("CONTROL_PANEL_V3_PORT", 8767),
        job_timeout_seconds=_positive_int("CONTROL_PANEL_V3_JOB_TIMEOUT_SECONDS", 7200),
        ecm_url=os.getenv("CONTROL_PANEL_V3_ECM_URL", "https://ecm.chunwo.com/navigator/").strip(),
        worker_timeout_seconds=_positive_int("CONTROL_PANEL_V3_WORKER_TIMEOUT_SECONDS", 180),
        excel_filter_timeout_seconds=_positive_int("CONTROL_PANEL_V3_EXCEL_FILTER_TIMEOUT_SECONDS", 90),
        step0_headed=_boolean("CONTROL_PANEL_V3_STEP0_HEADED", True),
        step0_login_if_needed=_boolean("CONTROL_PANEL_V3_STEP0_LOGIN_IF_NEEDED", True),
        step0_timeout_seconds=_positive_int("CONTROL_PANEL_V3_STEP0_TIMEOUT_SECONDS", 60),
        step0_auth_timeout_seconds=_positive_int("CONTROL_PANEL_V3_STEP0_AUTH_TIMEOUT_SECONDS", 120),
        step0_failure_browser_hold_seconds=_nonnegative_int("CONTROL_PANEL_V3_STEP0_FAILURE_BROWSER_HOLD_SECONDS", 120),
    )
