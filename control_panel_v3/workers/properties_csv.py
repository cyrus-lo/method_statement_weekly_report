"""V3-owned validation for the ECM Step 0 export CSV contract."""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path

CSV_HEADERS = (
    "Document title", "Document date", "Name", "Document revision", "Modified By",
    "Major Version", "SMM ID", "Modified On", "Outgoing ref. no.", "Path", "Link",
)
REQUIRED_ROW_FIELDS = ("Document title", "Document date", "Name", "Major Version")


class CsvValidationError(ValueError):
    """The downloaded CSV does not meet the Step 1 input contract."""


def parse_ecm_datetime(value: str, label: str, row_number: int):
    text = value.strip()
    for pattern in ("%m/%d/%Y, %I:%M %p", "%m/%d/%Y %I:%M %p"):
        try:
            return datetime.strptime(text, pattern)
        except ValueError:
            pass
    raise CsvValidationError(f"Row {row_number}: {label} is not a supported ECM date/time: {value!r}.")


def validate_properties_csv(path: Path) -> int:
    if not path.is_file() or path.stat().st_size == 0:
        raise CsvValidationError("Downloaded ECM export is missing or empty.")
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with path.open("r", encoding=encoding, newline="") as stream:
                reader = csv.DictReader(stream)
                headers = tuple((header or "").strip() for header in (reader.fieldnames or ()))
                if headers != CSV_HEADERS:
                    raise CsvValidationError(f"CSV headers do not match the approved 11-column contract: {headers!r}.")
                count = 0
                for row_number, row in enumerate(reader, start=2):
                    if None in row:
                        raise CsvValidationError(f"Row {row_number}: unexpected extra CSV values.")
                    if all(not (row.get(header) or "").strip() for header in CSV_HEADERS):
                        continue
                    missing = [header for header in REQUIRED_ROW_FIELDS if not (row.get(header) or "").strip()]
                    if missing:
                        raise CsvValidationError(f"Row {row_number}: required fields are blank: {', '.join(missing)}.")
                    parse_ecm_datetime(row["Document date"], "Document date", row_number)
                    if (row.get("Modified On") or "").strip():
                        parse_ecm_datetime(row["Modified On"], "Modified On", row_number)
                    try:
                        major_version = int(row["Major Version"].strip())
                    except ValueError as error:
                        raise CsvValidationError(f"Row {row_number}: Major Version must be a whole number.") from error
                    if major_version < 0:
                        raise CsvValidationError(f"Row {row_number}: Major Version must not be negative.")
                    count += 1
                if count == 0:
                    raise CsvValidationError("ECM export contains no usable data rows.")
                return count
        except UnicodeDecodeError:
            continue
    raise CsvValidationError("ECM export is not a supported text CSV encoding.")
