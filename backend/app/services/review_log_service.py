import csv
from pathlib import Path
from typing import Any, Dict, List

from app.core.constants import REVIEW_LOG_COLUMNS


def _sanitize_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value)


def _normalize_row(row: Dict[str, Any]) -> Dict[str, str]:
    """
    Ensure every review log row conforms to the unified log schema.
    Missing fields are filled with empty strings.
    """
    return {
        field: _sanitize_value(row.get(field, ""))
        for field in REVIEW_LOG_COLUMNS
    }


def write_review_log(log_rows: list[dict], output_path: Path):
    """
    Writes review log as CSV using the unified review-log schema.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=REVIEW_LOG_COLUMNS)
        writer.writeheader()

        for row in log_rows:
            writer.writerow(_normalize_row(row))