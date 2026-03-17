from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

from app.core.config import (
    LOG_DIR,
    PROCESSED_DIR,
    PROCESSED_FILE_SUFFIX,
    REVIEW_LOG_FILE_SUFFIX,
)
from app.engine.autofill_engine import autofill_workbook
from app.engine.layout_hints import get_workbook_fingerprint
from app.engine.master_loader import load_master_data, load_synonym_mapping
from app.engine.workbook_scanner import scan_workbook
from app.services.file_service import get_uploaded_file_path
from app.services.review_log_service import write_review_log


_MASTER_DATA_CACHE: Optional[Dict[str, Any]] = None
_SYNONYM_CACHE: Optional[Dict[str, Any]] = None


def _get_cached_master_data() -> Dict[str, Any]:
    global _MASTER_DATA_CACHE
    if _MASTER_DATA_CACHE is None:
        _MASTER_DATA_CACHE = load_master_data()
    return _MASTER_DATA_CACHE


def _get_cached_synonyms() -> Dict[str, Any]:
    global _SYNONYM_CACHE
    if _SYNONYM_CACHE is None:
        _SYNONYM_CACHE = load_synonym_mapping()
    return _SYNONYM_CACHE


def _safe_status(row: Dict[str, Any]) -> str:
    return str(row.get("status", "")).strip().upper()


def _count_status(log_rows: List[Dict[str, Any]], status: str) -> int:
    wanted = status.strip().upper()
    return sum(1 for row in log_rows if _safe_status(row) == wanted)


def _build_output_paths(file_id: str) -> tuple[Path, Path]:
    filled_output_path = PROCESSED_DIR / f"{file_id}{PROCESSED_FILE_SUFFIX}"
    review_log_path = LOG_DIR / f"{file_id}{REVIEW_LOG_FILE_SUFFIX}"
    return filled_output_path, review_log_path


def warm_processing_cache() -> None:
    _get_cached_master_data()
    _get_cached_synonyms()


def clear_processing_cache() -> None:
    global _MASTER_DATA_CACHE, _SYNONYM_CACHE
    _MASTER_DATA_CACHE = None
    _SYNONYM_CACHE = None


def process_uploaded_file(file_id: str) -> dict:
    """
    End-to-end processing flow for an uploaded vendor registration workbook.

    Steps:
    1. resolve uploaded file path
    2. load cached master data and synonym mapping
    3. scan workbook
    4. compute lightweight fingerprint
    5. autofill workbook
    6. save filled workbook
    7. write review log
    8. return processing summary
    """
    input_file_path = get_uploaded_file_path(file_id)

    if input_file_path is None or not input_file_path.exists():
        raise FileNotFoundError(f"No uploaded file found for file_id={file_id}")

    master_data = _get_cached_master_data()
    synonyms = _get_cached_synonyms()

    wb, scan_result = scan_workbook(input_file_path)
    fingerprint = get_workbook_fingerprint(wb)

    wb, log_rows = autofill_workbook(
        wb=wb,
        findings=scan_result,
        master_data=master_data,
        synonyms=synonyms,
    )

    filled_output_path, review_log_path = _build_output_paths(file_id)

    wb.save(filled_output_path)
    write_review_log(log_rows, review_log_path)

    filled_count = _count_status(log_rows, "FILLED")
    skipped_count = _count_status(log_rows, "SKIPPED")
    unmatched_count = _count_status(log_rows, "UNMATCHED")
    error_count = _count_status(log_rows, "ERROR")

    scan_findings = scan_result.get("findings", []) if isinstance(scan_result, dict) else []
    scanned_label_count = len(scan_findings)

    return {
        "file_id": file_id,
        "input_file": str(input_file_path),
        "filled_file": str(filled_output_path),
        "review_log": str(review_log_path),
        "template_fingerprint": fingerprint,
        "scanned_label_count": scanned_label_count,
        "total_logged_items": len(log_rows),
        "filled_count": filled_count,
        "skipped_count": skipped_count,
        "unmatched_count": unmatched_count,
        "error_count": error_count,
    }