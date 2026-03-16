from __future__ import annotations

from datetime import date, datetime
from typing import Any, Dict, List, Tuple

from app.core.config import DEFAULT_PROJECT_ROWS_TO_FILL
from app.core.constants import (
    MATCH_TYPE_PROJECT_TABLE,
    REVIEW_STATUS_FILLED,
    REVIEW_STATUS_SKIPPED,
    SECTION_PROJECTS,
    SKIP_REASON_EXCEPTION,
    SKIP_REASON_EXISTING_VALUE,
    SKIP_REASON_FORMULA_CELL,
    SKIP_REASON_NO_TARGET,
    TABLE_TYPE_PROJECTS,
)
from app.engine.layout_hints import detect_project_table_mode
from app.engine.project_selector import select_projects
from app.engine.table_detectors import detect_project_tables
from app.engine.target_cell_resolver import get_merged_anchor_cell


PROJECT_FIELD_TO_COLUMN_KEY = {
    "serial_no": None,
    "project_name": "project_name",
    "client": "client",
    "location": "location",
    "category": "category",
    "start_date": "start_date",
    "end_date": "end_date",
    "status": "status",
    "value": "value",
    "contact_name": "contact_name",
    "contact_phone": "contact_phone",
}


def _safe_preview(value: Any, max_len: int = 120) -> str:
    if value is None:
        return ""
    text = str(value)
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _is_formula_cell(cell) -> bool:
    value = getattr(cell, "value", None)
    return isinstance(value, str) and value.startswith("=")


def _is_meaningful_value(value: Any) -> bool:
    if value is None:
        return False
    return str(value).strip() != ""


def _format_date_value(value: Any) -> Any:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, date):
        return value.strftime("%d-%m-%Y")

    return value


def _format_project_scalar(value: Any) -> Any:
    if value is None:
        return None

    if isinstance(value, float) and value.is_integer():
        return int(value)

    return value


def _normalize_status(value: Any) -> str:
    text = _safe_text(value).lower()
    if not text:
        return ""

    if any(token in text for token in ["ongoing", "under execution", "in progress", "running", "current"]):
        return "ongoing"

    if any(token in text for token in ["completed", "executed", "finished", "closed", "done"]):
        return "completed"

    return text


def _derive_project_location(project: Dict[str, Any]) -> Any:
    for key in ("location", "site_location"):
        value = project.get(key)
        if _is_meaningful_value(value):
            return value

    city = project.get("city")
    state = project.get("state")

    city_text = _safe_text(city)
    state_text = _safe_text(state)

    if city_text and state_text:
        return f"{city_text}, {state_text}"
    if city_text:
        return city_text
    if state_text:
        return state_text

    return None


def _derive_project_value(project: Dict[str, Any]) -> Any:
    for key in (
        "value",
        "project_value",
        "contract_value",
        "value_of_work_order",
        "work_order_value",
        "volume_of_work",
    ):
        value = project.get(key)
        if _is_meaningful_value(value):
            return value
    return None


def _normalize_project_record(project: Dict[str, Any]) -> Dict[str, Any]:
    normalized = dict(project or {})

    if not _is_meaningful_value(normalized.get("project_name")):
        for key in ("name_of_project", "project", "work_name"):
            if _is_meaningful_value(normalized.get(key)):
                normalized["project_name"] = normalized.get(key)
                break

    if not _is_meaningful_value(normalized.get("client")):
        for key in ("client_name", "customer"):
            if _is_meaningful_value(normalized.get(key)):
                normalized["client"] = normalized.get(key)
                break

    if not _is_meaningful_value(normalized.get("category")):
        for key in ("type_of_work", "nature_of_work", "scope_of_work"):
            if _is_meaningful_value(normalized.get(key)):
                normalized["category"] = normalized.get(key)
                break

    if not _is_meaningful_value(normalized.get("start_date")):
        for key in ("order_date", "commencement_date", "po_date"):
            if _is_meaningful_value(normalized.get(key)):
                normalized["start_date"] = normalized.get(key)
                break

    if not _is_meaningful_value(normalized.get("end_date")):
        for key in ("completion_date", "work_completion_date"):
            if _is_meaningful_value(normalized.get(key)):
                normalized["end_date"] = normalized.get(key)
                break

    if not _is_meaningful_value(normalized.get("value")):
        derived_value = _derive_project_value(normalized)
        if _is_meaningful_value(derived_value):
            normalized["value"] = derived_value

    if not _is_meaningful_value(normalized.get("location")):
        derived_location = _derive_project_location(normalized)
        if _is_meaningful_value(derived_location):
            normalized["location"] = derived_location

    return normalized


def _get_project_value(
    project: Dict[str, Any],
    source_key: str,
    logical_col: str,
    project_index: int | None = None,
) -> Any:
    if logical_col == "serial_no":
        return project_index

    if source_key == "location":
        return _derive_project_location(project)

    if source_key == "value":
        return _derive_project_value(project)

    value = project.get(source_key)
    if not _is_meaningful_value(value):
        return None

    if source_key in {"start_date", "end_date"}:
        return _format_date_value(value)

    if source_key == "status":
        normalized_status = _normalize_status(value)
        return normalized_status.title() if normalized_status in {"ongoing", "completed"} else value

    return _format_project_scalar(value)


def _build_log_row(
    sheet_name: str,
    header_row: int,
    target_cell: str = "",
    target_merged_range: str = "",
    filled_value: Any = "",
    status: str = REVIEW_STATUS_SKIPPED,
    note: str = "",
    reason: str = "",
    table_mode: str = "",
    source_field_key: str = "__projects__",
    project_index: int | None = None,
):
    return {
        "sheet": sheet_name,
        "sheet_name": sheet_name,
        "label_cell": f"ROW-{header_row}",
        "label_text": "Project Table",
        "normalized_label": "project table",
        "cell_type": "TABLE_HEADER",
        "active_section": SECTION_PROJECTS,
        "section": SECTION_PROJECTS,
        "mapped_field_key": source_field_key,
        "field_key": source_field_key,
        "match_type": MATCH_TYPE_PROJECT_TABLE,
        "confidence": 100,
        "semantic_confidence": 1.0,
        "layout_confidence": 1.0,
        "resolver": "project_table_filler",
        "layout_type": TABLE_TYPE_PROJECTS,
        "table_type": TABLE_TYPE_PROJECTS,
        "target_cell": target_cell,
        "target_merged_range": target_merged_range,
        "filled_value": filled_value,
        "value_preview": _safe_preview(filled_value),
        "status": status,
        "write_result": status,
        "note": note,
        "reason": reason or note,
        "table_mode": table_mode,
        "project_index": project_index,
    }


def _get_projects_from_master(master_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Keep this tolerant because the master loader may store projects
    in different keys depending on current implementation.
    """
    if not isinstance(master_data, dict):
        return []

    for key in ("__projects__", "projects", "project_master"):
        value = master_data.get(key)
        if isinstance(value, list):
            return [_normalize_project_record(item) for item in value if isinstance(item, dict)]

    return []


def _available_project_rows(
    sheet,
    header_row: int,
    column_map: Dict[str, int],
    max_scan_rows: int = 16,
) -> List[int]:
    """
    Find rows below the project header that can still accept project data.

    Important improvement:
    A row is considered available even if some cells are already filled,
    as long as at least one mapped non-serial column is blank.
    This is critical for PQ tables where serial no. may already exist.
    """
    available_rows: List[int] = []

    start_row = header_row + 1
    end_row = min(sheet.max_row, header_row + max_scan_rows)

    non_serial_columns = {
        logical_col: col_idx
        for logical_col, col_idx in column_map.items()
        if logical_col != "serial_no"
    }

    if not non_serial_columns:
        return available_rows

    for row_idx in range(start_row, end_row + 1):
        has_blank_target = False
        row_has_any_structure = False

        for logical_col, col_idx in non_serial_columns.items():
            anchor = get_merged_anchor_cell(sheet, sheet.cell(row=row_idx, column=col_idx))
            row_has_any_structure = True

            if not _is_meaningful_value(anchor.value) and not _is_formula_cell(anchor):
                has_blank_target = True

        if row_has_any_structure and has_blank_target:
            available_rows.append(row_idx)

    return available_rows


def _mode_matches_project(project: Dict[str, Any], mode: str) -> bool:
    normalized_mode = _safe_text(mode).lower() or "all"
    if normalized_mode == "all":
        return True

    project_status = _normalize_status(project.get("status"))

    if normalized_mode == "ongoing":
        return project_status == "ongoing" or project_status == ""

    if normalized_mode == "completed":
        return project_status == "completed" or project_status == ""

    return True


def _select_projects_for_table(
    master_data: Dict[str, Any],
    raw_master_projects: List[Dict[str, Any]],
    mode: str,
    limit: int,
) -> List[Dict[str, Any]]:
    projects = select_projects(master_data, mode=mode, limit=limit)
    normalized_projects = [_normalize_project_record(item) for item in projects if isinstance(item, dict)]

    if normalized_projects:
        return normalized_projects

    mode_filtered = [p for p in raw_master_projects if _mode_matches_project(p, mode)]
    if mode_filtered:
        return mode_filtered[:limit]

    return raw_master_projects[:limit]


def _fill_single_project_cell(
    sheet,
    row_idx: int,
    col_idx: int,
    value: Any,
) -> Tuple[bool, str, str]:
    """
    Returns:
        success, coordinate, merged_range
    Raises no exception outside; caller handles result/logging.
    """
    anchor = get_merged_anchor_cell(sheet, sheet.cell(row=row_idx, column=col_idx))
    merged_range = None

    for merged in sheet.merged_cells.ranges:
        if anchor.coordinate in merged:
            merged_range = str(merged)
            break

    if _is_formula_cell(anchor):
        return False, anchor.coordinate, merged_range or ""

    if _is_meaningful_value(anchor.value):
        return False, anchor.coordinate, merged_range or ""

    anchor.value = value
    return True, anchor.coordinate, merged_range or ""


def fill_project_tables(wb, master_data):
    """
    Detects project tables and fills them deterministically from Project_Master data.

    Improvements:
    - supports partially empty project rows
    - fills row-wise per cell instead of requiring fully blank row
    - stronger mode fallback
    - no fill color applied
    """
    log_rows: List[Dict[str, Any]] = []
    master_projects = _get_projects_from_master(master_data)

    for sheet in wb.worksheets:
        detections = detect_project_tables(sheet)

        for detection in detections:
            header_row = detection["header_row"]
            column_map = detection["column_map"]

            mode = detection.get("project_mode") or detect_project_table_mode(sheet, header_row) or "all"

            available_rows = _available_project_rows(
                sheet=sheet,
                header_row=header_row,
                column_map=column_map,
                max_scan_rows=max(DEFAULT_PROJECT_ROWS_TO_FILL * 3, 16),
            )

            if not available_rows:
                log_rows.append(
                    _build_log_row(
                        sheet_name=sheet.title,
                        header_row=header_row,
                        status=REVIEW_STATUS_SKIPPED,
                        note="NO_SAFE_TARGET",
                        reason=SKIP_REASON_NO_TARGET,
                        table_mode=mode,
                    )
                )
                continue

            projects = _select_projects_for_table(
                master_data=master_data,
                raw_master_projects=master_projects,
                mode=mode,
                limit=min(len(available_rows), max(DEFAULT_PROJECT_ROWS_TO_FILL, len(available_rows))),
            )

            if not projects:
                log_rows.append(
                    _build_log_row(
                        sheet_name=sheet.title,
                        header_row=header_row,
                        status=REVIEW_STATUS_SKIPPED,
                        note="NO_MASTER_VALUE",
                        reason="NO_MASTER_VALUE",
                        table_mode=mode,
                    )
                )
                continue

            rows_to_use = available_rows[: min(len(available_rows), len(projects))]

            for project_index, (row_idx, project) in enumerate(zip(rows_to_use, projects), start=1):
                normalized_project = _normalize_project_record(project)

                for logical_col, col_idx in column_map.items():
                    source_key = PROJECT_FIELD_TO_COLUMN_KEY.get(logical_col, logical_col)

                    if logical_col == "serial_no":
                        value = project_index
                    else:
                        if not source_key:
                            continue
                        value = _get_project_value(
                            project=normalized_project,
                            source_key=source_key,
                            logical_col=logical_col,
                            project_index=project_index,
                        )

                    if value in [None, ""]:
                        continue

                    try:
                        success, coordinate, merged_range = _fill_single_project_cell(
                            sheet=sheet,
                            row_idx=row_idx,
                            col_idx=col_idx,
                            value=value,
                        )
                    except Exception as exc:
                        target_anchor = get_merged_anchor_cell(sheet, sheet.cell(row=row_idx, column=col_idx))
                        log_rows.append(
                            _build_log_row(
                                sheet_name=sheet.title,
                                header_row=header_row,
                                target_cell=target_anchor.coordinate,
                                filled_value="",
                                status=REVIEW_STATUS_SKIPPED,
                                note=f"WRITE_ERROR: {str(exc)}",
                                reason=SKIP_REASON_EXCEPTION,
                                table_mode=mode,
                                source_field_key=f"project.{source_key or logical_col}",
                                project_index=project_index,
                            )
                        )
                        continue

                    if not success:
                        anchor = get_merged_anchor_cell(sheet, sheet.cell(row=row_idx, column=col_idx))
                        reason = (
                            SKIP_REASON_FORMULA_CELL
                            if _is_formula_cell(anchor)
                            else SKIP_REASON_EXISTING_VALUE
                        )
                        note = "FORMULA_CELL" if _is_formula_cell(anchor) else "TARGET_ALREADY_HAS_VALUE"

                        log_rows.append(
                            _build_log_row(
                                sheet_name=sheet.title,
                                header_row=header_row,
                                target_cell=coordinate,
                                target_merged_range=merged_range,
                                filled_value=anchor.value,
                                status=REVIEW_STATUS_SKIPPED,
                                note=note,
                                reason=reason,
                                table_mode=mode,
                                source_field_key=f"project.{source_key or logical_col}",
                                project_index=project_index,
                            )
                        )
                        continue

                    log_rows.append(
                        _build_log_row(
                            sheet_name=sheet.title,
                            header_row=header_row,
                            target_cell=coordinate,
                            target_merged_range=merged_range,
                            filled_value=value,
                            status=REVIEW_STATUS_FILLED,
                            note=f"MODE:{mode}",
                            reason="FILLED",
                            table_mode=mode,
                            source_field_key=f"project.{source_key or logical_col}",
                            project_index=project_index,
                        )
                    )

    return wb, log_rows