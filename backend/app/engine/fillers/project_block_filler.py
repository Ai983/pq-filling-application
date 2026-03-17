from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

from app.core.config import MIN_LAYOUT_CONFIDENCE
from app.core.constants import (
    CELL_TYPE_SUBFIELD,
    MATCH_TYPE_SECTION_CONTEXT,
    REVIEW_STATUS_FILLED,
    REVIEW_STATUS_SKIPPED,
    SECTION_PROJECTS,
    SKIP_REASON_EXCEPTION,
    SKIP_REASON_EXISTING_VALUE,
    SKIP_REASON_FORMULA_CELL,
    SKIP_REASON_LOW_CONFIDENCE,
    SKIP_REASON_NO_TARGET,
    TABLE_TYPE_PROJECTS,
)
from app.engine.project_selector import select_projects
from app.engine.target_cell_resolver import (
    get_merged_anchor_cell,
    resolve_target_cell,
)
from app.engine.utils import normalize_text


PROJECT_BLOCK_FIELD_ALIASES: Dict[str, str] = {
    "project": "project_name",
    "project name": "project_name",
    "name of project": "project_name",
    "location": "location",
    "site location": "location",
    "city": "location",
    "state": "location",
    "area in sft": "area_sft",
    "area in sq ft": "area_sft",
    "area in sqft": "area_sft",
    "area sq ft": "area_sft",
    "built up area": "area_sft",
    "builtup area": "area_sft",
    "awarded amount in inr": "value",
    "awarded amount": "value",
    "awarded value": "value",
    "value of work": "value",
    "work order value": "value",
    "value of work order": "value",
    "contract value": "value",
    "project value": "value",
    "value": "value",
    "pmc of the project": "pmc_name",
    "pmc": "pmc_name",
    "pmc name": "pmc_name",
    "type of project": "category",
    "type of work": "category",
    "nature of work": "category",
    "category": "category",
    "client reference": "client_reference",
    "client reference name mobile number and email id": "client_reference",
    "client reference name mobile number email id": "client_reference",
    "client reference name mobile no and email id": "client_reference",
    "client reference name mobile no email id": "client_reference",
    "client reference name mobile number": "client_reference",
    "client reference name contact number email": "client_reference",
    "client name": "client",
    "client": "client",
    "customer": "client",
    "start date": "start_date",
    "commencement date": "start_date",
    "order date": "start_date",
    "completion date": "end_date",
    "end date": "end_date",
    "status": "status",
}

PROJECT_SECTION_TERMINATORS = [
    "please specify if you have any other business interests",
    "business interests",
    "litigation",
    "arbitration",
    "banker",
    "blacklisted",
    "why do you wish to work with us",
]

KNOWN_LOCATION_HINTS = [
    "bangalore",
    "bengaluru",
    "mumbai",
    "delhi",
    "gurgaon",
    "gurugram",
    "noida",
    "pune",
    "hyderabad",
    "chennai",
]


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _safe_preview(value: Any, max_len: int = 120) -> str:
    text = _safe_text(value)
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def _is_formula_cell(cell) -> bool:
    value = getattr(cell, "value", None)
    return isinstance(value, str) and value.startswith("=")


def _is_meaningful_value(value: Any) -> bool:
    return _safe_text(value) != ""


def _normalize_text(value: Any) -> str:
    return normalize_text(_safe_text(value))


def _normalize_headerish_text(value: Any) -> str:
    text = _normalize_text(value)
    if not text:
        return ""
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = " ".join(text.split()).strip()
    return text


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

    if any(token in text for token in ["ongoing", "under execution", "under progress", "in progress", "running", "current"]):
        return "ongoing"

    if any(token in text for token in ["completed", "executed", "finished", "closed", "done"]):
        return "completed"

    return text


def _normalize_bucket(value: Any) -> str:
    text = _normalize_headerish_text(value)
    if not text:
        return ""
    return text.upper().replace(" ", "_")


def _derive_project_location(project: Dict[str, Any]) -> Any:
    for key in ("location", "site_location"):
        value = project.get(key)
        if _is_meaningful_value(value):
            return value

    city = _safe_text(project.get("city"))
    state = _safe_text(project.get("state"))

    if city and state:
        return f"{city}, {state}"
    if city:
        return city
    if state:
        return state

    return None


def _derive_project_value(project: Dict[str, Any]) -> Any:
    for key in (
        "value",
        "project_value",
        "contract_value",
        "value_of_work_order",
        "work_order_value",
        "volume_of_work",
        "awarded_amount",
        "awarded_value",
    ):
        value = project.get(key)
        if _is_meaningful_value(value):
            return value
    return None


def _derive_project_area(project: Dict[str, Any]) -> Any:
    for key in (
        "area_sft",
        "area_sqft",
        "area_sq_ft",
        "area_in_sft",
        "area_in_sqft",
        "built_up_area",
        "builtup_area",
        "project_area",
        "area",
    ):
        value = project.get(key)
        if _is_meaningful_value(value):
            return value
    return None


def _derive_project_pmc(project: Dict[str, Any]) -> Any:
    for key in (
        "pmc_name",
        "pmc",
        "consultant",
        "project_management_consultant",
    ):
        value = project.get(key)
        if _is_meaningful_value(value):
            return value
    return None


def _derive_client_reference(project: Dict[str, Any]) -> Any:
    name = _safe_text(project.get("contact_name"))
    phone = _safe_text(project.get("contact_phone"))
    email = _safe_text(project.get("contact_email"))

    if not name:
        for key in ("client_contact", "contact_person", "reference_name"):
            maybe = _safe_text(project.get(key))
            if maybe:
                name = maybe
                break

    if not email:
        for key in ("email", "contact_mail", "contact_email_id", "reference_email"):
            maybe = _safe_text(project.get(key))
            if maybe:
                email = maybe
                break

    parts = [part for part in [name, phone, email] if part]
    if parts:
        return " | ".join(parts)

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
        for key in ("type_of_work", "nature_of_work", "scope_of_work", "type_of_project"):
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

    if not _is_meaningful_value(normalized.get("contact_name")):
        for key in ("client_contact", "contact_person", "reference_name"):
            if _is_meaningful_value(normalized.get(key)):
                normalized["contact_name"] = normalized.get(key)
                break

    if not _is_meaningful_value(normalized.get("value")):
        derived_value = _derive_project_value(normalized)
        if _is_meaningful_value(derived_value):
            normalized["value"] = derived_value

    if not _is_meaningful_value(normalized.get("location")):
        derived_location = _derive_project_location(normalized)
        if _is_meaningful_value(derived_location):
            normalized["location"] = derived_location

    if not _is_meaningful_value(normalized.get("area_sft")):
        derived_area = _derive_project_area(normalized)
        if _is_meaningful_value(derived_area):
            normalized["area_sft"] = derived_area

    if not _is_meaningful_value(normalized.get("pmc_name")):
        derived_pmc = _derive_project_pmc(normalized)
        if _is_meaningful_value(derived_pmc):
            normalized["pmc_name"] = derived_pmc

    if not _is_meaningful_value(normalized.get("client_reference")):
        derived_reference = _derive_client_reference(normalized)
        if _is_meaningful_value(derived_reference):
            normalized["client_reference"] = derived_reference

    normalized["bucket"] = _normalize_bucket(normalized.get("bucket"))
    return normalized


def _get_projects_from_master(master_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    if not isinstance(master_data, dict):
        return []

    for key in ("__projects__", "projects", "project_master"):
        value = master_data.get(key)
        if isinstance(value, list):
            return [_normalize_project_record(item) for item in value if isinstance(item, dict)]

    return []


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


def _matches_location_filter(project: Dict[str, Any], location_hint: str) -> bool:
    hint = _normalize_headerish_text(location_hint)
    if not hint:
        return True

    haystack = " ".join(
        [
            _safe_text(project.get("location")),
            _safe_text(project.get("site_location")),
            _safe_text(project.get("city")),
            _safe_text(project.get("state")),
        ]
    )
    haystack_norm = _normalize_headerish_text(haystack)
    return hint in haystack_norm


def _infer_bucket_hint(mode: str, location_hint: str = "") -> str:
    mode_norm = _safe_text(mode).lower()
    location_norm = _normalize_headerish_text(location_hint)

    if mode_norm == "ongoing":
        return "ONGOING"

    if mode_norm == "completed" and location_norm in {"bangalore", "bengaluru"}:
        return "BANGALORE_COMPLETED"

    if mode_norm == "completed":
        return "COMPLETED"

    return ""


def _dedupe_projects(projects: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen = set()
    output = []

    for project in projects:
        key = (
            _normalize_headerish_text(project.get("project_name")),
            _normalize_headerish_text(project.get("client")),
            _normalize_headerish_text(project.get("location")),
        )
        if key in seen:
            continue
        seen.add(key)
        output.append(project)

    return output


def _select_projects_for_blocks(
    master_data: Dict[str, Any],
    mode: str,
    limit: int,
    location_hint: str = "",
) -> List[Dict[str, Any]]:
    bucket_hint = _infer_bucket_hint(mode=mode, location_hint=location_hint)
    raw_master_projects = _get_projects_from_master(master_data)

    if bucket_hint:
        bucket_projects = [p for p in raw_master_projects if _normalize_bucket(p.get("bucket")) == bucket_hint]
        if location_hint:
            bucket_projects = [p for p in bucket_projects if _matches_location_filter(p, location_hint)]
        bucket_projects = _dedupe_projects(bucket_projects)
        if bucket_projects:
            return bucket_projects[:limit]

    selected = select_projects(master_data, mode=mode, limit=max(limit * 3, limit))
    normalized_selected = [_normalize_project_record(item) for item in selected if isinstance(item, dict)]

    if location_hint:
        location_filtered = [p for p in normalized_selected if _matches_location_filter(p, location_hint)]
        location_filtered = _dedupe_projects(location_filtered)
        if location_filtered:
            return location_filtered[:limit]

    normalized_selected = _dedupe_projects(normalized_selected)
    if normalized_selected:
        return normalized_selected[:limit]

    mode_filtered = [p for p in raw_master_projects if _mode_matches_project(p, mode)]

    if location_hint:
        location_filtered = [p for p in mode_filtered if _matches_location_filter(p, location_hint)]
        location_filtered = _dedupe_projects(location_filtered)
        if location_filtered:
            return location_filtered[:limit]

    mode_filtered = _dedupe_projects(mode_filtered)
    if mode_filtered:
        return mode_filtered[:limit]

    if location_hint:
        location_filtered = [p for p in raw_master_projects if _matches_location_filter(p, location_hint)]
        location_filtered = _dedupe_projects(location_filtered)
        if location_filtered:
            return location_filtered[:limit]

    return _dedupe_projects(raw_master_projects)[:limit]


def _row_text(sheet, row_idx: int) -> str:
    parts: List[str] = []
    for col_idx in range(1, sheet.max_column + 1):
        text = _safe_text(sheet.cell(row=row_idx, column=col_idx).value)
        if text:
            parts.append(text)
    return " ".join(parts).strip()


def _normalized_row_text(sheet, row_idx: int) -> str:
    return _normalize_headerish_text(_row_text(sheet, row_idx))


def _find_project_anchor_in_row(sheet, row_idx: int) -> Optional[Tuple[int, str, int]]:
    """
    Returns:
        (column_index, raw_label, project_index)
    """
    for col_idx in range(1, sheet.max_column + 1):
        raw = _safe_text(sheet.cell(row=row_idx, column=col_idx).value)
        if not raw:
            continue

        normalized = _normalize_headerish_text(raw)
        match = re.search(r"\bproject\b\s*(\d+)\b", normalized)
        if match:
            print(
                f"DEBUG: project anchor matched at row {row_idx}, "
                f"col {col_idx}, raw='{raw}', normalized='{normalized}'"
            )
            return col_idx, raw, int(match.group(1))

    return None


def _infer_block_field_from_label(raw_label: str, is_anchor: bool = False) -> Optional[str]:
    if is_anchor:
        return "project_name"

    normalized = _normalize_headerish_text(raw_label)
    if not normalized:
        return None

    if normalized in PROJECT_BLOCK_FIELD_ALIASES:
        return PROJECT_BLOCK_FIELD_ALIASES[normalized]

    for alias, field_key in PROJECT_BLOCK_FIELD_ALIASES.items():
        if normalized == alias or normalized in alias or alias in normalized:
            return field_key

    return None


def _infer_section_mode_and_location(sheet, start_row: int, lookback: int = 8) -> Tuple[str, str]:
    context_lines: List[str] = []

    for row_idx in range(max(1, start_row - lookback), start_row):
        text = _normalized_row_text(sheet, row_idx)
        if text:
            context_lines.append(text)

    context = " ".join(context_lines)

    mode = "all"
    if any(token in context for token in ["ongoing", "under execution", "under progress", "in progress", "running"]):
        mode = "ongoing"
    elif any(token in context for token in ["completed", "executed", "finished"]):
        mode = "completed"

    location_hint = ""
    for city in KNOWN_LOCATION_HINTS:
        if city in context:
            location_hint = city
            break

    return mode, location_hint


def _looks_like_new_numbered_section(sheet, row_idx: int) -> bool:
    row_text = _row_text(sheet, row_idx)
    if not row_text:
        return False

    normalized = _normalize_headerish_text(row_text)

    if any(phrase in normalized for phrase in PROJECT_SECTION_TERMINATORS):
        return True

    if re.search(r"\bproject\b\s*(\d+)\b", normalized):
        return False

    for field_alias in PROJECT_BLOCK_FIELD_ALIASES.keys():
        if field_alias in normalized:
            return False

    if re.match(r"^\d+\b", normalized):
        return True

    return False


def _collect_project_blocks(sheet) -> List[Dict[str, Any]]:
    """
    Detect repeated vertical blocks like:
        Project 1
        Location
        Area in Sft
        Awarded Amount in INR
        ...
    """
    detections: List[Dict[str, Any]] = []
    row_idx = 1

    while row_idx <= sheet.max_row:
        row_text = _row_text(sheet, row_idx)
        if row_text and "project" in _normalize_headerish_text(row_text):
            print(f"DEBUG: candidate row {row_idx} -> {row_text}")

        anchor = _find_project_anchor_in_row(sheet, row_idx)
        if not anchor:
            row_idx += 1
            continue

        anchor_col, anchor_label, project_index = anchor
        mode, location_hint = _infer_section_mode_and_location(sheet, row_idx)

        block_fields: List[Dict[str, Any]] = [
            {
                "row": row_idx,
                "col": anchor_col,
                "label_text": anchor_label,
                "field_key": "project_name",
                "is_anchor": True,
            }
        ]

        next_row = row_idx + 1
        max_scan_until = min(sheet.max_row, row_idx + 12)

        while next_row <= max_scan_until:
            if _looks_like_new_numbered_section(sheet, next_row):
                break

            next_anchor = _find_project_anchor_in_row(sheet, next_row)
            if next_anchor:
                break

            captured_any = False
            for col_idx in range(1, sheet.max_column + 1):
                raw = _safe_text(sheet.cell(row=next_row, column=col_idx).value)
                if not raw:
                    continue

                field_key = _infer_block_field_from_label(raw, is_anchor=False)
                if not field_key:
                    continue

                block_fields.append(
                    {
                        "row": next_row,
                        "col": col_idx,
                        "label_text": raw,
                        "field_key": field_key,
                        "is_anchor": False,
                    }
                )
                captured_any = True
                break

            if not captured_any:
                next_row_text = _normalized_row_text(sheet, next_row)

                if next_row_text and any(phrase in next_row_text for phrase in PROJECT_SECTION_TERMINATORS):
                    break

                if not next_row_text:
                    blank_gap = 0
                    probe_row = next_row
                    while probe_row <= max_scan_until and not _normalized_row_text(sheet, probe_row):
                        blank_gap += 1
                        probe_row += 1
                    if blank_gap >= 2:
                        break

            next_row += 1

        if len(block_fields) >= 2:
            detections.append(
                {
                    "sheet": sheet.title,
                    "anchor_row": row_idx,
                    "anchor_col": anchor_col,
                    "project_index": project_index,
                    "mode": mode,
                    "location_hint": location_hint,
                    "fields": block_fields,
                }
            )

        row_idx = next_row

    return detections


def _get_project_field_value(project: Dict[str, Any], field_key: str, project_index: int) -> Any:
    if field_key == "serial_no":
        return project_index

    if field_key == "project_name":
        value = project.get("project_name")
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "client":
        value = project.get("client")
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "location":
        value = _derive_project_location(project)
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "area_sft":
        value = _derive_project_area(project)
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "value":
        value = _derive_project_value(project)
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "pmc_name":
        value = _derive_project_pmc(project)
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "category":
        value = project.get("category")
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "client_reference":
        value = _derive_client_reference(project)
        return _format_project_scalar(value) if _is_meaningful_value(value) else None

    if field_key == "start_date":
        value = project.get("start_date")
        return _format_date_value(value) if _is_meaningful_value(value) else None

    if field_key == "end_date":
        value = project.get("end_date")
        return _format_date_value(value) if _is_meaningful_value(value) else None

    if field_key == "status":
        value = project.get("status")
        if not _is_meaningful_value(value):
            return None
        normalized = _normalize_status(value)
        return normalized.title() if normalized in {"ongoing", "completed"} else value

    value = project.get(field_key)
    return _format_project_scalar(value) if _is_meaningful_value(value) else None


def _build_log_row(
    sheet_name: str,
    label_cell: str,
    label_text: str,
    normalized_label: str,
    mapped_field_key: str = "",
    target_cell: str = "",
    target_merged_range: str = "",
    filled_value: Any = "",
    status: str = REVIEW_STATUS_SKIPPED,
    note: str = "",
    reason: str = "",
    resolver: str = "project_block_filler",
    layout_type: str = "project_block",
    layout_confidence: float = 1.0,
    project_index: Optional[int] = None,
    project_mode: str = "",
):
    return {
        "sheet": sheet_name,
        "sheet_name": sheet_name,
        "label_cell": label_cell,
        "label_text": label_text,
        "normalized_label": normalized_label,
        "cell_type": CELL_TYPE_SUBFIELD,
        "active_section": SECTION_PROJECTS,
        "section": SECTION_PROJECTS,
        "mapped_field_key": mapped_field_key,
        "field_key": mapped_field_key,
        "match_type": MATCH_TYPE_SECTION_CONTEXT,
        "confidence": 100,
        "semantic_confidence": 1.0,
        "layout_confidence": layout_confidence,
        "total_confidence": 1.0,
        "resolver": resolver,
        "layout_type": layout_type,
        "table_type": TABLE_TYPE_PROJECTS,
        "target_cell": target_cell,
        "target_merged_range": target_merged_range,
        "filled_value": filled_value,
        "value_preview": _safe_preview(filled_value),
        "status": status,
        "write_result": status,
        "note": note,
        "reason": reason or note,
        "project_index": project_index,
        "project_mode": project_mode,
    }


def _resolve_target_for_block_field(sheet, row_idx: int, col_idx: int, label_text: str) -> Dict[str, Any]:
    return resolve_target_cell(
        sheet=sheet,
        label_row=row_idx,
        label_col=col_idx,
        label_text=label_text,
        cell_type=CELL_TYPE_SUBFIELD,
    )


def _write_block_field(
    sheet,
    row_idx: int,
    col_idx: int,
    label_text: str,
    value: Any,
) -> Tuple[str, str, float, str, str, Any]:
    resolution = _resolve_target_for_block_field(sheet, row_idx, col_idx, label_text)
    target_cell = resolution.get("target_cell")
    target_coordinate = resolution.get("target_coordinate", "")
    target_merged_range = resolution.get("target_merged_range", "") or ""
    layout_confidence = float(resolution.get("layout_confidence", 0.0))
    target_score = int(resolution.get("score", 0))

    if target_cell is None:
        return REVIEW_STATUS_SKIPPED, SKIP_REASON_NO_TARGET, layout_confidence, target_coordinate, target_merged_range, ""

    if layout_confidence < MIN_LAYOUT_CONFIDENCE or target_score < int(MIN_LAYOUT_CONFIDENCE * 100):
        return REVIEW_STATUS_SKIPPED, SKIP_REASON_LOW_CONFIDENCE, layout_confidence, target_coordinate, target_merged_range, ""

    anchor = get_merged_anchor_cell(sheet, target_cell)

    if _is_formula_cell(anchor):
        return REVIEW_STATUS_SKIPPED, SKIP_REASON_FORMULA_CELL, layout_confidence, anchor.coordinate, target_merged_range, ""

    existing_value = getattr(anchor, "value", None)
    if _is_meaningful_value(existing_value):
        return REVIEW_STATUS_SKIPPED, SKIP_REASON_EXISTING_VALUE, layout_confidence, anchor.coordinate, target_merged_range, existing_value

    anchor.value = value
    return REVIEW_STATUS_FILLED, "FILLED", layout_confidence, anchor.coordinate, target_merged_range, value


def fill_project_blocks(wb, master_data):
    print("DEBUG: project_block_filler is running")

    log_rows: List[Dict[str, Any]] = []

    for sheet in wb.worksheets:
        print(f"DEBUG: scanning sheet for project blocks -> {sheet.title}")

        project_blocks = _collect_project_blocks(sheet)

        print(f"DEBUG: detected project blocks in {sheet.title}: {len(project_blocks)}")
        for block in project_blocks:
            print(
                "DEBUG: block",
                {
                    "anchor_row": block.get("anchor_row"),
                    "anchor_col": block.get("anchor_col"),
                    "project_index": block.get("project_index"),
                    "mode": block.get("mode"),
                    "location_hint": block.get("location_hint"),
                    "field_count": len(block.get("fields", [])),
                }
            )

        if not project_blocks:
            continue

        section_groups: List[List[Dict[str, Any]]] = []
        current_group: List[Dict[str, Any]] = []

        for block in project_blocks:
            if not current_group:
                current_group.append(block)
                continue

            prev = current_group[-1]
            same_mode = prev.get("mode", "all") == block.get("mode", "all")
            same_location = prev.get("location_hint", "") == block.get("location_hint", "")
            close_enough = abs(int(block.get("anchor_row", 0)) - int(prev.get("anchor_row", 0))) <= 20

            if same_mode and same_location and close_enough:
                current_group.append(block)
            else:
                section_groups.append(current_group)
                current_group = [block]

        if current_group:
            section_groups.append(current_group)

        for group in section_groups:
            mode = group[0].get("mode", "all")
            location_hint = group[0].get("location_hint", "")

            print(
                "DEBUG: selecting projects for group",
                {
                    "sheet": sheet.title,
                    "mode": mode,
                    "location_hint": location_hint,
                    "group_size": len(group),
                }
            )

            projects = _select_projects_for_blocks(
                master_data=master_data,
                mode=mode,
                limit=len(group),
                location_hint=location_hint,
            )

            print(f"DEBUG: selected projects count = {len(projects)}")
            for idx, project in enumerate(projects, start=1):
                print(
                    "DEBUG: selected project",
                    idx,
                    {
                        "project_name": project.get("project_name"),
                        "client": project.get("client"),
                        "location": project.get("location"),
                        "bucket": project.get("bucket"),
                        "status": project.get("status"),
                    }
                )

            if not projects:
                for block in group:
                    anchor_row = int(block["anchor_row"])
                    anchor_col = int(block["anchor_col"])
                    label_text = _safe_text(sheet.cell(row=anchor_row, column=anchor_col).value)
                    log_rows.append(
                        _build_log_row(
                            sheet_name=sheet.title,
                            label_cell=sheet.cell(row=anchor_row, column=anchor_col).coordinate,
                            label_text=label_text,
                            normalized_label=_normalize_headerish_text(label_text),
                            mapped_field_key="__projects__",
                            status=REVIEW_STATUS_SKIPPED,
                            note="NO_MASTER_VALUE",
                            reason="NO_MASTER_VALUE",
                            project_index=block.get("project_index"),
                            project_mode=mode,
                        )
                    )
                continue

            for selected_idx, block in enumerate(group, start=1):
                if selected_idx > len(projects):
                    break

                project = _normalize_project_record(projects[selected_idx - 1])

                for field_item in block.get("fields", []):
                    row_idx = int(field_item["row"])
                    col_idx = int(field_item["col"])
                    label_text = field_item["label_text"]
                    field_key = field_item["field_key"]

                    value = _get_project_field_value(
                        project=project,
                        field_key=field_key,
                        project_index=selected_idx,
                    )

                    print(
                        "DEBUG: field evaluation",
                        {
                            "row": row_idx,
                            "col": col_idx,
                            "label_text": label_text,
                            "field_key": field_key,
                            "derived_value": value,
                            "project_index": selected_idx,
                        }
                    )

                    if value in (None, ""):
                        continue

                    label_cell = sheet.cell(row=row_idx, column=col_idx).coordinate
                    normalized_label = _normalize_headerish_text(label_text)

                    try:
                        (
                            write_status,
                            write_reason,
                            layout_confidence,
                            target_coordinate,
                            target_merged_range,
                            existing_or_written_value,
                        ) = _write_block_field(
                            sheet=sheet,
                            row_idx=row_idx,
                            col_idx=col_idx,
                            label_text=label_text,
                            value=value,
                        )
                    except Exception as exc:
                        print(f"DEBUG: write exception at {label_cell} -> {exc}")
                        log_rows.append(
                            _build_log_row(
                                sheet_name=sheet.title,
                                label_cell=label_cell,
                                label_text=label_text,
                                normalized_label=normalized_label,
                                mapped_field_key=f"project.{field_key}",
                                status=REVIEW_STATUS_SKIPPED,
                                note=f"WRITE_ERROR: {str(exc)}",
                                reason=SKIP_REASON_EXCEPTION,
                                project_index=selected_idx,
                                project_mode=mode,
                            )
                        )
                        continue

                    print(
                        "DEBUG: write result",
                        {
                            "label_cell": label_cell,
                            "target_cell": target_coordinate,
                            "write_status": write_status,
                            "write_reason": write_reason,
                            "layout_confidence": layout_confidence,
                            "value": existing_or_written_value,
                        }
                    )

                    if write_status == REVIEW_STATUS_FILLED:
                        log_rows.append(
                            _build_log_row(
                                sheet_name=sheet.title,
                                label_cell=label_cell,
                                label_text=label_text,
                                normalized_label=normalized_label,
                                mapped_field_key=f"project.{field_key}",
                                target_cell=target_coordinate,
                                target_merged_range=target_merged_range,
                                filled_value=existing_or_written_value,
                                status=REVIEW_STATUS_FILLED,
                                note=f"FILLED_MODE:{mode}",
                                reason="FILLED",
                                layout_confidence=layout_confidence,
                                project_index=selected_idx,
                                project_mode=mode,
                            )
                        )
                    else:
                        note = {
                            SKIP_REASON_NO_TARGET: "NO_SAFE_TARGET",
                            SKIP_REASON_LOW_CONFIDENCE: "LOW_LAYOUT_CONFIDENCE",
                            SKIP_REASON_FORMULA_CELL: "FORMULA_CELL",
                            SKIP_REASON_EXISTING_VALUE: "TARGET_ALREADY_HAS_VALUE",
                        }.get(write_reason, write_reason)

                        log_rows.append(
                            _build_log_row(
                                sheet_name=sheet.title,
                                label_cell=label_cell,
                                label_text=label_text,
                                normalized_label=normalized_label,
                                mapped_field_key=f"project.{field_key}",
                                target_cell=target_coordinate,
                                target_merged_range=target_merged_range,
                                filled_value=existing_or_written_value,
                                status=REVIEW_STATUS_SKIPPED,
                                note=note,
                                reason=write_reason,
                                layout_confidence=layout_confidence,
                                project_index=selected_idx,
                                project_mode=mode,
                            )
                        )

    return wb, log_rows