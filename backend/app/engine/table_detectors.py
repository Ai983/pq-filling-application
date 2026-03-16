from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Tuple

from app.core.config import MIN_TABLE_HEADER_MATCHES
from app.core.constants import (
    FINANCIAL_ROW_KEYWORDS,
    PROJECT_TABLE_HEADER_KEYWORDS,
    RESOURCE_ROW_KEYWORDS,
    TABLE_TYPE_COLUMN_HEADER,
    TABLE_TYPE_COMPLIANCE,
    TABLE_TYPE_FINANCIAL,
    TABLE_TYPE_PROJECTS,
    TABLE_TYPE_RESOURCE,
    TABLE_TYPE_ROW_LABEL,
    TABLE_TYPE_UNKNOWN,
)
from app.engine.utils import normalize_text


# ---------------------------------------------------------------------
# Deterministic row maps
# ---------------------------------------------------------------------

TURNOVER_ROW_MAP = {
    "fy 2019-20": "financial.turnover.fy2019_20",
    "2019-20": "financial.turnover.fy2019_20",
    "2019-2020": "financial.turnover.fy2019_20",
    "fy 2020-21": "financial.turnover.fy2020_21",
    "2020-21": "financial.turnover.fy2020_21",
    "2020-2021": "financial.turnover.fy2020_21",
    "fy 2021-22": "financial.turnover.fy2021_22",
    "2021-22": "financial.turnover.fy2021_22",
    "2021-2022": "financial.turnover.fy2021_22",
    "fy 2022-23": "financial.turnover.fy2022_23",
    "2022-23": "financial.turnover.fy2022_23",
    "2022-2023": "financial.turnover.fy2022_23",
    "fy 2023-24": "financial.turnover.fy2023_24",
    "2023-24": "financial.turnover.fy2023_24",
    "2023-2024": "financial.turnover.fy2023_24",
    "fy 2024-25": "financial.turnover.fy2024_25",
    "2024-25": "financial.turnover.fy2024_25",
    "2024-2025": "financial.turnover.fy2024_25",
}

MANPOWER_ROW_MAP = {
    "engineers": "resource.manpower.engineers",
    "engineer": "resource.manpower.engineers",
    "technical staff": "resource.manpower.engineers",
    "number of technical staff": "resource.manpower.engineers",
    "no of technical staff": "resource.manpower.engineers",
    "no. of technical staff": "resource.manpower.engineers",
    "supervisors": "resource.manpower.supervisors",
    "supervisor": "resource.manpower.supervisors",
    "skilled labour": "resource.manpower.skilled_labour",
    "skilled labor": "resource.manpower.skilled_labour",
    "skilled workers": "resource.manpower.skilled_labour",
    "unskilled labour": "resource.manpower.unskilled_labour",
    "unskilled labor": "resource.manpower.unskilled_labour",
    "unskilled workers": "resource.manpower.unskilled_labour",
    "staff strength": "resource.manpower.total_staff",
    "total staff": "resource.manpower.total_staff",
    "total manpower": "resource.manpower.total_staff",
    "capacity to mobilize manpower": "resource.manpower.total_staff",
    "manpower capacity": "resource.manpower.total_staff",
}

MACHINERY_ROW_MAP = {
    "machine details": "resource.machinery.details",
    "machinery details": "resource.machinery.details",
    "equipment details": "resource.machinery.details",
    "tools & tackles": "resource.machinery.details",
    "tools and tackles": "resource.machinery.details",
    "list of machinery": "resource.machinery.details",
    "machinery": "resource.machinery.details",
    "available plant & machinery": "resource.machinery.details",
    "available plant and machinery": "resource.machinery.details",
    "plant & machinery": "resource.machinery.details",
    "plant and machinery": "resource.machinery.details",
    "factory production capacity": "resource.machinery.details",
    "production capacity": "resource.machinery.details",
}

COMPLIANCE_ROW_MAP = {
    "iso certified": "compliance.iso_certified",
    "quality policy": "compliance.quality_policy",
    "safety policy": "compliance.safety_policy",
    "ohs policy": "compliance.ohs_policy",
    "ohse policy": "compliance.ohs_policy",
    "litigation": "compliance.litigation",
    "arbitration": "compliance.arbitration",
    "audit reports": "compliance.audit_reports",
    "msme": "tax.msme",
    "msme registration number": "tax.msme",
    "msme registration no": "tax.msme",
    "pf": "tax.pf",
    "pf registration no": "tax.pf",
    "pf registration number": "tax.pf",
    "esi": "tax.esi",
    "esi registration no": "tax.esi",
    "pan": "tax.pan",
    "pan number": "tax.pan",
    "gst": "tax.gst.primary",
    "gst number": "tax.gst.primary",
    "gstin": "tax.gst.primary",
}

PROJECT_HEADER_ALIASES = {
    "sr no": "serial_no",
    "sr. no": "serial_no",
    "s no": "serial_no",
    "serial no": "serial_no",
    "serial number": "serial_no",
    "project name": "project_name",
    "name of project": "project_name",
    "project": "project_name",
    "work name": "project_name",
    "client": "client",
    "client name": "client",
    "customer": "client",
    "type of work": "category",
    "nature of work": "category",
    "scope of work": "category",
    "volume of work": "value",
    "value of work order": "value",
    "work order value": "value",
    "contract value": "value",
    "awarded value": "value",
    "project value": "value",
    "value": "value",
    "location": "location",
    "city": "location",
    "state": "location",
    "site location": "location",
    "start date": "start_date",
    "commencement date": "start_date",
    "order date": "start_date",
    "po date": "start_date",
    "completion date": "end_date",
    "end date": "end_date",
    "work completion date": "end_date",
    "status": "status",
    "contact person": "contact_name",
    "contact": "contact_name",
    "contact no": "contact_phone",
    "contact number": "contact_phone",
    "phone": "contact_phone",
    "mobile": "contact_phone",
}

for canonical_key, aliases in PROJECT_TABLE_HEADER_KEYWORDS.items():
    for alias in aliases:
        alias_norm = normalize_text(alias)
        if alias_norm and alias_norm not in PROJECT_HEADER_ALIASES:
            PROJECT_HEADER_ALIASES[alias_norm] = canonical_key

for field_key, aliases in FINANCIAL_ROW_KEYWORDS.items():
    for alias in aliases:
        alias_norm = normalize_text(alias)
        if alias_norm:
            TURNOVER_ROW_MAP.setdefault(alias_norm, field_key)

for field_key, aliases in RESOURCE_ROW_KEYWORDS.items():
    for alias in aliases:
        alias_norm = normalize_text(alias)
        if alias_norm:
            if field_key.startswith("resource.manpower."):
                MANPOWER_ROW_MAP.setdefault(alias_norm, field_key)
            else:
                MACHINERY_ROW_MAP.setdefault(alias_norm, field_key)


# ---------------------------------------------------------------------
# Low-level helpers
# ---------------------------------------------------------------------

def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _normalize_headerish_text(text: str) -> str:
    normalized = normalize_text(text)
    if not normalized:
        return ""
    normalized = normalized.replace("&", " and ")
    normalized = re.sub(r"\bno\b", "number", normalized)
    normalized = re.sub(r"\bsr\b", "serial", normalized)
    normalized = " ".join(normalized.split()).strip()
    return normalized


def _row_nonempty_cells(sheet, row_idx: int) -> List[Tuple[int, str]]:
    result: List[Tuple[int, str]] = []

    for col_idx in range(1, sheet.max_column + 1):
        value = sheet.cell(row=row_idx, column=col_idx).value
        text = _safe_text(value)
        if text:
            result.append((col_idx, _normalize_headerish_text(text)))

    return result


def _count_blank_cells_in_row(sheet, row_idx: int, start_col: int, end_col: int) -> int:
    blank_count = 0
    for col_idx in range(start_col, end_col + 1):
        value = sheet.cell(row=row_idx, column=col_idx).value
        if not _safe_text(value):
            blank_count += 1
    return blank_count


def _has_multiple_blank_targets_to_right(sheet, row_idx: int, label_col: int, scan_width: int = 4) -> bool:
    blank_count = 0
    for offset in range(1, scan_width + 1):
        col = label_col + offset
        if col > sheet.max_column:
            break
        value = sheet.cell(row=row_idx, column=col).value
        if not _safe_text(value):
            blank_count += 1
    return blank_count >= 1


def _normalize_project_header(text: str) -> Optional[str]:
    normalized = _normalize_headerish_text(text)
    if not normalized:
        return None

    if normalized in PROJECT_HEADER_ALIASES:
        return PROJECT_HEADER_ALIASES[normalized]

    for alias, canonical in PROJECT_HEADER_ALIASES.items():
        if alias in normalized or normalized in alias:
            return canonical

    return None


def _normalize_year_token(text: str) -> str:
    normalized = _normalize_headerish_text(text)
    if not normalized:
        return ""

    normalized = normalized.replace("financial year", "fy")
    normalized = re.sub(r"\bfy\b", "fy ", normalized)
    normalized = re.sub(r"[^0-9a-zA-Z]+", " ", normalized)
    normalized = " ".join(normalized.split()).strip()

    match = re.search(r"(20\d{2})\s*(20\d{2}|\d{2})", normalized)
    if match:
        first = match.group(1)
        second_raw = match.group(2)
        second = second_raw if len(second_raw) == 4 else f"{first[:2]}{second_raw}"
        short_second = second[-2:]
        return f"{first}-{short_second}"

    return normalized


def _year_field_key_from_text(text: str) -> Optional[str]:
    year_token = _normalize_year_token(text)
    if not year_token:
        return None

    if year_token in TURNOVER_ROW_MAP:
        return TURNOVER_ROW_MAP[year_token]

    for alias, field_key in TURNOVER_ROW_MAP.items():
        if year_token == alias or year_token in alias or alias in year_token:
            return field_key

    return None


def _contains_financial_context(text: str) -> bool:
    normalized = _normalize_headerish_text(text)
    markers = [
        "turnover",
        "annual turnover",
        "annual turn over",
        "financial",
        "revenue",
    ]
    return any(marker in normalized for marker in markers)


def _contains_project_section_context(text: str) -> Optional[str]:
    normalized = _normalize_headerish_text(text)
    if not normalized:
        return None

    if "under execution" in normalized or "ongoing" in normalized or "under progress" in normalized:
        return "ongoing"
    if "executed" in normalized or "completed" in normalized:
        return "completed"
    if "project" in normalized:
        return "all"
    return None


def _looks_like_turnover_heading(text: str) -> bool:
    normalized = _normalize_headerish_text(text)
    return bool(normalized) and (
        "annual turnover" in normalized
        or "annual turn over" in normalized
        or normalized == "turnover"
        or "turnover cr" in normalized
    )


# ---------------------------------------------------------------------
# Simple table row detection
# ---------------------------------------------------------------------

def detect_table_field_key(normalized_label: str) -> str | None:
    label = _normalize_headerish_text(normalized_label)

    if label in TURNOVER_ROW_MAP:
        return TURNOVER_ROW_MAP[label]

    year_field_key = _year_field_key_from_text(label)
    if year_field_key:
        return year_field_key

    if label in MANPOWER_ROW_MAP:
        return MANPOWER_ROW_MAP[label]

    if label in MACHINERY_ROW_MAP:
        return MACHINERY_ROW_MAP[label]

    if label in COMPLIANCE_ROW_MAP:
        return COMPLIANCE_ROW_MAP[label]

    return None


def detect_table_type(normalized_label: str) -> str:
    label = _normalize_headerish_text(normalized_label)

    if label in TURNOVER_ROW_MAP or _year_field_key_from_text(label):
        return TABLE_TYPE_FINANCIAL

    if label in MANPOWER_ROW_MAP or label in MACHINERY_ROW_MAP:
        return TABLE_TYPE_RESOURCE

    if label in COMPLIANCE_ROW_MAP:
        return TABLE_TYPE_COMPLIANCE

    return TABLE_TYPE_UNKNOWN


def is_turnover_row(normalized_label: str) -> bool:
    label = _normalize_headerish_text(normalized_label)
    return label in TURNOVER_ROW_MAP or _year_field_key_from_text(label) is not None


def is_manpower_row(normalized_label: str) -> bool:
    return _normalize_headerish_text(normalized_label) in MANPOWER_ROW_MAP


def is_machinery_row(normalized_label: str) -> bool:
    return _normalize_headerish_text(normalized_label) in MACHINERY_ROW_MAP


def is_compliance_row(normalized_label: str) -> bool:
    return _normalize_headerish_text(normalized_label) in COMPLIANCE_ROW_MAP


def detect_row_label_tables(sheet) -> List[Dict[str, Any]]:
    detections: List[Dict[str, Any]] = []
    current_group: List[Dict[str, Any]] = []
    recent_context_type: Optional[str] = None
    recent_context_row: Optional[int] = None

    for row_idx in range(1, sheet.max_row + 1):
        matched_items: List[Dict[str, Any]] = []

        row_text_joined = " ".join(
            _normalize_headerish_text(_safe_text(sheet.cell(row=row_idx, column=col).value))
            for col in range(1, sheet.max_column + 1)
            if _safe_text(sheet.cell(row=row_idx, column=col).value)
        ).strip()

        if row_text_joined:
            if _contains_financial_context(row_text_joined):
                recent_context_type = TABLE_TYPE_FINANCIAL
                recent_context_row = row_idx
            elif "capacity" in row_text_joined or "technical staff" in row_text_joined:
                recent_context_type = TABLE_TYPE_RESOURCE
                recent_context_row = row_idx

        for col_idx in range(1, sheet.max_column + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            text = _safe_text(value)
            if not text:
                continue

            normalized = _normalize_headerish_text(text)
            field_key = detect_table_field_key(normalized)
            if not field_key:
                continue

            if not _has_multiple_blank_targets_to_right(sheet, row_idx, col_idx):
                continue

            detected_type = detect_table_type(normalized)

            if detected_type == TABLE_TYPE_FINANCIAL and recent_context_type == TABLE_TYPE_FINANCIAL:
                if recent_context_row is not None and (row_idx - recent_context_row) <= 8:
                    detected_type = TABLE_TYPE_FINANCIAL

            matched_items.append(
                {
                    "sheet": sheet.title,
                    "row": row_idx,
                    "label_col": col_idx,
                    "label_text": text,
                    "normalized_label": normalized,
                    "field_key": field_key,
                    "table_type": detected_type,
                }
            )

        if matched_items:
            current_group.extend(matched_items)
        else:
            if len(current_group) >= 1:
                detections.append(
                    {
                        "sheet": sheet.title,
                        "table_type": TABLE_TYPE_ROW_LABEL,
                        "rows": current_group[:],
                    }
                )
            current_group = []

    if len(current_group) >= 1:
        detections.append(
            {
                "sheet": sheet.title,
                "table_type": TABLE_TYPE_ROW_LABEL,
                "rows": current_group[:],
            }
        )

    return detections


# ---------------------------------------------------------------------
# Horizontal financial detection
# ---------------------------------------------------------------------

def detect_horizontal_financial_tables(sheet) -> List[Dict[str, Any]]:
    detections: List[Dict[str, Any]] = []

    for row_idx in range(1, sheet.max_row + 1):
        row_cells = _row_nonempty_cells(sheet, row_idx)
        if not row_cells:
            continue

        heading_found = any(_looks_like_turnover_heading(text) for _, text in row_cells)
        if not heading_found:
            continue

        for year_row in range(row_idx, min(sheet.max_row, row_idx + 3) + 1):
            year_map: Dict[str, int] = {}
            detected_headers: List[Dict[str, Any]] = []

            for col_idx, text in _row_nonempty_cells(sheet, year_row):
                field_key = _year_field_key_from_text(text)
                if field_key:
                    year_map[field_key] = col_idx
                    detected_headers.append(
                        {
                            "column": col_idx,
                            "raw_header": _safe_text(sheet.cell(row=year_row, column=col_idx).value),
                            "normalized_header": text,
                            "field_key": field_key,
                        }
                    )

            if len(year_map) >= 2:
                detections.append(
                    {
                        "sheet": sheet.title,
                        "table_type": TABLE_TYPE_FINANCIAL,
                        "layout": "horizontal_year_headers",
                        "heading_row": row_idx,
                        "header_row": year_row,
                        "value_row": min(sheet.max_row, year_row + 1),
                        "year_column_map": year_map,
                        "detected_headers": detected_headers,
                    }
                )
                break

    return detections


# ---------------------------------------------------------------------
# Project / column-header table detection
# ---------------------------------------------------------------------

def detect_project_tables(sheet) -> List[Dict[str, Any]]:
    """
    Detect project tables even when there is NO project_name column.

    This is important for PQ forms where the table only contains:
    client / type of work / value / order date / completion date
    """
    detections = []

    for row_idx in range(1, sheet.max_row + 1):
        row_values = _row_nonempty_cells(sheet, row_idx)
        if not row_values:
            continue

        column_map: Dict[str, int] = {}
        detected_headers: List[Dict[str, Any]] = []

        for col_idx, text in row_values:
            canonical = _normalize_project_header(text)
            if canonical:
                column_map[canonical] = col_idx
                detected_headers.append(
                    {
                        "column": col_idx,
                        "raw_header": _safe_text(sheet.cell(row=row_idx, column=col_idx).value),
                        "normalized_header": text,
                        "canonical_key": canonical,
                    }
                )

        if column_map:
            min_col = min(column_map.values())
            max_col = max(column_map.values())
        else:
            min_col = 1
            max_col = min(5, sheet.max_column)

        blank_cells_below = _count_blank_cells_in_row(
            sheet=sheet,
            row_idx=min(row_idx + 1, sheet.max_row),
            start_col=min_col,
            end_col=max_col,
        )

        project_mode = "all"
        context_text_parts = []
        for scan_row in range(max(1, row_idx - 4), row_idx + 1):
            row_text = " ".join(
                _safe_text(sheet.cell(row=scan_row, column=col).value)
                for col in range(1, min(sheet.max_column, max_col + 2) + 1)
                if _safe_text(sheet.cell(row=scan_row, column=col).value)
            ).strip()
            if row_text:
                context_text_parts.append(row_text)

        context_text = " ".join(context_text_parts)
        mode_hint = _contains_project_section_context(context_text)
        if mode_hint:
            project_mode = mode_hint

        required_match_count = MIN_TABLE_HEADER_MATCHES

        has_strong_project_signature = (
            "client" in column_map
            and (
                "value" in column_map
                or "category" in column_map
                or "start_date" in column_map
                or "end_date" in column_map
            )
        )

        has_project_name = "project_name" in column_map

        # Accept either:
        # 1) classic project table with project_name
        # 2) PQ execution table without project_name but with strong project signature
        is_valid_project_table = (
            len(column_map) >= required_match_count
            and (has_project_name or has_strong_project_signature)
        )

        if is_valid_project_table:
            detections.append(
                {
                    "sheet": sheet.title,
                    "table_type": TABLE_TYPE_PROJECTS,
                    "header_row": row_idx,
                    "start_row": row_idx + 1 if row_idx < sheet.max_row else row_idx,
                    "column_map": column_map,
                    "detected_headers": detected_headers,
                    "below_blank_count": blank_cells_below,
                    "project_mode": project_mode,
                }
            )

    return detections


def detect_column_header_tables(sheet) -> List[Dict[str, Any]]:
    tables: List[Dict[str, Any]] = []

    project_tables = detect_project_tables(sheet)
    for item in project_tables:
        tables.append(
            {
                **item,
                "table_type": TABLE_TYPE_COLUMN_HEADER,
                "subtype": TABLE_TYPE_PROJECTS,
            }
        )

    return tables


# ---------------------------------------------------------------------
# Combined sheet-level detection
# ---------------------------------------------------------------------

def detect_sheet_tables(sheet) -> Dict[str, Any]:
    row_label_tables = detect_row_label_tables(sheet)
    project_tables = detect_project_tables(sheet)
    column_header_tables = detect_column_header_tables(sheet)
    horizontal_financial_tables = detect_horizontal_financial_tables(sheet)

    return {
        "sheet": sheet.title,
        "row_label_tables": row_label_tables,
        "project_tables": project_tables,
        "column_header_tables": column_header_tables,
        "horizontal_financial_tables": horizontal_financial_tables,
    }