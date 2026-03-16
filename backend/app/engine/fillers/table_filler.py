from __future__ import annotations

from typing import Any, Dict, List

from app.core.config import (
    MIN_LAYOUT_CONFIDENCE,
    MIN_TOTAL_CONFIDENCE_FOR_TABLE_FILL,
)
from app.core.constants import (
    MATCH_TYPE_TABLE_CONTEXT,
    REVIEW_STATUS_FILLED,
    REVIEW_STATUS_SKIPPED,
    REVIEW_STATUS_UNMATCHED,
    SKIP_REASON_EXCEPTION,
    SKIP_REASON_EXISTING_VALUE,
    SKIP_REASON_FORMULA_CELL,
    SKIP_REASON_LOW_CONFIDENCE,
    SKIP_REASON_NO_MATCH,
    SKIP_REASON_NO_TARGET,
)
from app.engine.master_loader import get_master_value, get_master_value_variants
from app.engine.skip_rules import should_skip_no_master_value
from app.engine.table_detectors import detect_table_field_key, detect_table_type
from app.engine.target_cell_resolver import (
    get_merged_anchor_cell,
    resolve_table_value_cell,
)
from app.engine.utils import is_blank, normalize_text


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


def _normalize_yes_no(value: Any) -> Any:
    if value is None:
        return None

    normalized = normalize_text(str(value))
    if normalized in {"yes", "y", "true", "available", "applicable", "provided"}:
        return "Yes"
    if normalized in {"no", "n", "false", "not available", "not applicable", "nil"}:
        return "No"
    return value


def _format_numeric_like(value: Any) -> Any:
    if value is None:
        return None

    if isinstance(value, float) and value.is_integer():
        return int(value)

    return value


def _format_table_value(field_key: str, value: Any, value_variants: Dict[str, Any] | None = None) -> Any:
    """
    Conservative deterministic formatting for row-table values.
    """
    if value is None:
        return None

    if value_variants is None:
        value_variants = {}

    preferred = value_variants.get("preferred_display")
    if preferred not in (None, ""):
        return preferred

    if field_key.startswith("compliance.") or field_key in {
        "tax.msme",
        "tax.pf",
        "tax.esi",
    }:
        return _normalize_yes_no(value)

    if field_key.startswith("financial.turnover.") or field_key.startswith("resource.manpower."):
        return _format_numeric_like(value)

    return value


def _combined_table_confidence(semantic_confidence: float, layout_confidence: float) -> float:
    return round((semantic_confidence * 0.60) + (layout_confidence * 0.40), 4)


def _build_log_row(
    item: dict,
    mapped_field_key: str = "",
    match_type: str = "",
    confidence: int = 0,
    semantic_confidence: float = 0.0,
    layout_confidence: float = 0.0,
    total_confidence: float = 0.0,
    target_cell: str = "",
    target_merged_range: str = "",
    filled_value: Any = "",
    status: str = REVIEW_STATUS_SKIPPED,
    note: str = "",
    resolver: str = "",
    layout_type: str = "",
    reason: str = "",
    table_type: str = "",
):
    return {
        "sheet": item.get("sheet", ""),
        "sheet_name": item.get("sheet", ""),
        "label_cell": item.get("cell", ""),
        "label_text": item.get("value", ""),
        "normalized_label": item.get("normalized_value", ""),
        "cell_type": item.get("cell_type", ""),
        "active_section": item.get("active_section", ""),
        "section": item.get("active_section", ""),
        "mapped_field_key": mapped_field_key,
        "field_key": mapped_field_key,
        "match_type": match_type,
        "confidence": confidence,
        "semantic_confidence": semantic_confidence,
        "layout_confidence": layout_confidence,
        "total_confidence": total_confidence,
        "resolver": resolver,
        "layout_type": layout_type,
        "table_type": table_type,
        "target_cell": target_cell,
        "target_merged_range": target_merged_range,
        "filled_value": filled_value,
        "value_preview": _safe_preview(filled_value),
        "status": status,
        "write_result": status,
        "note": note,
        "reason": reason or note,
    }


def _candidate_same_row_targets(sheet, row: int, label_col: int, max_scan: int = 6) -> List[Any]:
    candidates: List[Any] = []

    for offset in range(1, max_scan + 1):
        col = label_col + offset
        if col > sheet.max_column:
            break
        candidates.append(sheet.cell(row=row, column=col))

    unique = []
    seen = set()
    for cell in candidates:
        anchor = get_merged_anchor_cell(sheet, cell)
        if anchor.coordinate in seen:
            continue
        seen.add(anchor.coordinate)
        unique.append(anchor)

    return unique


def _find_alternate_blank_target(sheet, row: int, label_col: int, primary_coordinate: str = ""):
    for anchor in _candidate_same_row_targets(sheet, row=row, label_col=label_col, max_scan=6):
        if primary_coordinate and anchor.coordinate == primary_coordinate:
            continue
        if _is_formula_cell(anchor):
            continue
        if is_blank(anchor.value):
            return anchor
    return None


def fill_table_row_field(item, wb, master_data):
    sheet = wb[item["sheet"]]
    normalized_label = item.get("normalized_value", "")

    field_key = detect_table_field_key(normalized_label)
    table_type = detect_table_type(normalized_label)

    if not field_key:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_UNMATCHED,
            note="UNMATCHED",
            reason=SKIP_REASON_NO_MATCH,
            table_type=table_type,
        )

    value = get_master_value(master_data, field_key)
    value_variants = get_master_value_variants(master_data, field_key)

    skip, reason = should_skip_no_master_value(value)
    if skip:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            total_confidence=1.0,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
            table_type=table_type,
        )

    value = _format_table_value(field_key, value, value_variants=value_variants)

    resolution = resolve_table_value_cell(
        sheet=sheet,
        row=item["row"],
        label_col=item["column"],
    )

    target_cell = resolution.get("target_cell")
    target_coordinate = resolution.get("target_coordinate", "")
    target_merged_range = resolution.get("target_merged_range", "") or ""
    layout_confidence = float(resolution.get("layout_confidence", 0.0))
    resolver = resolution.get("resolver", "")
    layout_type = resolution.get("layout_type", "")
    target_score = int(resolution.get("score", 0))
    resolution_reason = resolution.get("reason", "")

    total_confidence = _combined_table_confidence(1.0, layout_confidence)

    if target_cell is None:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="NO_SAFE_TARGET",
            reason=resolution_reason or SKIP_REASON_NO_TARGET,
            table_type=table_type,
        )

    if layout_confidence < MIN_LAYOUT_CONFIDENCE:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            target_cell=target_coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="LOW_LAYOUT_CONFIDENCE",
            reason=SKIP_REASON_LOW_CONFIDENCE,
            table_type=table_type,
        )

    if target_score < int(MIN_TOTAL_CONFIDENCE_FOR_TABLE_FILL * 100) - 5:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            target_cell=target_coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="LOW_LAYOUT_SCORE",
            reason=SKIP_REASON_LOW_CONFIDENCE,
            table_type=table_type,
        )

    if total_confidence < MIN_TOTAL_CONFIDENCE_FOR_TABLE_FILL:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            target_cell=target_coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="LOW_TOTAL_CONFIDENCE",
            reason=SKIP_REASON_LOW_CONFIDENCE,
            table_type=table_type,
        )

    anchor = get_merged_anchor_cell(sheet, target_cell)

    if _is_formula_cell(anchor) or not is_blank(anchor.value):
        alternate_anchor = _find_alternate_blank_target(
            sheet=sheet,
            row=item["row"],
            label_col=item["column"],
            primary_coordinate=anchor.coordinate,
        )
        if alternate_anchor is not None:
            anchor = alternate_anchor
            target_coordinate = anchor.coordinate
            target_merged_range = ""

    if _is_formula_cell(anchor):
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            target_cell=anchor.coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="FORMULA_CELL",
            reason=SKIP_REASON_FORMULA_CELL,
            table_type=table_type,
        )

    existing_value = getattr(anchor, "value", None)
    if not is_blank(existing_value):
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            target_cell=anchor.coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            filled_value=existing_value,
            status=REVIEW_STATUS_SKIPPED,
            note="TARGET_ALREADY_HAS_VALUE",
            reason=SKIP_REASON_EXISTING_VALUE,
            table_type=table_type,
        )

    try:
        anchor.value = value
    except Exception as exc:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_TABLE_CONTEXT,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            target_cell=anchor.coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note=f"WRITE_ERROR: {str(exc)}",
            reason=SKIP_REASON_EXCEPTION,
            table_type=table_type,
        )

    return _build_log_row(
        item,
        mapped_field_key=field_key,
        match_type=MATCH_TYPE_TABLE_CONTEXT,
        confidence=100,
        semantic_confidence=1.0,
        layout_confidence=layout_confidence,
        total_confidence=total_confidence,
        target_cell=anchor.coordinate,
        target_merged_range=target_merged_range,
        resolver=resolver,
        layout_type=layout_type,
        filled_value=value,
        status=REVIEW_STATUS_FILLED,
        note="FILLED",
        reason="FILLED",
        table_type=table_type,
    )