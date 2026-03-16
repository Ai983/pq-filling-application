from __future__ import annotations

from typing import Any, Dict

from app.core.config import (
    MIN_LAYOUT_CONFIDENCE,
    MIN_TOTAL_CONFIDENCE_TO_FILL,
)
from app.core.constants import (
    CELL_TYPE_SUBFIELD,
    MATCH_TYPE_SECTION_CONTEXT,
    REVIEW_STATUS_FILLED,
    REVIEW_STATUS_SKIPPED,
    REVIEW_STATUS_UNMATCHED,
    SECTION_UNKNOWN,
    SKIP_REASON_EXCEPTION,
    SKIP_REASON_EXISTING_VALUE,
    SKIP_REASON_FORMULA_CELL,
    SKIP_REASON_LOW_CONFIDENCE,
    SKIP_REASON_NO_MATCH,
    SKIP_REASON_NO_TARGET,
    SKIP_REASON_UNSUPPORTED_LAYOUT,
)
from app.engine.section_resolver import resolve_contextual_field_key_with_fallback
from app.engine.skip_rules import (
    should_skip_by_cell_type,
    should_skip_low_confidence,
    should_skip_no_master_value,
)
from app.engine.target_cell_resolver import (
    get_merged_anchor_cell,
    resolve_target_cell,
)


def _safe_preview(value: Any, max_len: int = 120) -> str:
    if value is None:
        return ""
    text = str(value)
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def _is_formula_cell(cell) -> bool:
    value = getattr(cell, "value", None)
    return isinstance(value, str) and value.startswith("=")


def _extract_master_value(master_data: Dict[str, Any], field_key: str):
    if not isinstance(master_data, dict) or not field_key:
        return None
    return master_data.get(field_key)


def _build_log_row(
    item: dict,
    mapped_field_key: str = "",
    match_type: str = "",
    confidence: int = 0,
    semantic_confidence: float = 0.0,
    layout_confidence: float = 0.0,
    target_cell: str = "",
    target_merged_range: str = "",
    filled_value: Any = "",
    status: str = REVIEW_STATUS_SKIPPED,
    note: str = "",
    resolver: str = "",
    layout_type: str = "",
    reason: str = "",
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
        "resolver": resolver,
        "layout_type": layout_type,
        "target_cell": target_cell,
        "target_merged_range": target_merged_range,
        "filled_value": filled_value,
        "value_preview": _safe_preview(filled_value),
        "status": status,
        "write_result": status,
        "note": note,
        "reason": reason or note,
    }


def fill_section_block_field(item, wb, master_data):
    """
    Fills repeated contextual subfields such as:
    - Name
    - Designation
    - Mobile
    - Email

    based on active section such as:
    - owner
    - project_contact
    - accounts_contact
    """
    sheet = wb[item["sheet"]]
    normalized_label = item.get("normalized_value", "")
    cell_type = item.get("cell_type", "")
    active_section = item.get("active_section", SECTION_UNKNOWN)

    # ---------------------------------------------------------
    # 1. skip checks
    # ---------------------------------------------------------
    skip, reason = should_skip_by_cell_type(cell_type)
    if skip:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
        )

    if cell_type != CELL_TYPE_SUBFIELD:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_SKIPPED,
            note="UNSUPPORTED_CELL_TYPE",
            reason=SKIP_REASON_UNSUPPORTED_LAYOUT,
        )

    if active_section == SECTION_UNKNOWN:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_SKIPPED,
            note="SECTION_REQUIRED",
            reason=SKIP_REASON_NO_MATCH,
        )

    # ---------------------------------------------------------
    # 2. contextual field resolution
    # ---------------------------------------------------------
    field_key, resolution_mode = resolve_contextual_field_key_with_fallback(
        active_section,
        normalized_label,
    )

    if not field_key:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_UNMATCHED,
            note="UNMATCHED",
            reason=SKIP_REASON_NO_MATCH,
        )

    # deterministic contextual mapping is high-confidence by design
    confidence = 100
    semantic_confidence = 1.0
    match_type = MATCH_TYPE_SECTION_CONTEXT

    skip, reason = should_skip_low_confidence(confidence, threshold=int(MIN_TOTAL_CONFIDENCE_TO_FILL * 100))
    if skip:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=SKIP_REASON_LOW_CONFIDENCE,
        )

    # ---------------------------------------------------------
    # 3. master value lookup
    # ---------------------------------------------------------
    value = _extract_master_value(master_data, field_key)

    skip, reason = should_skip_no_master_value(value)
    if skip:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
        )

    # ---------------------------------------------------------
    # 4. target resolution
    # ---------------------------------------------------------
    resolution = resolve_target_cell(
        sheet=sheet,
        label_row=item["row"],
        label_col=item["column"],
        label_text=item.get("value", ""),
        cell_type=cell_type,
    )

    target_cell = resolution.get("target_cell")
    target_coordinate = resolution.get("target_coordinate", "")
    target_merged_range = resolution.get("target_merged_range", "") or ""
    layout_confidence = float(resolution.get("layout_confidence", 0.0))
    resolver = resolution.get("resolver", "")
    layout_type = resolution.get("layout_type", "")
    target_score = int(resolution.get("score", 0))
    resolution_reason = resolution.get("reason", "")

    if target_cell is None:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            layout_confidence=layout_confidence,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="NO_SAFE_TARGET",
            reason=resolution_reason or SKIP_REASON_NO_TARGET,
        )

    if layout_confidence < MIN_LAYOUT_CONFIDENCE or target_score < int(MIN_LAYOUT_CONFIDENCE * 100):
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            layout_confidence=layout_confidence,
            target_cell=target_coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="LOW_LAYOUT_CONFIDENCE",
            reason=SKIP_REASON_LOW_CONFIDENCE,
        )

    # ---------------------------------------------------------
    # 5. safe anchor checks
    # ---------------------------------------------------------
    anchor = get_merged_anchor_cell(sheet, target_cell)

    if _is_formula_cell(anchor):
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            layout_confidence=layout_confidence,
            target_cell=anchor.coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="FORMULA_CELL",
            reason=SKIP_REASON_FORMULA_CELL,
        )

    existing_value = getattr(anchor, "value", None)
    if existing_value not in (None, "") and str(existing_value).strip() != "":
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            layout_confidence=layout_confidence,
            target_cell=anchor.coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            filled_value=existing_value,
            status=REVIEW_STATUS_SKIPPED,
            note="TARGET_ALREADY_HAS_VALUE",
            reason=SKIP_REASON_EXISTING_VALUE,
        )

    # ---------------------------------------------------------
    # 6. write
    # ---------------------------------------------------------
    try:
        anchor.value = value
    except Exception as exc:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            layout_confidence=layout_confidence,
            target_cell=anchor.coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note=f"WRITE_ERROR: {str(exc)}",
            reason=SKIP_REASON_EXCEPTION,
        )

    return _build_log_row(
        item,
        mapped_field_key=field_key,
        match_type=match_type,
        confidence=confidence,
        semantic_confidence=semantic_confidence,
        layout_confidence=layout_confidence,
        target_cell=anchor.coordinate,
        target_merged_range=target_merged_range,
        resolver=resolver,
        layout_type=layout_type,
        filled_value=value,
        status=REVIEW_STATUS_FILLED,
        note=f"FILLED_{resolution_mode}",
        reason="FILLED",
    )