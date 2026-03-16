from __future__ import annotations

from typing import Any, Dict

from app.core.config import (
    MIN_LAYOUT_CONFIDENCE,
    MIN_TOTAL_CONFIDENCE_TO_FILL,
)
from app.core.constants import (
    CELL_TYPE_UNKNOWN,
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
from app.engine.label_matcher import match_label_detailed
from app.engine.master_loader import get_master_value, get_master_value_variants
from app.engine.skip_rules import (
    should_skip_by_cell_type,
    should_skip_generic_subfield,
    should_skip_low_confidence,
    should_skip_no_master_value,
    should_skip_subfield_in_current_batch,
)
from app.engine.target_cell_resolver import (
    get_merged_anchor_cell,
    resolve_target_cell,
)
from app.engine.utils import is_blank, normalize_text


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


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _normalize_yes_no(value: Any) -> Any:
    if value is None:
        return None

    normalized = normalize_text(str(value))
    if normalized in {"yes", "y", "true", "available", "applicable", "provided"}:
        return "Yes"
    if normalized in {"no", "n", "false", "not available", "not applicable", "nil"}:
        return "No"
    return value


def _format_scalar_value(field_key: str, value: Any, value_variants: Dict[str, Any] | None = None) -> Any:
    """
    Conservative deterministic formatting.
    Avoid aggressive transformation.
    """
    if value is None:
        return None

    if value_variants is None:
        value_variants = {}

    # explicit preferred display format support if provided in master variants
    preferred = value_variants.get("preferred_display")
    if preferred not in (None, ""):
        return preferred

    if field_key.startswith("compliance.") or field_key in {
        "tax.msme",
        "tax.pf",
        "tax.esi",
    }:
        return _normalize_yes_no(value)

    return value


def _combined_confidence(semantic_confidence: float, layout_confidence: float) -> float:
    """
    Weighted confidence:
    semantic accuracy matters more than layout.
    """
    return round((semantic_confidence * 0.68) + (layout_confidence * 0.32), 4)


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
        "target_cell": target_cell,
        "target_merged_range": target_merged_range,
        "filled_value": filled_value,
        "value_preview": _safe_preview(filled_value),
        "status": status,
        "write_result": status,
        "note": note,
        "reason": reason or note,
    }


def fill_simple_field(item, wb, master_data, synonyms):
    """
    Deterministic filler for simple one-value fields.

    Flow:
    1. skip-rule checks
    2. semantic label matching
    3. master value lookup with alias fallback
    4. layout-aware target resolution
    5. safe write to resolved anchor cell
    """
    sheet = wb[item["sheet"]]
    label_text = item.get("value", "")
    normalized_label = item.get("normalized_value", "")
    cell_type = item.get("cell_type", CELL_TYPE_UNKNOWN)
    active_section = item.get("active_section", "")

    # ---------------------------------------------------------
    # 1. pre-skip checks
    # ---------------------------------------------------------
    skip, reason = should_skip_by_cell_type(cell_type)
    if skip:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
        )

    skip, reason = should_skip_generic_subfield(normalized_label, active_section)
    if skip:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
        )

    skip, reason = should_skip_subfield_in_current_batch(cell_type)
    if skip:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
        )

    # ---------------------------------------------------------
    # 2. semantic matching
    # ---------------------------------------------------------
    match_result = match_label_detailed(
        label_text=label_text,
        synonyms=synonyms,
        active_section=active_section,
    )

    field_key = match_result.get("field_key")
    match_type = match_result.get("match_type", "")
    semantic_confidence = float(match_result.get("semantic_confidence", 0.0))
    confidence = int(match_result.get("score", 0))

    if not field_key:
        return _build_log_row(
            item,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            total_confidence=0.0,
            status=REVIEW_STATUS_UNMATCHED,
            note="UNMATCHED",
            reason=SKIP_REASON_NO_MATCH,
        )

    # semantic threshold first
    skip, reason = should_skip_low_confidence(
        confidence,
        threshold=int(MIN_TOTAL_CONFIDENCE_TO_FILL * 100),
    )
    if skip:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            total_confidence=semantic_confidence,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=SKIP_REASON_LOW_CONFIDENCE,
        )

    # ---------------------------------------------------------
    # 3. master data lookup with alias fallback
    # ---------------------------------------------------------
    value = get_master_value(master_data, field_key)
    value_variants = get_master_value_variants(master_data, field_key)

    skip, reason = should_skip_no_master_value(value)
    if skip:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            total_confidence=semantic_confidence,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
        )

    value = _format_scalar_value(field_key, value, value_variants=value_variants)

    # ---------------------------------------------------------
    # 4. target resolution
    # ---------------------------------------------------------
    resolution = resolve_target_cell(
        sheet=sheet,
        label_row=item["row"],
        label_col=item["column"],
        label_text=label_text,
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

    total_confidence = _combined_confidence(semantic_confidence, layout_confidence)

    if target_cell is None:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
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
            total_confidence=total_confidence,
            target_cell=target_coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="LOW_LAYOUT_CONFIDENCE",
            reason=SKIP_REASON_LOW_CONFIDENCE,
        )

    if total_confidence < MIN_TOTAL_CONFIDENCE_TO_FILL:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
            layout_confidence=layout_confidence,
            total_confidence=total_confidence,
            target_cell=target_coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="LOW_TOTAL_CONFIDENCE",
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
            total_confidence=total_confidence,
            target_cell=anchor.coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="FORMULA_CELL",
            reason=SKIP_REASON_FORMULA_CELL,
        )

    existing_value = getattr(anchor, "value", None)
    if not is_blank(existing_value):
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=match_type,
            confidence=confidence,
            semantic_confidence=semantic_confidence,
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
            total_confidence=total_confidence,
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
        total_confidence=total_confidence,
        target_cell=anchor.coordinate,
        target_merged_range=target_merged_range,
        resolver=resolver,
        layout_type=layout_type,
        filled_value=value,
        status=REVIEW_STATUS_FILLED,
        note="FILLED",
        reason="FILLED",
    )