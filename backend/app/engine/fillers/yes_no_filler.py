from __future__ import annotations

from typing import Any, Optional

from app.core.config import (
    MIN_LAYOUT_CONFIDENCE,
    MIN_TOTAL_CONFIDENCE_FOR_YES_NO_FILL,
)
from app.core.constants import (
    MATCH_TYPE_YES_NO_RULE,
    REVIEW_STATUS_FILLED,
    REVIEW_STATUS_SKIPPED,
    REVIEW_STATUS_UNMATCHED,
    SKIP_REASON_EXCEPTION,
    SKIP_REASON_EXISTING_VALUE,
    SKIP_REASON_FORMULA_CELL,
    SKIP_REASON_LOW_CONFIDENCE,
    SKIP_REASON_NO_MATCH,
    SKIP_REASON_NO_TARGET,
    YES_NO_REPRESENTATIONS,
)
from app.engine.skip_rules import should_skip_no_master_value
from app.engine.target_cell_resolver import (
    get_merged_anchor_cell,
    resolve_target_cell,
)


YES_NO_FIELD_MAP = {
    "iso": "compliance.iso_certified",
    "iso certified": "compliance.iso_certified",
    "quality policy": "compliance.quality_policy",
    "safety policy": "compliance.safety_policy",
    "ohse": "compliance.ohs_policy",
    "ohs": "compliance.ohs_policy",
    "ohse policy": "compliance.ohs_policy",
    "ohs policy": "compliance.ohs_policy",
    "litigation": "compliance.litigation",
    "arbitration": "compliance.arbitration",
    "msme": "tax.msme",
    "audit report": "compliance.audit_reports",
    "audit reports": "compliance.audit_reports",
    "workshop": "resource.workshop_available",
    "workshop facilities": "resource.workshop_available",
    "pf": "tax.pf",
    "esi": "tax.esi",
}

YES_NO_REPRESENTATION_HINTS = {
    "compliance.iso_certified": "yes_no_words",
    "compliance.quality_policy": "yes_no_words",
    "compliance.safety_policy": "yes_no_words",
    "compliance.ohs_policy": "yes_no_words",
    "compliance.litigation": "yes_no_words",
    "compliance.arbitration": "yes_no_words",
    "compliance.audit_reports": "yes_no_words",
    "resource.workshop_available": "yes_no_words",
    "tax.msme": "yes_no_words",
    "tax.pf": "yes_no_words",
    "tax.esi": "yes_no_words",
}


def _safe_preview(value: Any, max_len: int = 120) -> str:
    if value is None:
        return ""
    text = str(value)
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def _normalize_bool_like_value(value: Any) -> Optional[str]:
    """
    Convert master value into canonical 'yes' or 'no' when possible.
    """
    if value is None:
        return None

    text = str(value).strip().lower()
    if not text:
        return None

    yes_tokens = {"yes", "y", "true", "1", "applicable", "available", "provided"}
    no_tokens = {"no", "n", "false", "0", "not applicable", "not available", "not provided"}

    if text in yes_tokens:
        return "yes"
    if text in no_tokens:
        return "no"

    if "yes" == text or text.startswith("yes"):
        return "yes"
    if "no" == text or text.startswith("no"):
        return "no"

    return None


def _map_representation(field_key: str, raw_value: Any) -> Any:
    """
    Convert canonical yes/no into template-friendly representation.
    For now, default to yes_no_words unless field-specific rule says otherwise.
    """
    canonical = _normalize_bool_like_value(raw_value)
    if canonical is None:
        return raw_value

    representation_key = YES_NO_REPRESENTATION_HINTS.get(field_key, "yes_no_words")
    representation = YES_NO_REPRESENTATIONS.get(representation_key, YES_NO_REPRESENTATIONS["yes_no_words"])

    if canonical == "yes":
        return representation["yes"]
    return representation["no"]


def _is_formula_cell(cell) -> bool:
    value = getattr(cell, "value", None)
    return isinstance(value, str) and value.startswith("=")


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


def resolve_yes_no_field_key(normalized_label: str) -> str | None:
    normalized_label = str(normalized_label or "").strip().lower()
    if not normalized_label:
        return None

    for keyword, field_key in YES_NO_FIELD_MAP.items():
        if keyword in normalized_label:
            return field_key
    return None


def fill_yes_no_field(item, wb, master_data):
    """
    Deterministic filler for compliance/statutory yes-no style fields.
    """
    sheet = wb[item["sheet"]]
    normalized_label = item.get("normalized_value", "")

    field_key = resolve_yes_no_field_key(normalized_label)
    if not field_key:
        return _build_log_row(
            item,
            status=REVIEW_STATUS_UNMATCHED,
            note="UNMATCHED",
            reason=SKIP_REASON_NO_MATCH,
        )

    raw_value = master_data.get(field_key)

    skip, reason = should_skip_no_master_value(raw_value)
    if skip:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_YES_NO_RULE,
            confidence=100,
            semantic_confidence=1.0,
            status=REVIEW_STATUS_SKIPPED,
            note=reason,
            reason=reason,
        )

    value = _map_representation(field_key, raw_value)

    resolution = resolve_target_cell(
        sheet=sheet,
        label_row=item["row"],
        label_col=item["column"],
        label_text=item.get("value", ""),
        cell_type=item.get("cell_type", ""),
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
            match_type=MATCH_TYPE_YES_NO_RULE,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="NO_SAFE_TARGET",
            reason=resolution_reason or SKIP_REASON_NO_TARGET,
        )

    if layout_confidence < MIN_LAYOUT_CONFIDENCE or target_score < int(MIN_TOTAL_CONFIDENCE_FOR_YES_NO_FILL * 100) - 5:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_YES_NO_RULE,
            confidence=100,
            semantic_confidence=1.0,
            layout_confidence=layout_confidence,
            target_cell=target_coordinate,
            target_merged_range=target_merged_range,
            resolver=resolver,
            layout_type=layout_type,
            status=REVIEW_STATUS_SKIPPED,
            note="LOW_LAYOUT_CONFIDENCE",
            reason=SKIP_REASON_LOW_CONFIDENCE,
        )

    anchor = get_merged_anchor_cell(sheet, target_cell)

    if _is_formula_cell(anchor):
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_YES_NO_RULE,
            confidence=100,
            semantic_confidence=1.0,
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
            match_type=MATCH_TYPE_YES_NO_RULE,
            confidence=100,
            semantic_confidence=1.0,
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

    try:
        anchor.value = value
    except Exception as exc:
        return _build_log_row(
            item,
            mapped_field_key=field_key,
            match_type=MATCH_TYPE_YES_NO_RULE,
            confidence=100,
            semantic_confidence=1.0,
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
        match_type=MATCH_TYPE_YES_NO_RULE,
        confidence=100,
        semantic_confidence=1.0,
        layout_confidence=layout_confidence,
        target_cell=anchor.coordinate,
        target_merged_range=target_merged_range,
        resolver=resolver,
        layout_type=layout_type,
        filled_value=value,
        status=REVIEW_STATUS_FILLED,
        note="FILLED",
        reason="FILLED",
    )