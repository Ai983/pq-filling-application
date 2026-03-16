from __future__ import annotations

from typing import Any, Tuple

from app.core.constants import (
    CELL_TYPE_DECLARATION,
    CELL_TYPE_HEADER_BAND,
    CELL_TYPE_INSTRUCTION,
    CELL_TYPE_SECTION_HEADER,
    CELL_TYPE_SUBFIELD,
    CELL_TYPE_TABLE_HEADER,
    PLACEHOLDER_VALUES,
    SECTION_UNKNOWN,
)
from app.engine.utils import normalize_text


GENERIC_SUBFIELDS = {
    "name",
    "designation",
    "mobile",
    "mobile no",
    "mobile number",
    "contact no",
    "contact number",
    "phone",
    "telephone",
    "tel",
    "email",
    "email id",
    "e mail",
    "mail id",
    "address",
    "location",
}


def should_skip_by_cell_type(cell_type: str) -> tuple[bool, str | None]:
    """
    Skip cell types that should never be handled by the generic simple filler.
    """
    if cell_type == CELL_TYPE_SECTION_HEADER:
        return True, "UNSUPPORTED_CELL_TYPE"

    if cell_type == CELL_TYPE_TABLE_HEADER:
        return True, "UNSUPPORTED_CELL_TYPE"

    if cell_type == CELL_TYPE_INSTRUCTION:
        return True, "UNSUPPORTED_CELL_TYPE"

    if cell_type == CELL_TYPE_HEADER_BAND:
        return True, "UNSUPPORTED_CELL_TYPE"

    if cell_type == CELL_TYPE_DECLARATION:
        return True, "UNSUPPORTED_CELL_TYPE"

    return False, None


def should_skip_generic_subfield(normalized_label: str, active_section: str) -> tuple[bool, str | None]:
    """
    Generic labels like 'name' or 'email' are too risky unless section context exists.
    """
    normalized_label = normalize_text(normalized_label)

    if normalized_label in GENERIC_SUBFIELDS and active_section == SECTION_UNKNOWN:
        return True, "GENERIC_SUBFIELD_SKIPPED"

    return False, None


def should_skip_subfield_in_current_batch(cell_type: str) -> tuple[bool, str | None]:
    """
    Section subfields are no longer skipped globally.
    They are now handled by section_block_filler.
    """
    return False, None


def should_skip_low_confidence(confidence: int, threshold: int = 90) -> tuple[bool, str | None]:
    if confidence < threshold:
        return True, "LOW_LABEL_CONFIDENCE"

    return False, None


def should_skip_no_master_value(value: Any) -> tuple[bool, str | None]:
    if value is None:
        return True, "NO_MASTER_VALUE"

    if isinstance(value, str):
        cleaned = value.strip()
        if cleaned == "":
            return True, "NO_MASTER_VALUE"
        if cleaned.lower() in PLACEHOLDER_VALUES:
            return True, "NO_MASTER_VALUE"

    return False, None


def should_skip_existing_value(value: Any) -> tuple[bool, str | None]:
    if value is None:
        return False, None

    text = str(value).strip()
    if text == "":
        return False, None

    if text.lower() in PLACEHOLDER_VALUES:
        return False, None

    return True, "TARGET_ALREADY_HAS_VALUE"


def should_skip_formula_cell(cell) -> tuple[bool, str | None]:
    value = getattr(cell, "value", None)
    if isinstance(value, str) and value.startswith("="):
        return True, "FORMULA_CELL"

    return False, None


def should_skip_empty_label(normalized_label: str) -> tuple[bool, str | None]:
    if not normalize_text(normalized_label):
        return True, "EMPTY_LABEL"

    return False, None


def should_skip_declaration_like_text(normalized_label: str, max_len: int = 140) -> tuple[bool, str | None]:
    """
    Long declaration / paragraph style labels should not be treated
    as normal simple fields in generic mode.
    """
    normalized_label = normalize_text(normalized_label)
    if not normalized_label:
        return True, "EMPTY_LABEL"

    declaration_markers = {
        "hereby declare",
        "we hereby declare",
        "certify that",
        "undertake that",
        "authorized signatory",
        "authorised signatory",
    }

    if len(normalized_label) > max_len:
        return True, "DECLARATION_LIKE_TEXT"

    for marker in declaration_markers:
        if marker in normalized_label:
            return True, "DECLARATION_LIKE_TEXT"

    return False, None