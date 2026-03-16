from app.core.constants import (
    CELL_TYPE_DECLARATION,
    CELL_TYPE_INSTRUCTION,
    CELL_TYPE_SECTION_HEADER,
    CELL_TYPE_SIMPLE_FIELD,
    CELL_TYPE_SUBFIELD,
    CELL_TYPE_TABLE_HEADER,
    CELL_TYPE_UNKNOWN,
)
from app.engine.utils import (
    is_likely_declaration,
    is_likely_instruction,
    is_likely_section_header,
    is_likely_simple_field,
    is_likely_subfield,
    is_likely_table_header,
    normalize_text,
)


def classify_cell_text(text: str) -> str:
    """
    Conservative deterministic classifier for visible Excel text.

    Priority matters:
    - instruction before section/simple
    - declaration before simple
    - section before subfield
    - subfield before table/simple
    """
    normalized = normalize_text(text)

    if not normalized:
        return CELL_TYPE_UNKNOWN

    if is_likely_instruction(normalized):
        return CELL_TYPE_INSTRUCTION

    if is_likely_declaration(normalized):
        return CELL_TYPE_DECLARATION

    if is_likely_section_header(normalized):
        return CELL_TYPE_SECTION_HEADER

    if is_likely_subfield(normalized):
        return CELL_TYPE_SUBFIELD

    if is_likely_table_header(normalized):
        return CELL_TYPE_TABLE_HEADER

    if is_likely_simple_field(normalized):
        return CELL_TYPE_SIMPLE_FIELD

    return CELL_TYPE_UNKNOWN