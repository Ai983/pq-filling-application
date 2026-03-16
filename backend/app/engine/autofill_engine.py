from __future__ import annotations

from typing import Any, Dict, List, Set, Tuple

from app.core.constants import (
    CELL_TYPE_SIMPLE_FIELD,
    CELL_TYPE_SUBFIELD,
    CELL_TYPE_UNKNOWN,
    MATCH_TYPE_TABLE_CONTEXT,
    REVIEW_STATUS_FILLED,
    REVIEW_STATUS_SKIPPED,
    REVIEW_STATUS_UNMATCHED,
    SKIP_REASON_EXCEPTION,
    SKIP_REASON_EXISTING_VALUE,
    SKIP_REASON_FORMULA_CELL,
    SKIP_REASON_NO_TARGET,
    SKIP_REASON_UNSUPPORTED_LAYOUT,
    TABLE_TYPE_FINANCIAL,
)
from app.engine.fillers.project_table_filler import fill_project_tables
from app.engine.fillers.section_block_filler import fill_section_block_field
from app.engine.fillers.simple_field_filler import fill_simple_field
from app.engine.fillers.table_filler import fill_table_row_field
from app.engine.fillers.yes_no_filler import fill_yes_no_field
from app.engine.master_loader import get_master_value
from app.engine.section_resolver import enrich_findings_with_sections
from app.engine.table_detectors import (
    detect_sheet_tables,
    detect_table_field_key,
    is_compliance_row,
    is_machinery_row,
    is_manpower_row,
    is_turnover_row,
)
from app.engine.target_cell_resolver import get_merged_anchor_cell


def _normalize_findings_input(findings: Any) -> List[Dict[str, Any]]:
    """
    Supports both:
    1) old input: findings is already a list
    2) new input: findings is scan_result dict with 'findings'
    """
    if isinstance(findings, list):
        return findings

    if isinstance(findings, dict):
        nested = findings.get("findings")
        if isinstance(nested, list):
            return nested

    return []


def _make_skip_log(item: Dict[str, Any], note: str, reason: str = "") -> Dict[str, Any]:
    return {
        "sheet": item.get("sheet", ""),
        "sheet_name": item.get("sheet", ""),
        "label_cell": item.get("cell", ""),
        "label_text": item.get("value", ""),
        "normalized_label": item.get("normalized_value", ""),
        "cell_type": item.get("cell_type", ""),
        "active_section": item.get("active_section", ""),
        "section": item.get("active_section", ""),
        "mapped_field_key": "",
        "field_key": "",
        "match_type": "",
        "confidence": 0,
        "semantic_confidence": 0,
        "layout_confidence": 0,
        "total_confidence": 0,
        "resolver": "",
        "layout_type": "",
        "table_type": "",
        "target_cell": "",
        "target_merged_range": "",
        "filled_value": "",
        "value_preview": "",
        "status": REVIEW_STATUS_SKIPPED,
        "write_result": REVIEW_STATUS_SKIPPED,
        "note": note,
        "reason": reason or note,
    }


def _make_exception_log(item: Dict[str, Any], exc: Exception, note_prefix: str = "ENGINE_EXCEPTION") -> Dict[str, Any]:
    return {
        "sheet": item.get("sheet", ""),
        "sheet_name": item.get("sheet", ""),
        "label_cell": item.get("cell", ""),
        "label_text": item.get("value", ""),
        "normalized_label": item.get("normalized_value", ""),
        "cell_type": item.get("cell_type", ""),
        "active_section": item.get("active_section", ""),
        "section": item.get("active_section", ""),
        "mapped_field_key": "",
        "field_key": "",
        "match_type": "",
        "confidence": 0,
        "semantic_confidence": 0,
        "layout_confidence": 0,
        "total_confidence": 0,
        "resolver": "",
        "layout_type": "",
        "table_type": "",
        "target_cell": "",
        "target_merged_range": "",
        "filled_value": "",
        "value_preview": "",
        "status": REVIEW_STATUS_SKIPPED,
        "write_result": REVIEW_STATUS_SKIPPED,
        "note": f"{note_prefix}: {str(exc)}",
        "reason": SKIP_REASON_EXCEPTION,
    }


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


def _finding_identity(item: Dict[str, Any]) -> Tuple[str, str]:
    return (
        str(item.get("sheet", "")),
        str(item.get("cell", "")),
    )


def is_likely_yes_no_candidate(item: dict) -> bool:
    """
    Conservative yes/no candidate detection.
    We allow both keyword-based detection and known compliance/statutory row labels.
    """
    normalized = item.get("normalized_value", "") or ""
    normalized = str(normalized).strip().lower()

    if not normalized:
        return False

    if is_compliance_row(normalized):
        return True

    yes_no_keywords = [
        "yes",
        "no",
        "iso",
        "litigation",
        "arbitration",
        "policy",
        "workshop",
        "msme",
        "audit report",
        "audit reports",
        "quality",
        "safety",
        "ohse",
        "ohs",
        "applicable",
        "available",
        "provided",
        "pf",
        "esi",
    ]
    return any(keyword in normalized for keyword in yes_no_keywords)


def _is_table_like_candidate(item: Dict[str, Any]) -> bool:
    normalized = item.get("normalized_value", "")
    if not normalized:
        return False

    return (
        is_turnover_row(normalized)
        or is_manpower_row(normalized)
        or is_machinery_row(normalized)
        or is_compliance_row(normalized)
        or detect_table_field_key(normalized) is not None
    )


def _sort_findings_for_fill(findings: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Prefer deterministic order:
    1. row/column order
    2. table-like rows earlier if same physical position
    """
    def sort_key(item: Dict[str, Any]):
        row = int(item.get("row", 10**9) or 10**9)
        col = int(item.get("column", 10**9) or 10**9)
        normalized = str(item.get("normalized_value", "") or "")
        table_bias = 0 if detect_table_field_key(normalized) else 1
        return (str(item.get("sheet", "")), row, col, table_bias)

    return sorted(findings, key=sort_key)


def _build_horizontal_financial_log(
    sheet_name: str,
    heading_row: int,
    field_key: str,
    target_cell: str = "",
    filled_value: Any = "",
    status: str = REVIEW_STATUS_SKIPPED,
    note: str = "",
    reason: str = "",
):
    return {
        "sheet": sheet_name,
        "sheet_name": sheet_name,
        "label_cell": f"ROW-{heading_row}",
        "label_text": "Annual Turnover",
        "normalized_label": "annual turnover",
        "cell_type": "TABLE_HEADER",
        "active_section": "financial",
        "section": "financial",
        "mapped_field_key": field_key,
        "field_key": field_key,
        "match_type": MATCH_TYPE_TABLE_CONTEXT,
        "confidence": 100,
        "semantic_confidence": 1.0,
        "layout_confidence": 1.0,
        "total_confidence": 1.0,
        "resolver": "horizontal_financial_filler",
        "layout_type": "horizontal_year_headers",
        "table_type": TABLE_TYPE_FINANCIAL,
        "target_cell": target_cell,
        "target_merged_range": "",
        "filled_value": filled_value,
        "value_preview": _safe_preview(filled_value),
        "status": status,
        "write_result": status,
        "note": note,
        "reason": reason or note,
    }


def _fill_horizontal_financial_tables(wb, master_data):
    """
    Fills sections like:
        Annual Turn Over (Cr)
           2023-2024 | 2022-2023 | ...
           <values row below>

    This is separate from row-label table filling.
    """
    log_rows: List[Dict[str, Any]] = []

    for sheet in wb.worksheets:
        sheet_tables = detect_sheet_tables(sheet)
        horizontal_tables = sheet_tables.get("horizontal_financial_tables", []) or []

        for table in horizontal_tables:
            value_row = int(table.get("value_row") or 0)
            heading_row = int(table.get("heading_row") or value_row)
            year_column_map = table.get("year_column_map", {}) or {}

            if not value_row or not year_column_map:
                continue

            for field_key, col_idx in year_column_map.items():
                value = get_master_value(master_data, field_key)

                if value in (None, ""):
                    log_rows.append(
                        _build_horizontal_financial_log(
                            sheet_name=sheet.title,
                            heading_row=heading_row,
                            field_key=field_key,
                            status=REVIEW_STATUS_SKIPPED,
                            note="NO_MASTER_VALUE",
                            reason="NO_MASTER_VALUE",
                        )
                    )
                    continue

                anchor = get_merged_anchor_cell(sheet, sheet.cell(row=value_row, column=col_idx))

                if _is_formula_cell(anchor):
                    log_rows.append(
                        _build_horizontal_financial_log(
                            sheet_name=sheet.title,
                            heading_row=heading_row,
                            field_key=field_key,
                            target_cell=anchor.coordinate,
                            filled_value="",
                            status=REVIEW_STATUS_SKIPPED,
                            note="FORMULA_CELL",
                            reason=SKIP_REASON_FORMULA_CELL,
                        )
                    )
                    continue

                existing_value = getattr(anchor, "value", None)
                if existing_value not in (None, "") and str(existing_value).strip() != "":
                    log_rows.append(
                        _build_horizontal_financial_log(
                            sheet_name=sheet.title,
                            heading_row=heading_row,
                            field_key=field_key,
                            target_cell=anchor.coordinate,
                            filled_value=existing_value,
                            status=REVIEW_STATUS_SKIPPED,
                            note="TARGET_ALREADY_HAS_VALUE",
                            reason=SKIP_REASON_EXISTING_VALUE,
                        )
                    )
                    continue

                try:
                    anchor.value = value
                except Exception as exc:
                    log_rows.append(
                        _build_horizontal_financial_log(
                            sheet_name=sheet.title,
                            heading_row=heading_row,
                            field_key=field_key,
                            target_cell=anchor.coordinate,
                            filled_value="",
                            status=REVIEW_STATUS_SKIPPED,
                            note=f"WRITE_ERROR: {str(exc)}",
                            reason=SKIP_REASON_EXCEPTION,
                        )
                    )
                    continue

                log_rows.append(
                    _build_horizontal_financial_log(
                        sheet_name=sheet.title,
                        heading_row=heading_row,
                        field_key=field_key,
                        target_cell=anchor.coordinate,
                        filled_value=value,
                        status=REVIEW_STATUS_FILLED,
                        note="FILLED",
                        reason="FILLED",
                    )
                )

    return wb, log_rows


def autofill_workbook(wb, findings, master_data, synonyms):
    """
    Main deterministic autofill orchestration.

    Strategy order:
    1. Table-like row labels
    2. Section-aware repeated blocks
    3. Yes/No and compliance/statutory fields
    4. Simple fields
    5. Horizontal financial tables
    6. Project tables after cell-by-cell fill
    """
    log_rows: List[Dict[str, Any]] = []

    normalized_findings = _normalize_findings_input(findings)
    normalized_findings = enrich_findings_with_sections(normalized_findings)
    normalized_findings = _sort_findings_for_fill(normalized_findings)

    processed_cells: Set[Tuple[str, str]] = set()

    for item in normalized_findings:
        try:
            identity = _finding_identity(item)
            if identity in processed_cells:
                continue

            cell_type = item.get("cell_type", CELL_TYPE_UNKNOWN)

            # ---------------------------------------------------------
            # 1. Table-like row labels
            # ---------------------------------------------------------
            if _is_table_like_candidate(item):
                row_result = fill_table_row_field(item, wb, master_data)

                if isinstance(row_result, dict):
                    log_rows.append(row_result)
                    processed_cells.add(identity)
                else:
                    log_rows.append(
                        _make_skip_log(
                            item,
                            note="TABLE_FILLER_RETURNED_INVALID_RESULT",
                            reason=SKIP_REASON_UNSUPPORTED_LAYOUT,
                        )
                    )
                    processed_cells.add(identity)
                continue

            # ---------------------------------------------------------
            # 2. Section-aware repeated blocks
            # ---------------------------------------------------------
            if cell_type == CELL_TYPE_SUBFIELD:
                row_result = fill_section_block_field(item, wb, master_data)

                if isinstance(row_result, dict):
                    log_rows.append(row_result)
                    processed_cells.add(identity)
                else:
                    log_rows.append(
                        _make_skip_log(
                            item,
                            note="SECTION_BLOCK_FILLER_RETURNED_INVALID_RESULT",
                            reason=SKIP_REASON_UNSUPPORTED_LAYOUT,
                        )
                    )
                    processed_cells.add(identity)
                continue

            # ---------------------------------------------------------
            # 3. Yes/No sections
            # ---------------------------------------------------------
            if is_likely_yes_no_candidate(item):
                yes_no_row = fill_yes_no_field(item, wb, master_data)

                if isinstance(yes_no_row, dict):
                    if yes_no_row.get("status") != REVIEW_STATUS_UNMATCHED:
                        log_rows.append(yes_no_row)
                        processed_cells.add(identity)
                        continue
                else:
                    log_rows.append(
                        _make_skip_log(
                            item,
                            note="YES_NO_FILLER_RETURNED_INVALID_RESULT",
                            reason=SKIP_REASON_UNSUPPORTED_LAYOUT,
                        )
                    )
                    processed_cells.add(identity)
                    continue

            # ---------------------------------------------------------
            # 4. Simple fields
            # ---------------------------------------------------------
            if cell_type in {CELL_TYPE_SIMPLE_FIELD, CELL_TYPE_UNKNOWN, ""}:
                simple_row = fill_simple_field(item, wb, master_data, synonyms)

                if isinstance(simple_row, dict):
                    log_rows.append(simple_row)
                    processed_cells.add(identity)
                else:
                    log_rows.append(
                        _make_skip_log(
                            item,
                            note="SIMPLE_FILLER_RETURNED_INVALID_RESULT",
                            reason=SKIP_REASON_UNSUPPORTED_LAYOUT,
                        )
                    )
                    processed_cells.add(identity)
                continue

            # ---------------------------------------------------------
            # 5. Later/unsupported types
            # ---------------------------------------------------------
            log_rows.append(
                _make_skip_log(
                    item,
                    note="UNSUPPORTED_CELL_TYPE",
                    reason=SKIP_REASON_UNSUPPORTED_LAYOUT,
                )
            )
            processed_cells.add(identity)

        except Exception as exc:
            log_rows.append(_make_exception_log(item, exc))
            processed_cells.add(_finding_identity(item))

    # -------------------------------------------------------------
    # 6. Horizontal financial fill
    # -------------------------------------------------------------
    try:
        wb, financial_log_rows = _fill_horizontal_financial_tables(wb, master_data)
        if isinstance(financial_log_rows, list):
            log_rows.extend(financial_log_rows)
    except Exception as exc:
        log_rows.append(
            {
                "sheet": "",
                "sheet_name": "",
                "label_cell": "",
                "label_text": "",
                "normalized_label": "",
                "cell_type": "",
                "active_section": "",
                "section": "",
                "mapped_field_key": "",
                "field_key": "",
                "match_type": "",
                "confidence": 0,
                "semantic_confidence": 0,
                "layout_confidence": 0,
                "total_confidence": 0,
                "resolver": "",
                "layout_type": "",
                "table_type": TABLE_TYPE_FINANCIAL,
                "target_cell": "",
                "target_merged_range": "",
                "filled_value": "",
                "value_preview": "",
                "status": REVIEW_STATUS_SKIPPED,
                "write_result": REVIEW_STATUS_SKIPPED,
                "note": f"HORIZONTAL_FINANCIAL_FILL_EXCEPTION: {str(exc)}",
                "reason": SKIP_REASON_EXCEPTION,
            }
        )

    # -------------------------------------------------------------
    # 7. Project table fill happens after cell-by-cell fills
    # -------------------------------------------------------------
    try:
        wb, project_log_rows = fill_project_tables(wb, master_data)
        if isinstance(project_log_rows, list):
            log_rows.extend(project_log_rows)
    except Exception as exc:
        log_rows.append(
            {
                "sheet": "",
                "sheet_name": "",
                "label_cell": "",
                "label_text": "",
                "normalized_label": "",
                "cell_type": "",
                "active_section": "",
                "section": "",
                "mapped_field_key": "",
                "field_key": "",
                "match_type": "",
                "confidence": 0,
                "semantic_confidence": 0,
                "layout_confidence": 0,
                "total_confidence": 0,
                "resolver": "",
                "layout_type": "",
                "table_type": "",
                "target_cell": "",
                "target_merged_range": "",
                "filled_value": "",
                "value_preview": "",
                "status": REVIEW_STATUS_SKIPPED,
                "write_result": REVIEW_STATUS_SKIPPED,
                "note": f"PROJECT_TABLE_FILL_EXCEPTION: {str(exc)}",
                "reason": SKIP_REASON_EXCEPTION,
            }
        )

    return wb, log_rows