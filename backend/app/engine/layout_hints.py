from __future__ import annotations

from typing import Any, Dict, List

from app.engine.utils import normalize_text


def get_workbook_fingerprint(wb) -> str:
    """
    Lightweight deterministic workbook fingerprint.

    Used for:
    - template hinting
    - future template profile lookup
    - debugging recurring vendor formats
    """
    parts: List[str] = []

    for sheet in wb.worksheets:
        parts.append(sheet.title.strip().lower())
        parts.append(f"{sheet.max_row}x{sheet.max_column}")
        parts.append(f"merged:{len(sheet.merged_cells.ranges)}")

        # include first few visible text anchors for stronger identity
        visible_texts = []
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                text = str(cell.value).strip()
                if not text:
                    continue
                visible_texts.append(normalize_text(text))
                if len(visible_texts) >= 10:
                    break
            if len(visible_texts) >= 10:
                break

        parts.extend(visible_texts)

    return "|".join(parts)


def detect_nearby_context_text(sheet, row: int, lookback: int = 5, lookahead: int = 0) -> str:
    """
    Collect surrounding visible text near a row to infer section/table intent.
    """
    texts: List[str] = []

    start_row = max(1, row - lookback)
    end_row = min(sheet.max_row, row + lookahead)

    for r in range(start_row, end_row + 1):
        for c in range(1, min(sheet.max_column, 10) + 1):
            value = sheet.cell(r, c).value
            if value is not None:
                text = str(value).strip()
                if text:
                    texts.append(text)

    return normalize_text(" ".join(texts))


def detect_project_table_mode(sheet, header_row: int) -> str:
    """
    Decide whether the table should prefer ongoing/completed/all projects.
    """
    context = detect_nearby_context_text(sheet, header_row, lookback=6, lookahead=1)

    if "ongoing" in context or "running projects" in context or "in progress" in context:
        return "ongoing"

    if "completed" in context or "completed projects" in context or "finished projects" in context:
        return "completed"

    return "all"


def detect_layout_family(sheet) -> str:
    """
    Very light layout family classifier for debugging / future profile matching.
    """
    merged_count = len(sheet.merged_cells.ranges)

    if merged_count >= 20:
        return "merged_heavy_form"

    if sheet.max_column >= 10 and sheet.max_row >= 20:
        return "tabular_form"

    if sheet.max_column <= 6:
        return "compact_form"

    return "mixed_form"


def detect_section_bias(sheet, row: int) -> str:
    """
    Infer dominant nearby section intent using nearby text only.
    """
    context = detect_nearby_context_text(sheet, row, lookback=6, lookahead=1)

    checks = [
        ("projects", ["project", "client", "location", "completed", "ongoing", "work order"]),
        ("financial", ["turnover", "financial", "audited", "balance sheet"]),
        ("banking", ["bank", "account", "ifsc", "branch"]),
        ("tax", ["gst", "pan", "pf", "esi", "msme", "registration"]),
        ("compliance", ["iso", "policy", "litigation", "arbitration", "audit"]),
        ("resource", ["engineers", "supervisors", "manpower", "machinery", "workshop"]),
        ("contacts", ["contact", "mobile", "email", "designation"]),
    ]

    for section_name, keywords in checks:
        if any(keyword in context for keyword in keywords):
            return section_name

    return "unknown"