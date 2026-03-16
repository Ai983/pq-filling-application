from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple

from app.core.constants import (
    SECTION_ACCOUNTS_CONTACT,
    SECTION_BANKING,
    SECTION_BILLING,
    SECTION_CERTIFICATION,
    SECTION_COMPLIANCE,
    SECTION_COMPANY,
    SECTION_DECLARATION,
    SECTION_FINANCIAL,
    SECTION_KEYWORDS,
    SECTION_OWNER,
    SECTION_PROJECT_CONTACT,
    SECTION_PROJECTS,
    SECTION_RESOURCE,
    SECTION_TAX,
    SECTION_TECHNICAL,
    SECTION_UNKNOWN,
    SECTION_WORKSHOP,
)
from app.engine.utils import normalize_text


# ---------------------------------------------------------------------
# Deterministic section hints
# ---------------------------------------------------------------------

SECTION_HINT_MAP: Dict[str, List[str]] = {
    SECTION_OWNER: [
        "key position holder",
        "principal",
        "owner",
        "director",
        "managing director",
        "authorized signatory",
        "authorised signatory",
        "proprietor",
        "partner",
        "contact detail of key position holder",
        "contact details of key position holder",
        "contact person",
        "promoter",
        "representative",
    ],
    SECTION_PROJECT_CONTACT: [
        "project location",
        "site contact",
        "site office",
        "project contact",
        "contact person at project location",
        "key position holder in project location",
        "execution contact",
        "project coordinator",
        "site engineer",
    ],
    SECTION_ACCOUNTS_CONTACT: [
        "accounts",
        "finance contact",
        "billing contact",
        "commercial contact",
        "accounts contact",
        "payment contact",
        "invoice contact",
    ],
    SECTION_BANKING: [
        "bank details",
        "banking",
        "account details",
        "bank particulars",
        "bank account",
    ],
    SECTION_WORKSHOP: [
        "workshop",
        "factory",
        "warehouse",
        "facilities",
    ],
    SECTION_RESOURCE: [
        "resource",
        "resources",
        "manpower",
        "machinery",
        "plant & machinery",
        "equipment",
        "staff strength",
        "infrastructure",
        "workshop available",
    ],
    SECTION_COMPLIANCE: [
        "compliance",
        "quality",
        "safety",
        "ohse",
        "ohs",
        "iso",
        "litigation",
        "arbitration",
        "policy",
        "audit",
        "legal",
    ],
    SECTION_CERTIFICATION: [
        "certification",
        "certifications",
        "certificate",
        "iso certificate",
    ],
    SECTION_FINANCIAL: [
        "turnover",
        "financial",
        "annual turnover",
        "audited",
        "balance sheet",
        "profit and loss",
        "financial details",
    ],
    SECTION_PROJECTS: [
        "project details",
        "ongoing projects",
        "completed projects",
        "similar works",
        "client details",
        "past experience",
        "work orders",
        "credentials",
        "project credentials",
    ],
    SECTION_COMPANY: [
        "general information",
        "company information",
        "registration details",
        "organisation details",
        "organization details",
        "firm details",
        "company details",
        "vendor details",
        "general details",
    ],
    SECTION_TAX: [
        "tax",
        "gst",
        "pan",
        "pf",
        "esi",
        "statutory",
        "registration",
        "statutory details",
        "taxation",
    ],
    SECTION_DECLARATION: [
        "declaration",
        "undertaking",
        "certify",
        "authorized signatory",
        "authorised signatory",
        "seal and signature",
        "signature with stamp",
    ],
    SECTION_TECHNICAL: [
        "technical",
        "engineering",
        "site execution",
        "project team",
        "technical contact",
    ],
    SECTION_BILLING: [
        "billing",
        "invoice",
        "payment terms",
        "commercial terms",
    ],
}

# Merge in shared keyword hints from constants, without losing local priority
for section_name, keywords in SECTION_KEYWORDS.items():
    existing = SECTION_HINT_MAP.get(section_name, [])
    for keyword in keywords:
        if keyword not in existing:
            existing.append(keyword)
    SECTION_HINT_MAP[section_name] = existing


# ---------------------------------------------------------------------
# Contextual subfield mapping
# ---------------------------------------------------------------------

SECTION_FIELD_MAP: Dict[str, Dict[str, str]] = {
    SECTION_OWNER: {
        "name": "contact.owner.name",
        "contact person": "contact.owner.name",
        "person name": "contact.owner.name",
        "designation": "contact.owner.designation",
        "role": "contact.owner.designation",
        "mobile": "contact.owner.mobile",
        "mobile no": "contact.owner.mobile",
        "mobile number": "contact.owner.mobile",
        "contact no": "contact.owner.mobile",
        "contact number": "contact.owner.mobile",
        "phone": "contact.owner.mobile",
        "telephone": "contact.owner.mobile",
        "tel": "contact.owner.mobile",
        "email": "contact.owner.email",
        "email id": "contact.owner.email",
        "e mail": "contact.owner.email",
        "mail id": "contact.owner.email",
    },
    SECTION_PROJECT_CONTACT: {
        "name": "contact.project.name",
        "contact person": "contact.project.name",
        "person name": "contact.project.name",
        "designation": "contact.project.designation",
        "role": "contact.project.designation",
        "mobile": "contact.project.mobile",
        "mobile no": "contact.project.mobile",
        "mobile number": "contact.project.mobile",
        "contact no": "contact.project.mobile",
        "contact number": "contact.project.mobile",
        "phone": "contact.project.mobile",
        "telephone": "contact.project.mobile",
        "tel": "contact.project.mobile",
        "email": "contact.project.email",
        "email id": "contact.project.email",
        "e mail": "contact.project.email",
        "mail id": "contact.project.email",
    },
    SECTION_ACCOUNTS_CONTACT: {
        "name": "contact.accounts.name",
        "contact person": "contact.accounts.name",
        "person name": "contact.accounts.name",
        "designation": "contact.accounts.designation",
        "role": "contact.accounts.designation",
        "mobile": "contact.accounts.mobile",
        "mobile no": "contact.accounts.mobile",
        "mobile number": "contact.accounts.mobile",
        "contact no": "contact.accounts.mobile",
        "contact number": "contact.accounts.mobile",
        "phone": "contact.accounts.mobile",
        "telephone": "contact.accounts.mobile",
        "tel": "contact.accounts.mobile",
        "email": "contact.accounts.email",
        "email id": "contact.accounts.email",
        "e mail": "contact.accounts.email",
        "mail id": "contact.accounts.email",
    },
    SECTION_BANKING: {
        "bank name": "bank.name",
        "name of bank": "bank.name",
        "account number": "bank.account_number",
        "a c no": "bank.account_number",
        "ac no": "bank.account_number",
        "ifsc": "bank.ifsc",
        "ifsc code": "bank.ifsc",
        "branch": "bank.branch",
        "branch name": "bank.branch",
    },
    SECTION_TAX: {
        "pan": "tax.pan",
        "pan number": "tax.pan",
        "gst": "tax.gst.primary",
        "gst number": "tax.gst.primary",
        "gst no": "tax.gst.primary",
        "gstin": "tax.gst.primary",
        "pf": "tax.pf",
        "pf number": "tax.pf",
        "esi": "tax.esi",
        "esi number": "tax.esi",
        "msme": "tax.msme",
        "msme number": "tax.msme",
    },
    SECTION_FINANCIAL: {
        "fy 2024-25": "financial.turnover.fy2024_25",
        "2024-25": "financial.turnover.fy2024_25",
        "fy 2023-24": "financial.turnover.fy2023_24",
        "2023-24": "financial.turnover.fy2023_24",
        "fy 2022-23": "financial.turnover.fy2022_23",
        "2022-23": "financial.turnover.fy2022_23",
        "fy 2021-22": "financial.turnover.fy2021_22",
        "2021-22": "financial.turnover.fy2021_22",
    },
    SECTION_RESOURCE: {
        "engineers": "resource.manpower.engineers",
        "engineer": "resource.manpower.engineers",
        "supervisors": "resource.manpower.supervisors",
        "supervisor": "resource.manpower.supervisors",
        "skilled labour": "resource.manpower.skilled_labour",
        "skilled labor": "resource.manpower.skilled_labour",
        "unskilled labour": "resource.manpower.unskilled_labour",
        "unskilled labor": "resource.manpower.unskilled_labour",
        "total staff": "resource.manpower.total_staff",
        "total manpower": "resource.manpower.total_staff",
        "machinery": "resource.machinery.details",
        "machinery details": "resource.machinery.details",
        "workshop available": "resource.workshop_available",
    },
}


# ---------------------------------------------------------------------
# Basic section resolution
# ---------------------------------------------------------------------

def resolve_section_from_text(text: str) -> str:
    normalized = normalize_text(text)

    if not normalized:
        return SECTION_UNKNOWN

    # direct contains lookup
    for section_name, keywords in SECTION_HINT_MAP.items():
        for keyword in keywords:
            keyword_norm = normalize_text(keyword)
            if keyword_norm and keyword_norm in normalized:
                return section_name

    return SECTION_UNKNOWN


def score_section_from_text(text: str) -> Tuple[str, int]:
    """
    Returns:
        (best_section, score)
    score range:
        0..100
    """
    normalized = normalize_text(text)
    if not normalized:
        return SECTION_UNKNOWN, 0

    best_section = SECTION_UNKNOWN
    best_score = 0

    for section_name, keywords in SECTION_HINT_MAP.items():
        for keyword in keywords:
            keyword_norm = normalize_text(keyword)
            if not keyword_norm:
                continue

            if normalized == keyword_norm:
                return section_name, 100

            if keyword_norm in normalized:
                score = min(95, int((len(keyword_norm) / max(len(normalized), 1)) * 100) + 25)
                if score > best_score:
                    best_section = section_name
                    best_score = score

    return best_section, best_score


# ---------------------------------------------------------------------
# Section-aware contextual mapping for generic labels
# ---------------------------------------------------------------------

def resolve_contextual_field_key(section: str, normalized_label: str) -> str | None:
    """
    Maps a generic subfield like 'name' or 'email' to a contextual field key
    based on the active section.
    """
    if not section or section == SECTION_UNKNOWN:
        return None

    normalized_label = normalize_text(normalized_label)
    if not normalized_label:
        return None

    section_map = SECTION_FIELD_MAP.get(section, {})
    return section_map.get(normalized_label)


def resolve_contextual_field_key_with_fallback(
    section: str,
    normalized_label: str,
) -> Tuple[Optional[str], str]:
    """
    Returns:
        (field_key, resolution_mode)

    resolution_mode:
        - EXACT_CONTEXT
        - GENERIC_CONTEXT
        - UNRESOLVED
    """
    field_key = resolve_contextual_field_key(section, normalized_label)
    if field_key:
        return field_key, "EXACT_CONTEXT"

    # Small generic cleanup fallback
    compact = normalize_text(normalized_label).replace(".", "").replace("  ", " ").strip()

    field_key = resolve_contextual_field_key(section, compact)
    if field_key:
        return field_key, "GENERIC_CONTEXT"

    return None, "UNRESOLVED"


# ---------------------------------------------------------------------
# Sheet / scan driven section propagation helpers
# ---------------------------------------------------------------------

def infer_active_section_for_row(
    row_index: int,
    section_headers: List[Dict[str, Any]],
    max_lookback_rows: int = 8,
) -> str:
    """
    Finds the nearest previous section header within a row window.
    """
    best_section = SECTION_UNKNOWN
    best_distance = None

    for header in section_headers:
        header_row = int(header.get("row", 0))
        section_name = header.get("active_section") or resolve_section_from_text(header.get("value", ""))

        if not section_name or section_name == SECTION_UNKNOWN:
            continue

        if header_row <= row_index:
            distance = row_index - header_row
            if distance <= max_lookback_rows:
                if best_distance is None or distance < best_distance:
                    best_distance = distance
                    best_section = section_name

    return best_section


def enrich_findings_with_sections(
    findings: List[Dict[str, Any]],
    max_lookback_rows: int = 8,
) -> List[Dict[str, Any]]:
    """
    Backfills or strengthens active_section in scan findings using detected
    section header rows.

    This is helpful when sheets contain repeated contact blocks and the scanner
    captured labels before the final active_section assignment stabilized.
    """
    by_sheet: Dict[str, List[Dict[str, Any]]] = {}

    for item in findings:
        by_sheet.setdefault(item.get("sheet", ""), []).append(item)

    enriched: List[Dict[str, Any]] = []

    for sheet_name, rows in by_sheet.items():
        section_headers = [
            item for item in rows
            if item.get("cell_type") == "SECTION_HEADER"
        ]

        for item in rows:
            current_section = item.get("active_section") or SECTION_UNKNOWN
            if current_section == SECTION_UNKNOWN:
                inferred = infer_active_section_for_row(
                    row_index=int(item.get("row", 0)),
                    section_headers=section_headers,
                    max_lookback_rows=max_lookback_rows,
                )
                item = dict(item)
                item["active_section"] = inferred
            enriched.append(item)

    return enriched