from __future__ import annotations

from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

from rapidfuzz import fuzz, process

from app.core.config import FUZZY_MATCH_THRESHOLD
from app.core.constants import (
    MATCH_TYPE_EMPTY,
    MATCH_TYPE_EXACT,
    MATCH_TYPE_FUZZY,
    MATCH_TYPE_NORMALIZED_EXACT,
    MATCH_TYPE_SECTION_CONTEXT,
    MATCH_TYPE_TOKEN_CONTAINS,
    MATCH_TYPE_UNMATCHED,
    SECTION_ACCOUNTS_CONTACT,
    SECTION_BANKING,
    SECTION_COMPANY,
    SECTION_COMPLIANCE,
    SECTION_FINANCIAL,
    SECTION_OWNER,
    SECTION_PROJECT_CONTACT,
    SECTION_PROJECTS,
    SECTION_RESOURCE,
    SECTION_TAX,
    SECTION_TECHNICAL,
    SECTION_UNKNOWN,
)
from app.engine.utils import normalize_text


# ---------------------------------------------------------------------
# Domain phrase normalization helpers
# These improve deterministic matching for PQ/vendor registration forms.
# ---------------------------------------------------------------------
PHRASE_ALIASES: Dict[str, str] = {
    "reg no": "registration number",
    "reg number": "registration number",
    "registration no": "registration number",
    "registration no.": "registration number",
    "regn no": "registration number",
    "regn no.": "registration number",
    "gst no": "gst number",
    "gst no.": "gst number",
    "pan no": "pan number",
    "pan no.": "pan number",
    "pf no": "pf number",
    "pf no.": "pf number",
    "esi no": "esi number",
    "esi no.": "esi number",
    "esic no": "esi number",
    "esic no.": "esi number",
    "msme no": "msme number",
    "msme no.": "msme number",
    "udyam no": "udyam number",
    "udyam no.": "udyam number",
    "yr of establishment": "year of establishment",
    "year of estb": "year of establishment",
    "year of estd": "year of establishment",
    "type of co": "type of company",
    "co type": "company type",
    "owner md": "owner managing director",
    "md name": "managing director name",
    "contact details": "contact information",
    "annual turn over": "annual turnover",
    "turn over": "turnover",
    "work order value": "value of work order",
    "no of": "number of",
    "list of major projects executed": "major projects executed",
    "list of projects under execution": "projects under execution",
}


STOPWORDS = {
    "the",
    "a",
    "an",
    "of",
    "for",
    "and",
    "or",
    "to",
    "with",
    "by",
    "in",
    "at",
    "on",
    "as",
    "from",
    "is",
    "are",
    "be",
    "details",
    "detail",
    "information",
    "info",
}


IMPORTANT_TOKENS = {
    "company",
    "legal",
    "name",
    "entity",
    "type",
    "business",
    "address",
    "head",
    "office",
    "factory",
    "works",
    "plant",
    "gst",
    "pan",
    "pf",
    "esi",
    "msme",
    "udyam",
    "bank",
    "account",
    "ifsc",
    "branch",
    "owner",
    "managing",
    "director",
    "md",
    "contact",
    "project",
    "client",
    "work",
    "order",
    "value",
    "turnover",
    "financial",
    "technical",
    "staff",
    "engineers",
    "supervisors",
    "machinery",
    "capacity",
    "certificate",
    "certificates",
    "registration",
    "number",
    "year",
    "establishment",
}


def _ensure_list(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    if isinstance(value, tuple):
        return [str(v).strip() for v in value if str(v).strip()]
    text = str(value).strip()
    return [text] if text else []


def _apply_phrase_aliases(text: str) -> str:
    if not text:
        return ""
    updated = f" {text} "
    for source, target in sorted(PHRASE_ALIASES.items(), key=lambda x: len(x[0]), reverse=True):
        updated = updated.replace(f" {source} ", f" {target} ")
    return " ".join(updated.split()).strip()


def _normalize_for_matching(text: str) -> str:
    normalized = normalize_text(text)
    if not normalized:
        return ""
    normalized = _apply_phrase_aliases(normalized)
    normalized = " ".join(normalized.split()).strip()
    return normalized


def _normalize_synonyms_map(
    synonyms: Any,
) -> Tuple[Dict[str, str], Dict[str, Dict[str, Any]]]:
    """
    Supports both:
    1. old style:
        {
            "company name": "company.legal_name",
            "gst number": "tax.gst.primary"
        }

    2. richer style:
        {
            "company name": {
                "field_key": "company.legal_name",
                "category": "company",
                "risk_level": "low",
                "section_affinity": ["company"]
            }
        }

    Returns:
        flat_map: normalized label -> field_key
        metadata_map: normalized label -> metadata dict
    """
    flat_map: Dict[str, str] = {}
    metadata_map: Dict[str, Dict[str, Any]] = {}

    if not isinstance(synonyms, dict):
        return flat_map, metadata_map

    for raw_label, raw_value in synonyms.items():
        normalized_label = _normalize_for_matching(str(raw_label))
        if not normalized_label:
            continue

        if isinstance(raw_value, dict):
            field_key = str(raw_value.get("field_key", "")).strip()
            if not field_key:
                continue

            flat_map[normalized_label] = field_key
            metadata_map[normalized_label] = {
                "field_key": field_key,
                "category": raw_value.get("category"),
                "risk_level": raw_value.get("risk_level"),
                "notes": raw_value.get("notes"),
                "section_affinity": _ensure_list(raw_value.get("section_affinity")),
                "layout_hint": raw_value.get("layout_hint"),
                "match_priority": raw_value.get("match_priority"),
                "value_type": raw_value.get("value_type"),
            }
        else:
            field_key = str(raw_value).strip()
            if not field_key:
                continue

            flat_map[normalized_label] = field_key
            metadata_map[normalized_label] = {
                "field_key": field_key,
                "category": None,
                "risk_level": None,
                "notes": None,
                "section_affinity": [],
                "layout_hint": None,
                "match_priority": None,
                "value_type": None,
            }

    return flat_map, metadata_map


def _tokenize(text: str) -> List[str]:
    normalized = _normalize_for_matching(text)
    if not normalized:
        return []
    return [token for token in normalized.split() if token]


def _content_tokens(text: str) -> List[str]:
    tokens = _tokenize(text)
    return [t for t in tokens if t not in STOPWORDS]


def _token_contains_score(label_a: str, label_b: str) -> int:
    """
    Conservative but domain-aware token-overlap score.
    returns 0..100
    """
    tokens_a = set(_content_tokens(label_a))
    tokens_b = set(_content_tokens(label_b))

    if not tokens_a or not tokens_b:
        return 0

    intersection = tokens_a.intersection(tokens_b)
    if not intersection:
        return 0

    # Base overlap
    coverage_a = len(intersection) / len(tokens_a)
    coverage_b = len(intersection) / len(tokens_b)
    base_score = min(coverage_a, coverage_b) * 100.0

    # Stronger signal if important tokens overlap
    important_overlap = len(intersection.intersection(IMPORTANT_TOKENS))
    if important_overlap:
        base_score += min(12, important_overlap * 4)

    # Penalize if one side has too many unmatched meaningful tokens
    unmatched_a = len(tokens_a - intersection)
    unmatched_b = len(tokens_b - intersection)
    penalty = min(12, (unmatched_a + unmatched_b) * 2)

    final_score = max(0, min(100, int(round(base_score - penalty))))
    return final_score


def _phrase_compression_variants(text: str) -> List[str]:
    """
    Generates a few deterministic semantic variants for better matching.
    """
    normalized = _normalize_for_matching(text)
    if not normalized:
        return []

    variants = {normalized}

    replacements = [
        ("type of company", "company type"),
        ("type of business", "business type"),
        ("year of establishment", "incorporation year"),
        ("head office address", "office address"),
        ("factory address", "plant address"),
        ("contact information", "contact"),
        ("annual turnover", "turnover"),
        ("value of work order", "work order value"),
        ("number of technical staff", "technical staff"),
        ("owner managing director name", "owner name"),
        ("managing director name", "director name"),
    ]

    for source, target in replacements:
        if source in normalized:
            variants.add(normalized.replace(source, target).strip())
        if target in normalized:
            variants.add(normalized.replace(target, source).strip())

    return [v for v in variants if v]


def _field_section_affinity(field_key: str) -> List[str]:
    """
    Field-family-based fallback section affinity.
    Used when explicit section_affinity is absent in synonym metadata.
    """
    if not field_key:
        return []

    field_key = field_key.strip().lower()

    if field_key.startswith("company."):
        return [SECTION_COMPANY]
    if field_key.startswith("tax."):
        return [SECTION_TAX, SECTION_COMPLIANCE]
    if field_key.startswith("bank."):
        return [SECTION_BANKING]
    if field_key.startswith("financial."):
        return [SECTION_FINANCIAL]
    if field_key.startswith("resource."):
        return [SECTION_RESOURCE]
    if field_key.startswith("compliance."):
        return [SECTION_COMPLIANCE]

    if field_key.startswith("contact.owner."):
        return [SECTION_OWNER, SECTION_COMPANY]
    if field_key.startswith("contact.accounts."):
        return [SECTION_ACCOUNTS_CONTACT, SECTION_BANKING, SECTION_FINANCIAL]
    if field_key.startswith("contact.project."):
        return [SECTION_PROJECT_CONTACT, SECTION_PROJECTS, SECTION_TECHNICAL]

    if field_key.startswith("project.") or field_key.startswith("projects."):
        return [SECTION_PROJECTS]

    return [SECTION_UNKNOWN]


def _section_bonus(active_section: Optional[str], field_key: str, metadata: Dict[str, Any]) -> Tuple[int, bool]:
    """
    Returns:
        bonus_points, section_matched
    """
    if not active_section or active_section == SECTION_UNKNOWN or not field_key:
        return 0, False

    explicit_affinity = _ensure_list(metadata.get("section_affinity"))
    affinity = explicit_affinity or _field_section_affinity(field_key)

    if active_section in affinity:
        return 10, True

    return 0, False


def _priority_bonus(metadata: Dict[str, Any]) -> int:
    priority = str(metadata.get("match_priority") or "").strip().lower()
    if priority == "high":
        return 4
    if priority == "medium":
        return 2
    return 0


def _best_fuzzy_match(normalized_label: str, flat_map: Dict[str, str]) -> Tuple[Optional[str], int]:
    choices = list(flat_map.keys())
    if not choices:
        return None, 0

    match = process.extractOne(normalized_label, choices, scorer=fuzz.token_sort_ratio)
    if not match:
        return None, 0

    matched_label, score, _ = match
    return matched_label, int(score)


def _best_token_match(normalized_label: str, flat_map: Dict[str, str]) -> Tuple[Optional[str], int]:
    best_label = None
    best_score = 0

    for candidate in flat_map.keys():
        score = _token_contains_score(normalized_label, candidate)
        if score > best_score:
            best_label = candidate
            best_score = score

    return best_label, best_score


def _candidate_from_match(
    matched_label: str,
    match_type: str,
    base_score: int,
    flat_map: Dict[str, str],
    metadata_map: Dict[str, Dict[str, Any]],
    active_section: Optional[str],
) -> Dict[str, Any]:
    field_key = flat_map[matched_label]
    metadata = metadata_map.get(matched_label, {})
    section_bonus, section_matched = _section_bonus(active_section, field_key, metadata)
    priority_bonus = _priority_bonus(metadata)

    final_score = min(100, base_score + section_bonus + priority_bonus)

    return {
        "field_key": field_key,
        "match_type": match_type,
        "base_score": int(base_score),
        "final_score": int(final_score),
        "matched_label": matched_label,
        "section_bonus_applied": section_matched,
        "priority_bonus_applied": priority_bonus > 0,
        "metadata": metadata,
    }


def match_label(
    label_text: str,
    synonyms: dict,
    active_section: Optional[str] = None,
) -> tuple[str | None, str, int]:
    """
    Backward-compatible matcher.

    Returns:
        (field_key, match_type, score)
    """
    result = match_label_detailed(
        label_text=label_text,
        synonyms=synonyms,
        active_section=active_section,
    )
    return (
        result["field_key"],
        result["match_type"],
        int(result["score"]),
    )


def match_label_detailed(
    label_text: str,
    synonyms: dict,
    active_section: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Rich deterministic label matcher.

    Output example:
    {
        "label_text": "Company Name",
        "normalized_label": "company name",
        "field_key": "company.legal_name",
        "match_type": "EXACT",
        "score": 100,
        "semantic_confidence": 1.0,
        "matched_label": "company name",
        "section_used": "company",
        "section_bonus_applied": False,
        "metadata": {...}
    }
    """
    normalized = _normalize_for_matching(label_text)

    if not normalized:
        return {
            "label_text": label_text,
            "normalized_label": "",
            "field_key": None,
            "match_type": MATCH_TYPE_EMPTY,
            "score": 0,
            "semantic_confidence": 0.0,
            "matched_label": None,
            "section_used": active_section or SECTION_UNKNOWN,
            "section_bonus_applied": False,
            "metadata": {},
        }

    flat_map, metadata_map = _normalize_synonyms_map(synonyms)

    if not flat_map:
        return {
            "label_text": label_text,
            "normalized_label": normalized,
            "field_key": None,
            "match_type": MATCH_TYPE_UNMATCHED,
            "score": 0,
            "semantic_confidence": 0.0,
            "matched_label": None,
            "section_used": active_section or SECTION_UNKNOWN,
            "section_bonus_applied": False,
            "metadata": {},
        }

    # 1) Direct normalized exact and semantic variants
    exact_variants = _phrase_compression_variants(normalized)
    for variant in exact_variants:
        if variant in flat_map:
            field_key = flat_map[variant]
            metadata = metadata_map.get(variant, {})
            _, section_matched = _section_bonus(active_section, field_key, metadata)

            return {
                "label_text": label_text,
                "normalized_label": normalized,
                "field_key": field_key,
                "match_type": MATCH_TYPE_NORMALIZED_EXACT if variant != normalize_text(label_text) else MATCH_TYPE_EXACT,
                "score": 100,
                "semantic_confidence": 1.0,
                "matched_label": variant,
                "section_used": active_section or SECTION_UNKNOWN,
                "section_bonus_applied": section_matched,
                "metadata": metadata,
            }

    candidates: List[Dict[str, Any]] = []

    # 2) Token overlap candidates from all variants
    for variant in exact_variants:
        token_match_label, token_score = _best_token_match(variant, flat_map)
        if token_match_label and token_score >= 82:
            candidates.append(
                _candidate_from_match(
                    matched_label=token_match_label,
                    match_type=MATCH_TYPE_TOKEN_CONTAINS,
                    base_score=token_score,
                    flat_map=flat_map,
                    metadata_map=metadata_map,
                    active_section=active_section,
                )
            )

    # 3) Fuzzy candidates from all variants
    for variant in exact_variants:
        fuzzy_match_label, fuzzy_score = _best_fuzzy_match(variant, flat_map)
        if fuzzy_match_label and fuzzy_score >= FUZZY_MATCH_THRESHOLD:
            candidates.append(
                _candidate_from_match(
                    matched_label=fuzzy_match_label,
                    match_type=MATCH_TYPE_FUZZY,
                    base_score=fuzzy_score,
                    flat_map=flat_map,
                    metadata_map=metadata_map,
                    active_section=active_section,
                )
            )

    if not candidates:
        return {
            "label_text": label_text,
            "normalized_label": normalized,
            "field_key": None,
            "match_type": MATCH_TYPE_UNMATCHED,
            "score": 0,
            "semantic_confidence": 0.0,
            "matched_label": None,
            "section_used": active_section or SECTION_UNKNOWN,
            "section_bonus_applied": False,
            "metadata": {},
        }

    # Remove weak ambiguous duplicates by keeping best per field_key
    deduped: Dict[str, Dict[str, Any]] = {}
    for item in candidates:
        key = item["field_key"]
        existing = deduped.get(key)
        if existing is None or item["final_score"] > existing["final_score"]:
            deduped[key] = item

    ranked = list(deduped.values())
    ranked.sort(
        key=lambda item: (
            item["final_score"],
            1 if item["section_bonus_applied"] else 0,
            1 if item["priority_bonus_applied"] else 0,
            1 if item["match_type"] == MATCH_TYPE_TOKEN_CONTAINS else 0,
        ),
        reverse=True,
    )

    best = ranked[0]

    # Ambiguity guard:
    # if top two are too close and point to different field keys, reject.
    if len(ranked) >= 2:
        second = ranked[1]
        if (
            best["field_key"] != second["field_key"]
            and abs(best["final_score"] - second["final_score"]) <= 3
            and best["final_score"] < 92
        ):
            return {
                "label_text": label_text,
                "normalized_label": normalized,
                "field_key": None,
                "match_type": MATCH_TYPE_UNMATCHED,
                "score": 0,
                "semantic_confidence": 0.0,
                "matched_label": None,
                "section_used": active_section or SECTION_UNKNOWN,
                "section_bonus_applied": False,
                "metadata": {},
            }

    match_type = best["match_type"]
    if best["section_bonus_applied"] and match_type in {MATCH_TYPE_FUZZY, MATCH_TYPE_TOKEN_CONTAINS}:
        match_type = MATCH_TYPE_SECTION_CONTEXT

    return {
        "label_text": label_text,
        "normalized_label": normalized,
        "field_key": best["field_key"],
        "match_type": match_type,
        "score": int(best["final_score"]),
        "semantic_confidence": round(best["final_score"] / 100.0, 4),
        "matched_label": best["matched_label"],
        "section_used": active_section or SECTION_UNKNOWN,
        "section_bonus_applied": bool(best["section_bonus_applied"]),
        "metadata": best["metadata"],
    }


def build_field_to_labels_index(synonyms: dict) -> Dict[str, List[str]]:
    """
    Utility for debugging / teaching workflows.
    Returns:
        {
            "company.legal_name": ["company name", "name of firm", ...]
        }
    """
    flat_map, _ = _normalize_synonyms_map(synonyms)
    grouped: Dict[str, List[str]] = defaultdict(list)

    for label, field_key in flat_map.items():
        grouped[field_key].append(label)

    return dict(grouped)