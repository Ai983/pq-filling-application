from __future__ import annotations

import re
from typing import Any


def normalize_text(text: Any) -> str:
    """
    Conservative normalization for Excel label matching.

    Goals:
    - normalize casing and spacing
    - keep meaningful separators where useful
    - normalize common business-form variants
    """
    if text is None:
        return ""

    text = str(text).strip().lower()

    # normalize whitespace
    text = re.sub(r"[\n\r\t]+", " ", text)

    # normalize common punctuation variants
    text = text.replace("–", "-").replace("—", "-")
    text = text.replace("“", '"').replace("”", '"').replace("’", "'")

    # common semantic normalizations
    replacements = {
        "e-mail": "email",
        "email address": "email",
        "email id": "email id",
        "e mail": "email",
        "gstin no": "gstin",
        "gst no.": "gst no",
        "gst number": "gst number",
        "pan no.": "pan no",
        "mobile no.": "mobile no",
        "contact no.": "contact no",
        "telephone no.": "telephone",
        "a/c no": "account number",
        "a/c no.": "account number",
        "ac no": "account number",
        "a c no": "account number",
        "&": " and ",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)

    # remove unsupported characters but keep common label structure chars
    text = re.sub(r"[^a-z0-9 /&()\-:,\.]", "", text)

    # collapse punctuation spacing
    text = re.sub(r"\s*:\s*", ": ", text)
    text = re.sub(r"\s*,\s*", ", ", text)

    # collapse repeated spaces
    text = re.sub(r"\s+", " ", text)

    # trim trailing punctuation noise
    text = text.strip(" .,:;-")

    return text.strip()


def is_blank(value: Any) -> bool:
    return value is None or str(value).strip() == ""


def is_numeric_like(text: Any) -> bool:
    if text is None:
        return False

    text = str(text).strip()
    if not text:
        return False

    compact = (
        text.replace(",", "")
        .replace(".", "")
        .replace("%", "")
        .replace("-", "")
        .replace("/", "")
        .replace(":", "")
    )
    return compact.isdigit()


def word_count(text: Any) -> int:
    normalized = normalize_text(text)
    if not normalized:
        return 0
    return len(normalized.split())


def is_likely_instruction(text: Any) -> bool:
    normalized = normalize_text(text)

    if not normalized:
        return False

    instruction_markers = [
        "please provide",
        "tick whichever is applicable",
        "attach",
        "enclose",
        "if yes",
        "if any",
        "please mention",
        "specify details",
        "give details",
        "write in brief",
        "state whether",
        "mention details",
        "wherever applicable",
        "kindly provide",
        "please attach",
        "please enclose",
    ]

    return any(marker in normalized for marker in instruction_markers)


def is_likely_section_header(text: Any) -> bool:
    normalized = normalize_text(text)

    if not normalized:
        return False

    section_markers = [
        "contact detail",
        "contact details",
        "details of",
        "particulars of",
        "bank details",
        "financial details",
        "project details",
        "workshop facilities",
        "machinery details",
        "staff strength",
        "turnover",
        "statutory details",
        "registration details",
        "general information",
        "company details",
        "organisation details",
        "organization details",
        "key position holder",
        "principal owner",
        "project location",
        "compliance",
        "quality",
        "safety",
        "organization structure",
        "declaration",
        "undertaking",
        "projects executed",
        "completed projects",
        "ongoing projects",
    ]

    if any(marker in normalized for marker in section_markers):
        return True

    wc = word_count(normalized)

    # Short heading-like lines often act as sections
    if 2 <= wc <= 10:
        if normalized.endswith(":"):
            return True

    return False


def is_likely_subfield(text: Any) -> bool:
    normalized = normalize_text(text)

    if not normalized:
        return False

    subfields = {
        "name",
        "designation",
        "mobile",
        "mobile no",
        "mobile number",
        "contact no",
        "contact number",
        "phone",
        "telephone",
        "email",
        "email id",
        "fax",
        "address",
        "location",
        "role",
    }

    return normalized in subfields


def is_likely_table_header(text: Any) -> bool:
    normalized = normalize_text(text)

    if not normalized:
        return False

    table_headers = {
        "project name",
        "name of project",
        "client name",
        "client",
        "customer",
        "location",
        "city",
        "state",
        "start date",
        "commencement date",
        "completion date",
        "end date",
        "status",
        "value",
        "contract value",
        "awarded value",
        "project value",
        "area",
        "remarks",
        "year",
        "fy",
        "description",
        "qty",
        "quantity",
        "make",
        "model",
        "capacity",
        "contact person",
        "contact number",
    }

    return normalized in table_headers


def is_likely_simple_field(text: Any) -> bool:
    normalized = normalize_text(text)

    if not normalized:
        return False

    simple_markers = [
        "registered name",
        "legal name",
        "name of firm",
        "company name",
        "name of company",
        "registered address",
        "business address",
        "communication address",
        "office address",
        "head office address",
        "year of registration",
        "year of incorporation",
        "year of establishment",
        "pan",
        "pan no",
        "pan number",
        "gst",
        "gst no",
        "gst number",
        "gstin",
        "pf",
        "pf registration number",
        "esi",
        "esi registration number",
        "bank name",
        "branch name",
        "account number",
        "ifsc",
        "ifsc code",
        "website",
        "email",
        "email address",
        "phone number",
        "contact number",
        "type of firm",
        "constitution of business",
        "type of company",
        "nature of business",
    ]

    return any(marker == normalized for marker in simple_markers)


def is_likely_declaration(text: Any) -> bool:
    normalized = normalize_text(text)

    if not normalized:
        return False

    declaration_markers = [
        "hereby declare",
        "we hereby declare",
        "we certify",
        "certify that",
        "undertake that",
        "authorized signatory",
        "authorised signatory",
        "signature with seal",
        "seal and signature",
    ]

    if any(marker in normalized for marker in declaration_markers):
        return True

    if word_count(normalized) > 25:
        return True

    return False


def is_likely_label(text: Any) -> bool:
    if text is None:
        return False

    normalized = normalize_text(text)

    if not normalized:
        return False

    if len(normalized) > 140:
        return False

    if is_numeric_like(normalized):
        return False

    ignore_values = {
        "yes",
        "no",
        "na",
        "n/a",
        "-",
        "--",
        "---",
    }
    if normalized in ignore_values:
        return False

    # obvious declaration paragraphs are not regular label candidates
    if is_likely_declaration(normalized) and word_count(normalized) > 25:
        return False

    return True