from __future__ import annotations

from typing import Any, Dict, List, Optional

import pandas as pd

from app.core.config import resolve_master_data_file, resolve_synonym_file
from app.core.constants import (
    MASTER_PROJECTS_KEY,
    MASTER_STANDARD_TEXT_BLOCKS_KEY,
    MASTER_TEMPLATE_HINTS_KEY,
    MASTER_VALUE_VARIANTS_KEY,
)
from app.engine.utils import normalize_text


# ---------------------------------------------------------------------
# Master key aliases
# ---------------------------------------------------------------------
MASTER_FIELD_ALIASES: Dict[str, List[str]] = {
    "company.legal_name": [
        "company.name",
        "company.company_name",
        "company.registered_name",
        "company.organization_name",
        "company.organisation_name",
    ],
    "company.entity_type": [
        "company.type",
        "company.company_type",
        "company.constitution_type",
        "company.legal_entity_type",
    ],
    "company.incorporation_year": [
        "company.year_of_establishment",
        "company.establishment_year",
        "company.year_established",
        "company.founding_year",
    ],
    "company.business_type": [
        "company.nature_of_business",
        "company.business_nature",
        "company.type_of_business",
        "company.core_business_type",
    ],
    "company.address": [
        "company.office_address",
        "company.head_office_address",
        "company.registered_address",
        "company.business_address",
        "company.corporate_address",
    ],
    "company.factory_address": [
        "company.works_address",
        "company.plant_address",
        "company.site_address",
        "company.manufacturing_address",
        "company.office_address",
        "company.business_address",
    ],
    "company.phone": [
        "company.mobile",
        "company.contact_number",
        "company.phone_number",
        "company.telephone",
        "company.tel",
        "company.office_phone",
    ],
    "company.email": [
        "company.mail",
        "company.email_id",
        "company.contact_email",
        "company.office_email",
    ],
    "company.website": [
        "company.web",
        "company.web_site",
        "company.url",
    ],
    "tax.pan": [
        "company.pan",
        "statutory.pan",
        "tax.pan_number",
    ],
    "tax.gst.primary": [
        "tax.gst",
        "company.gst",
        "statutory.gst",
        "tax.gstin",
        "tax.gst_number",
    ],
    "tax.pf": [
        "company.pf",
        "statutory.pf",
        "tax.pf_number",
        "tax.pf_registration_no",
        "tax.provident_fund",
    ],
    "tax.esi": [
        "company.esi",
        "statutory.esi",
        "tax.esi_number",
        "tax.esic",
        "tax.esic_number",
    ],
    "tax.msme": [
        "company.msme",
        "company.msme_uam",
        "company.udyam",
        "tax.msme_uam",
        "tax.udyam",
        "statutory.msme",
    ],
    "bank.name": [
        "bank.bank_name",
        "bank.account_bank_name",
    ],
    "bank.account_number": [
        "bank.account_no",
        "bank.ac_no",
        "bank.a_c_no",
    ],
    "bank.ifsc": [
        "bank.ifsc_code",
    ],
    "bank.branch": [
        "bank.branch_name",
    ],
    "resource.manpower.engineers": [
        "resource.engineering_staff",
        "resource.technical_staff",
        "resource.engineering_staff_pan_india",
        "resource.engineering_staff_project_location",
        "resource.no_of_engineers",
    ],
    "resource.manpower.supervisors": [
        "resource.supervisory_staff",
        "resource.no_of_supervisors",
    ],
    "resource.manpower.skilled_labour": [
        "resource.skilled_labour",
        "resource.skilled_workers",
        "resource.skilled_workmen",
    ],
    "resource.manpower.unskilled_labour": [
        "resource.unskilled_labour",
        "resource.unskilled_workers",
        "resource.unskilled_workmen",
    ],
    "resource.manpower.total_staff": [
        "resource.total_manpower",
        "resource.total_staff_strength",
    ],
    "resource.machinery.details": [
        "resource.plant_machinery",
        "resource.available_plant_machinery",
        "resource.machinery",
    ],
    "resource.workshop_available": [
        "resource.workshop",
        "resource.own_workshop",
    ],
    "contact.owner.name": [
        "contact.promoter.name",
        "contact.md.name",
        "contact.owner_md.name",
        "contact.director.name",
    ],
    "contact.owner.designation": [
        "contact.promoter.designation",
        "contact.md.designation",
        "contact.owner_md.designation",
        "contact.director.designation",
    ],
    "contact.owner.mobile": [
        "contact.promoter.mobile",
        "contact.md.mobile",
        "contact.owner_md.mobile",
        "contact.director.mobile",
    ],
    "contact.owner.email": [
        "contact.promoter.email",
        "contact.md.email",
        "contact.owner_md.email",
        "contact.director.email",
    ],
    "contact.project.name": [
        "contact.projects.name",
        "contact.project_head.name",
        "contact.execution.name",
    ],
    "contact.project.designation": [
        "contact.projects.designation",
        "contact.project_head.designation",
        "contact.execution.designation",
    ],
    "contact.project.mobile": [
        "contact.projects.mobile",
        "contact.project_head.mobile",
        "contact.execution.mobile",
    ],
    "contact.project.email": [
        "contact.projects.email",
        "contact.project_head.email",
        "contact.execution.email",
    ],
    "contact.accounts.name": [
        "contact.finance.name",
        "contact.billing.name",
    ],
    "contact.accounts.designation": [
        "contact.finance.designation",
        "contact.billing.designation",
    ],
    "contact.accounts.mobile": [
        "contact.finance.mobile",
        "contact.billing.mobile",
    ],
    "contact.accounts.email": [
        "contact.finance.email",
        "contact.billing.email",
    ],
}


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------
def _safe_str(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _safe_value(value: Any):
    if pd.isna(value):
        return None
    return value


def _row_has_meaningful_data(row_dict: Dict[str, Any]) -> bool:
    for value in row_dict.values():
        if value is None:
            continue
        if str(value).strip() != "":
            return True
    return False


def _normalize_sheet_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def _first_present(row: pd.Series, *keys: str) -> Any:
    for key in keys:
        if key in row.index:
            value = row.get(key)
            if not pd.isna(value):
                return value
    return None


def _build_normalized_master_index(master_data: Dict[str, Any]) -> Dict[str, str]:
    index: Dict[str, str] = {}
    for key in master_data.keys():
        if not isinstance(key, str):
            continue
        normalized = normalize_text(key)
        if normalized and normalized not in index:
            index[normalized] = key
    return index


def _finalize_master_alias_index(master_data: Dict[str, Any]) -> Dict[str, List[str]]:
    normalized_index = _build_normalized_master_index(master_data)
    alias_index: Dict[str, List[str]] = {}

    all_candidate_keys = set(master_data.keys())
    for canonical, aliases in MASTER_FIELD_ALIASES.items():
        ordered_candidates: List[str] = []

        if canonical in all_candidate_keys:
            ordered_candidates.append(canonical)
        else:
            normalized_canonical = normalize_text(canonical)
            actual = normalized_index.get(normalized_canonical)
            if actual:
                ordered_candidates.append(actual)

        for alias in aliases:
            actual_alias = alias if alias in all_candidate_keys else normalized_index.get(normalize_text(alias))
            if actual_alias and actual_alias not in ordered_candidates:
                ordered_candidates.append(actual_alias)

        if ordered_candidates:
            alias_index[canonical] = ordered_candidates

    return alias_index


def _normalize_project_columns(df: pd.DataFrame) -> Dict[str, str]:
    """
    normalized_col_name -> actual dataframe column name
    """
    return {normalize_text(col): col for col in df.columns}


def _find_projectish_column(column_index: Dict[str, str], *candidates: str) -> Optional[str]:
    for candidate in candidates:
        normalized_candidate = normalize_text(candidate)
        if normalized_candidate in column_index:
            return column_index[normalized_candidate]

    for normalized_col, actual_col in column_index.items():
        for candidate in candidates:
            normalized_candidate = normalize_text(candidate)
            if normalized_candidate and (
                normalized_candidate in normalized_col or normalized_col in normalized_candidate
            ):
                return actual_col

    return None


def _sheet_looks_like_project_sheet(sheet_name: str, df: pd.DataFrame) -> bool:
    normalized_sheet_name = normalize_text(sheet_name)
    if normalized_sheet_name in {
        "project master",
        "projects",
        "project details",
        "project detail",
        "completed projects",
        "ongoing projects",
        "executed projects",
        "projects under execution",
        "major projects",
    }:
        return True

    column_index = _normalize_project_columns(df)

    projectish_hits = 0
    signals = [
        "client name",
        "client",
        "type of work",
        "nature of work",
        "scope of work",
        "value of work order",
        "work order value",
        "volume of work",
        "order date",
        "work completion date",
        "completion date",
        "project name",
        "name of project",
    ]

    for signal in signals:
        if _find_projectish_column(column_index, signal):
            projectish_hits += 1

    return projectish_hits >= 3


def _infer_project_status_from_sheet(sheet_name: str) -> Optional[str]:
    normalized_sheet_name = normalize_text(sheet_name)
    if any(token in normalized_sheet_name for token in ["ongoing", "under execution", "running", "in progress"]):
        return "ongoing"
    if any(token in normalized_sheet_name for token in ["completed", "executed", "finished"]):
        return "completed"
    return None


def _derive_project_name_from_row(row: pd.Series, column_index: Dict[str, str]) -> Optional[str]:
    explicit_project_col = _find_projectish_column(
        column_index,
        "project_name",
        "project name",
        "name_of_project",
        "name of project",
        "project",
        "work name",
    )
    if explicit_project_col:
        value = row.get(explicit_project_col)
        if not pd.isna(value) and str(value).strip():
            return str(value).strip()

    client_col = _find_projectish_column(column_index, "client", "client name", "customer")
    category_col = _find_projectish_column(column_index, "type of work", "nature of work", "scope of work", "category")

    client_value = ""
    category_value = ""

    if client_col:
        value = row.get(client_col)
        if not pd.isna(value):
            client_value = str(value).strip()

    if category_col:
        value = row.get(category_col)
        if not pd.isna(value):
            category_value = str(value).strip()

    if client_value and category_value:
        return f"{client_value} - {category_value}"
    if client_value:
        return client_value
    if category_value:
        return category_value

    return None


def _derive_project_location_from_row(row: pd.Series, column_index: Dict[str, str]) -> Optional[str]:
    location_col = _find_projectish_column(column_index, "location", "site location")
    if location_col:
        value = row.get(location_col)
        if not pd.isna(value) and str(value).strip():
            return str(value).strip()

    city_col = _find_projectish_column(column_index, "city", "location city")
    state_col = _find_projectish_column(column_index, "state", "location state")

    city_text = ""
    state_text = ""

    if city_col:
        value = row.get(city_col)
        if not pd.isna(value):
            city_text = str(value).strip()

    if state_col:
        value = row.get(state_col)
        if not pd.isna(value):
            state_text = str(value).strip()

    if city_text and state_text:
        return f"{city_text}, {state_text}"
    if city_text:
        return city_text
    if state_text:
        return state_text
    return None


def _build_project_record_from_row(
    row: pd.Series,
    column_index: Dict[str, str],
    sheet_level_status: Optional[str] = None,
) -> Dict[str, Any]:
    project_name_col = _find_projectish_column(
        column_index,
        "project_name",
        "project name",
        "name_of_project",
        "name of project",
        "project",
        "work name",
    )
    client_col = _find_projectish_column(column_index, "client", "client name", "customer")
    city_col = _find_projectish_column(column_index, "city", "location city")
    state_col = _find_projectish_column(column_index, "state", "location state")
    location_col = _find_projectish_column(column_index, "location", "site location")
    status_col = _find_projectish_column(column_index, "status", "project status")
    bucket_col = _find_projectish_column(column_index, "bucket", "project bucket", "project_group")
    start_date_col = _find_projectish_column(
        column_index,
        "start_date",
        "start date",
        "order_date",
        "order date",
        "commencement_date",
        "commencement date",
        "po date",
    )
    end_date_col = _find_projectish_column(
        column_index,
        "end_date",
        "end date",
        "completion_date",
        "completion date",
        "work_completion_date",
        "work completion date",
    )
    value_col = _find_projectish_column(
        column_index,
        "value",
        "project value",
        "contract value",
        "value of work order",
        "work order value",
        "volume of work",
        "awarded value",
    )
    category_col = _find_projectish_column(
        column_index,
        "category",
        "type of work",
        "nature of work",
        "scope of work",
        "type of project",
    )
    contact_name_col = _find_projectish_column(
        column_index,
        "contact_name",
        "contact name",
        "contact person",
        "client contact name",
        "client_contact",
        "client contact",
    )
    contact_phone_col = _find_projectish_column(
        column_index,
        "contact_phone",
        "contact phone",
        "client contact phone",
        "contact number",
        "contact no",
        "mobile",
        "phone",
    )
    contact_email_col = _find_projectish_column(
        column_index,
        "contact_email",
        "contact email",
        "client contact email",
        "email",
        "email id",
        "mail id",
    )
    pmc_col = _find_projectish_column(
        column_index,
        "pmc_name",
        "pmc",
        "pmc name",
        "project management consultant",
        "consultant",
    )
    area_col = _find_projectish_column(
        column_index,
        "area_sft",
        "area sqft",
        "area in sft",
        "area in sqft",
        "built up area",
        "builtup area",
        "project area",
        "area",
    )

    record = {
        "project_name": _safe_value(row.get(project_name_col)) if project_name_col else None,
        "client": _safe_value(row.get(client_col)) if client_col else None,
        "city": _safe_value(row.get(city_col)) if city_col else None,
        "state": _safe_value(row.get(state_col)) if state_col else None,
        "location": _safe_value(row.get(location_col)) if location_col else None,
        "status": _safe_value(row.get(status_col)) if status_col else sheet_level_status,
        "bucket": _safe_value(row.get(bucket_col)) if bucket_col else None,
        "start_date": _safe_value(row.get(start_date_col)) if start_date_col else None,
        "end_date": _safe_value(row.get(end_date_col)) if end_date_col else None,
        "value": _safe_value(row.get(value_col)) if value_col else None,
        "category": _safe_value(row.get(category_col)) if category_col else None,
        "contact_name": _safe_value(row.get(contact_name_col)) if contact_name_col else None,
        "contact_phone": _safe_value(row.get(contact_phone_col)) if contact_phone_col else None,
        "contact_email": _safe_value(row.get(contact_email_col)) if contact_email_col else None,
        "pmc_name": _safe_value(row.get(pmc_col)) if pmc_col else None,
        "area_sft": _safe_value(row.get(area_col)) if area_col else None,
    }

    if not record["project_name"]:
        record["project_name"] = _derive_project_name_from_row(row, column_index)

    if not record["location"]:
        record["location"] = _derive_project_location_from_row(row, column_index)

    return record


# ---------------------------------------------------------------------
# Public lookup helpers for engine use
# ---------------------------------------------------------------------
def get_master_value(master_data: Dict[str, Any], field_key: str, default: Any = None) -> Any:
    if not field_key:
        return default

    if field_key in master_data:
        value = master_data.get(field_key)
        if value is not None and str(value).strip() != "":
            return value

    normalized_index = master_data.get("__normalized_master_index__", {})
    if isinstance(normalized_index, dict):
        actual_key = normalized_index.get(normalize_text(field_key))
        if actual_key:
            value = master_data.get(actual_key)
            if value is not None and str(value).strip() != "":
                return value

    alias_index = master_data.get("__master_alias_index__", {})
    if isinstance(alias_index, dict):
        candidate_keys = alias_index.get(field_key, [])
        for candidate_key in candidate_keys:
            value = master_data.get(candidate_key)
            if value is not None and str(value).strip() != "":
                return value

        for canonical_key, candidates in alias_index.items():
            if field_key == canonical_key or field_key in candidates:
                for candidate_key in candidates:
                    value = master_data.get(candidate_key)
                    if value is not None and str(value).strip() != "":
                        return value

    return default


def get_master_value_variants(master_data: Dict[str, Any], field_key: str) -> Dict[str, Any]:
    variants = master_data.get(MASTER_VALUE_VARIANTS_KEY, {}) or {}
    if field_key in variants:
        return variants[field_key]

    alias_index = master_data.get("__master_alias_index__", {})
    if isinstance(alias_index, dict):
        for canonical_key, candidate_keys in alias_index.items():
            if field_key == canonical_key or field_key in candidate_keys:
                if canonical_key in variants:
                    return variants[canonical_key]
                for candidate in candidate_keys:
                    if candidate in variants:
                        return variants[candidate]

    return {}


# ---------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------
def load_master_data() -> dict:
    master_file = resolve_master_data_file()
    if not master_file.exists():
        raise FileNotFoundError(f"Master data file not found: {master_file}")

    master_data: Dict[str, Any] = {
        MASTER_PROJECTS_KEY: [],
        MASTER_TEMPLATE_HINTS_KEY: {},
        MASTER_VALUE_VARIANTS_KEY: {},
        MASTER_STANDARD_TEXT_BLOCKS_KEY: {},
    }

    workbook = pd.ExcelFile(master_file)

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(master_file, sheet_name=sheet_name)
        df = _normalize_sheet_columns(df)

        normalized_sheet_name = normalize_text(sheet_name)

        # ---------------------------------------------------------
        # 1. Standard scalar sheets
        # ---------------------------------------------------------
        if "field_key" in df.columns and "value" in df.columns:
            for _, row in df.iterrows():
                field_key = _safe_str(row.get("field_key"))
                value = _safe_value(row.get("value"))

                if field_key:
                    master_data[field_key] = value

        # ---------------------------------------------------------
        # 2. Project master structured sheet
        # ---------------------------------------------------------
        if _sheet_looks_like_project_sheet(sheet_name, df):
            records: List[Dict[str, Any]] = []
            column_index = _normalize_project_columns(df)
            sheet_level_status = _infer_project_status_from_sheet(sheet_name)

            for _, row in df.iterrows():
                record = _build_project_record_from_row(
                    row=row,
                    column_index=column_index,
                    sheet_level_status=sheet_level_status,
                )

                if _row_has_meaningful_data(record):
                    records.append(record)

            if records:
                master_data[MASTER_PROJECTS_KEY].extend(records)

        # ---------------------------------------------------------
        # 3. Optional template hints sheet
        # ---------------------------------------------------------
        if normalized_sheet_name == "template hints":
            hints: Dict[str, Dict[str, Any]] = {}
            required = {"fingerprint", "hint_key", "hint_value"}

            if required.issubset(set(df.columns)):
                for _, row in df.iterrows():
                    fingerprint = _safe_str(row.get("fingerprint"))
                    hint_key = _safe_str(row.get("hint_key"))
                    hint_value = _safe_value(row.get("hint_value"))

                    if fingerprint and hint_key:
                        hints.setdefault(fingerprint, {})
                        hints[fingerprint][hint_key] = hint_value

            master_data[MASTER_TEMPLATE_HINTS_KEY] = hints

        # ---------------------------------------------------------
        # 4. Optional value variants sheet
        # ---------------------------------------------------------
        if normalized_sheet_name == "value variants":
            variants: Dict[str, Dict[str, Any]] = {}

            required = {"field_key", "variant_type", "variant_value"}
            if required.issubset(set(df.columns)):
                for _, row in df.iterrows():
                    field_key = _safe_str(row.get("field_key"))
                    variant_type = _safe_str(row.get("variant_type"))
                    variant_value = _safe_value(row.get("variant_value"))

                    if field_key and variant_type:
                        variants.setdefault(field_key, {})
                        variants[field_key][variant_type] = variant_value

            master_data[MASTER_VALUE_VARIANTS_KEY] = variants

        # ---------------------------------------------------------
        # 5. Optional standard text blocks sheet
        # ---------------------------------------------------------
        if normalized_sheet_name == "standard text blocks":
            text_blocks: Dict[str, Any] = {}

            if {"text_key", "value"}.issubset(set(df.columns)):
                for _, row in df.iterrows():
                    text_key = _safe_str(row.get("text_key"))
                    value = _safe_value(row.get("value"))

                    if text_key:
                        text_blocks[text_key] = value

            elif {"field_key", "value"}.issubset(set(df.columns)):
                for _, row in df.iterrows():
                    field_key = _safe_str(row.get("field_key"))
                    value = _safe_value(row.get("value"))

                    if field_key:
                        text_blocks[field_key] = value

            master_data[MASTER_STANDARD_TEXT_BLOCKS_KEY] = text_blocks

    # deduplicate projects conservatively
    if isinstance(master_data.get(MASTER_PROJECTS_KEY), list):
        deduped_projects: List[Dict[str, Any]] = []
        seen = set()

        for record in master_data[MASTER_PROJECTS_KEY]:
            if not isinstance(record, dict):
                continue

            signature = (
                _safe_str(record.get("project_name")),
                _safe_str(record.get("client")),
                _safe_str(record.get("start_date")),
                _safe_str(record.get("end_date")),
                _safe_str(record.get("value")),
            )
            if signature in seen:
                continue
            seen.add(signature)
            deduped_projects.append(record)

        master_data[MASTER_PROJECTS_KEY] = deduped_projects

    master_data["__normalized_master_index__"] = _build_normalized_master_index(master_data)
    master_data["__master_alias_index__"] = _finalize_master_alias_index(master_data)

    return master_data


def load_synonym_mapping() -> dict:
    synonym_file = resolve_synonym_file()
    if not synonym_file.exists():
        raise FileNotFoundError(f"Synonym mapping file not found: {synonym_file}")

    workbook = pd.ExcelFile(synonym_file)

    target_sheet = "Synonyms" if "Synonyms" in workbook.sheet_names else workbook.sheet_names[0]
    df = pd.read_excel(synonym_file, sheet_name=target_sheet)
    df = _normalize_sheet_columns(df)

    synonyms: Dict[str, Dict[str, Any]] = {}

    for _, row in df.iterrows():
        label_text = _safe_str(row.get("label_text"))
        field_key = _safe_str(row.get("field_key"))

        if not label_text or not field_key:
            continue

        normalized_label = normalize_text(label_text)
        if not normalized_label:
            continue

        section_affinity_raw = _safe_str(row.get("section_affinity"))
        section_affinity = []

        if section_affinity_raw:
            section_affinity = [
                part.strip()
                for part in section_affinity_raw.split(",")
                if part.strip()
            ]

        synonyms[normalized_label] = {
            "field_key": field_key,
            "category": _safe_str(row.get("category")) or None,
            "risk_level": _safe_str(row.get("risk_level")) or None,
            "notes": _safe_str(row.get("notes")) or None,
            "section_affinity": section_affinity,
            "layout_hint": _safe_str(row.get("layout_hint")) or None,
            "match_priority": _safe_str(row.get("match_priority")) or None,
            "value_type": _safe_str(row.get("value_type")) or None,
        }

    return synonyms