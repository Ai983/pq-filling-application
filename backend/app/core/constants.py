"""
Application-wide constants for the Vendor Registration Autofill System.

These constants are intentionally deterministic and conservative.
The generic autofill engine should prefer skipping over unsafe filling.
"""

# -------------------------------------------------------------------
# Review / processing statuses
# -------------------------------------------------------------------

REVIEW_STATUS_FILLED = "FILLED"
REVIEW_STATUS_SKIPPED = "SKIPPED"
REVIEW_STATUS_UNMATCHED = "UNMATCHED"
REVIEW_STATUS_ERROR = "ERROR"
REVIEW_STATUS_PROFILE_FILLED = "PROFILE_FILLED"


# -------------------------------------------------------------------
# Match types
# -------------------------------------------------------------------

MATCH_TYPE_EXACT = "EXACT"
MATCH_TYPE_FUZZY = "FUZZY"
MATCH_TYPE_EMPTY = "EMPTY"
MATCH_TYPE_UNMATCHED = "UNMATCHED"
MATCH_TYPE_SECTION_CONTEXT = "SECTION_CONTEXT"
MATCH_TYPE_TABLE_CONTEXT = "TABLE_CONTEXT"
MATCH_TYPE_YES_NO_RULE = "YES_NO_RULE"
MATCH_TYPE_PROJECT_TABLE = "PROJECT_TABLE"
MATCH_TYPE_TEMPLATE_PROFILE = "TEMPLATE_PROFILE"
MATCH_TYPE_NORMALIZED_EXACT = "NORMALIZED_EXACT"
MATCH_TYPE_TOKEN_CONTAINS = "TOKEN_CONTAINS"
MATCH_TYPE_ALIAS = "ALIAS"


# -------------------------------------------------------------------
# Cell types
# -------------------------------------------------------------------

CELL_TYPE_SECTION_HEADER = "SECTION_HEADER"
CELL_TYPE_SIMPLE_FIELD = "SIMPLE_FIELD"
CELL_TYPE_SUBFIELD = "SUBFIELD"
CELL_TYPE_TABLE_HEADER = "TABLE_HEADER"
CELL_TYPE_TABLE_ROW_LABEL = "TABLE_ROW_LABEL"
CELL_TYPE_INSTRUCTION = "INSTRUCTION"
CELL_TYPE_VALUE = "VALUE"
CELL_TYPE_EMPTY = "EMPTY"
CELL_TYPE_DECLARATION = "DECLARATION"
CELL_TYPE_HEADER_BAND = "HEADER_BAND"
CELL_TYPE_UNKNOWN = "UNKNOWN"


# -------------------------------------------------------------------
# Logical sections
# -------------------------------------------------------------------

SECTION_OWNER = "owner"
SECTION_PROJECT_CONTACT = "project_contact"
SECTION_ACCOUNTS_CONTACT = "accounts_contact"
SECTION_BANKING = "banking"
SECTION_WORKSHOP = "workshop"
SECTION_COMPLIANCE = "compliance"
SECTION_FINANCIAL = "financial"
SECTION_PROJECTS = "projects"
SECTION_COMPANY = "company"
SECTION_TAX = "tax"
SECTION_RESOURCE = "resource"
SECTION_CERTIFICATION = "certification"
SECTION_DECLARATION = "declaration"
SECTION_TECHNICAL = "technical"
SECTION_BILLING = "billing"
SECTION_UNKNOWN = "unknown"


# -------------------------------------------------------------------
# Master data special keys
# -------------------------------------------------------------------

MASTER_PROJECTS_KEY = "__projects__"
MASTER_TEMPLATE_HINTS_KEY = "__template_hints__"
MASTER_VALUE_VARIANTS_KEY = "__value_variants__"
MASTER_STANDARD_TEXT_BLOCKS_KEY = "__standard_text_blocks__"


# -------------------------------------------------------------------
# Layout types / local pattern classification
# -------------------------------------------------------------------

LAYOUT_INLINE_RIGHT = "inline_right"
LAYOUT_INLINE_MERGED_RIGHT = "inline_merged_right"
LAYOUT_BELOW_ANSWER = "below_answer"
LAYOUT_BELOW_MERGED = "below_merged"
LAYOUT_ROW_TABLE = "row_table"
LAYOUT_COLUMN_TABLE = "column_table"
LAYOUT_SECTION_BLOCK = "section_block"
LAYOUT_YES_NO_OPTION = "yes_no_option"
LAYOUT_DECLARATION = "declaration"
LAYOUT_UNKNOWN = "unknown"


# -------------------------------------------------------------------
# Resolver names
# -------------------------------------------------------------------

RESOLVER_RIGHT_CELL = "right_cell_resolver"
RESOLVER_RIGHT_MERGED = "right_merged_resolver"
RESOLVER_BELOW_CELL = "below_cell_resolver"
RESOLVER_ROW_TABLE = "row_table_resolver"
RESOLVER_HEADER_ANSWER_ROW = "header_answer_row_resolver"
RESOLVER_YES_NO = "yes_no_resolver"
RESOLVER_TEMPLATE_PROFILE = "template_profile_resolver"
RESOLVER_UNKNOWN = "unknown_resolver"

RESOLVER_PRIORITY_ORDER = [
    RESOLVER_TEMPLATE_PROFILE,
    RESOLVER_RIGHT_MERGED,
    RESOLVER_RIGHT_CELL,
    RESOLVER_BELOW_CELL,
    RESOLVER_ROW_TABLE,
    RESOLVER_HEADER_ANSWER_ROW,
    RESOLVER_YES_NO,
]


# -------------------------------------------------------------------
# Confidence levels and thresholds
# -------------------------------------------------------------------

CONFIDENCE_HIGH = 0.90
CONFIDENCE_MEDIUM = 0.75
CONFIDENCE_LOW = 0.55
CONFIDENCE_MIN_FILL = 0.72
CONFIDENCE_MIN_PROFILE_FILL = 0.85
CONFIDENCE_MIN_TABLE_FILL = 0.78
CONFIDENCE_MIN_YES_NO_FILL = 0.80


# -------------------------------------------------------------------
# Risk levels
# -------------------------------------------------------------------

RISK_LOW = "low"
RISK_MEDIUM = "medium"
RISK_HIGH = "high"

SAFE_FIELD_RISK_LEVELS = {RISK_LOW, RISK_MEDIUM}


# -------------------------------------------------------------------
# Common placeholder / non-meaningful values
# -------------------------------------------------------------------

PLACEHOLDER_VALUES = {
    "",
    "-",
    "--",
    "---",
    "n/a",
    "na",
    "not available",
    "nil",
    "none",
    ".",
    ":",
    "tbd",
    "to be updated",
    "to be filled",
    "fill here",
    "enter here",
    "details to be provided",
}

PLACEHOLDER_TOKENS = {
    "please fill",
    "fill this",
    "fill here",
    "enter",
    "details",
    "to be filled",
    "to be provided",
    "input",
}


# -------------------------------------------------------------------
# Boolean / yes-no value normalization
# -------------------------------------------------------------------

YES_TOKENS = {"yes", "y", "true", "applicable", "available", "provided"}
NO_TOKENS = {"no", "n", "false", "not applicable", "not available", "not provided"}

YES_NO_REPRESENTATIONS = {
    "yes_no_words": {"yes": "Yes", "no": "No"},
    "y_n": {"yes": "Y", "no": "N"},
    "applicable_not_applicable": {"yes": "Applicable", "no": "Not Applicable"},
    "available_not_available": {"yes": "Available", "no": "Not Available"},
    "provided_not_provided": {"yes": "Provided", "no": "Not Provided"},
}


# -------------------------------------------------------------------
# Template / profile constants
# -------------------------------------------------------------------

PROFILE_MODE_EXACT = "EXACT"
PROFILE_MODE_FALLBACK = "FALLBACK"
PROFILE_MODE_DISABLED = "DISABLED"

PROFILE_FILE_EXTENSION = ".json"


# -------------------------------------------------------------------
# Table types
# -------------------------------------------------------------------

TABLE_TYPE_ROW_LABEL = "row_label"
TABLE_TYPE_COLUMN_HEADER = "column_header"
TABLE_TYPE_PROJECTS = "projects"
TABLE_TYPE_FINANCIAL = "financial"
TABLE_TYPE_RESOURCE = "resource"
TABLE_TYPE_COMPLIANCE = "compliance"
TABLE_TYPE_UNKNOWN = "unknown"


# -------------------------------------------------------------------
# Project ordering / selection modes
# -------------------------------------------------------------------

PROJECT_SELECTION_HIGHEST_VALUE = "highest_value"
PROJECT_SELECTION_NEWEST = "newest"
PROJECT_SELECTION_COMPLETED_FIRST = "completed_first"
PROJECT_SELECTION_ONGOING_FIRST = "ongoing_first"
PROJECT_SELECTION_BALANCED = "balanced"


# -------------------------------------------------------------------
# Common section keyword hints
# -------------------------------------------------------------------

SECTION_KEYWORDS = {
    SECTION_OWNER: [
        "owner",
        "proprietor",
        "director",
        "managing director",
        "md",
        "authorized person",
        "authorised person",
        "contact person",
        "promoter",
        "partner",
    ],
    SECTION_PROJECT_CONTACT: [
        "project contact",
        "project coordinator",
        "project team",
        "site contact",
        "execution contact",
        "technical contact",
    ],
    SECTION_ACCOUNTS_CONTACT: [
        "accounts",
        "accounts contact",
        "billing",
        "billing contact",
        "finance contact",
        "accounts department",
    ],
    SECTION_BANKING: [
        "bank details",
        "banking",
        "bank particulars",
        "bank account",
    ],
    SECTION_COMPLIANCE: [
        "compliance",
        "legal",
        "statutory",
        "statutory details",
        "registration details",
    ],
    SECTION_FINANCIAL: [
        "financial",
        "turnover",
        "annual turnover",
        "financial details",
    ],
    SECTION_PROJECTS: [
        "projects",
        "project credentials",
        "work orders",
        "experience",
        "past projects",
        "completed projects",
        "ongoing projects",
    ],
    SECTION_COMPANY: [
        "company",
        "company details",
        "firm details",
        "organization details",
        "organisation details",
        "general information",
    ],
    SECTION_TAX: [
        "tax",
        "gst",
        "pan",
        "pf",
        "esi",
        "statutory",
        "registration",
    ],
    SECTION_RESOURCE: [
        "resource",
        "resources",
        "manpower",
        "machinery",
        "staff strength",
        "infrastructure",
        "workshop",
    ],
    SECTION_DECLARATION: [
        "declaration",
        "undertaking",
        "certification",
        "authorized signatory",
        "authorised signatory",
    ],
}


# -------------------------------------------------------------------
# Table header hints
# -------------------------------------------------------------------

PROJECT_TABLE_HEADER_KEYWORDS = {
    "project_name": ["project", "name of project", "project name", "work name"],
    "client": ["client", "customer", "principal employer"],
    "location": ["location", "city", "state", "site location"],
    "value": ["value", "contract value", "order value", "project value"],
    "start_date": ["start date", "commencement date"],
    "end_date": ["end date", "completion date"],
    "status": ["status"],
    "contact_name": ["contact", "contact person", "client contact"],
    "contact_phone": ["phone", "mobile", "contact number"],
}

FINANCIAL_ROW_KEYWORDS = {
    "financial.turnover.fy2024_25": ["fy 2024-25", "2024-25", "2024 25"],
    "financial.turnover.fy2023_24": ["fy 2023-24", "2023-24", "2023 24"],
    "financial.turnover.fy2022_23": ["fy 2022-23", "2022-23", "2022 23"],
    "financial.turnover.fy2021_22": ["fy 2021-22", "2021-22", "2021 22"],
}

RESOURCE_ROW_KEYWORDS = {
    "resource.manpower.engineers": ["engineers", "no of engineers", "engineer"],
    "resource.manpower.supervisors": ["supervisors", "no of supervisors", "supervisor"],
    "resource.manpower.skilled_labour": ["skilled labour", "skilled labor"],
    "resource.manpower.unskilled_labour": ["unskilled labour", "unskilled labor"],
    "resource.manpower.total_staff": ["total staff", "total manpower", "total employees"],
    "resource.workshop_available": ["workshop available", "own workshop"],
    "resource.machinery.details": ["machinery", "plant & machinery", "equipment details"],
}


# -------------------------------------------------------------------
# Generic scan / layout boundaries
# -------------------------------------------------------------------

DEFAULT_MAX_RIGHT_SCAN = 6
DEFAULT_MAX_DOWN_SCAN = 4
DEFAULT_MAX_LEFT_SCAN = 2
DEFAULT_MAX_UP_SCAN = 2

MAX_LABEL_LENGTH = 120
MAX_DECLARATION_LENGTH = 500
MIN_SECTION_HEADER_LENGTH = 3


# -------------------------------------------------------------------
# Skip reasons
# -------------------------------------------------------------------

SKIP_REASON_NO_MATCH = "no_match"
SKIP_REASON_LOW_CONFIDENCE = "low_confidence"
SKIP_REASON_NO_TARGET = "no_target"
SKIP_REASON_AMBIGUOUS_TARGET = "ambiguous_target"
SKIP_REASON_EXISTING_VALUE = "existing_value"
SKIP_REASON_FORMULA_CELL = "formula_cell"
SKIP_REASON_PROTECTED_CELL = "protected_cell"
SKIP_REASON_UNSAFE_MERGED_TARGET = "unsafe_merged_target"
SKIP_REASON_UNSUPPORTED_LAYOUT = "unsupported_layout"
SKIP_REASON_UNSUPPORTED_DECLARATION = "unsupported_declaration"
SKIP_REASON_VALUE_NOT_FOUND = "value_not_found"
SKIP_REASON_TABLE_NOT_FOUND = "table_not_found"
SKIP_REASON_PROFILE_NOT_FOUND = "profile_not_found"
SKIP_REASON_EXCEPTION = "exception"


# -------------------------------------------------------------------
# Safe write result codes
# -------------------------------------------------------------------

WRITE_RESULT_FILLED = "filled"
WRITE_RESULT_SKIPPED_EXISTING = "skipped_existing_value"
WRITE_RESULT_SKIPPED_FORMULA = "skipped_formula"
WRITE_RESULT_SKIPPED_PROTECTED = "skipped_protected"
WRITE_RESULT_SKIPPED_AMBIGUOUS = "skipped_ambiguous_target"
WRITE_RESULT_SKIPPED_CONFLICT = "skipped_conflict"
WRITE_RESULT_SKIPPED_INVALID = "skipped_invalid_target"
WRITE_RESULT_FAILED = "failed"


# -------------------------------------------------------------------
# Canonical field families
# -------------------------------------------------------------------

FIELD_FAMILY_COMPANY = "company"
FIELD_FAMILY_CONTACT = "contact"
FIELD_FAMILY_TAX = "tax"
FIELD_FAMILY_BANK = "bank"
FIELD_FAMILY_COMPLIANCE = "compliance"
FIELD_FAMILY_FINANCIAL = "financial"
FIELD_FAMILY_RESOURCE = "resource"
FIELD_FAMILY_PROJECT = "project"
FIELD_FAMILY_TEXT = "text"
FIELD_FAMILY_UNKNOWN = "unknown"


# -------------------------------------------------------------------
# Fields that are usually safe in generic mode
# -------------------------------------------------------------------

GENERIC_SAFE_FIELDS = {
    "company.legal_name",
    "company.entity_type",
    "company.incorporation_year",
    "company.business_type",
    "company.address",
    "company.phone",
    "company.email",
    "company.website",
    "contact.owner.name",
    "contact.owner.designation",
    "contact.owner.mobile",
    "contact.owner.email",
    "contact.project.name",
    "contact.project.designation",
    "contact.project.mobile",
    "contact.project.email",
    "contact.accounts.name",
    "contact.accounts.designation",
    "contact.accounts.mobile",
    "contact.accounts.email",
    "tax.pan",
    "tax.gst.primary",
    "tax.pf",
    "tax.esi",
    "tax.msme",
    "bank.name",
    "bank.account_number",
    "bank.ifsc",
    "bank.branch",
    "compliance.iso_certified",
    "compliance.quality_policy",
    "compliance.safety_policy",
    "compliance.ohs_policy",
    "compliance.litigation",
    "compliance.arbitration",
    "compliance.audit_reports",
    "financial.turnover.fy2024_25",
    "financial.turnover.fy2023_24",
    "financial.turnover.fy2022_23",
    "financial.turnover.fy2021_22",
    "resource.manpower.engineers",
    "resource.manpower.supervisors",
    "resource.manpower.skilled_labour",
    "resource.manpower.unskilled_labour",
    "resource.manpower.total_staff",
    "resource.machinery.details",
    "resource.workshop_available",
}


# -------------------------------------------------------------------
# Fields that should usually be profile-based or carefully handled
# -------------------------------------------------------------------

PROFILE_PREFERRED_FIELDS = {
    "company.registration_number",
    "company.nature_of_business",
    "company.core_capabilities",
    "declaration.authorized_signatory_name",
    "declaration.place",
    "declaration.date",
    "declaration.signature",
}


# -------------------------------------------------------------------
# Normalization stopwords for labels
# -------------------------------------------------------------------

LABEL_NOISE_TOKENS = {
    "please",
    "kindly",
    "enter",
    "mention",
    "provide",
    "details",
    "detail",
    "of",
    "the",
    "a",
    "an",
}


# -------------------------------------------------------------------
# Common file / sheet naming hints
# -------------------------------------------------------------------

LIKELY_TEMPLATE_SHEET_NAMES = {
    "vendor registration",
    "vendor form",
    "pq form",
    "prequalification",
    "pre-qualification",
    "company credential",
    "credentials",
}


# -------------------------------------------------------------------
# Review log column order
# -------------------------------------------------------------------

REVIEW_LOG_COLUMNS = [
    "sheet_name",
    "label_cell",
    "label_text",
    "normalized_label",
    "cell_type",
    "section",
    "field_key",
    "match_type",
    "semantic_confidence",
    "layout_type",
    "resolver",
    "target_cell",
    "target_merged_range",
    "value_preview",
    "write_result",
    "status",
    "reason",
]