from pathlib import Path

# ---------------------------------------------------------
# Base paths
# ---------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent.parent

APP_NAME = "Vendor Autofill Backend"
APP_VERSION = "0.2.0"


# ---------------------------------------------------------
# Core directories
# ---------------------------------------------------------

DATA_DIR = BASE_DIR / "data"
MASTER_DIR = DATA_DIR / "master"
UPLOAD_DIR = DATA_DIR / "uploads"
PROCESSED_DIR = DATA_DIR / "processed"
LOG_DIR = DATA_DIR / "logs"

ENGINE_DIR = BASE_DIR / "app" / "engine"
TEMPLATE_PROFILE_DIR = ENGINE_DIR / "template_profiles"


# ---------------------------------------------------------
# Master data files
# ---------------------------------------------------------

# Keep fallback names so the app remains compatible if old files exist.
MASTER_DATA_FILE = MASTER_DIR / "master_data.xlsx"
SYNONYM_FILE = MASTER_DIR / "synonym_mapping.xlsx"

# Preferred aligned files uploaded by you.
MASTER_DATA_FILE_ALIGNED = MASTER_DIR / "master_data_FINAL_aligned.xlsx"
SYNONYM_FILE_ALIGNED = MASTER_DIR / "synonym_mapping_FINAL_aligned.xlsx"


# ---------------------------------------------------------
# Upload / processing settings
# ---------------------------------------------------------

ALLOWED_EXTENSIONS = {".xlsx"}
MAX_UPLOAD_SIZE_MB = 20
MAX_UPLOAD_SIZE_BYTES = MAX_UPLOAD_SIZE_MB * 1024 * 1024

# Preserve workbook layout and avoid risky actions.
PRESERVE_EXISTING_STYLES = True
PRESERVE_MERGED_CELLS = True
ALLOW_OVERWRITE_EXISTING_VALUES = False
ALLOW_PROFILE_OVERWRITE_PLACEHOLDERS_ONLY = True
SKIP_FORMULA_CELLS = True
SKIP_PROTECTED_LIKE_CELLS = True


# ---------------------------------------------------------
# Matching thresholds
# ---------------------------------------------------------

# Keep fuzzy threshold relatively high because wrong autofill is worse than skip.
FUZZY_MATCH_THRESHOLD = 88
NORMALIZED_EXACT_MATCH_SCORE = 100
TOKEN_MATCH_MIN_SCORE = 90

# Final fill confidence thresholds
MIN_SEMANTIC_CONFIDENCE = 0.72
MIN_LAYOUT_CONFIDENCE = 0.70
MIN_TOTAL_CONFIDENCE_TO_FILL = 0.72
MIN_TOTAL_CONFIDENCE_FOR_TABLE_FILL = 0.78
MIN_TOTAL_CONFIDENCE_FOR_YES_NO_FILL = 0.80
MIN_TOTAL_CONFIDENCE_FOR_PROFILE_FILL = 0.85


# ---------------------------------------------------------
# Layout scanning limits
# ---------------------------------------------------------

MAX_RIGHT_SCAN = 6
MAX_DOWN_SCAN = 4
MAX_LEFT_SCAN = 2
MAX_UP_SCAN = 2

# Candidate extraction limits
MAX_LABEL_TEXT_LENGTH = 120
MAX_DECLARATION_TEXT_LENGTH = 500
MIN_SECTION_HEADER_TEXT_LENGTH = 3

# Table and block inference
MIN_TABLE_HEADER_MATCHES = 3
MIN_ROW_TABLE_GROUP_SIZE = 2
MAX_HEADER_TO_ANSWER_ROW_GAP = 1
MAX_SECTION_LOOKBACK_ROWS = 8


# ---------------------------------------------------------
# Template profile settings
# ---------------------------------------------------------

ENABLE_TEMPLATE_PROFILES = True
ENABLE_GENERIC_FALLBACK = True
TEMPLATE_PROFILE_EXTENSION = ".json"

# Fingerprint tuning
FINGERPRINT_TOP_TEXT_CELL_COUNT = 25
FINGERPRINT_REQUIRED_LABEL_MATCHES = 2


# ---------------------------------------------------------
# Project table fill settings
# ---------------------------------------------------------

DEFAULT_PROJECT_ROWS_TO_FILL = 5
PROJECT_SELECTION_MODE = "balanced"

# Possible values supported later:
# - balanced
# - highest_value
# - newest
# - completed_first
# - ongoing_first


# ---------------------------------------------------------
# Review log settings
# ---------------------------------------------------------

REVIEW_LOG_FILE_SUFFIX = "_review_log.xlsx"
PROCESSED_FILE_SUFFIX = "_filled.xlsx"


# ---------------------------------------------------------
# Runtime helper functions
# ---------------------------------------------------------

def resolve_master_data_file() -> Path:
    """
    Prefer aligned master file if available, otherwise fall back
    to the old default file path.
    """
    if MASTER_DATA_FILE_ALIGNED.exists():
        return MASTER_DATA_FILE_ALIGNED
    return MASTER_DATA_FILE


def resolve_synonym_file() -> Path:
    """
    Prefer aligned synonym file if available, otherwise fall back
    to the old default file path.
    """
    if SYNONYM_FILE_ALIGNED.exists():
        return SYNONYM_FILE_ALIGNED
    return SYNONYM_FILE


# ---------------------------------------------------------
# Ensure required directories exist
# ---------------------------------------------------------

for path in [
    DATA_DIR,
    MASTER_DIR,
    UPLOAD_DIR,
    PROCESSED_DIR,
    LOG_DIR,
    TEMPLATE_PROFILE_DIR,
]:
    path.mkdir(parents=True, exist_ok=True)