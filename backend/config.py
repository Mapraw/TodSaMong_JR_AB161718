import os
from pathlib import Path
from .shared.standard_paths import category_standard_path

# Base Paths
BACKEND_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BACKEND_DIR.parent
BASE_DIR = str(PROJECT_DIR)
DATABASE_DIR = os.getenv("TODSAMONG_DATABASE_DIR", "")

# Optional external data paths. The app still runs when these are unset and
# users can upload price schedules directly from the UI.
MASTER_DATA_PATH = os.getenv("TODSAMONG_MASTER_DATA_PATH", "")
PRICE_SCHEDULE_DEFAULT_PATH = os.getenv("TODSAMONG_PRICE_SCHEDULE_PATH", "")
MFR_LIST_PATH = os.getenv(
    "TODSAMONG_MFR_LIST_PATH",
    category_standard_path("ab18", "Manufacturer List (MAY2026).xls"),
)

# Category Identification Cells
IDENTITY_CELLS = {
    "AB16_TERM": {"header_cell": "A1", "header_text": "Cable terminations", "id_cell": "K5"},
    "AB16_CLEAT": {"header_cell": "A1", "header_text": "Cable Cleats", "id_cell": "K6"},
    "AB17": {"header_cell": "A1", "header_text": "XLPE Power cable", "id_cell": "K5"},
    "AB18": {"header_cell": "A2", "header_text": "LOW VOLTAGE CABLE AND CONDUCTOR", "id_cell": "L8"},
}

# Common Proposal Data Mappings
COMMON_CELLS = {
    "procurement_ref": "F8",
    "bidder": "C11",
    "schedule_no": "L8",
    "item_no": "L11",
}

# Country Normalization
MASTER_COUNTRIES = {
    "TH": "THAILAND", "THA": "THAILAND", "THAI": "THAILAND", "THAIALND":"THAILAND",
    "CN": "CHINA", "CHN": "CHINA", "PRC": "CHINA",
    "IN": "INDIA", "IND": "INDIA",
    "JP": "JAPAN", "JPN": "JAPAN",
    "KR": "KOREA", "KOR": "KOREA", "SOUTH KOREA": "KOREA",
    "VN": "VIETNAM", "VNM": "VIETNAM",
}

# Items to exclude from summaries and missing checks
UNWANTED_TERMS = [
    "TOTAL PRICE", "SUMMARY", "TRANSPORTATION", "GRAND TOTAL", 
    "TOTAL PRICE FOR SCHEDULE", "CONSTRUCTION AND INSTALLATION",
    "PRICE OF SCHEDULE", "PRICE FOR SCHEDULE"
]

# Fuzzy Thresholds
BIDDER_MATCH_THRESHOLD = 95  # Increased from 85 to prevent over-merging (e.g. Siemens vs Precise)
MFR_MATCH_THRESHOLD = 85

# Manual Bidder Aliases (Optional explicit mapping)
BIDDER_ALIASES = {
    "BENYHAPA": "BENYAPHA",
}

