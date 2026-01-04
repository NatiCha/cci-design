"""
Configuration constants and environment setup.
"""

import os
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

# =============================================================================
# PATHS
# =============================================================================

PROJECT_ROOT = Path(__file__).parent.parent.parent
DB_PATH = PROJECT_ROOT / "data" / "db" / "cci-timesheets.db"
OUTPUT_DIR = PROJECT_ROOT / "output"
TEMPLATES_DIR = PROJECT_ROOT / "data" / "templates"

# =============================================================================
# EMAIL CONFIGURATION
# =============================================================================

FROM_EMAIL = "charles@landslidelogic.com"
TO_EMAIL = "charles@landslidelogic.com"
ERROR_EMAIL = "charles.squeri@gmail.com"

# =============================================================================
# CALENDAR CONFIGURATION
# =============================================================================

CALENDAR_PATTERN = " TIME CARD"  # e.g., "CES TIME CARD"

# =============================================================================
# VALIDATION CODES
# =============================================================================

VALID_PHASE_CODES = {"PD", "SD", "DD", "CD", "CA", "M", "NA"}
VALID_TASK_CODES = {"BD", "DP", "PM", "3-D", "D-D", "M", "NA"}

# Non-project names (matched case-insensitively by prefix before ':')
NON_PROJECT_NAMES = {"office", "vacation", "holiday", "sick", "personal time"}

# =============================================================================
# REPORT CONFIGURATION
# =============================================================================

DETAIL_HEADERS = ["Project ID", "Date", "Employee", "Hours", "Task", "Phase", "WID"]
DETAIL_HEADERS_EDIT = [
    "Project ID", "Date", "Employee", "Hours",
    "Hours Adjusted", "Total Adjusted Hours",
    "Task", "Phase", "WID"
]
MIN_TABLE_ROWS = 12  # Minimum rows for better display in Numbers when few events

BILLABLE_GOALS_ROW_LABELS = [
    "Billable goal",
    "Gross billable hours adjusted",
    "% of goal",
    "Non-project hours",
    "Net billable hours this period"
]

# =============================================================================
# MS GRAPH CREDENTIALS (from environment)
# =============================================================================

GRAPH_TENANT_ID = os.environ.get("MICROSOFT_GRAPH_TENANT_ID", "")
GRAPH_APP_ID = os.environ.get("MICROSOFT_GRAPH_APP_ID", "")
GRAPH_CLIENT_SECRET = os.environ.get("MICROSOFT_GRAPH_CLIENT_SECRET", "")

# =============================================================================
# API CONFIGURATION
# =============================================================================

CCI_API_KEY = os.environ.get("CCI_API_KEY", "")
API_HOST = os.environ.get("API_HOST", "0.0.0.0")
API_PORT = int(os.environ.get("API_PORT", "8000"))
API_DEBUG = os.environ.get("API_DEBUG", "false").lower() == "true"
MAX_UPLOAD_SIZE_MB = int(os.environ.get("MAX_UPLOAD_SIZE_MB", "50"))
MAX_UPLOAD_SIZE_BYTES = MAX_UPLOAD_SIZE_MB * 1024 * 1024
API_VERSION = "1.0.0"
