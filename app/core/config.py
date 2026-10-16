"""
Centralized project configuration and root directory resolution.
"""

from pathlib import Path

# Points directly to the project root directory
ROOT_DIR = Path(__file__).resolve().parents[2]

# Standard runtime directories created directly in the root directory
UPLOADS_DIR = ROOT_DIR / "uploads"
LOGS_DIR = ROOT_DIR / "logs"
ERRORS_DIR = ROOT_DIR / "errors"
AUTOMATION_PROFILE_DIR = ROOT_DIR / "chrome_automation_profile"

# Ensure core operational directories exist immediately
LOGS_DIR.mkdir(parents=True, exist_ok=True)
UPLOADS_DIR.mkdir(parents=True, exist_ok=True)

# Subdirectories inside uploads/
NDC_REPORTS_DIR = UPLOADS_DIR / "NDC_Reports"
FF_REPORTS_DIR = UPLOADS_DIR / "FF_Reports"
ELCM_REPORTS_DIR = UPLOADS_DIR / "ELCM_Reports"
PREJOINING_REPORTS_DIR = UPLOADS_DIR / "Pre_Joining_Reports"

# Subdirectories inside errors/
NDC_ERRORS_DIR = ERRORS_DIR / "NDC_Errors"
FF_ERRORS_DIR = ERRORS_DIR / "FF_Errors"
ELCM_ERRORS_DIR = ERRORS_DIR / "ELCM_Errors"
PREJOINING_ERRORS_DIR = ERRORS_DIR / "Pre_Joining_Errors"

# Helper function to create directories on-demand only when needed
def ensure_dir(directory: Path) -> Path:
    """Ensure a directory exists on-demand and return its path."""
    directory.mkdir(parents=True, exist_ok=True)
    return directory

# ==== Centralized Scheduler Timings (12-hour AM/PM or 24-hour format) ====
# Set direct schedule timings here:
# NDC: 10 AM, 1 PM, 3 PM, 5 PM, 7 PM
NDC_SCHEDULE_TIMES = ["10:00 AM", "1:00 PM", "3:00 PM", "5:00 PM", "7:00 PM"]

# F&F: 11 AM, 4 PM
FNF_SCHEDULE_TIMES = ["11:00 AM", "4:00 PM"]

# ELCM: 10:30 AM
ELCM_SCHEDULE_TIMES = ["10:30 AM"]
