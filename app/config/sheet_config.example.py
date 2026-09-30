"""Copy this file to sheet_config.py (git-ignored) and paste the sheet link.

Use the CSV export form of the tab that holds the patient list:
    https://docs.google.com/spreadsheets/d/<SHEET_ID>/export?format=csv&gid=<TAB_GID>
The sheet contains patient details: never commit the real link.
"""
SHEET_CSV_URL = ""

# Preferred: the Apps Script web app URL (ends in /exec), from deploying
# tools/apps_script/Code.gs. When set, it is used instead of SHEET_CSV_URL and
# the sheet can be made private.
APPS_SCRIPT_URL = ""

# Fallback when APPS_SCRIPT_URL isn't set (the script returns the QC values
# itself): the sequencing QC tab (Anderson_ID / After Data / Q30), CSV export:
#     https://docs.google.com/spreadsheets/d/<SHEET_ID>/export?format=csv&gid=<TAB_GID>
# Fills the Sequence data attributes on Fetch. The tab must be shared as
# "anyone with the link".
QC_CSV_URL = ""
