"""
Configuration settings for the SharePoint Automation
"""
import os

# Base directories - FIXED to point to project root, not src
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Go up from src/ to project root
DATA_DIR = os.path.join(BASE_DIR, "data")

# File paths
USER_PROFILE = os.environ.get('USERPROFILE', '')
SYNCED_FILE_PATH = os.path.join(USER_PROFILE, 'DPDHL', 'SM Team - SG - AD EDS, MFA, GSN VS AD, GSN VS ER Weekly Report', 'Weekly Report 2025 - Copy.xlsx')

# Data files - Now correctly pointing to project_root/data/
GSN_DATA_FILE = os.path.join(DATA_DIR, "gsn_data.json")
AD_RESULTS_FILE = os.path.join(DATA_DIR, "ad_results.json")
AD_COMPARISON_FILE = os.path.join(DATA_DIR, "ad_comparison_results.json")

# OneDrive sync paths - will be checked in order
ONEDRIVE_PATHS = [
    os.path.join(USER_PROFILE, 'OneDrive - Deutsche Post DHL', 'Shared Documents', 'testteam372', 'Shared Documents'),
    os.path.join(USER_PROFILE, 'OneDrive - Deutsche Post DHL', 'Teams', 'testteam372', 'Shared Documents'),
    os.path.join(USER_PROFILE, 'OneDrive - Deutsche Post DHL', 'testteam372', 'Shared Documents'),
    os.path.join(USER_PROFILE, 'OneDrive', 'Shared Documents', 'testteam372'),
    os.path.join(USER_PROFILE, 'OneDrive', 'testteam372', 'Shared Documents')
]

# Search patterns for files
FILE_PATTERNS = {
    'gsn': 'alm_hardware*',  # This will match alm_hardware.xlsx, alm_hardware (48).xlsx, etc.
    'er': 'data*'            # This will match data.xlsx, data(2).xlsx, data(45).xlsx, etc.
}

# Active Directory search parameters
AD_SEARCH = {
    'ldap_filter': "(&(&(objectCategory=computer)(objectClass=computer)(&(cn=SG*)(!cn=SGD*)(!cn=SGG*)(!cn=SGSAH*)(!cn=SGSI*)(!cn=SGSR*)(!cn=SGT*))))",
    'search_base': "OU=SCO,OU=EXP,OU=SG,OU=Prod,OU=Computers,OU=NGWS,DC=kul-dc,DC=dhl,DC=com"
}

# Colors for console output (ANSI color codes)
COLORS = {
    'WHITE': '\033[37m',
    'YELLOW': '\033[33m',
    'GREEN': '\033[32m',
    'CYAN': '\033[36m',
    'MAGENTA': '\033[35m',
    'RED': '\033[31m',
    'RESET': '\033[0m'
}