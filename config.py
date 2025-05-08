"""
Configuration settings for the SharePoint Automation
"""
import os

# File paths
USER_PROFILE = os.environ.get('USERPROFILE', '')
SYNCED_FILE_PATH = os.path.join(USER_PROFILE, 'DPDHL', 'SM Team - SG - AD EDS, MFA, GSN VS AD, GSN VS ER Weekly Report', 'Weekly Report 2025 - Copy.xlsx')

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
    'gsn': 'alm_hardware*',
    'er': 'data*'
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