"""
Logging utilities for the SharePoint Automation
"""
import datetime
from config import COLORS

def write_log(message, color="WHITE"):
    """
    Write a log message with timestamp and color
    
    Args:
        message (str): Message to log
        color (str): Color name as defined in config.COLORS
    """
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    color_code = COLORS.get(color.upper(), COLORS['WHITE'])
    print(f"{color_code}[{timestamp}] {message}{COLORS['RESET']}")