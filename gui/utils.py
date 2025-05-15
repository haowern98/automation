"""
Python utilities for the GUI components of SharePoint Automation
"""
from .settings_dialog import show_settings_dialog

# Function to open settings dialog
def open_settings():
    """
    Open the settings dialog
    
    Returns:
        bool: True if settings were saved, False otherwise
    """
    return show_settings_dialog()