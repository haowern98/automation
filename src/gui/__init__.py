"""
GUI Package

Graphical user interface components for the SharePoint Automation application.

This package is organized into:
- tabs/: Individual tab components
- widgets/: Reusable GUI widgets
- Main application dialogs and utilities
"""

# Import main application functions
from .tabbed_app import show_tabbed_date_range_selection, show_enhanced_date_range_selection, parse_date_range_string
from .date_selector import show_date_range_selection, parse_date_range_string as parse_date_range_string_legacy
from .settings_dialog import show_settings_dialog, get_settings, SettingsManager

# Import tab classes for direct access if needed
from .tabs import DateRangeTab, DateRangeResult, SettingsTab, WeeklyReportTab

# Import utility functions
from .utils import open_settings

__all__ = [
    # Main application functions
    'show_tabbed_date_range_selection',
    'show_enhanced_date_range_selection', 
    'show_date_range_selection',
    'parse_date_range_string',
    'parse_date_range_string_legacy',
    
    # Settings functions
    'show_settings_dialog',
    'get_settings',
    'SettingsManager',
    
    # Tab components
    'DateRangeTab',
    'DateRangeResult',
    'SettingsTab', 
    'WeeklyReportTab',
    
    # Utilities
    'open_settings'
]