"""
GUI Tabs Package

Individual tab components for the SharePoint Automation application.
Updated to use the new DateRangeTab with integrated buttons.
"""

# Import from the updated files
from .date_range_tab import DateRangeTab, DateRangeResult
from .settings_tab import SettingsTab
from .weekly_report_tab import WeeklyReportTab

__all__ = [
    'DateRangeTab',
    'DateRangeResult', 
    'SettingsTab',
    'WeeklyReportTab'
]