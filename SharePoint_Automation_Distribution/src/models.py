"""
Shared data models and structures for SharePoint Automation
"""
import datetime


class DateRangeResult:
    """Class to hold date range selection results"""
    
    def __init__(self, start_date=None, end_date=None, formatted_date=""):
        """
        Initialize date range result
        
        Args:
            start_date (datetime.date): Start date
            end_date (datetime.date): End date
            formatted_date (str): Formatted date range string
        """
        self.start_date = start_date
        self.end_date = end_date
        self.date_range_formatted = formatted_date
        self.year = end_date.year if end_date else ""
        self.short_date_range = ""  # Kept for compatibility
        self.cancelled = False  # Flag to indicate if the dialog was cancelled
        self.user_terminated = False  # Flag to indicate if user terminated entire process
        self.use_auto_date = False  # Flag to indicate if user chose auto date
    
    @property
    def is_valid(self):
        """Check if the date range is valid"""
        return self.start_date is not None and self.end_date is not None