"""
Date Range Selection Tab

Tab for selecting date ranges for SharePoint automation reports.
"""
import datetime
from PyQt5.QtWidgets import QWidget, QGridLayout, QLabel, QDateEdit, QLineEdit
from PyQt5.QtCore import QDate
from PyQt5.QtGui import QFont


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


class DateRangeTab(QWidget):
    """Date range selection tab"""
    
    def __init__(self, parent=None, manual_mode=False):
        """
        Initialize the date range selector tab
        
        Args:
            parent: Parent widget
            manual_mode (bool): Whether the application is running in manual mode
        """
        super(DateRangeTab, self).__init__(parent)
        
        self.manual_mode = manual_mode
        
        # Create the result object
        self.result_obj = DateRangeResult()
        
        # Create layout
        layout = QGridLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Create widgets
        if manual_mode:
            title_text = "Please select the start and end dates for your report:"
        else:
            title_text = "Select your date range or choose an option below:"
        
        title_label = QLabel(title_text)
        title_font = QFont("Segoe UI", 10)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setWordWrap(True)
        
        start_date_label = QLabel("Start Date:")
        self.start_date_picker = QDateEdit()
        self.start_date_picker.setCalendarPopup(True)
        self.start_date_picker.setDate(QDate.currentDate())
        
        end_date_label = QLabel("End Date:")
        self.end_date_picker = QDateEdit()
        self.end_date_picker.setCalendarPopup(True)
        self.end_date_picker.setDate(QDate.currentDate().addDays(1))
        
        preview_label = QLabel("Preview:")
        self.preview_textbox = QLineEdit()
        self.preview_textbox.setReadOnly(True)
        
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: red;")
        
        # Add widgets to layout
        layout.addWidget(title_label, 0, 0, 1, 2)
        layout.addWidget(start_date_label, 1, 0)
        layout.addWidget(self.start_date_picker, 1, 1)
        layout.addWidget(end_date_label, 2, 0)
        layout.addWidget(self.end_date_picker, 2, 1)
        layout.addWidget(preview_label, 3, 0)
        layout.addWidget(self.preview_textbox, 3, 1)
        layout.addWidget(self.status_label, 4, 0, 1, 2)
        
        # Set spacing
        layout.setVerticalSpacing(10)
        
        # Connect signals
        self.start_date_picker.dateChanged.connect(self.update_preview)
        self.end_date_picker.dateChanged.connect(self.update_preview)
        
        # Initial preview update
        self.update_preview()
    
    def update_preview(self):
        """Update the date range preview"""
        start_date = self.start_date_picker.date().toPyDate()
        end_date = self.end_date_picker.date().toPyDate()
        
        # Validate dates
        if end_date < start_date:
            self.status_label.setText("End date cannot be earlier than start date.")
            return
        else:
            self.status_label.setText("")
        
        # Format the date range
        if start_date.month == end_date.month and start_date.year == end_date.year:
            date_range_formatted = f"{start_date.day}-{end_date.day} {start_date.strftime('%B')} {start_date.year}"
        else:
            date_range_formatted = f"{start_date.day} {start_date.strftime('%B')} - {end_date.day} {end_date.strftime('%B')} {end_date.year}"
        
        # Update preview
        self.preview_textbox.setText(date_range_formatted)
        
        # Update result object
        self.result_obj.start_date = start_date
        self.result_obj.end_date = end_date
        self.result_obj.date_range_formatted = date_range_formatted
        self.result_obj.year = str(end_date.year)