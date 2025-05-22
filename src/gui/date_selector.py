"""
Date range selection dialog for SharePoint Automation
"""
import sys
import datetime
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout, 
                             QLabel, QDateEdit, QLineEdit, QPushButton,
                             QGridLayout, QMessageBox)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor
from .settings_dialog import show_settings_dialog

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
        self.cancelled = False  # New flag to indicate if the dialog was cancelled
    
    @property
    def is_valid(self):
        """Check if the date range is valid"""
        return self.start_date is not None and self.end_date is not None

class DateRangeSelector(QDialog):
    """Date range selection dialog"""
    
    def __init__(self, parent=None, manual_mode=False):
        """
        Initialize the date range selector dialog
        
        Args:
            parent: Parent widget
            manual_mode (bool): Whether the application is running in manual mode
        """
        super(DateRangeSelector, self).__init__(parent)
        
        self.manual_mode = manual_mode
        
        self.setWindowTitle("SharePoint Automation - Date Range Selection")
        self.setFixedSize(460, 360)  # Increased height for settings button
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        # Create the result object
        self.result_obj = DateRangeResult()
        
        # Create layout
        layout = QGridLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Create widgets
        title_label = QLabel("Please select the start and end dates for your report:")
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
        
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        
        # Set cancel button text based on mode
        cancel_text = "Exit" if manual_mode else "Use Auto Date"
        self.cancel_button = QPushButton(cancel_text)
        
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.setAlignment(Qt.AlignCenter)
        
        # Create settings button layout
        settings_layout = QHBoxLayout()
        self.settings_button = QPushButton("Settings")
        self.settings_button.setIcon(self.style().standardIcon(self.style().SP_FileDialogInfoView))
        settings_layout.addWidget(self.settings_button)
        settings_layout.setAlignment(Qt.AlignRight)
        
        # Add widgets to layout
        layout.addWidget(title_label, 0, 0, 1, 2)
        layout.addWidget(start_date_label, 1, 0)
        layout.addWidget(self.start_date_picker, 1, 1)
        layout.addWidget(end_date_label, 2, 0)
        layout.addWidget(self.end_date_picker, 2, 1)
        layout.addWidget(preview_label, 3, 0)
        layout.addWidget(self.preview_textbox, 3, 1)
        layout.addWidget(self.status_label, 4, 0, 1, 2)
        layout.addLayout(button_layout, 5, 0, 1, 2)
        layout.addLayout(settings_layout, 6, 0, 1, 2)  # Add settings button at the bottom
        
        # Set spacing
        layout.setVerticalSpacing(10)
        
        # Connect signals
        self.start_date_picker.dateChanged.connect(self.update_preview)
        self.end_date_picker.dateChanged.connect(self.update_preview)
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.handle_cancel)
        self.settings_button.clicked.connect(self.open_settings)
        
        # Initial preview update
        self.update_preview()
    
    def update_preview(self):
        """Update the date range preview"""
        start_date = self.start_date_picker.date().toPyDate()
        end_date = self.end_date_picker.date().toPyDate()
        
        # Validate dates
        if end_date < start_date:
            self.status_label.setText("End date cannot be earlier than start date.")
            self.ok_button.setEnabled(False)
            return
        else:
            self.status_label.setText("")
            self.ok_button.setEnabled(True)
        
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
    
    def handle_cancel(self):
        """Handle the cancel button click differently based on mode"""
        if self.manual_mode:
            # In manual mode, exit the application
            self.result_obj.cancelled = True
            self.reject()
        else:
            # In auto mode, just reject the dialog
            # The main program will use automatic date calculation
            self.result_obj.cancelled = True
            self.reject()
    
    def open_settings(self):
        """Open the settings dialog"""
        settings_saved = show_settings_dialog()
        if settings_saved:
            # In a real implementation, you would reload settings here
            pass

def show_date_range_selection(manual_mode=False):
    """
    Show the date range selection dialog
    
    Args:
        manual_mode (bool): Whether the application is running in manual mode
        
    Returns:
        DateRangeResult: The selected date range or a result object with cancelled=True if cancelled
    """
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    dialog = DateRangeSelector(manual_mode=manual_mode)
    result = dialog.exec_() == QDialog.Accepted
    
    # Always return the result object, even when cancelled
    return dialog.result_obj

def parse_date_range_string(date_range_string):
    """
    Parse a date range string into a DateRangeResult object
    
    Args:
        date_range_string (str): The date range string to parse
        
    Returns:
        DateRangeResult: The parsed date range
    """
    result = DateRangeResult()
    result.date_range_formatted = date_range_string
    
    try:
        # Pattern 1: "15-17 April 2025"
        import re
        pattern1 = r'(\d+)-(\d+)\s+([A-Za-z]+)\s+(\d{4})'
        pattern2 = r'(\d+)\s+([A-Za-z]+)\s+-\s+(\d+)\s+([A-Za-z]+)\s+(\d{4})'
        
        match1 = re.match(pattern1, date_range_string)
        match2 = re.match(pattern2, date_range_string)
        
        if match1:
            start_day = int(match1.group(1))
            end_day = int(match1.group(2))
            month = match1.group(3)
            year = match1.group(4)
            
            # Create datetime objects
            start_date = datetime.datetime.strptime(f"{start_day} {month} {year}", "%d %B %Y").date()
            end_date = datetime.datetime.strptime(f"{end_day} {month} {year}", "%d %B %Y").date()
            
            # Set result properties
            result.start_date = start_date
            result.end_date = end_date
            result.year = year
            result.short_date_range = f"{start_day}-{end_day} {start_date.strftime('%b')}"
            return result
            
        elif match2:
            start_day = int(match2.group(1))
            start_month = match2.group(2)
            end_day = int(match2.group(3))
            end_month = match2.group(4)
            year = match2.group(5)
            
            # Create datetime objects
            start_date = datetime.datetime.strptime(f"{start_day} {start_month} {year}", "%d %B %Y").date()
            end_date = datetime.datetime.strptime(f"{end_day} {end_month} {year}", "%d %B %Y").date()
            
            # Set result properties
            result.start_date = start_date
            result.end_date = end_date
            result.year = year
            result.short_date_range = f"{start_day} {start_date.strftime('%b')} - {end_day} {end_date.strftime('%b')}"
            return result
            
        else:
            return None
            
    except Exception as e:
        print(f"Error parsing date range string: {str(e)}")
        return None

# Test if running as standalone
if __name__ == "__main__":
    date_range = show_date_range_selection()
    if date_range and not date_range.cancelled:
        print(f"Selected date range: {date_range.date_range_formatted}")
        print(f"Start date: {date_range.start_date}")
        print(f"End date: {date_range.end_date}")
        print(f"Year: {date_range.year}")
    else:
        print("Date selection was cancelled.")