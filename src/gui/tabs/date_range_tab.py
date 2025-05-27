"""
Date Range Selection Tab with Buttons

Updated tab for selecting date ranges with mode-specific buttons integrated.
"""
import datetime
from PyQt5.QtWidgets import QWidget, QGridLayout, QLabel, QDateEdit, QLineEdit, QHBoxLayout, QPushButton
from PyQt5.QtCore import QDate, pyqtSignal
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
    """Date range selection tab with integrated buttons"""
    
    # Signals for communicating with the main dialog
    ok_clicked = pyqtSignal()
    exit_clicked = pyqtSignal()
    use_auto_date_clicked = pyqtSignal()
    
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
            title_text = "Excel Data Processing - Select date range to extract and process data into the Weekly Report Excel file:"
        else:
            title_text = "Excel Data Processing - Select a date range or use Auto Date:"
        
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
        
        # Create buttons based on mode
        self.button_layout = self._create_buttons()
        
        # Add widgets to layout
        layout.addWidget(title_label, 0, 0, 1, 2)
        layout.addWidget(start_date_label, 1, 0)
        layout.addWidget(self.start_date_picker, 1, 1)
        layout.addWidget(end_date_label, 2, 0)
        layout.addWidget(self.end_date_picker, 2, 1)
        layout.addWidget(preview_label, 3, 0)
        layout.addWidget(self.preview_textbox, 3, 1)
        layout.addWidget(self.status_label, 4, 0, 1, 2)
        
        # Add button layout to the grid layout
        layout.addLayout(self.button_layout, 5, 0, 1, 2)
        
        # Set spacing
        layout.setVerticalSpacing(10)
        
        # Connect signals
        self.start_date_picker.dateChanged.connect(self.update_preview)
        self.end_date_picker.dateChanged.connect(self.update_preview)
        
        # Initial preview update
        self.update_preview()
    
    def _create_buttons(self):
        """Create buttons based on mode"""
        button_layout = QHBoxLayout()
        
        if self.manual_mode:
            # Manual mode: Just OK and Exit buttons centered
            self.ok_button = QPushButton("OK")
            self.exit_button = QPushButton("Exit")
            
            button_layout.addStretch()
            button_layout.addWidget(self.ok_button)
            button_layout.addWidget(self.exit_button)
            button_layout.addStretch()
            
            # Connect signals
            self.ok_button.clicked.connect(self.ok_clicked.emit)
            self.exit_button.clicked.connect(self.exit_clicked.emit)
            
        else:
            # Auto mode: OK, Use Auto Date, and Exit buttons
            self.ok_button = QPushButton("OK")
            self.use_auto_date_button = QPushButton("Use Auto Date")
            self.exit_button = QPushButton("Exit")
            
            # Layout: OK (left), center area with Use Auto Date and Exit
            button_layout.addWidget(self.ok_button)
            button_layout.addStretch()
            button_layout.addWidget(self.use_auto_date_button)
            button_layout.addWidget(self.exit_button)
            button_layout.addStretch()
            
            # Connect signals
            self.ok_button.clicked.connect(self.ok_clicked.emit)
            self.use_auto_date_button.clicked.connect(self.use_auto_date_clicked.emit)
            self.exit_button.clicked.connect(self.exit_clicked.emit)
        
        return button_layout
    
    def update_preview(self):
        """Update the date range preview"""
        start_date = self.start_date_picker.date().toPyDate()
        end_date = self.end_date_picker.date().toPyDate()
        
        # Validate dates
        if end_date < start_date:
            self.status_label.setText("End date cannot be earlier than start date.")
            self._set_ok_button_enabled(False)
            return
        else:
            self.status_label.setText("")
            self._set_ok_button_enabled(True)
        
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
    
    def _set_ok_button_enabled(self, enabled):
        """Enable or disable the OK button"""
        if hasattr(self, 'ok_button'):
            self.ok_button.setEnabled(enabled)
    
    def set_buttons_enabled(self, enabled):
        """Enable or disable all buttons"""
        if hasattr(self, 'ok_button'):
            self.ok_button.setEnabled(enabled)
        if hasattr(self, 'exit_button'):
            self.exit_button.setEnabled(enabled)
        if hasattr(self, 'use_auto_date_button'):
            self.use_auto_date_button.setEnabled(enabled)