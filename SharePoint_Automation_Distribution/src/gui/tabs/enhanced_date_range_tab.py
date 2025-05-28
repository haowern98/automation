"""
Enhanced Date Range Tab with Processing Options and Status Display

This module provides an enhanced date range tab with:
- Standard date range selection
- Processing options
- Progress and status display
- File status indicators
- Error display
"""
import os
import datetime
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, 
                            QLabel, QDateEdit, QLineEdit, QPushButton, 
                            QGroupBox, QCheckBox, QProgressBar, QScrollArea,
                            QFrame)
from PyQt5.QtCore import Qt, QDate, pyqtSignal
from PyQt5.QtGui import QFont
from src.models import DateRangeResult

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


class EnhancedDateRangeTab(QWidget):
    """Enhanced date range tab with processing options and status display"""
    
    # Signals for communicating with the main dialog
    ok_clicked = pyqtSignal()
    exit_clicked = pyqtSignal()
    use_auto_date_clicked = pyqtSignal()
    
    def __init__(self, parent=None, manual_mode=False):
        """
        Initialize the enhanced date range tab
        
        Args:
            parent: Parent widget
            manual_mode (bool): Whether the application is running in manual mode
        """
        super(EnhancedDateRangeTab, self).__init__(parent)
        
        self.manual_mode = manual_mode
        self.processing = False
        self.error_logs = []
        self.warning_logs = []
        
        # Create result object
        self.result_obj = DateRangeResult()
        
        # Create main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(10)
        
        # Create title label with exact text from screenshot
        title_label = QLabel("Excel Data Processing - Select date range to extract and process data into the Weekly Report Excel file:")
        title_font = QFont()
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setWordWrap(True)
        main_layout.addWidget(title_label)
        
        # Add spacing after title
        main_layout.addSpacing(20)
        
        # Create date selection grid layout
        date_grid = QGridLayout()
        date_grid.setVerticalSpacing(10)
        
        # Start date
        start_date_label = QLabel("Start Date:")
        self.start_date_picker = QDateEdit()
        self.start_date_picker.setCalendarPopup(True)
        self.start_date_picker.setDate(QDate.currentDate())
        
        # End date
        end_date_label = QLabel("End Date:")
        self.end_date_picker = QDateEdit()
        self.end_date_picker.setCalendarPopup(True)
        self.end_date_picker.setDate(QDate.currentDate().addDays(1))
        
        # Preview
        preview_label = QLabel("Preview:")
        self.preview_textbox = QLineEdit()
        self.preview_textbox.setReadOnly(True)
        
        # Status label for validation errors
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: red;")
        
        # Add widgets to grid layout
        date_grid.addWidget(start_date_label, 0, 0)
        date_grid.addWidget(self.start_date_picker, 0, 1)
        date_grid.addWidget(end_date_label, 1, 0)
        date_grid.addWidget(self.end_date_picker, 1, 1)
        date_grid.addWidget(preview_label, 2, 0)
        date_grid.addWidget(self.preview_textbox, 2, 1)
        date_grid.addWidget(self.status_label, 3, 0, 1, 2)
        
        # Set column stretch factors
        date_grid.setColumnStretch(0, 0)  # Label column doesn't stretch
        date_grid.setColumnStretch(1, 1)  # Input column stretches
        
        # Add date grid to main layout
        main_layout.addLayout(date_grid)
        
        # Add spacing before processing options
        main_layout.addSpacing(20)
        
        # Create processing options section (collapsible/expandable)
        self.options_group = QGroupBox("Processing Options")
        self.options_group.setCheckable(True)
        self.options_group.setChecked(True)  # Expanded by default
        options_layout = QVBoxLayout(self.options_group)
        
        # Create option checkboxes
        self.gsn_er_checkbox = QCheckBox("Process GSN vs ER comparison")
        self.gsn_er_checkbox.setChecked(True)
        
        self.gsn_ad_checkbox = QCheckBox("Process GSN vs AD comparison")
        self.gsn_ad_checkbox.setChecked(True)
        
        self.er_nologon_checkbox = QCheckBox("Generate ER No Logon report (31-60 days)")
        self.er_nologon_checkbox.setChecked(True)
        
        self.backup_checkbox = QCheckBox("Create backup before processing")
        self.backup_checkbox.setChecked(False)
        
        # Add checkboxes to options layout
        options_layout.addWidget(self.gsn_er_checkbox)
        options_layout.addWidget(self.gsn_ad_checkbox)
        options_layout.addWidget(self.er_nologon_checkbox)
        options_layout.addWidget(self.backup_checkbox)
        
        # Add options group to main layout
        main_layout.addWidget(self.options_group)
        
        # Add status section (collapsible/expandable)
        self.status_group = QGroupBox("Status & Progress")
        self.status_group.setCheckable(True)
        self.status_group.setChecked(True)  # Expanded by default
        status_layout = QVBoxLayout(self.status_group)
        
        # Progress bar and label
        progress_layout = QHBoxLayout()
        progress_layout.addWidget(QLabel("Progress:"))
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        self.progress_percent = QLabel("0%")
        progress_layout.addWidget(self.progress_percent)
        
        status_layout.addLayout(progress_layout)
        
        # Current step
        self.current_step_label = QLabel("Current Step: Waiting to start...")
        status_layout.addWidget(self.current_step_label)
        
        # Estimated time
        self.estimated_time_label = QLabel("Estimated Time: -")
        status_layout.addWidget(self.estimated_time_label)
        
        # Add separator line
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        status_layout.addWidget(separator)
        
        # File status section
        status_layout.addWidget(QLabel("File Status:"))
        
        # File status grid
        file_status_grid = QGridLayout()
        
        # GSN File status
        self.gsn_file_icon = QLabel("✓")
        self.gsn_file_label = QLabel("GSN File:")
        self.gsn_file_status = QLabel("Not found")
        
        file_status_grid.addWidget(self.gsn_file_icon, 0, 0)
        file_status_grid.addWidget(self.gsn_file_label, 0, 1)
        file_status_grid.addWidget(self.gsn_file_status, 0, 2)
        
        # ER File status
        self.er_file_icon = QLabel("✓")
        self.er_file_label = QLabel("ER File:")
        self.er_file_status = QLabel("Not found")
        
        file_status_grid.addWidget(self.er_file_icon, 1, 0)
        file_status_grid.addWidget(self.er_file_label, 1, 1)
        file_status_grid.addWidget(self.er_file_status, 1, 2)
        
        # Weekly Report status
        self.weekly_file_icon = QLabel("✓")
        self.weekly_file_label = QLabel("Weekly Report:")
        self.weekly_file_status = QLabel("Not found")
        
        file_status_grid.addWidget(self.weekly_file_icon, 2, 0)
        file_status_grid.addWidget(self.weekly_file_label, 2, 1)
        file_status_grid.addWidget(self.weekly_file_status, 2, 2)
        
        # AD Results status
        self.ad_file_icon = QLabel("!")
        self.ad_file_label = QLabel("AD Results:")
        self.ad_file_status = QLabel("Will be generated during processing")
        
        file_status_grid.addWidget(self.ad_file_icon, 3, 0)
        file_status_grid.addWidget(self.ad_file_label, 3, 1)
        file_status_grid.addWidget(self.ad_file_status, 3, 2)
        
        # Set column stretch
        file_status_grid.setColumnStretch(2, 1)
        
        status_layout.addLayout(file_status_grid)
        
        # Add separator before errors
        separator2 = QFrame()
        separator2.setFrameShape(QFrame.HLine)
        separator2.setFrameShadow(QFrame.Sunken)
        status_layout.addWidget(separator2)
        
        # Error section
        self.error_heading = QLabel("ERRORS:")
        status_layout.addWidget(self.error_heading)
        
        # Scrollable error list
        self.error_scroll = QScrollArea()
        self.error_scroll.setWidgetResizable(True)
        self.error_scroll.setMaximumHeight(100)
        self.error_scroll.setFrameShape(QFrame.StyledPanel)
        
        self.error_widget = QWidget()
        self.error_layout = QVBoxLayout(self.error_widget)
        self.error_layout.setAlignment(Qt.AlignTop)
        
        self.error_scroll.setWidget(self.error_widget)
        status_layout.addWidget(self.error_scroll)
        
        # Initially hide the error scroll area
        self.error_scroll.setVisible(False)
        self.error_heading.setVisible(False)
        
        # Add status group to main layout
        main_layout.addWidget(self.status_group)
        
        # Add stretch to push everything up
        main_layout.addStretch(1)
        
        # Create buttons layout
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        # Create OK button
        self.ok_button = QPushButton("OK")
        self.ok_button.clicked.connect(self.ok_clicked.emit)
        button_layout.addWidget(self.ok_button)
        
        # Add Use Auto Date button if not in manual mode
        if not self.manual_mode:
            self.use_auto_date_button = QPushButton("Use Auto Date")
            self.use_auto_date_button.clicked.connect(self.use_auto_date_clicked.emit)
            button_layout.addWidget(self.use_auto_date_button)
        
        # Create Exit button
        self.exit_button = QPushButton("Exit")
        self.exit_button.clicked.connect(self.exit_clicked.emit)
        button_layout.addWidget(self.exit_button)
        
        # Center the buttons
        button_container = QHBoxLayout()
        button_container.addStretch(1)
        button_container.addLayout(button_layout)
        button_container.addStretch(1)
        
        # Add buttons to main layout
        main_layout.addLayout(button_container)
        
        # Connect date change signals
        self.start_date_picker.dateChanged.connect(self.update_preview)
        self.end_date_picker.dateChanged.connect(self.update_preview)
        
        # Initial preview update
        self.update_preview()
        
        # Hide status section initially
        self.status_group.setVisible(False)
    
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
        
        # Store for result object
        self.result_obj.start_date = start_date
        self.result_obj.end_date = end_date
        self.result_obj.date_range_formatted = date_range_formatted
        self.result_obj.year = str(end_date.year)
    
    def set_progress(self, percent, step=None, time_remaining=None):
        """
        Set the progress indicators
        
        Args:
            percent (int): Progress percentage (0-100)
            step (str, optional): Current processing step
            time_remaining (str, optional): Estimated time remaining
        """
        # Update progress bar
        self.progress_bar.setValue(percent)
        self.progress_percent.setText(f"{percent}%")
        
        # Update current step if provided
        if step:
            self.current_step_label.setText(f"Current Step: {step}")
        
        # Update estimated time if provided
        if time_remaining:
            self.estimated_time_label.setText(f"Estimated Time: {time_remaining}")
        
        # Make sure status section is visible
        self.status_group.setChecked(True)
        self.status_group.setVisible(True)
    
    def set_file_status(self, file_type, found, name=None):
        """
        Set the status of a file
        
        Args:
            file_type (str): Type of file ('gsn', 'er', 'weekly', 'ad')
            found (bool): Whether the file was found
            name (str, optional): Name of the file if found
        """
        icon = None
        label = None
        status = None
        
        if file_type.lower() == 'gsn':
            icon = self.gsn_file_icon
            label = self.gsn_file_label
            status = self.gsn_file_status
        elif file_type.lower() == 'er':
            icon = self.er_file_icon
            label = self.er_file_label
            status = self.er_file_status
        elif file_type.lower() == 'weekly':
            icon = self.weekly_file_icon
            label = self.weekly_file_label
            status = self.weekly_file_status
        elif file_type.lower() == 'ad':
            icon = self.ad_file_icon
            label = self.ad_file_label
            status = self.ad_file_status
        else:
            return
        
        if found:
            icon.setText("✓")
            status.setText(f"Found ({name})" if name else "Found")
        else:
            icon.setText("✗")
            status.setText("Not found" if not name else name)
    
    def add_error(self, error_message):
        """
        Add an error message to the error list
        
        Args:
            error_message (str): Error message to display
        """
        # Create error label
        error_label = QLabel(error_message)
        error_label.setWordWrap(True)
        
        # Add to layout
        self.error_layout.addWidget(error_label)
        
        # Show error section
        self.error_heading.setVisible(True)
        self.error_scroll.setVisible(True)
        
        # Make sure status section is visible
        self.status_group.setChecked(True)
        self.status_group.setVisible(True)
    
    def clear_errors(self):
        """Clear all error messages"""
        # Remove all widgets from error layout
        while self.error_layout.count():
            item = self.error_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Hide error section
        self.error_heading.setVisible(False)
        self.error_scroll.setVisible(False)
    
    def start_processing_mode(self):
        """Enter processing mode - disable inputs, show status"""
        # Disable inputs
        self.start_date_picker.setEnabled(False)
        self.end_date_picker.setEnabled(False)
        self.gsn_er_checkbox.setEnabled(False)
        self.gsn_ad_checkbox.setEnabled(False)
        self.er_nologon_checkbox.setEnabled(False)
        self.backup_checkbox.setEnabled(False)
        
        # Disable buttons
        self.ok_button.setEnabled(False)
        if hasattr(self, 'use_auto_date_button'):
            self.use_auto_date_button.setEnabled(False)
        self.exit_button.setEnabled(False)
        
        # Show status section
        self.status_group.setChecked(True)
        self.status_group.setVisible(True)
        
        # Clear errors
        self.clear_errors()
        
        # Set processing flag
        self.processing = True
    
    def end_processing_mode(self, success=True):
        """
        Exit processing mode - re-enable inputs
        
        Args:
            success (bool): Whether processing was successful
        """
        # Re-enable inputs
        self.start_date_picker.setEnabled(True)
        self.end_date_picker.setEnabled(True)
        self.gsn_er_checkbox.setEnabled(True)
        self.gsn_ad_checkbox.setEnabled(True)
        self.er_nologon_checkbox.setEnabled(True)
        self.backup_checkbox.setEnabled(True)
        
        # Re-enable buttons
        self.ok_button.setEnabled(True)
        if hasattr(self, 'use_auto_date_button'):
            self.use_auto_date_button.setEnabled(True)
        self.exit_button.setEnabled(True)
        
        # Set final status
        if success:
            self.set_progress(100, "Processing complete!", "")
        else:
            self.set_progress(0, "Processing failed", "")
        
        # Set processing flag
        self.processing = False
    
    def get_processing_options(self):
        """
        Get the selected processing options
        
        Returns:
            dict: Dictionary of processing options
        """
        return {
            'process_gsn_er': self.gsn_er_checkbox.isChecked(),
            'process_gsn_ad': self.gsn_ad_checkbox.isChecked(),
            'generate_er_nologon': self.er_nologon_checkbox.isChecked(),
            'create_backup': self.backup_checkbox.isChecked()
        }