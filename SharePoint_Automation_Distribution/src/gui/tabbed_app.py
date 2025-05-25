"""
SharePoint Automation - Enhanced Main Application with 3-button auto mode and proper termination
"""
import sys
import datetime
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout,
                             QLabel, QDateEdit, QLineEdit, QPushButton,
                             QGridLayout, QTabWidget, QWidget, QMessageBox,
                             QGroupBox, QFileDialog, QProgressBar, QComboBox, QCheckBox)
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5.QtGui import QFont, QColor, QIcon

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

class SettingsTab(QWidget):
    """Settings tab widget with nested tabs"""
    
    def __init__(self, parent=None):
        """Initialize the settings tab"""
        super(SettingsTab, self).__init__(parent)
        
        # Create main layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create internal tab widget for settings
        self.settings_tabs = QTabWidget()
        
        # Create the General tab
        general_tab = self._create_general_tab()
        
        # Create the File Paths tab
        file_paths_tab = self._create_file_paths_tab()
        
        # Add tabs to the settings tab widget
        self.settings_tabs.addTab(general_tab, "General")
        self.settings_tabs.addTab(file_paths_tab, "File Paths")
        
        # Add the settings tabs to the main layout
        main_layout.addWidget(self.settings_tabs)
    
    def _create_general_tab(self):
        """Create the general settings tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Import settings manager
        from src.gui.settings_dialog import SettingsManager
        self.settings_manager = SettingsManager()
        
        # Create a heading
        heading = QLabel("General Settings")
        heading_font = QFont("Segoe UI", 12)
        heading_font.setBold(True)
        heading.setFont(heading_font)
        
        # Create a description
        description = QLabel("Configure general application settings")
        description.setWordWrap(True)
        
        # Create timeout settings group
        timeout_group = QGroupBox("Auto Mode Timeout")
        timeout_layout = QGridLayout(timeout_group)
        
        # Add timeout dropdown
        timeout_layout.addWidget(QLabel("Auto mode timeout:"), 0, 0)
        self.timeout_dropdown = QComboBox()
        self.timeout_dropdown.addItems(["10 seconds", "20 seconds", "30 seconds", "45 seconds", "60 seconds", "90 seconds", "120 seconds"])
        
        # Set default selection to 30 seconds (index 2)
        self.timeout_dropdown.setCurrentIndex(2)
        
        timeout_layout.addWidget(self.timeout_dropdown, 0, 1)
        
        # Add help text
        timeout_help = QLabel("Time before the date selection dialog automatically uses auto date in auto mode")
        timeout_help.setStyleSheet("color: gray; font-size: 10px;")
        timeout_layout.addWidget(timeout_help, 1, 0, 1, 2)
        
        # Create debug settings group
        debug_group = QGroupBox("Debug Settings")
        debug_layout = QGridLayout(debug_group)
        
        # Add terminal visibility checkbox
        self.show_terminal_checkbox = QCheckBox("Show terminal window")
        debug_layout.addWidget(self.show_terminal_checkbox, 0, 0)
        
        # Add help text
        terminal_help = QLabel("Show command terminal for debugging (requires restart)")
        terminal_help.setStyleSheet("color: gray; font-size: 10px;")
        debug_layout.addWidget(terminal_help, 1, 0)
        
        # Add widgets to layout
        layout.addWidget(heading)
        layout.addWidget(description)
        layout.addSpacing(20)
        layout.addWidget(timeout_group)
        layout.addSpacing(10)
        layout.addWidget(debug_group)
        layout.addStretch(1)  # Push everything up, button will be at the bottom
        
        # Add save button
        save_button = QPushButton("Save Settings")
        save_button.setDefault(False)  # Set to False to remove default status
        save_button.setFocusPolicy(Qt.NoFocus)
        save_button.clicked.connect(self.save_settings)
        layout.addWidget(save_button)
        
        # Load current timeout setting if available
        self.load_timeout_setting()
        self.load_debug_settings()
        
        return tab
    
    def _create_file_paths_tab(self):
        """Create the file paths settings tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Create a heading
        heading = QLabel("File Search Configuration")
        heading_font = QFont("Segoe UI", 12)
        heading_font.setBold(True)
        heading.setFont(heading_font)
        
        # Create a description
        description = QLabel("Configure directories to search for files and file name patterns")
        description.setWordWrap(True)
        
        # Add widgets to layout
        layout.addWidget(heading)
        layout.addWidget(description)
        layout.addSpacing(20)
        
        # Create GSN File Settings Group
        gsn_group = QGroupBox("GSN File Settings")
        gsn_layout = QGridLayout(gsn_group)
        
        # GSN Search Directory
        gsn_layout.addWidget(QLabel("Search Directory:"), 0, 0)
        self.gsn_directory_edit = QLineEdit()
        self.gsn_directory_edit.setPlaceholderText("Enter path to directory to search for GSN files")
        gsn_layout.addWidget(self.gsn_directory_edit, 0, 1)
        
        gsn_browse_button = QPushButton("Browse...")
        gsn_browse_button.clicked.connect(self.browse_gsn_directory)
        gsn_layout.addWidget(gsn_browse_button, 0, 2)
        
        # GSN File Pattern
        gsn_layout.addWidget(QLabel("File Name Pattern:"), 1, 0)
        self.gsn_pattern_edit = QLineEdit()
        self.gsn_pattern_edit.setPlaceholderText("Enter file name pattern (e.g., alm_hardware)")
        gsn_layout.addWidget(self.gsn_pattern_edit, 1, 1, 1, 2)
        
        # Add info label
        gsn_info = QLabel("Matches ANY file that starts with the File Name Pattern and ends with .xlsx")
        gsn_info.setStyleSheet("color: gray; font-size: 10px;")
        gsn_layout.addWidget(gsn_info, 2, 1, 1, 2)
        
        layout.addWidget(gsn_group)
        
        # Create ER File Settings Group
        er_group = QGroupBox("ER File Settings")
        er_layout = QGridLayout(er_group)
        
        # ER Search Directory
        er_layout.addWidget(QLabel("Search Directory:"), 0, 0)
        self.er_directory_edit = QLineEdit()
        self.er_directory_edit.setPlaceholderText("Enter path to directory to search for ER files")
        er_layout.addWidget(self.er_directory_edit, 0, 1)
        
        er_browse_button = QPushButton("Browse...")
        er_browse_button.clicked.connect(self.browse_er_directory)
        er_layout.addWidget(er_browse_button, 0, 2)
        
        # ER File Pattern
        er_layout.addWidget(QLabel("File Name Pattern:"), 1, 0)
        self.er_pattern_edit = QLineEdit()
        self.er_pattern_edit.setPlaceholderText("Enter file name pattern (e.g., data)")
        er_layout.addWidget(self.er_pattern_edit, 1, 1, 1, 2)
        
        # Add info label
        er_info = QLabel("Matches ANY file that starts with the File Name Pattern and ends with .xlsx")
        er_info.setStyleSheet("color: gray; font-size: 10px;")
        er_layout.addWidget(er_info, 2, 1, 1, 2)
        
        layout.addWidget(er_group)
        
        # Add save button
        save_button = QPushButton("Save Settings")
        save_button.setDefault(False)  # Set to False to remove default status
        save_button.clicked.connect(self.save_settings)
        
        # Add stretch to push everything up and button to bottom
        layout.addStretch(1)
        layout.addWidget(save_button)
        
        # Load current settings
        self.load_current_settings()
        
        return tab
    
    def browse_gsn_directory(self):
        """Browse for GSN search directory"""
        import os
        
        current_path = self.gsn_directory_edit.text()
        if not current_path:
            current_path = os.environ.get('USERPROFILE', '')
        
        directory = QFileDialog.getExistingDirectory(
            self, "Select GSN Search Directory", current_path)
        
        if directory:
            self.gsn_directory_edit.setText(directory)
    
    def browse_er_directory(self):
        """Browse for ER search directory"""
        import os
        
        current_path = self.er_directory_edit.text()
        if not current_path:
            current_path = os.environ.get('USERPROFILE', '')
        
        directory = QFileDialog.getExistingDirectory(
            self, "Select ER Search Directory", current_path)
        
        if directory:
            self.er_directory_edit.setText(directory)
    
    def load_current_settings(self):
        """Load current settings into the dialog"""
        # Load GSN settings
        gsn_dir = self.settings_manager.get('file_paths', 'gsn_search_directory', '')
        gsn_pattern = self.settings_manager.get('file_paths', 'gsn_file_pattern', 'alm_hardware')
        
        self.gsn_directory_edit.setText(gsn_dir)
        self.gsn_pattern_edit.setText(gsn_pattern)
        
        # Load ER settings
        er_dir = self.settings_manager.get('file_paths', 'er_search_directory', '')
        er_pattern = self.settings_manager.get('file_paths', 'er_file_pattern', 'data')
        
        self.er_directory_edit.setText(er_dir)
        self.er_pattern_edit.setText(er_pattern)
    
    def load_timeout_setting(self):
        """Load timeout setting from settings"""
        # Default to 30 seconds if not set
        timeout_value = self.settings_manager.get('general', 'auto_mode_timeout', '30')
        
        # Find the matching index in the dropdown
        for i in range(self.timeout_dropdown.count()):
            if timeout_value in self.timeout_dropdown.itemText(i):
                self.timeout_dropdown.setCurrentIndex(i)
                break
    
    def load_debug_settings(self):
        """Load debug settings"""
        # Load terminal visibility setting
        show_terminal = self.settings_manager.get('general', 'show_terminal', False)
        self.show_terminal_checkbox.setChecked(show_terminal)
    
    def save_settings(self):
        """Save all settings"""
        try:
            # Get values from UI controls
            
            # Save timeout setting
            timeout_text = self.timeout_dropdown.currentText()
            timeout_value = timeout_text.split()[0]  # Extract just the number
            self.settings_manager.set('general', 'auto_mode_timeout', timeout_value)
            
            # Save terminal visibility setting if the checkbox exists
            if hasattr(self, 'show_terminal_checkbox'):
                show_terminal = self.show_terminal_checkbox.isChecked()
                self.settings_manager.set('general', 'show_terminal', show_terminal)
            
            # Get file path settings
            gsn_dir = self.gsn_directory_edit.text().strip()
            er_dir = self.er_directory_edit.text().strip()
            gsn_pattern = self.gsn_pattern_edit.text().strip()
            er_pattern = self.er_pattern_edit.text().strip()
            
            # Validate file path settings
            import os
            if gsn_dir and not os.path.exists(gsn_dir):
                QMessageBox.warning(self, "Invalid Directory", 
                                f"GSN search directory does not exist:\n{gsn_dir}")
                return
            
            if er_dir and not os.path.exists(er_dir):
                QMessageBox.warning(self, "Invalid Directory", 
                                f"ER search directory does not exist:\n{er_dir}")
                return
            
            # Check if patterns are not empty
            if not gsn_pattern:
                QMessageBox.warning(self, "Invalid Pattern", 
                                "GSN file pattern cannot be empty")
                return
            
            if not er_pattern:
                QMessageBox.warning(self, "Invalid Pattern", 
                                "ER file pattern cannot be empty")
                return
            
            # Save file path settings
            self.settings_manager.set('file_paths', 'gsn_search_directory', gsn_dir)
            self.settings_manager.set('file_paths', 'er_search_directory', er_dir)
            self.settings_manager.set('file_paths', 'gsn_file_pattern', gsn_pattern)
            self.settings_manager.set('file_paths', 'er_file_pattern', er_pattern)
            
            # Save settings to file
            if self.settings_manager.save_settings():
                # Show notification about terminal visibility if it changed
                if hasattr(self, 'show_terminal_checkbox'):
                    show_terminal = self.show_terminal_checkbox.isChecked()
                    # Check if setting exists and has changed
                    if self.settings_manager.get('general', 'show_terminal', False) != show_terminal:
                        QMessageBox.information(self, "Settings Saved",
                                            "Settings have been saved successfully!\n\n"
                                            "Note: The terminal visibility setting has changed.\n"
                                            "Please restart the application for this change to take effect.")
                        return
                        
                # Standard success message
                QMessageBox.information(self, "Settings Saved", 
                                    "Settings have been saved successfully!")
            else:
                QMessageBox.critical(self, "Save Error", 
                                    "Failed to save settings to file!")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving settings:\n{str(e)}")

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

class EnhancedSharePointAutomationApp(QDialog):
    """Enhanced main application dialog with 3-button auto mode and proper termination"""
    
    def __init__(self, parent=None, manual_mode=False, timeout_seconds=30):
        """
        Initialize the main application dialog
        
        Args:
            parent: Parent widget
            manual_mode (bool): Whether the application is running in manual mode
            timeout_seconds (int): Timeout in seconds for auto mode
        """
        super(EnhancedSharePointAutomationApp, self).__init__(parent)
        
        self.manual_mode = manual_mode
        self.timeout_seconds = timeout_seconds
        self.process_terminated = False
        self.timed_out = False
        self.remaining_seconds = timeout_seconds
        
        self.setWindowTitle("SharePoint Automation")
        self.setFixedSize(550, 600)  # Increased height for timeout indicator
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        # Override close event to handle X button termination
        self.closeEvent = self.handle_close_event
        
        # Create main layout
        layout = QVBoxLayout(self)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Create tabs
        self.date_range_tab = DateRangeTab(manual_mode=manual_mode)
        self.settings_tab = SettingsTab()
        
        # Add tabs to tab widget
        self.tab_widget.addTab(self.date_range_tab, "Date Range Selector")
        self.tab_widget.addTab(self.settings_tab, "Settings")
        
        # Add tab widget to layout
        layout.addWidget(self.tab_widget)
        
        # Add timeout indicator for auto mode
        if not manual_mode:
            self.timeout_layout = QHBoxLayout()
            
            # Countdown label
            self.countdown_label = QLabel(f"⏱️ Auto date will be used in {self.remaining_seconds} seconds")
            self.countdown_label.setAlignment(Qt.AlignCenter)
            self.countdown_label.setStyleSheet("color: #666; font-size: 11px; padding: 5px;")
            
            # Progress bar
            self.progress_bar = QProgressBar()
            self.progress_bar.setMaximum(timeout_seconds)
            self.progress_bar.setValue(timeout_seconds)
            self.progress_bar.setTextVisible(False)
            self.progress_bar.setMaximumHeight(8)
            
            self.timeout_layout.addWidget(self.countdown_label)
            layout.addLayout(self.timeout_layout)
            layout.addWidget(self.progress_bar)
            
            # Setup timer for countdown
            self.countdown_timer = QTimer()
            self.countdown_timer.timeout.connect(self.update_countdown)
            self.countdown_timer.start(1000)  # Update every second
            
            # Setup timeout timer
            self.timeout_timer = QTimer()
            self.timeout_timer.setSingleShot(True)
            self.timeout_timer.timeout.connect(self.handle_timeout)
            self.timeout_timer.start(timeout_seconds * 1000)
        
        # Add buttons based on mode
        self.create_buttons()
        
        # Add button layout to main layout
        layout.addLayout(self.button_layout)
    
    def create_buttons(self):
        """Create buttons based on mode"""
        self.button_layout = QHBoxLayout()
        
        if self.manual_mode:
            # Manual mode: Just OK and Exit buttons centered
            self.ok_button = QPushButton("OK")
            self.exit_button = QPushButton("Exit")
            
            self.button_layout.addStretch()
            self.button_layout.addWidget(self.ok_button)
            self.button_layout.addWidget(self.exit_button)
            self.button_layout.addStretch()
            
            # Connect signals
            self.ok_button.clicked.connect(self.accept)
            self.exit_button.clicked.connect(self.handle_terminate_process)
            
        else:
            # Auto mode: OK, Use Auto Date, and Exit buttons with default styling
            self.ok_button = QPushButton("OK")
            self.use_auto_date_button = QPushButton("Use Auto Date")
            self.exit_button = QPushButton("Exit")
            
            # Layout: OK (left), center area with Use Auto Date and Exit
            self.button_layout.addWidget(self.ok_button)
            self.button_layout.addStretch()
            self.button_layout.addWidget(self.use_auto_date_button)
            self.button_layout.addWidget(self.exit_button)
            self.button_layout.addStretch()
            
            # Connect signals
            self.ok_button.clicked.connect(self.accept)
            self.use_auto_date_button.clicked.connect(self.handle_use_auto_date)
            self.exit_button.clicked.connect(self.handle_terminate_process)
    
    def update_countdown(self):
        """Update the countdown timer display"""
        if not self.manual_mode:
            self.remaining_seconds -= 1
            self.countdown_label.setText(f"⏱️ Auto date will be used in {self.remaining_seconds} seconds")
            self.progress_bar.setValue(self.remaining_seconds)
            
            if self.remaining_seconds <= 0:
                self.countdown_timer.stop()
    
    def handle_timeout(self):
        """Handle timeout in auto mode"""
        if not self.manual_mode:
            self.timed_out = True
            self.countdown_timer.stop()
            self.countdown_label.setText("⏱️ Timeout - using auto date calculation")
            self.progress_bar.setValue(0)
            
            # Set flags for auto date usage
            date_range = self.date_range_tab.result_obj
            date_range.cancelled = True
            date_range.use_auto_date = True
            date_range.user_terminated = False
            
            # Close dialog and proceed with auto date
            self.reject()
    
    def handle_terminate_process(self):
        """Handle process termination"""
        self.process_terminated = True
        
        # Stop timers if running
        if hasattr(self, 'countdown_timer'):
            self.countdown_timer.stop()
        if hasattr(self, 'timeout_timer'):
            self.timeout_timer.stop()
        
        # Set termination flags
        date_range = self.date_range_tab.result_obj
        date_range.cancelled = True
        date_range.user_terminated = True
        
        self.reject()
    
    def handle_use_auto_date(self):
        """Handle use auto date button"""
        # Stop timers
        if hasattr(self, 'countdown_timer'):
            self.countdown_timer.stop()
        if hasattr(self, 'timeout_timer'):
            self.timeout_timer.stop()
        
        # Set auto date flags
        date_range = self.date_range_tab.result_obj
        date_range.cancelled = True
        date_range.use_auto_date = True
        date_range.user_terminated = False
        
        self.reject()
    
    def handle_close_event(self, event):
        """Handle the X button close event - always terminate the process"""
        self.handle_terminate_process()
        event.accept()
    
    def get_date_range_result(self):
        """Get the date range result from the date range tab"""
        return self.date_range_tab.result_obj

def show_tabbed_date_range_selection(manual_mode=False, timeout_seconds=30):
    """
    Show the enhanced tabbed date range selection dialog
    
    Args:
        manual_mode (bool): Whether the application is running in manual mode
        timeout_seconds (int): Timeout in seconds for auto mode
        
    Returns:
        DateRangeResult: The selected date range or a result object with appropriate flags
    """
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    dialog = EnhancedSharePointAutomationApp(manual_mode=manual_mode, timeout_seconds=timeout_seconds)
    result = dialog.exec_() == QDialog.Accepted
    
    # Always return the result object with appropriate flags set
    return dialog.get_date_range_result()

# Keep the old function name for compatibility
def show_enhanced_date_range_selection(manual_mode=False, timeout_seconds=30):
    """Alias for show_tabbed_date_range_selection"""
    return show_tabbed_date_range_selection(manual_mode, timeout_seconds)

# Compatibility function - just delegates to the enhanced version
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