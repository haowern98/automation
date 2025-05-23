"""
SharePoint Automation - Main Application with tabbed interface
"""
import sys
import datetime
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout,
                             QLabel, QDateEdit, QLineEdit, QPushButton,
                             QGridLayout, QTabWidget, QWidget, QMessageBox,
                             QGroupBox, QFileDialog)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QColor

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
        
        # Create a heading
        heading = QLabel("General Settings")
        heading_font = QFont("Segoe UI", 12)
        heading_font.setBold(True)
        heading.setFont(heading_font)
        
        # Create a description
        description = QLabel("Configure general application settings")
        description.setWordWrap(True)
        
        # Placeholder for future settings
        placeholder = QLabel("Settings will be available in future versions")
        placeholder.setStyleSheet("color: gray; font-style: italic;")
        placeholder.setAlignment(Qt.AlignCenter)
        
        # Add widgets to layout
        layout.addWidget(heading)
        layout.addWidget(description)
        layout.addSpacing(20)
        layout.addWidget(placeholder)
        layout.addStretch(1)
        
        return tab
    
    def _create_file_paths_tab(self):
        """Create the file paths settings tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Import settings manager
        from src.gui.settings_dialog import SettingsManager
        self.settings_manager = SettingsManager()
        
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
        save_button.clicked.connect(self.save_settings)
        layout.addWidget(save_button)
        
        # Add stretch to push everything up
        layout.addStretch(1)
        
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
    
    def save_settings(self):
        """Save settings"""
        import os
        
        try:
            # Validate inputs
            gsn_dir = self.gsn_directory_edit.text().strip()
            er_dir = self.er_directory_edit.text().strip()
            gsn_pattern = self.gsn_pattern_edit.text().strip()
            er_pattern = self.er_pattern_edit.text().strip()
            
            # Check if directories exist
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
            
            # Save settings
            self.settings_manager.set('file_paths', 'gsn_search_directory', gsn_dir)
            self.settings_manager.set('file_paths', 'er_search_directory', er_dir)
            self.settings_manager.set('file_paths', 'gsn_file_pattern', gsn_pattern)
            self.settings_manager.set('file_paths', 'er_file_pattern', er_pattern)
            
            # Save to file
            if self.settings_manager.save_settings():
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

class SharePointAutomationApp(QDialog):
    """Main application dialog with tabbed interface"""
    
    def __init__(self, parent=None, manual_mode=False):
        """
        Initialize the main application dialog
        
        Args:
            parent: Parent widget
            manual_mode (bool): Whether the application is running in manual mode
        """
        super(SharePointAutomationApp, self).__init__(parent)
        
        self.manual_mode = manual_mode
        
        self.setWindowTitle("SharePoint Automation")
        self.setFixedSize(550, 500)  # Increased height for settings content
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
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
        
        # Add buttons at the bottom
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        
        # Set cancel button text based on mode
        cancel_text = "Exit" if manual_mode else "Use Auto Date"
        self.cancel_button = QPushButton(cancel_text)
        
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.setAlignment(Qt.AlignCenter)
        
        # Add button layout to main layout
        layout.addLayout(button_layout)
        
        # Connect signals
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.handle_cancel)
    
    def handle_cancel(self):
        """Handle the cancel button click differently based on mode"""
        if self.manual_mode:
            # In manual mode, exit the application
            self.date_range_tab.result_obj.cancelled = True
            self.reject()
        else:
            # In auto mode, just reject the dialog
            # The main program will use automatic date calculation
            self.date_range_tab.result_obj.cancelled = True
            self.reject()
    
    def get_date_range_result(self):
        """Get the date range result from the date range tab"""
        return self.date_range_tab.result_obj

def show_tabbed_date_range_selection(manual_mode=False):
    """
    Show the tabbed date range selection dialog
    
    Args:
        manual_mode (bool): Whether the application is running in manual mode
        
    Returns:
        DateRangeResult: The selected date range or a result object with cancelled=True if cancelled
    """
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    dialog = SharePointAutomationApp(manual_mode=manual_mode)
    result = dialog.exec_() == QDialog.Accepted
    
    # Always return the result object, even when cancelled
    return dialog.get_date_range_result()

# Just for compatibility with existing code
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