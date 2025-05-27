"""
SharePoint Automation - Main Application with Buttons in Date Range Tab

This is the main application dialog with buttons moved to the date range tab.
The progress bar remains in the main dialog.
"""
import sys
import datetime
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout,
                             QTabWidget, QProgressBar, QLabel)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont

# Import individual tabs - using the updated DateRangeTab
from src.gui.tabs.date_range_tab import DateRangeTab, DateRangeResult
from src.gui.tabs import SettingsTab, WeeklyReportTab


class EnhancedSharePointAutomationApp(QDialog):
    """Enhanced main application dialog with buttons in date range tab"""
    
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
        self.timeout_disabled = False  # Initialize this flag
        
        self.setWindowTitle("SharePoint Automation")
        self.resize(700, 650)  # Set initial size but allow resizing
        self.setMinimumSize(600, 500)  # Set minimum size to prevent it from being too small
        # Only remove the help button, keep minimize, maximize, close buttons
        self.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.WindowSystemMenuHint | 
                           Qt.WindowMinimizeButtonHint | Qt.WindowMaximizeButtonHint | 
                           Qt.WindowCloseButtonHint)
        
        # Override close event to handle X button termination
        self.closeEvent = self.handle_close_event
        
        # Create main layout
        layout = QVBoxLayout(self)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Create tabs using the organized structure
        self.date_range_tab = DateRangeTab(manual_mode=manual_mode)
        self.weekly_report_tab = WeeklyReportTab()
        self.settings_tab = SettingsTab()
        
        # Connect signals from the date range tab
        self.date_range_tab.ok_clicked.connect(self.handle_ok)
        self.date_range_tab.exit_clicked.connect(self.handle_exit)
        if not manual_mode:
            self.date_range_tab.use_auto_date_clicked.connect(self.handle_use_auto_date)
        
        # Add tabs to tab widget in desired order
        self.tab_widget.addTab(self.date_range_tab, "Extract and Process Data to Excel")
        self.tab_widget.addTab(self.weekly_report_tab, "Generate Report from Excel")
        self.tab_widget.addTab(self.settings_tab, "Settings")
        
        # Add tab widget to layout
        layout.addWidget(self.tab_widget)
        
        # Add timeout indicator for auto mode (progress bar remains here)
        if not manual_mode:
            # Check if timeout is disabled (very large number)
            self.timeout_disabled = timeout_seconds >= 999999
            
            if not self.timeout_disabled:
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
            else:
                # Timeout is disabled, show a message instead
                disabled_label = QLabel("⏱️ Auto timeout is disabled")
                disabled_label.setAlignment(Qt.AlignCenter)
                disabled_label.setStyleSheet("color: #666; font-size: 11px; padding: 5px;")
                layout.addWidget(disabled_label)
    
    def handle_ok(self):
        """Handle OK button click from date range tab"""
        self.accept()
    
    def handle_exit(self):
        """Handle Exit button click from date range tab"""
        if self.manual_mode:
            self.handle_terminate_process()
        else:
            self.handle_terminate_process()
    
    def handle_use_auto_date(self):
        """Handle Use Auto Date button click from date range tab"""
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
    
    def update_countdown(self):
        """Update the countdown timer display"""
        if not self.manual_mode and not self.timeout_disabled:
            self.remaining_seconds -= 1
            self.countdown_label.setText(f"⏱️ Auto date will be used in {self.remaining_seconds} seconds")
            self.progress_bar.setValue(self.remaining_seconds)
            
            if self.remaining_seconds <= 0:
                self.countdown_timer.stop()
    
    def handle_timeout(self):
        """Handle timeout in auto mode"""
        if not self.manual_mode and not self.timeout_disabled:
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