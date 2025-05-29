"""
Loading Screen for SharePoint Automation

Shows progress during Excel initialization (shutdown and warm-up).
"""
import sys
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QProgressBar, QPushButton, QApplication)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPixmap
from src.utils.logger import write_log


class ExcelInitializationThread(QThread):
    """Background thread for Excel initialization only"""
    
    # Signals for communicating with the loading screen
    progress_updated = pyqtSignal(int, str)  # progress percentage, status message
    finished = pyqtSignal(bool, str)  # success, error_message
    
    def __init__(self, manual_mode=False, debug_mode=False):
        """Initialize the thread"""
        super().__init__()
        self.manual_mode = manual_mode
        self.debug_mode = debug_mode
        self.cancelled = False
    
    def cancel(self):
        """Cancel the initialization"""
        self.cancelled = True
    
    def run(self):
        """Run the Excel initialization process"""
        try:
            # Import here to avoid circular imports
            from src.utils.app_controller import check_run_date, check_excel_processes, warm_up_excel
            
            # Step 1: Check run date (if not manual mode)
            if not self.cancelled:
                self.progress_updated.emit(10, "Checking run date conditions...")
                if not self.manual_mode:
                    if not check_run_date():
                        self.finished.emit(False, "Not a designated run day")
                        return
                else:
                    write_log("Running in manual mode: skipping date checks", "YELLOW")
            
            # Step 2: Check Excel processes
            if not self.cancelled:
                self.progress_updated.emit(30, "Checking Excel processes...")
                excel_count = check_excel_processes()
                
                if excel_count > 0:
                    self.progress_updated.emit(50, f"Terminating {excel_count} Excel processes...")
                    check_excel_processes(terminate_all=True)
                    # Small delay to let processes close
                    self.msleep(2000)
                else:
                    self.progress_updated.emit(50, "No Excel processes found running")
                
            # Step 3: Warm up Excel
            if not self.cancelled:
                self.progress_updated.emit(80, "Warming up Excel application...")
                warm_up_excel()
            
            # Step 4: Complete
            if not self.cancelled:
                self.progress_updated.emit(100, "Excel initialization complete!")
                self.msleep(500)  # Brief pause to show completion
                self.finished.emit(True, "")
            
        except Exception as e:
            error_msg = f"Excel initialization failed: {str(e)}"
            write_log(error_msg, "RED")
            self.finished.emit(False, error_msg)


class ExcelLoadingScreen(QDialog):
    """Loading screen dialog for Excel initialization"""
    
    def __init__(self, manual_mode=False, debug_mode=False, parent=None):
        """Initialize the loading screen"""
        super(ExcelLoadingScreen, self).__init__(parent)
        
        self.manual_mode = manual_mode
        self.debug_mode = debug_mode
        self.initialization_successful = False
        self.error_message = ""
        
        # Set up the dialog
        self.setWindowTitle("SharePoint Automation - Initializing Excel")
        self.setFixedSize(500, 200)
        self.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint)  # Remove close button
        
        # Create layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(15)
        
        # Title label
        title_label = QLabel("SharePoint Automation")
        title_font = QFont("Segoe UI", 16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        # Status label
        self.status_label = QLabel("Initializing Excel application...")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #666; font-size: 12px;")
        layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)
        
        # Button layout
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        # Cancel button
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.setMinimumWidth(80)
        self.cancel_button.clicked.connect(self.cancel_initialization)
        button_layout.addWidget(self.cancel_button)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # Initialize the background thread
        self.init_thread = ExcelInitializationThread(manual_mode, debug_mode)
        self.init_thread.progress_updated.connect(self.update_progress)
        self.init_thread.finished.connect(self.initialization_finished)
        
        # Center the dialog on screen
        self.center_on_screen()
    
    def center_on_screen(self):
        """Center the dialog on the screen"""
        from PyQt5.QtWidgets import QDesktopWidget
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move(
            (screen.width() - size.width()) // 2,
            (screen.height() - size.height()) // 2
        )
    
    def start_initialization(self):
        """Start the Excel initialization process"""
        write_log("Starting Excel initialization...", "YELLOW")
        self.init_thread.start()
    
    def update_progress(self, progress, status):
        """Update the progress bar and status"""
        self.progress_bar.setValue(progress)
        self.status_label.setText(status)
        write_log(f"Excel Initialization: {progress}% - {status}", "CYAN")
    
    def initialization_finished(self, success, error_message):
        """Handle initialization completion"""
        self.initialization_successful = success
        self.error_message = error_message
        
        if success:
            write_log("Excel initialization completed successfully", "GREEN")
            self.accept()  # Close dialog with success
        else:
            write_log(f"Excel initialization failed: {error_message}", "RED")
            self.status_label.setText(f"Error: {error_message}")
            self.status_label.setStyleSheet("color: red; font-size: 12px;")
            self.cancel_button.setText("Exit")
    
    def cancel_initialization(self):
        """Cancel the initialization process"""
        if self.init_thread.isRunning():
            write_log("Cancelling Excel initialization...", "YELLOW")
            self.init_thread.cancel()
            self.init_thread.wait(3000)  # Wait up to 3 seconds for thread to finish
            
            if self.init_thread.isRunning():
                write_log("Force terminating initialization thread", "RED")
                self.init_thread.terminate()
        
        # Set termination flag in app controller
        from src.utils.app_controller import terminate_process
        terminate_process()
        
        self.reject()  # Close dialog with cancellation
    
    def closeEvent(self, event):
        """Handle close event"""
        # Don't allow closing during initialization unless it's finished
        if self.init_thread.isRunning():
            self.cancel_initialization()
        event.accept()


def show_loading_during_excel_init(manual_mode=False, debug_mode=False):
    """
    Show loading screen during Excel initialization only
    
    Args:
        manual_mode (bool): Whether running in manual mode
        debug_mode (bool): Whether running in debug mode
        
    Returns:
        bool: True if Excel initialization was successful, False if cancelled or failed
    """
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    # Create and show loading screen
    loading_screen = ExcelLoadingScreen(manual_mode, debug_mode)
    loading_screen.show()
    
    # Start initialization in background
    loading_screen.start_initialization()
    
    # Run the dialog
    result = loading_screen.exec_()
    
    # Return success status
    return result == QDialog.Accepted and loading_screen.initialization_successful


def show_loading_screen_and_initialize(manual_mode=False, debug_mode=False):
    """
    Legacy function kept for compatibility - redirects to Excel initialization
    
    Args:
        manual_mode (bool): Whether running in manual mode
        debug_mode (bool): Whether running in debug mode
        
    Returns:
        bool: True if initialization was successful, False if cancelled or failed
    """
    return show_loading_during_excel_init(manual_mode, debug_mode)