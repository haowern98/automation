"""
Settings Tab

Tab for configuring application settings including file paths and general preferences.
"""
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QTabWidget, QLabel, QGroupBox,
                             QGridLayout, QComboBox, QCheckBox, QPushButton, QLineEdit,
                             QFileDialog, QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


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