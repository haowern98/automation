"""
Enhanced Settings dialog for SharePoint Automation
"""
import os
import json
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QGridLayout, 
                             QTabWidget, QWidget, QGroupBox, QLineEdit,
                             QCheckBox, QFileDialog, QMessageBox, QComboBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QTextEdit  

class SettingsManager:
    """Manages application settings"""
    
    def __init__(self, settings_file="settings.json"):
        """Initialize settings manager"""
        self.settings_file = settings_file
        self.default_settings = {
            "file_paths": {
                "gsn_search_directory": os.path.join(os.environ.get('USERPROFILE', ''), 'Downloads'),
                "er_search_directory": os.path.join(os.environ.get('USERPROFILE', ''), 'Downloads'),
                "gsn_file_pattern": "alm_hardware",
                "er_file_pattern": "data",
                "weekly_report_file_path": os.path.join(
            os.environ.get('USERPROFILE', ''),
            'DPDHL',
            'SM Team - SG - AD EDS, MFA, GSN VS AD, GSN VS ER Weekly Report',
            'Weekly Report 2025 - Copy.xlsx'
                 ),
                "output_directory": os.path.join(
                os.environ.get('USERPROFILE', ''),
                'OneDrive - DPDHL',
                'Documents',
                'weeklyreportlog'
                )
            },
            "general": {
                "auto_mode_timeout": "30",
                "show_terminal": False,
                "section_keywords": [                    # <-- ADD THIS
                    "Applied MFA Method",
                    "ARP Invalid", 
                    "Accounts with Manager",
                    "No AD",
                    "GID assigned",
                    "Accounts with",
                    "Manager/ARP"
                ]
            }
        }
        self.settings = self.load_settings()
    
    def load_settings(self):
        """Load settings from file"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    loaded_settings = json.load(f)
                    # Merge with defaults to ensure all keys exist
                    settings = self.default_settings.copy()
                    self._update_dict_recursive(settings, loaded_settings)
                    return settings
            else:
                return self.default_settings.copy()
        except Exception as e:
            print(f"Error loading settings: {e}")
            return self.default_settings.copy()
    
    def _update_dict_recursive(self, d, u):
        """Recursively update a dictionary with another dictionary"""
        for k, v in u.items():
            if isinstance(v, dict):
                d[k] = self._update_dict_recursive(d.get(k, {}), v)
            else:
                d[k] = v
        return d
    
    def save_settings(self):
        """Save settings to file"""
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(self.settings, f, indent=4)
            return True
        except Exception as e:
            print(f"Error saving settings: {e}")
            return False
    
    def get(self, category, key, default=None):
        """Get a setting value"""
        try:
            return self.settings.get(category, {}).get(key, default)
        except Exception as e:
            print(f"Error getting setting {category}.{key}: {e}")
            return default
    
    def set(self, category, key, value):
        """Set a setting value"""
        if category not in self.settings:
            self.settings[category] = {}
        self.settings[category][key] = value

    def get_section_keywords(self):
        """Get section keywords as a list"""
        keywords = self.get('general', 'section_keywords', [])
        if not keywords:  # Fallback to defaults if empty
            keywords = self.default_settings['general']['section_keywords']
            return keywords

    def set_section_keywords(self, keywords_list):
        """Set section keywords from a list"""
        self.set('general', 'section_keywords', keywords_list)

class SettingsDialog(QDialog):
    """Enhanced Settings dialog for SharePoint Automation"""
    
    def __init__(self, parent=None):
        """Initialize the settings dialog"""
        super(SettingsDialog, self).__init__(parent)
        
        self.setWindowTitle("SharePoint Automation - Settings")
        self.setFixedSize(650, 500)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        # Initialize settings manager
        self.settings_manager = SettingsManager()
        
        # Create the main layout
        main_layout = QVBoxLayout(self)
        
        # Create a tab widget for different settings categories
        tab_widget = QTabWidget()
        
        # Create tabs
        general_tab = self._create_general_tab()
        file_paths_tab = self._create_file_paths_tab()
        
        # Add tabs to the tab widget
        tab_widget.addTab(general_tab, "General")
        tab_widget.addTab(file_paths_tab, "File Paths")
        
        # Add the tab widget to the main layout
        main_layout.addWidget(tab_widget)
        
        # Create buttons layout
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("Save Settings")
        self.cancel_button = QPushButton("Cancel")
        self.reset_button = QPushButton("Reset to Defaults")
        
        button_layout.addWidget(self.reset_button)
        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        
        # Add buttons to the main layout
        main_layout.addLayout(button_layout)
        
        # Connect button signals
        self.ok_button.clicked.connect(self.save_settings)
        self.cancel_button.clicked.connect(self.reject)
        self.reset_button.clicked.connect(self.reset_to_defaults)
        
        # Load current settings into the dialog
        self.load_current_settings()
    
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
    
    # ===== NEW SECTION: Section Keywords =====
    # Create section keywords group
    keywords_group = QGroupBox("Section Keywords")
    keywords_layout = QGridLayout(keywords_group)
    
    # Add section keywords text area
    keywords_layout.addWidget(QLabel("MFA Section Headers:"), 0, 0)
    self.section_keywords_edit = QTextEdit()
    self.section_keywords_edit.setMaximumHeight(120)  # Limit height
    self.section_keywords_edit.setPlaceholderText("Enter keywords that identify section headers (one per line)")
    keywords_layout.addWidget(self.section_keywords_edit, 0, 1)
    
    # Add help text for keywords
    keywords_help = QLabel("Keywords that identify section headers (blue rows) in MFA reports. One keyword per line.")
    keywords_help.setStyleSheet("color: gray; font-size: 10px;")
    keywords_help.setWordWrap(True)
    keywords_layout.addWidget(keywords_help, 1, 0, 1, 2)
    
    # Add reset keywords button
    reset_keywords_button = QPushButton("Reset to Defaults")
    reset_keywords_button.setMaximumWidth(120)
    reset_keywords_button.clicked.connect(self.reset_section_keywords)
    keywords_layout.addWidget(reset_keywords_button, 2, 1, alignment=Qt.AlignRight)
    # ===== END NEW SECTION =====
    
    # Add widgets to layout
    layout.addWidget(heading)
    layout.addWidget(description)
    layout.addSpacing(20)
    layout.addWidget(timeout_group)
    layout.addSpacing(10)  # Add spacing between groups
    layout.addWidget(keywords_group)  # Add the new keywords group
    layout.addStretch(1)
    
    # Load timeout setting
    self.load_timeout_setting()
    # Load section keywords
    self.load_section_keywords()  # New method to load keywords
    
    return tab

# ===== ADD THESE NEW METHODS TO THE SettingsDialog CLASS =====

def load_section_keywords(self):
    """Load section keywords from settings into the text area"""
    keywords = self.settings_manager.get_section_keywords()
    # Join keywords with newlines for display in text area
    keywords_text = '\n'.join(keywords)
    self.section_keywords_edit.setPlainText(keywords_text)

def reset_section_keywords(self):
    """Reset section keywords to defaults"""
    default_keywords = self.settings_manager.default_settings['general']['section_keywords']
    keywords_text = '\n'.join(default_keywords)
    self.section_keywords_edit.setPlainText(keywords_text)

    def get_section_keywords_from_ui(self):
        """Get section keywords from the UI text area as a list"""
        text = self.section_keywords_edit.toPlainText().strip()
        if not text:
            return []
        
        # Split by lines and clean up
        keywords = []
        for line in text.split('\n'):
            keyword = line.strip()
            if keyword:  # Only add non-empty lines
                keywords.append(keyword)
        
        return keywords

    # ===== UPDATE THE save_settings METHOD =====
    # Add this section to the existing save_settings method, after the timeout saving:

    # Save section keywords (ADD THIS SECTION)
    try:
        section_keywords = self.get_section_keywords_from_ui()
        if not section_keywords:
            QMessageBox.warning(self, "Invalid Keywords", 
                            "Section keywords cannot be empty. Please add at least one keyword.")
            return
        
        self.settings_manager.set_section_keywords(section_keywords)
    except Exception as e:
        QMessageBox.warning(self, "Keywords Error", 
                        f"Error processing section keywords:\n{str(e)}")
        return

    # ===== UPDATE THE reset_to_defaults METHOD =====
    # Add this line to the existing reset_to_defaults method:

    # Reset section keywords to defaults (ADD THIS LINE)
    self.reset_section_keywords()

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

        # Create Output Directory Settings Group
        output_group = QGroupBox("Output Directory Settings")
        output_layout = QGridLayout(output_group)

        # Output Directory Path
        output_layout.addWidget(QLabel("TXT Output Directory:"), 0, 0)
        self.output_directory_edit = QLineEdit()
        self.output_directory_edit.setPlaceholderText("Enter path to directory for saving TXT files")
        output_layout.addWidget(self.output_directory_edit, 0, 1)

        self.output_browse_button = QPushButton("Browse...")
        self.output_browse_button.clicked.connect(self.browse_output_directory)
        output_layout.addWidget(self.output_browse_button, 0, 2)

        # Add info label
        output_info = QLabel("Directory where generated TXT files will be saved")
        output_info.setStyleSheet("color: gray; font-size: 10px;")
        output_layout.addWidget(output_info, 1, 1, 1, 2)

        layout.addWidget(output_group)

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
        
        self.gsn_browse_button = QPushButton("Browse...")
        self.gsn_browse_button.clicked.connect(self.browse_gsn_directory)
        gsn_layout.addWidget(self.gsn_browse_button, 0, 2)
        
        # GSN File Pattern
        gsn_layout.addWidget(QLabel("File Name Pattern:"), 1, 0)
        self.gsn_pattern_edit = QLineEdit()
        self.gsn_pattern_edit.setPlaceholderText("Enter file name pattern (e.g., alm_hardware)")
        gsn_layout.addWidget(self.gsn_pattern_edit, 1, 1, 1, 2)
        
        # Add info label
        gsn_info = QLabel("Pattern will match files like: alm_hardware.xlsx, alm_hardware(2).xlsx, etc.")
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
        
        self.er_browse_button = QPushButton("Browse...")
        self.er_browse_button.clicked.connect(self.browse_er_directory)
        er_layout.addWidget(self.er_browse_button, 0, 2)
        
        # ER File Pattern
        er_layout.addWidget(QLabel("File Name Pattern:"), 1, 0)
        self.er_pattern_edit = QLineEdit()
        self.er_pattern_edit.setPlaceholderText("Enter file name pattern (e.g., data)")
        er_layout.addWidget(self.er_pattern_edit, 1, 1, 1, 2)
        
        # Add info label
        er_info = QLabel("Pattern will match files like: data.xlsx, data(2).xlsx, etc.")
        er_info.setStyleSheet("color: gray; font-size: 10px;")
        er_layout.addWidget(er_info, 2, 1, 1, 2)
        
        layout.addWidget(er_group)
        
        # Add stretch to push everything up
        layout.addStretch(1)
        
        return tab

    def browse_output_directory(self):
        """Browse for output directory"""
        current_path = self.output_directory_edit.text()
        if not current_path:
            current_path = os.environ.get('USERPROFILE', '')
        
        directory = QFileDialog.getExistingDirectory(
            self, "Select Output Directory", current_path)
        
        if directory:
             self.output_directory_edit.setText(directory)        
    
    def browse_gsn_directory(self):
        """Browse for GSN search directory"""
        current_path = self.gsn_directory_edit.text()
        if not current_path:
            current_path = os.environ.get('USERPROFILE', '')
        
        directory = QFileDialog.getExistingDirectory(
            self, "Select GSN Search Directory", current_path)
        
        if directory:
            self.gsn_directory_edit.setText(directory)
    
    def browse_er_directory(self):
        """Browse for ER search directory"""
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

        # Load Output Directory settings
        output_dir = self.settings_manager.get('file_paths', 'output_directory', '')
        self.output_directory_edit.setText(output_dir)
    
    def load_timeout_setting(self):
        """Load timeout setting from settings"""
        # Default to 30 seconds if not set
        timeout_value = self.settings_manager.get('general', 'auto_mode_timeout', '30')
        
        # Find the matching index in the dropdown
        for i in range(self.timeout_dropdown.count()):
            if timeout_value in self.timeout_dropdown.itemText(i):
                self.timeout_dropdown.setCurrentIndex(i)
                break
    
    def save_settings(self):
        """Save settings and close dialog"""
        try:
            # Validate inputs
            gsn_dir = self.gsn_directory_edit.text().strip()
            er_dir = self.er_directory_edit.text().strip()
            gsn_pattern = self.gsn_pattern_edit.text().strip()
            er_pattern = self.er_pattern_edit.text().strip()
            output_dir = self.output_directory_edit.text().strip()
            
            # Check if directories exist
            if gsn_dir and not os.path.exists(gsn_dir):
                QMessageBox.warning(self, "Invalid Directory", 
                                   f"GSN search directory does not exist:\n{gsn_dir}")
                return
            
            if er_dir and not os.path.exists(er_dir):
                QMessageBox.warning(self, "Invalid Directory", 
                                   f"ER search directory does not exist:\n{er_dir}")
                return

            if output_dir and not os.path.exists(output_dir):
                QMessageBox.warning(self, "Invalid Directory", 
                                   f"Output directory does not exist:\n{output_dir}")
            
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
            self.settings_manager.set('file_paths', 'output_directory', output_dir) 
            
            # Save timeout setting
            timeout_text = self.timeout_dropdown.currentText()
            timeout_value = timeout_text.split()[0]  # Extract just the number
            self.settings_manager.set('general', 'auto_mode_timeout', timeout_value)
            
            # Save to file
            if self.settings_manager.save_settings():
                QMessageBox.information(self, "Settings Saved", 
                                       "Settings have been saved successfully!")
                self.accept()
            else:
                QMessageBox.critical(self, "Save Error", 
                                    "Failed to save settings to file!")
        
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred while saving settings:\n{str(e)}")
    
    def reset_to_defaults(self):
        """Reset settings to defaults"""
        reply = QMessageBox.question(self, "Reset Settings", 
                                    "Are you sure you want to reset all settings to defaults?",
                                    QMessageBox.Yes | QMessageBox.No,
                                    QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # Reset to default values
            downloads_path = os.path.join(os.environ.get('USERPROFILE', ''), 'Downloads')
            
            self.gsn_directory_edit.setText(downloads_path)
            self.er_directory_edit.setText(downloads_path)
            self.gsn_pattern_edit.setText("alm_hardware")
            self.er_pattern_edit.setText("data")
            
            # Reset timeout to default (30 seconds)
            self.timeout_dropdown.setCurrentIndex(2)  # "30 seconds" is at index 2

def show_settings_dialog():
    """
    Show the settings dialog
    
    Returns:
        bool: True if settings were saved, False otherwise
    """
    dialog = SettingsDialog()
    return dialog.exec_() == QDialog.Accepted

def get_settings():
    """
    Get current settings
    
    Returns:
        SettingsManager: Settings manager instance
    """
    return SettingsManager()

# Test the dialog if run directly
if __name__ == "__main__":
    import sys
    from PyQt5.QtWidgets import QApplication
    
    app = QApplication(sys.argv)
    
    if show_settings_dialog():
        print("Settings were saved!")
    else:
        print("Settings dialog was cancelled.")
    
    app.quit()