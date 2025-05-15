"""
Settings dialog for SharePoint Automation
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, 
                             QLabel, QPushButton, QGridLayout, 
                             QTabWidget, QWidget, QGroupBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class SettingsDialog(QDialog):
    """Settings dialog for SharePoint Automation"""
    
    def __init__(self, parent=None):
        """Initialize the settings dialog"""
        super(SettingsDialog, self).__init__(parent)
        
        self.setWindowTitle("SharePoint Automation - Settings")
        self.setFixedSize(600, 400)  # Reduced height since we removed one tab
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        
        # Create the main layout
        main_layout = QVBoxLayout(self)
        
        # Create a tab widget for different settings categories
        tab_widget = QTabWidget()
        
        # Create tabs
        general_tab = self._create_general_tab()
        paths_tab = self._create_paths_tab()
        
        # Add tabs to the tab widget
        tab_widget.addTab(general_tab, "General")
        tab_widget.addTab(paths_tab, "File Paths")
        
        # Add the tab widget to the main layout
        main_layout.addWidget(tab_widget)
        
        # Create buttons layout
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("Save Settings")
        self.cancel_button = QPushButton("Cancel")
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.setAlignment(Qt.AlignCenter)
        
        # Add buttons to the main layout
        main_layout.addLayout(button_layout)
        
        # Connect button signals
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
    
    def _create_general_tab(self):
        """Create the general settings tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
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
    
    def _create_paths_tab(self):
        """Create the file paths settings tab"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Create a heading
        heading = QLabel("File Paths Settings")
        heading_font = QFont("Segoe UI", 12)
        heading_font.setBold(True)
        heading.setFont(heading_font)
        
        # Create a description
        description = QLabel("Configure file paths for GSN, ER, and SharePoint files")
        description.setWordWrap(True)
        
        # Placeholder for future settings
        placeholder = QLabel("Path configuration will be available in future versions")
        placeholder.setStyleSheet("color: gray; font-style: italic;")
        placeholder.setAlignment(Qt.AlignCenter)
        
        # Add widgets to layout
        layout.addWidget(heading)
        layout.addWidget(description)
        layout.addSpacing(20)
        layout.addWidget(placeholder)
        layout.addStretch(1)
        
        return tab

def show_settings_dialog():
    """
    Show the settings dialog
    
    Returns:
        bool: True if settings were saved, False otherwise
    """
    dialog = SettingsDialog()
    return dialog.exec_() == QDialog.Accepted