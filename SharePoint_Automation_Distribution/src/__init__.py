"""
SharePoint Automation Package

This package provides automation tools for SharePoint data processing,
including GSN vs ER comparison, AD data processing, and Excel report generation.

Main Components:
- processors: Data processing modules for GSN, ER, and AD data
- utils: Utility functions for logging, Excel operations, and data comparison
- gui: Graphical user interface components for date selection and settings

Usage:
    from src import run_sharepoint_automation
    run_sharepoint_automation(manual_mode=False)

Or run the main application:
    python -m src.main
"""

# Version information
__version__ = "1.0.0"
__author__ = "Your Name"
__description__ = "SharePoint Automation tool for GSN vs ER reports"

# Import main functions for easy access
from .utils.app_controller import run_sharepoint_automation
from .processors.gsn_processor import process_gsn_data
from .processors.er_processor import process_er_data
from .processors.ad_processor import process_ad_data, compare_gsn_with_ad
from .utils.comparison import compare_data_sets
from .gui.date_selector import show_date_range_selection, DateRangeResult
from .gui.tabbed_app import show_tabbed_date_range_selection

# Define what gets imported with "from src import *"
__all__ = [
    # Main application function
    'run_sharepoint_automation',
    
    # Data processors
    'process_gsn_data',
    'process_er_data', 
    'process_ad_data',
    'compare_gsn_with_ad',
    
    # Utilities
    'compare_data_sets',
    
    # GUI components
    'show_date_range_selection',
    'show_tabbed_date_range_selection',
    'DateRangeResult',
    
    # Version info
    '__version__',
    '__author__',
    '__description__'
]

# Package-level configuration
import os
import sys

# Add the src directory to Python path if needed
_src_dir = os.path.dirname(os.path.abspath(__file__))
if _src_dir not in sys.path:
    sys.path.insert(0, _src_dir)