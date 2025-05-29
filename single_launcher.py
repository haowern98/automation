#!/usr/bin/env python3
"""
Single Launcher for SharePoint Automation
This launcher ensures the application works for both manual and auto modes in a single EXE
"""
import sys
import os

# Add current directory to path so we can import src modules
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Import and run the main application with smart mode detection
if __name__ == "__main__":
    # Import and run main
    from src.main import main
    main()