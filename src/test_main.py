#!/usr/bin/env python3
"""
Test script to isolate the issue
"""
print("DEBUG: Starting test script")

import os
import sys
print("DEBUG: Basic imports successful")

# Add the parent directory to the Python path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)
print("DEBUG: Path setup complete")

try:
    from src.utils.logger import write_log
    print("DEBUG: logger import successful")
except Exception as e:
    print(f"DEBUG: logger import failed: {e}")

try:
    from src.config import DATA_DIR
    print("DEBUG: config import successful")
except Exception as e:
    print(f"DEBUG: config import failed: {e}")

try:
    print("DEBUG: About to import app_controller...")
    from src.utils.app_controller import run_sharepoint_automation_with_loading
    print("DEBUG: app_controller import successful")
except Exception as e:
    print(f"DEBUG: app_controller import failed: {e}")
    import traceback
    traceback.print_exc()

try:
    print("DEBUG: About to import PyQt5...")
    from PyQt5.QtWidgets import QApplication
    print("DEBUG: PyQt5 import successful")
except Exception as e:
    print(f"DEBUG: PyQt5 import failed: {e}")
    import traceback
    traceback.print_exc()

def test_main():
    print("DEBUG: In test_main function")
    return True

print("DEBUG: Function defined")

if __name__ == "__main__":
    print("DEBUG: About to call test_main")
    result = test_main()
    print(f"DEBUG: test_main returned: {result}")
    print("DEBUG: Script complete")