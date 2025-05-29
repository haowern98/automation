#!/usr/bin/env python3
"""
SharePoint Automation - Main Script
"""
import os
import sys
import argparse
from datetime import datetime

# Add the parent directory to the Python path so we can import src modules
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from PyQt5.QtWidgets import QApplication
from src.utils.logger import write_log
from src.config import DATA_DIR

print("DEBUG: All imports successful, about to define main()")

def main():
    """Main function to run the SharePoint automation"""
    print("DEBUG: Entered main function")
    
    # Record the start time
    start_time = datetime.now()
    write_log("Starting SharePoint Automation Script", "YELLOW")

    # Parse command line arguments
    parser = argparse.ArgumentParser(description='SharePoint Automation Script')
    parser.add_argument('--manual', action='store_true', 
                        help='Run in manual mode without date checks and with no timeout')
    parser.add_argument('--debug', action='store_true',
                        help='Run in debug mode with additional logging')
    args = parser.parse_args()
    
    print(f"DEBUG: Parsed args - manual: {args.manual}, debug: {args.debug}")
    
    # Ensure data directory exists
    os.makedirs(DATA_DIR, exist_ok=True)
    
    # Create QApplication for GUI components
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    print("DEBUG: About to call automation function")
    
    try:
        from src.utils.app_controller import run_sharepoint_automation_with_loading
        success = run_sharepoint_automation_with_loading(args.manual, args.debug)
        
        # APPLY TERMINAL SETTING AFTER LOADING SCREEN AND MAIN PROCESS COMPLETE
        # This ensures the loading screen is not affected
        print("DEBUG: Applying terminal visibility setting...")
        try:
            from src.utils.terminal_control import apply_terminal_setting
            terminal_applied = apply_terminal_setting()
            if terminal_applied:
                write_log("Terminal visibility setting applied successfully", "GREEN")
            else:
                write_log("Terminal visibility setting could not be applied", "YELLOW")
        except Exception as terminal_error:
            write_log(f"Error applying terminal setting: {str(terminal_error)}", "YELLOW")
        
        if success:
            write_log("SharePoint automation completed successfully", "GREEN")
        else:
            write_log("SharePoint automation was cancelled or failed", "YELLOW")
        
    except Exception as e:
        print(f"DEBUG: Exception in main: {str(e)}")
        write_log(f"An error occurred: {str(e)}", "RED")
        import traceback
        traceback.print_exc()
        
    finally:
        # Record the end time
        end_time = datetime.now()
        duration = end_time - start_time
        write_log(f"Total script execution time: {duration.total_seconds()} seconds", "CYAN")

print("DEBUG: main() function defined")

if __name__ == "__main__":
    print("DEBUG: About to call main()")
    main()
    print("DEBUG: main() completed")