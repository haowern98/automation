#!/usr/bin/env python3
"""
SharePoint Automation - Main Script
"""
import os
import sys
import argparse
from datetime import datetime
from src.utils.logger import write_log
from src.utils.app_controller import run_sharepoint_automation
from src.config import DATA_DIR

def main():
    """Main function to run the SharePoint automation"""
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
    
    # Ensure data directory exists
    os.makedirs(DATA_DIR, exist_ok=True)
    
    try:
        # Run the main automation process
        run_sharepoint_automation(args.manual, args.debug)
        
    except Exception as e:
        write_log(f"An error occurred: {str(e)}", "RED")
        import traceback
        write_log(traceback.format_exc(), "RED")
        
    finally:
        # Record the end time
        end_time = datetime.now()
        duration = end_time - start_time
        write_log(f"Total script execution time: {duration.total_seconds()} seconds", "CYAN")

if __name__ == "__main__":
    main()