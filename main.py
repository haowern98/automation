#!/usr/bin/env python3
"""
SharePoint Automation - Main Script
Enhanced with automated scheduling logic and 30-second GUI timeout
"""
import os
import sys
import time
import concurrent.futures
import calendar
import argparse
from datetime import datetime, timedelta
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtWidgets import QApplication

# Import modules from our package
from gui import show_date_range_selection, parse_date_range_string
from gui.date_selector import DateRangeResult  # Import directly from date_selector
from processors import process_gsn_data, process_er_data, process_ad_data
from utils import write_log, compare_data_sets, ExcelUpdater
from utils.excel_functions import ExcelApplication
from config import USER_PROFILE, SYNCED_FILE_PATH, FILE_PATTERNS, AD_SEARCH

def is_weekend(date=None):
    """
    Check if the given date is a weekend (Saturday or Sunday)
    
    Args:
        date (datetime.date, optional): Date to check, defaults to current date
        
    Returns:
        bool: True if it's a weekend
    """
    if date is None:
        date = datetime(2025, 7, 31).date() 
    
    # 5 represents Saturday, 6 represents Sunday in the weekday() function
    return date.weekday() >= 5

def is_last_day_of_month(date=None):
    """
    Check if the given date is the last day of the month
    
    Args:
        date (datetime.date, optional): Date to check, defaults to current date
        
    Returns:
        bool: True if it's the last day of the month
    """
    if date is None:
        date = datetime(2025, 7, 31).date()   
    
    # Get the last day of the current month
    last_day = calendar.monthrange(date.year, date.month)[1]
    return date.day == last_day

def is_friday(date=None):
    """
    Check if the given date is a Friday
    
    Args:
        date (datetime.date, optional): Date to check, defaults to current date
        
    Returns:
        bool: True if it's a Friday
    """
    if date is None:
        date = datetime(2025, 7, 31).date()   
        write_log(f"DEBUG - Date used: {date}", "YELLOW")
        write_log(f"DEBUG - Is Friday: {is_friday(date)}", "YELLOW")
        write_log(f"DEBUG - Is last day of month: {is_last_day_of_month(date)}", "YELLOW")
    
    # 4 represents Friday (0 is Monday in the weekday() function)
    return date.weekday() == 4

def get_monday_of_same_week(date=None):
    """
    Get the Monday of the same week as the given date
    
    Args:
        date (datetime.date, optional): Reference date, defaults to current date
        
    Returns:
        datetime.date: Date of the Monday of the same week
    """
    if date is None:
        date = datetime(2025, 7, 31).date()   
    
    # Calculate the number of days to subtract to get to Monday (weekday 0)
    days_to_subtract = date.weekday()
    
    # Return the Monday date
    return date - timedelta(days=days_to_subtract)

def get_date_range_based_on_day():
    """
    Automatically determine date range based on current day
    
    Returns:
        DateRangeResult: Date range object with start and end dates
    """
    current_date = datetime(2025, 7, 31).date()   
    
    # If it's the last day of the month
    if is_last_day_of_month(current_date):
        end_date = current_date
        # Get Monday of the same week
        start_date = get_monday_of_same_week(current_date)
        
        # If Monday is in a different month, use the first day of the current month
        if start_date.month != end_date.month:
            start_date = datetime(end_date.year, end_date.month, 1).date()
    
    # If it's a Friday
    elif is_friday(current_date):
        end_date = current_date
        # Get Monday of the same week
        start_date = get_monday_of_same_week(current_date)
        
        # If Monday is in a different month, use the first day of the current month
        if start_date.month != end_date.month:
            start_date = datetime(end_date.year, end_date.month, 1).date()
    
    # If it's neither Friday nor last day of month, return None
    else:
        return None
    
    # Create a DateRangeResult object
    result = DateRangeResult()
    result.start_date = start_date
    result.end_date = end_date
    
    # Format the date range string
    if start_date.month == end_date.month and start_date.year == end_date.year:
        # Same month format: "15-17 April 2025"
        date_range_formatted = f"{start_date.day}-{end_date.day} {start_date.strftime('%B')} {start_date.year}"
    else:
        # Different month format: "15 April - 17 May 2025"
        date_range_formatted = f"{start_date.day} {start_date.strftime('%B')} - {end_date.day} {end_date.strftime('%B')} {end_date.year}"
    
    result.date_range_formatted = date_range_formatted
    result.year = str(end_date.year)
    
    write_log(f"Automatically determined date range: {date_range_formatted}", "GREEN")
    return result

def show_date_range_with_timeout(timeout_seconds=30):
    """
    Show the date range selection dialog with a timeout
    
    Args:
        timeout_seconds (int): Number of seconds before timeout
        
    Returns:
        DateRangeResult or None: Selected date range or None if timed out/cancelled
    """
    from gui.date_selector import DateRangeSelector
    
    # Ensure we have a QApplication instance
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    # Create the dialog
    dialog = DateRangeSelector()
    
    # Create a timer for the timeout
    timer = QTimer()
    timer.setSingleShot(True)
    timer.setInterval(timeout_seconds * 1000)  # Convert to milliseconds
    
    # Connect the timer to close the dialog when it times out
    timer.timeout.connect(dialog.reject)
    
    # Start the timer
    write_log(f"Showing date range selection dialog with {timeout_seconds}-second timeout...", "YELLOW")
    timer.start()
    
    # Show the dialog
    result = None
    if dialog.exec_() == dialog.Accepted:
        result = dialog.result_obj
        timer.stop()  # Stop the timer if user made a selection
        write_log("User selected date range: " + result.date_range_formatted, "GREEN")
    else:
        # Dialog was closed either by timeout or user cancellation
        if not timer.isActive():
            write_log(f"Date range selection dialog timed out after {timeout_seconds} seconds", "YELLOW")
        else:
            write_log("User cancelled date range selection", "YELLOW")
    
    return result

def check_excel_processes(terminate_all=False):
    """
    Check for running Excel processes
    
    Args:
        terminate_all (bool): Whether to terminate all Excel processes
    
    Returns:
        int: Number of Excel processes found
    """
    try:
        import psutil
        excel_processes = []
        
        for proc in psutil.process_iter(['pid', 'name']):
            if 'excel' in proc.info['name'].lower():
                excel_processes.append(proc)
        
        if excel_processes:
            write_log(f"Found {len(excel_processes)} Excel processes running", "YELLOW")
            
            if terminate_all:
                for proc in excel_processes:
                    try:
                        proc.terminate()
                        write_log(f"Terminated Excel process: {proc.info['pid']}", "YELLOW")
                    except:
                        pass
                write_log("Attempted to terminate all Excel processes", "YELLOW")
        else:
            write_log("No Excel processes found running", "GREEN")
            
        return len(excel_processes)
    except Exception as e:
        write_log(f"Error checking Excel processes: {str(e)}", "RED")
        return 0

def warm_up_excel():
    """Warm up Excel to ensure it's ready for automation"""
    write_log("Warming up Excel...", "YELLOW")
    try:
        excel_app = ExcelApplication()
        # Create a dummy workbook to ensure Excel is responsive
        temp_wb = excel_app.excel.Workbooks.Add()
        time.sleep(1)  # Short delay
        temp_wb.Close(SaveChanges=False)
        excel_app.close()
        write_log("Excel warm-up successful", "GREEN")
        return True
    except Exception as e:
        write_log(f"Excel warm-up failed: {str(e)}", "RED")
        return False

def find_latest_file(root_dir, file_pattern, additional_search_dirs=None):
    """
    Find the latest file matching a pattern
    
    Args:
        root_dir (str): Root directory to search in
        file_pattern (str): File pattern to match
        additional_search_dirs (list): Additional directories to search in
        
    Returns:
        str or None: Path to the latest file or None if not found
    """
    try:
        latest_file = None
        latest_time = 0
        all_matching_files = []
        
        # First, expand the search to include specific paths
        search_dirs = [root_dir]
        if additional_search_dirs:
            search_dirs.extend(additional_search_dirs)
        
        # Function to process shortcuts and get real paths
        def resolve_shortcut(shortcut_path):
            if shortcut_path.lower().endswith('.lnk'):
                try:
                    import winshell
                    link = winshell.shortcut(shortcut_path)
                    return link.path
                except:
                    write_log(f"Failed to resolve shortcut: {shortcut_path}", "YELLOW")
                    return None
            return shortcut_path
        
        # Function to check if a filename matches our pattern, including numeric variations
        def matches_pattern(filename, pattern):
            import re
            
            # Handle 'data*.xlsx' pattern - match 'data.xlsx', 'data(2).xlsx', 'data (11).xlsx', etc.
            if pattern == 'data*':
                return re.match(r'data\s*(\(\d+\))?\.xlsx', filename, re.IGNORECASE) is not None
            
            # Handle 'alm_hardware*' pattern
            elif pattern == 'alm_hardware*':
                return re.match(r'alm_hardware\s*(\(\d+\))?\.xlsx', filename, re.IGNORECASE) is not None
            
            # Default case - simple startswith check
            else:
                base_pattern = pattern.replace('*', '')
                return filename.lower().startswith(base_pattern.lower())
        
        # Function to extract version number from filename
        def get_version_number(filename):
            import re
            match = re.search(r'\((\d+)\)', filename)
            if match:
                return int(match.group(1))
            return 0  # Files without a number are treated as version 0
        
        # Log how many directories we're searching
        write_log(f"Searching in {len(search_dirs)} directories", "CYAN")
        
        for search_dir in search_dirs:
            write_log(f"Searching in: {search_dir}", "CYAN")
            if not os.path.exists(search_dir):
                write_log(f"Search directory does not exist: {search_dir}", "YELLOW")
                continue
                
            # First, search for exact matches to handle specific paths
            if os.path.isfile(search_dir):
                if matches_pattern(os.path.basename(search_dir), file_pattern):
                    real_path = resolve_shortcut(search_dir)
                    if real_path and os.path.exists(real_path):
                        all_matching_files.append(real_path)
                        write_log(f"Found direct match: {real_path}", "GREEN")
            else:
                try:
                    # Search in directories, but limit depth to avoid excessive searching
                    max_depth = 2  # Adjust this as needed
                    base_depth = search_dir.count(os.sep)
                    
                    for dirpath, dirs, filenames in os.walk(search_dir):
                        # Check if we've gone too deep
                        current_depth = dirpath.count(os.sep) - base_depth
                        if current_depth > max_depth:
                            del dirs[:]  # Don't descend any deeper
                            continue
                        
                        for filename in filenames:
                            # First, do a simple check before running regex
                            if ('data' in filename.lower() or 'alm_hardware' in filename.lower()) and filename.lower().endswith('.xlsx'):
                                # Now check with the more precise pattern
                                if matches_pattern(filename, file_pattern):
                                    file_path = os.path.join(dirpath, filename)
                                    
                                    # If it's a shortcut, resolve to actual path
                                    real_path = resolve_shortcut(file_path)
                                    if not real_path:
                                        continue
                                        
                                    if os.path.exists(real_path):
                                        all_matching_files.append(real_path)
                                        file_time = os.path.getmtime(real_path)
                                        write_log(f"Found file: {real_path} (Modified: {time.ctime(file_time)})", "GREEN")
                except Exception as e:
                    write_log(f"Error searching directory {search_dir}: {str(e)}", "RED")
        
        write_log(f"Found {len(all_matching_files)} matching files in total", "CYAN")
        
        # If we found matching files, determine the best one to use
        if all_matching_files:
            # Logging all found files to debug
            write_log("All matching files:", "CYAN")
            for idx, file_path in enumerate(all_matching_files):
                filename = os.path.basename(file_path)
                version = get_version_number(filename)
                write_log(f"  {idx+1}. {filename} (Version: {version})", "WHITE")
            
            # First try to find the highest version number
            highest_version = -1
            highest_version_file = None
            
            for file_path in all_matching_files:
                filename = os.path.basename(file_path)
                version = get_version_number(filename)
                
                if version > highest_version:
                    highest_version = version
                    highest_version_file = file_path
            
            # If we found a versioned file, use it
            if highest_version > 0 and highest_version_file:
                latest_file = highest_version_file
                write_log(f"Selected highest version file: {latest_file} (Version: {highest_version})", "GREEN")
            else:
                # Otherwise fall back to the most recently modified file
                for file_path in all_matching_files:
                    file_time = os.path.getmtime(file_path)
                    if file_time > latest_time:
                        latest_time = file_time
                        latest_file = file_path
                
                write_log(f"Selected most recently modified file: {latest_file}", "GREEN")
        
        if not latest_file:
            write_log(f"No files found matching pattern '{file_pattern}'", "YELLOW")
        else:
            write_log(f"Latest file found: {latest_file}", "GREEN")
            
        return latest_file
    except Exception as e:
        write_log(f"Error finding latest file: {str(e)}", "RED")
        import traceback
        write_log(traceback.format_exc(), "RED")
        return None

def main():
    """Main function to run the SharePoint automation"""
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='SharePoint Automation Script')
    parser.add_argument('--manual', action='store_true', help='Run in manual mode without date checks and with no timeout')
    args = parser.parse_args()
    
    # Record the start time
    start_time = datetime.now()
    write_log("Starting SharePoint file access script (OneDrive sync method)", "YELLOW")
    
    try:
        # Skip date checks if manual mode is enabled
        if not args.manual:
            current_date = datetime(2025, 7, 31).date()   
            write_log(f"DEBUG - Date being used: {current_date}", "YELLOW")

            # Check if today is a weekend
            if is_weekend(current_date):
                write_log("Today is a weekend. Script will not run on weekends. Exiting.", "YELLOW")
                return

            write_log(f"DEBUG - Is Friday: {is_friday(current_date)}", "YELLOW")
            write_log(f"DEBUG - Is last day of month: {is_last_day_of_month(current_date)}", "YELLOW")
            
            # Check if today is a day to run (Friday or last day of month)
            is_run_day = is_friday(current_date) or is_last_day_of_month(current_date)
            
            if not is_run_day:
                write_log("Today is not a Friday or the last day of the month. Exiting.", "YELLOW")
                return
            
            day_type = "Friday" if is_friday(current_date) else "Last day of month"
            write_log(f"Today is a designated run day: {day_type}", "GREEN")
        else:
            write_log("Running in manual mode: skipping date checks", "YELLOW")
        
        # Check for Excel processes
        check_excel_processes()
        
        # Warm up Excel
        warm_up_excel()
        
        # Show the date selection dialog with or without timeout depending on mode
        date_range = None
        if args.manual:
            # In manual mode, use the original show_date_range_selection without timeout
            write_log("Showing date range selection dialog without timeout...", "YELLOW")
            date_range = show_date_range_selection()
        else:
            # In automatic mode, show dialog with 30-second timeout
            date_range = show_date_range_with_timeout(timeout_seconds=30)
        
        # If user cancelled or dialog timed out, use automatic date range
        if not date_range:
            write_log("No user input received. Using automatic date range.", "YELLOW")
            date_range = get_date_range_based_on_day()
            
            if not date_range:
                write_log("Could not determine appropriate date range. Exiting.", "RED")
                return
        
        write_log(f"Selected date range: {date_range.date_range_formatted}", "GREEN")
        
        # Define the root directories to start the search
        root_dir = os.path.join(USER_PROFILE)
        
        # Common locations to search for Excel files
        common_locations = [
            os.path.join(USER_PROFILE, "Downloads"),
            os.path.join(USER_PROFILE, "Desktop"),
            os.path.join(USER_PROFILE, "Documents"),
            os.path.join(USER_PROFILE, "OneDrive"),
            os.path.join(USER_PROFILE, "OneDrive - Deutsche Post DHL"),
            os.path.join(USER_PROFILE, "AppData", "Roaming", "Microsoft", "Windows", "Recent")
        ]
        
        # Use concurrent futures to search for files in parallel
        with concurrent.futures.ThreadPoolExecutor() as executor:
            # Start concurrent file search jobs
            write_log("Starting concurrent file search operations...", "YELLOW")
            
            alm_file_future = executor.submit(find_latest_file, root_dir, FILE_PATTERNS['gsn'], common_locations)
            data_file_future = executor.submit(find_latest_file, root_dir, FILE_PATTERNS['er'], common_locations)
            sharepoint_exists_future = executor.submit(os.path.exists, SYNCED_FILE_PATH)
            
            # Wait for all file search jobs to complete
            write_log("Waiting for file search operations to complete...", "YELLOW")
            
            excel_file_path = alm_file_future.result()
            data_file_path = data_file_future.result()
            sharepoint_exists = sharepoint_exists_future.result()
        
        # Validate required files exist
        if not excel_file_path:
            write_log(f"No files found matching pattern '{FILE_PATTERNS['gsn']}'. Please specify the file path manually.", "RED")
            # Allow manual input as fallback
            excel_file_path = input("Please enter the full path to the GSN Excel file (e.g., C:\\path\\to\\alm_hardware.xlsx): ").strip('"')
            if not os.path.exists(excel_file_path):
                write_log("Invalid file path. Exiting.", "RED")
                return
        
        if not data_file_path:
            write_log(f"No files found matching pattern '{FILE_PATTERNS['er']}'. Please specify the file path manually.", "RED")
            # Allow manual input as fallback
            data_file_path = input("Please enter the full path to the ER data file (e.g., C:\\path\\to\\data.xlsx): ").strip('"')
            if not os.path.exists(data_file_path):
                write_log("Invalid file path. Exiting.", "RED")
                return
        
        write_log(f"Found latest files: {excel_file_path} and {data_file_path}", "GREEN")
        
        # Check if the SharePoint file exists
        if not sharepoint_exists:
            write_log(f"SharePoint file not found at: {SYNCED_FILE_PATH}", "RED")
            write_log("Please ensure OneDrive sync is set up correctly.", "YELLOW")
            return
        
        write_log(f"Found SharePoint file at: {SYNCED_FILE_PATH}", "GREEN")
        
        # Start data processing in parallel
        write_log("Starting parallel data processing...", "YELLOW")
        
        with concurrent.futures.ThreadPoolExecutor() as executor:
            # Start Jobs
            gsn_job = executor.submit(process_gsn_data, excel_file_path)
            er_job = executor.submit(process_er_data, data_file_path)
            ad_job = executor.submit(process_ad_data, AD_SEARCH['ldap_filter'], AD_SEARCH['search_base'])
            
            # Wait for all jobs to complete
            write_log("Waiting for data processing jobs to complete...", "YELLOW")
            
            extracted_values = gsn_job.result()
            er_results = er_job.result()
            ad_computers = ad_job.result()
        
        # Extract data from results
        filtered_er_hostnames = er_results["FilteredERHostnames"]
        filtered_hostnames2 = er_results["FilteredHostnames2"]
        er_serial_number = er_results["ErSerialNumber"]
        
        # Display results summary
        write_log("Data processing completed:", "GREEN")
        write_log(f"- GSN Entries: {len(extracted_values)}", "WHITE")
        write_log(f"- ER Entries: {len(filtered_er_hostnames)}", "WHITE")
        write_log(f"- ER Entries (31-60 days): {len(filtered_hostnames2)}", "WHITE")
        write_log(f"- AD Computer Entries: {len(ad_computers)}", "WHITE")
        
        # Display ER No Logon Hostnames and Serial Numbers
        write_log("\n=========================================", "YELLOW")
        write_log("ER NO LOGON DETAILS (31-60 DAYS)", "YELLOW")
        write_log("=========================================", "YELLOW")
        
        write_log("\nHostname and Serial Number:", "MAGENTA")
        if filtered_hostnames2:
            for i in range(len(filtered_hostnames2)):
                hostname = filtered_hostnames2[i]
                sn = er_serial_number[i]
                write_log(f"  {hostname}   {sn}", "CYAN")
        else:
            write_log("  No devices found with login between 31-60 days", "MAGENTA")
        write_log("=========================================", "YELLOW")
        
        # Compare GSN and ER entries
        comparison_results = compare_data_sets(extracted_values, filtered_er_hostnames)
        missing_in_er = comparison_results["MissingInER"]
        missing_in_gsn = comparison_results["MissingInGSN"]
        
        # Update the Excel file with the GSN vs ER analysis using the synced file path
        write_log("Updating Excel file with comparison results...", "YELLOW")
        excel_updater = ExcelUpdater(SYNCED_FILE_PATH)
        excel_updater.analyze_excel_file(
            extracted_values,
            filtered_er_hostnames,
            ad_computers,
            date_range,
            missing_in_er,
            missing_in_gsn,
            filtered_hostnames2,
            er_serial_number
        )
        
        # Script succeeded
        write_log("Script completed successfully", "GREEN")
        write_log("OneDrive will automatically sync changes to SharePoint", "CYAN")
        
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