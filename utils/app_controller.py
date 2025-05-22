"""
Main application controller for SharePoint Automation
"""
import os
import sys
import time
import re
import concurrent.futures
from datetime import datetime

from utils.logger import write_log
from utils.excel_functions import ExcelApplication
from utils.comparison import compare_data_sets, ExcelUpdater
from gui.date_selector import DateRangeResult
from gui.tabbed_app import show_tabbed_date_range_selection
from processors.gsn_processor import process_gsn_data
from processors.er_processor import process_er_data
from processors.ad_processor import process_ad_data, compare_gsn_with_ad
from config import USER_PROFILE, SYNCED_FILE_PATH, FILE_PATTERNS, AD_SEARCH, DATA_DIR

def run_sharepoint_automation(manual_mode=False, debug_mode=False):
    """
    Run the SharePoint automation process
    
    Args:
        manual_mode (bool): Whether to run in manual mode
        debug_mode (bool): Whether to run in debug mode
        
    Returns:
        bool: Success status
    """
    try:
        # Skip date checks if manual mode is enabled
        if not manual_mode:
            if not check_run_date():
                return False
        else:
            write_log("Running in manual mode: skipping date checks", "YELLOW")
        
        # Check for Excel processes and warm up Excel
        manage_excel()
        
        # Get date range (through GUI or automatic calculation)
        date_range = get_date_range(manual_mode)
        if not date_range or not date_range.is_valid:
            write_log("No valid date range provided. Exiting.", "RED")
            return False
            
        write_log(f"Using date range: {date_range.date_range_formatted}", "GREEN")
        
        # Find required files
        file_paths = find_required_files()
        if not file_paths:
            return False
            
        # Process data
        data_results = process_data(file_paths)
        if not data_results:
            return False
            
        # Update Excel file with results
        success = update_excel_file(date_range, data_results)
        
        if success:
            write_log("SharePoint automation completed successfully", "GREEN")
            write_log("OneDrive will automatically sync changes to SharePoint", "CYAN")
        else:
            write_log("SharePoint automation completed with errors", "YELLOW")
            
        return success
        
    except Exception as e:
        write_log(f"Error in SharePoint automation: {str(e)}", "RED")
        import traceback
        write_log(traceback.format_exc(), "RED")
        return False

def check_run_date():
    """
    Check if today is a day to run the automation
    
    Returns:
        bool: True if automation should run, False otherwise
    """
    # In production, use: current_date = datetime.now().date()
    
    # TEST DATE - replace with datetime.now().date() in production
    current_date = datetime(2025, 8, 15).date()
    write_log(f"DEBUG - Test date being used: {current_date}", "YELLOW")

    # Check if today is a weekend
    if is_weekend(current_date):
        write_log("Test date is a weekend. Script will not run on weekends. Exiting.", "YELLOW")
        return False

    # Check if today is a day to run (Friday or last day of month)
    is_run_day = is_friday(current_date) or is_last_day_of_month(current_date)
    
    if not is_run_day:
        write_log("Test date is not a Friday or the last day of the month. Exiting.", "YELLOW")
        return False
    
    day_type = "Friday" if is_friday(current_date) else "Last day of month"
    write_log(f"Test date is a designated run day: {day_type}", "GREEN")
    return True

def is_weekend(date):
    """Check if the given date is a weekend (Saturday or Sunday)"""
    return date.weekday() >= 5

def is_friday(date):
    """Check if the given date is a Friday"""
    return date.weekday() == 4

def is_last_day_of_month(date):
    """Check if the given date is the last day of the month"""
    import calendar
    last_day = calendar.monthrange(date.year, date.month)[1]
    return date.day == last_day

def get_monday_of_same_week(date):
    """Get the Monday of the same week as the given date"""
    from datetime import timedelta
    days_to_subtract = date.weekday()
    return date - timedelta(days=days_to_subtract)

def manage_excel():
    """Check for Excel processes and warm up Excel"""
    check_excel_processes()
    warm_up_excel()

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

def get_date_range(manual_mode):
    """
    Get date range from user or calculate automatically
    
    Args:
        manual_mode (bool): Whether running in manual mode
        
    Returns:
        DateRangeResult: Selected date range
    """
    # Show the date selection dialog with or without timeout depending on mode
    date_range = None
    if manual_mode:
        # In manual mode, use the updated tabbed interface with manual_mode=True
        write_log("Showing date range selection dialog in manual mode...", "YELLOW")
        date_range = show_tabbed_date_range_selection(manual_mode=True)
        
        # If cancelled in manual mode, exit the program
        if date_range and date_range.cancelled:
            write_log("User chose to exit in manual mode", "YELLOW")
            return None
    else:
        # In automatic mode, show dialog with 30-second timeout
        date_range = show_date_range_with_timeout(timeout_seconds=30, manual_mode=False)
    
    # If cancelled or dialog timed out, use automatic date range in auto mode
    if date_range and date_range.cancelled and not manual_mode:
        write_log("User cancelled selection or timed out. Using automatic date range based on test date.", "YELLOW")
        # Use our test date-based calculation
        date_range = get_automatic_date_range()
        
    return date_range

def show_date_range_with_timeout(timeout_seconds=30, manual_mode=False):
    """
    Show the date range selection dialog with a timeout
    
    Args:
        timeout_seconds (int): Number of seconds before timeout
        manual_mode (bool): Whether running in manual mode
        
    Returns:
        DateRangeResult: Selected date range or None if cancelled
    """
    from PyQt5.QtCore import QTimer, Qt
    from PyQt5.QtWidgets import QApplication, QDialog
    from gui.tabbed_app import SharePointAutomationApp
    
    # Ensure we have a QApplication instance
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    # Create the dialog
    dialog = SharePointAutomationApp(manual_mode=manual_mode)
    
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
    result = dialog.exec_() == QDialog.Accepted
    timer.stop()  # Stop the timer
    
    # Get the result object from the dialog
    date_range = dialog.get_date_range_result()
    
    # Check if user cancelled or accepted
    if not result:
        if not timer.isActive():
            write_log(f"Date range selection dialog timed out after {timeout_seconds} seconds", "YELLOW")
        else:
            write_log("User cancelled date range selection", "YELLOW")
        
        # If in manual mode and cancelled, user wants to exit
        if manual_mode and date_range.cancelled:
            write_log("Manual mode: Exiting application as requested by user", "YELLOW")
            return None
        
        # In auto mode, mark as cancelled so we use auto date calculation
        date_range.cancelled = True
    else:
        write_log("User selected date range: " + date_range.date_range_formatted, "GREEN")
    
    return date_range

def get_automatic_date_range():
    """
    Automatically determine date range based on test day
    
    Returns:
        DateRangeResult: Date range object with start and end dates
    """
    from datetime import timedelta
    
    # TEST DATE - replace with datetime.now().date() in production
    current_date = datetime(2025, 8, 15).date()
    write_log(f"Using test date for auto calculation: {current_date.strftime('%Y-%m-%d')}", "YELLOW")
    
    # If it's the last day of the month
    if is_last_day_of_month(current_date):
        write_log("Test date is the last day of the month, calculating date range accordingly", "CYAN")
        end_date = current_date
        # Get Monday of the same week
        start_date = get_monday_of_same_week(current_date)
        write_log(f"Monday of the same week: {start_date.strftime('%Y-%m-%d')}", "CYAN")
        
        # If Monday is in a different month, use the first day of the current month
        if start_date.month != end_date.month:
            start_date = datetime(end_date.year, end_date.month, 1).date()
            write_log(f"Monday is in a different month, using first day of current month: {start_date.strftime('%Y-%m-%d')}", "CYAN")
    
    # If it's a Friday
    elif is_friday(current_date):
        write_log("Test date is a Friday, calculating date range accordingly", "CYAN")
        end_date = current_date
        # Get Monday of the same week
        start_date = get_monday_of_same_week(current_date)
        write_log(f"Monday of the same week: {start_date.strftime('%Y-%m-%d')}", "CYAN")
        
        # If Monday is in a different month, use the first day of the current month
        if start_date.month != end_date.month:
            start_date = datetime(end_date.year, end_date.month, 1).date()
            write_log(f"Monday is in a different month, using first day of current month: {start_date.strftime('%Y-%m-%d')}", "CYAN")
    
    # If it's neither Friday nor last day of month, return None
    else:
        write_log("Test date is neither Friday nor last day of month, cannot determine date range", "RED")
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

def find_required_files():
    """
    Find required files for the automation
    
    Returns:
        dict: Dictionary containing file paths or None if files not found
    """
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
            return None
    
    if not data_file_path:
        write_log(f"No files found matching pattern '{FILE_PATTERNS['er']}'. Please specify the file path manually.", "RED")
        # Allow manual input as fallback
        data_file_path = input("Please enter the full path to the ER data file (e.g., C:\\path\\to\\data.xlsx): ").strip('"')
        if not os.path.exists(data_file_path):
            write_log("Invalid file path. Exiting.", "RED")
            return None
    
    write_log(f"Found GSN file: {excel_file_path}", "GREEN")
    write_log(f"Found ER file: {data_file_path}", "GREEN")
    
    # Check if the SharePoint file exists
    if not sharepoint_exists:
        write_log(f"SharePoint file not found at: {SYNCED_FILE_PATH}", "RED")
        write_log("Please ensure OneDrive sync is set up correctly.", "YELLOW")
        return None
    
    write_log(f"Found SharePoint file at: {SYNCED_FILE_PATH}", "GREEN")
    
    return {
        'gsn_file': excel_file_path,
        'er_file': data_file_path,
        'sharepoint_file': SYNCED_FILE_PATH
    }

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

def process_data(file_paths):
    """
    Process data from the input files
    
    Args:
        file_paths (dict): Dictionary containing file paths
        
    Returns:
        dict: Dictionary containing processed data or None if processing failed
    """
    try:
        # Start data processing in parallel
        write_log("Starting parallel data processing...", "YELLOW")
        
        with concurrent.futures.ThreadPoolExecutor() as executor:
            # Start Jobs
            gsn_job = executor.submit(process_gsn_data, file_paths['gsn_file'])
            er_job = executor.submit(process_er_data, file_paths['er_file'])
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
        
        return {
            'gsn_entries': extracted_values,
            'er_entries': filtered_er_hostnames,
            'ad_entries': ad_computers,
            'missing_in_er': missing_in_er,
            'missing_in_gsn': missing_in_gsn,
            'filtered_hostnames2': filtered_hostnames2,
            'er_serial_number': er_serial_number
        }
        
    except Exception as e:
        write_log(f"Error processing data: {str(e)}", "RED")
        import traceback
        write_log(traceback.format_exc(), "RED")
        return None

def update_excel_file(date_range, data):
    """
    Update the Excel file with the comparison results
    
    Args:
        date_range (DateRangeResult): Date range for the report
        data (dict): Dictionary containing processed data
        
    Returns:
        bool: Success status
    """
    try:
        # Update the Excel file with the GSN vs ER analysis
        write_log("Updating Excel file with comparison results...", "YELLOW")
        excel_updater = ExcelUpdater(SYNCED_FILE_PATH)
        result = excel_updater.analyze_excel_file(
            data['gsn_entries'],
            data['er_entries'],
            data['ad_entries'],
            date_range,
            data['missing_in_er'],
            data['missing_in_gsn'],
            data['filtered_hostnames2'],
            data['er_serial_number']
        )
        
        if result:
            write_log("Excel file updated successfully", "GREEN")
        else:
            write_log("Failed to update Excel file", "RED")
            
        return result
        
    except Exception as e:
        write_log(f"Error updating Excel file: {str(e)}", "RED")
        import traceback
        write_log(traceback.format_exc(), "RED")
        return False