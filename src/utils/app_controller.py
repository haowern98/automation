"""
Enhanced Main application controller for SharePoint Automation with proper flow termination
"""
import os
import sys
import time
import re
import concurrent.futures
from datetime import datetime

from src.utils.logger import write_log
from src.utils.excel_functions import ExcelApplication
from src.utils.comparison import compare_data_sets, ExcelUpdater
from src.gui.date_selector import DateRangeResult
from src.gui.settings_dialog import get_settings
from src.processors.gsn_processor import process_gsn_data
from src.processors.er_processor import process_er_data
from src.processors.ad_processor import process_ad_data, compare_gsn_with_ad
from src.config import USER_PROFILE, SYNCED_FILE_PATH, FILE_PATTERNS, AD_SEARCH, DATA_DIR

# Global flag to track if user wants to terminate the entire process
_USER_TERMINATED = False

def run_sharepoint_automation(manual_mode=False, debug_mode=False):
    """
    Run the SharePoint automation process
    
    Args:
        manual_mode (bool): Whether to run in manual mode
        debug_mode (bool): Whether to run in debug mode
        
    Returns:
        bool: Success status
    """
    global _USER_TERMINATED
    _USER_TERMINATED = False  # Reset termination flag
    
    try:
        # Skip date checks if manual mode is enabled
        if not manual_mode:
            if not check_run_date():
                return False
        else:
            write_log("Running in manual mode: skipping date checks", "YELLOW")
        
        # Check for Excel processes and warm up Excel
        manage_excel()
        
        # Check if user terminated during Excel management
        if _USER_TERMINATED:
            write_log("Process terminated by user during Excel management", "YELLOW")
            return False
        
        # Get date range (through GUI or automatic calculation)
        date_range = get_date_range(manual_mode)
        if not date_range or not date_range.is_valid:
            if _USER_TERMINATED:
                write_log("Process terminated by user", "YELLOW")
                return False
            write_log("No valid date range provided. Exiting.", "RED")
            return False
            
        write_log(f"Using date range: {date_range.date_range_formatted}", "GREEN")
        
        # Check if user terminated after date selection
        if _USER_TERMINATED:
            write_log("Process terminated by user after date selection", "YELLOW")
            return False
        
        # Find required files
        file_paths = find_required_files()
        if not file_paths:
            return False
            
        # Check if user terminated during file search
        if _USER_TERMINATED:
            write_log("Process terminated by user during file search", "YELLOW")
            return False
        
        # Process data
        data_results = process_data(file_paths)
        if not data_results:
            return False
            
        # Check if user terminated during data processing
        if _USER_TERMINATED:
            write_log("Process terminated by user during data processing", "YELLOW")
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
        if _USER_TERMINATED:
            write_log("Process terminated by user", "YELLOW")
            return False
        write_log(f"Error in SharePoint automation: {str(e)}", "RED")
        import traceback
        write_log(traceback.format_exc(), "RED")
        return False

def terminate_process():
    """Set the global termination flag"""
    global _USER_TERMINATED
    _USER_TERMINATED = True
    write_log("User requested process termination", "YELLOW")

def check_run_date():
    """
    Check if today is a day to run the automation
    
    Returns:
        bool: True if automation should run, False otherwise
    """
    # In production, use: current_date = datetime.now().date()
    
    # TEST DATE - replace with datetime.now().date() in production
    current_date = datetime(2025, 6, 30).date()
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
    global _USER_TERMINATED
    
    # Import the enhanced UI
    from src.gui.tabbed_app import show_tabbed_date_range_selection
    
    # Show the date selection dialog
    date_range = None
    if manual_mode:
        # In manual mode, no timeout
        write_log("Showing date range selection dialog in manual mode...", "YELLOW")
        date_range = show_tabbed_date_range_selection(manual_mode=True)
        
        # If user terminated in manual mode, exit the program
        if date_range and date_range.user_terminated:
            write_log("User chose to exit in manual mode", "YELLOW")
            _USER_TERMINATED = True
            return None
    else:
        # In automatic mode, show dialog with 30-second timeout
        write_log("Showing date range selection dialog with 30-second timeout...", "YELLOW")
        write_log("Options: OK (use selected dates), Use Auto Date (skip to calculated dates), Cancel Process (terminate)", "CYAN")
        
        date_range = show_tabbed_date_range_selection(manual_mode=False, timeout_seconds=30)
    
    # Handle the different outcomes
    if date_range:
        # Check if user explicitly terminated the entire process
        if date_range.user_terminated:
            write_log("User explicitly terminated the process", "YELLOW")
            _USER_TERMINATED = True
            return None
        
        # Check if user chose to use auto date (either button or timeout)
        if date_range.use_auto_date or (date_range.cancelled and not date_range.user_terminated):
            if date_range.use_auto_date:
                write_log("User chose to use auto date calculation", "YELLOW")
            else:
                write_log("Dialog timed out. Using automatic date range calculation", "YELLOW")
            
            # Use automatic date calculation
            date_range = get_automatic_date_range()
            if not date_range:
                write_log("Failed to calculate automatic date range", "RED")
                return None
        
        # If we reach here with a valid date range, user chose OK with their selected dates
        elif not date_range.cancelled:
            write_log(f"User selected custom date range: {date_range.date_range_formatted}", "GREEN")
    
    return date_range

def get_automatic_date_range():
    """
    Automatically determine date range based on test day
    
    Returns:
        DateRangeResult: Date range object with start and end dates
    """
    from datetime import timedelta
    
    # TEST DATE - replace with datetime.now().date() in production
    current_date = datetime(2025, 6, 30).date()
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
    Find required files for the automation using settings configuration
    
    Returns:
        dict: Dictionary containing file paths or None if files not found
    """
    write_log("Loading settings configuration...", "CYAN")
    
    # Load settings
    settings = get_settings()
    
    # Get configured paths and patterns
    gsn_search_dir = settings.get('file_paths', 'gsn_search_directory', '')
    er_search_dir = settings.get('file_paths', 'er_search_directory', '')
    gsn_pattern = settings.get('file_paths', 'gsn_file_pattern', 'alm_hardware')
    er_pattern = settings.get('file_paths', 'er_file_pattern', 'data')
    
    write_log(f"GSN Search Directory: {gsn_search_dir}", "CYAN")
    write_log(f"GSN File Pattern: {gsn_pattern}", "CYAN")
    write_log(f"ER Search Directory: {er_search_dir}", "CYAN")
    write_log(f"ER File Pattern: {er_pattern}", "CYAN")
    
    # Prepare search directories
    gsn_search_dirs = []
    er_search_dirs = []
    
    # Add configured directories if they exist
    if gsn_search_dir and os.path.exists(gsn_search_dir):
        gsn_search_dirs.append(gsn_search_dir)
        write_log(f"Added GSN search directory: {gsn_search_dir}", "GREEN")
    else:
        write_log(f"GSN search directory not found or not configured: {gsn_search_dir}", "YELLOW")
    
    if er_search_dir and os.path.exists(er_search_dir):
        er_search_dirs.append(er_search_dir)
        write_log(f"Added ER search directory: {er_search_dir}", "GREEN")
    else:
        write_log(f"ER search directory not found or not configured: {er_search_dir}", "YELLOW")
    
    # Add fallback common locations if configured directories are not available
    fallback_locations = [
        os.path.join(USER_PROFILE, "Downloads"),
        os.path.join(USER_PROFILE, "Desktop"),
        os.path.join(USER_PROFILE, "Documents"),
        os.path.join(USER_PROFILE, "OneDrive"),
        os.path.join(USER_PROFILE, "OneDrive - Deutsche Post DHL")
    ]
    
    # Add fallback locations to search if primary directories are empty
    if not gsn_search_dirs:
        gsn_search_dirs.extend([loc for loc in fallback_locations if os.path.exists(loc)])
        write_log("Using fallback locations for GSN search", "YELLOW")
    
    if not er_search_dirs:
        er_search_dirs.extend([loc for loc in fallback_locations if os.path.exists(loc)])
        write_log("Using fallback locations for ER search", "YELLOW")
    
    # Use concurrent futures to search for files in parallel
    with concurrent.futures.ThreadPoolExecutor() as executor:
        # Start concurrent file search jobs
        write_log("Starting concurrent file search operations...", "YELLOW")
        
        gsn_file_future = executor.submit(find_latest_file_with_pattern, gsn_search_dirs, gsn_pattern)
        er_file_future = executor.submit(find_latest_file_with_pattern, er_search_dirs, er_pattern)
        sharepoint_exists_future = executor.submit(os.path.exists, SYNCED_FILE_PATH)
        
        # Wait for all file search jobs to complete
        write_log("Waiting for file search operations to complete...", "YELLOW")
        
        excel_file_path = gsn_file_future.result()
        data_file_path = er_file_future.result()
        sharepoint_exists = sharepoint_exists_future.result()
    
    # Validate required files exist
    if not excel_file_path:
        write_log(f"No GSN files found matching pattern '{gsn_pattern}' in configured directories.", "RED")
        # Allow manual input as fallback
        excel_file_path = input(f"Please enter the full path to the GSN Excel file (pattern: {gsn_pattern}*.xlsx): ").strip('"')
        if not os.path.exists(excel_file_path):
            write_log("Invalid file path. Exiting.", "RED")
            return None
    
    if not data_file_path:
        write_log(f"No ER files found matching pattern '{er_pattern}' in configured directories.", "RED")
        # Allow manual input as fallback
        data_file_path = input(f"Please enter the full path to the ER data file (pattern: {er_pattern}*.xlsx): ").strip('"')
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

def find_latest_file_with_pattern(search_directories, file_pattern):
    """
    Find the latest file matching a pattern in specified directories
    
    Args:
        search_directories (list): List of directories to search in
        file_pattern (str): File pattern to match (without *)
        
    Returns:
        str or None: Path to the latest file or None if not found
    """
    try:
        latest_file = None
        latest_time = 0
        all_matching_files = []
        
        write_log(f"Searching for pattern '{file_pattern}' in {len(search_directories)} directories", "CYAN")
        
        # Function to check if a filename matches our pattern
        def matches_pattern(filename, pattern):
            """Check if filename matches the pattern - more flexible matching"""
            # Create regex pattern: starts with pattern + any characters + .xlsx
            # Examples: "data.xlsx", "data(2).xlsx", "data 23-8-2025.xlsx", "data_latest.xlsx"
            regex_pattern = f"^{re.escape(pattern)}.*\\.xlsx$"
            return re.match(regex_pattern, filename, re.IGNORECASE) is not None
        
        # Function to extract version number from filename
        def get_version_number(filename):
            """Extract version number from filename - handles various formats"""
            # Look for version numbers in parentheses: (2), (11), etc.
            paren_match = re.search(r'\((\d+)\)', filename)
            if paren_match:
                return int(paren_match.group(1))
            
            # Look for version numbers after underscore: _v2, _version2, _2
            underscore_match = re.search(r'_(?:v|version)?(\d+)', filename, re.IGNORECASE)
            if underscore_match:
                return int(underscore_match.group(1))
            
            # Look for dates in filename and use as version (latest date = highest version)
            # Patterns: 23-8-2025, 2025-08-23, 08-23-2025, etc.
            date_patterns = [
                r'(\d{1,2})-(\d{1,2})-(\d{4})',  # 23-8-2025
                r'(\d{4})-(\d{1,2})-(\d{1,2})',  # 2025-8-23
                r'(\d{1,2})\.(\d{1,2})\.(\d{4})', # 23.8.2025
                r'(\d{4})\.(\d{1,2})\.(\d{1,2})'  # 2025.8.23
            ]
            
            for pattern in date_patterns:
                date_match = re.search(pattern, filename)
                if date_match:
                    groups = date_match.groups()
                    # Convert date to a comparable number
                    # For simplicity, just use the sum of all digits
                    return sum(int(group) for group in groups)
            
            return 0  # Files without identifiable version are treated as version 0
        
        for search_dir in search_directories:
            if not os.path.exists(search_dir):
                write_log(f"Search directory does not exist: {search_dir}", "YELLOW")
                continue
                
            write_log(f"Searching in: {search_dir}", "CYAN")
            
            try:
                # Search in the directory (limit depth to avoid excessive searching)
                max_depth = 2
                base_depth = search_dir.count(os.sep)
                
                for dirpath, dirs, filenames in os.walk(search_dir):
                    # Check if we've gone too deep
                    current_depth = dirpath.count(os.sep) - base_depth
                    if current_depth > max_depth:
                        del dirs[:]  # Don't descend any deeper
                        continue
                    
                    for filename in filenames:
                        # Check if filename matches our pattern
                        if filename.lower().endswith('.xlsx') and matches_pattern(filename, file_pattern):
                            file_path = os.path.join(dirpath, filename)
                            
                            if os.path.exists(file_path):
                                all_matching_files.append(file_path)
                                file_time = os.path.getmtime(file_path)
                                write_log(f"Found file: {filename} (Modified: {time.ctime(file_time)})", "GREEN")
                                
            except Exception as e:
                write_log(f"Error searching directory {search_dir}: {str(e)}", "RED")
        
        write_log(f"Found {len(all_matching_files)} matching files in total", "CYAN")
        
        # If we found matching files, determine the best one to use
        if all_matching_files:
            # Always use the most recently modified file (latest by modification date)
            for file_path in all_matching_files:
                file_time = os.path.getmtime(file_path)
                if file_time > latest_time:
                    latest_time = file_time
                    latest_file = file_path
            
            write_log(f"Selected most recently modified file: {os.path.basename(latest_file)} (Modified: {time.ctime(latest_time)})", "GREEN")
        
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