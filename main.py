#!/usr/bin/env python3
"""
SharePoint Automation - Main Script
Converted from PowerShell to Python
"""
import os
import sys
import time
import concurrent.futures
from datetime import datetime

# Import modules from our package
from gui import show_date_range_selection
from processors import process_gsn_data, process_er_data, process_ad_data
from utils import write_log, compare_data_sets, ExcelUpdater
from config import USER_PROFILE, SYNCED_FILE_PATH, FILE_PATTERNS, AD_SEARCH

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
    # Record the start time
    start_time = datetime.now()
    write_log("Starting SharePoint file access script (OneDrive sync method)", "YELLOW")

    try:
        # First, get the date range from the user
        write_log("Prompting for date range...", "YELLOW")
        date_range = show_date_range_selection()
        
        if not date_range:
            write_log("Date range selection was cancelled. Exiting.", "RED")
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