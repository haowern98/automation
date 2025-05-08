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

def find_latest_file(root_dir, file_pattern):
    """
    Find the latest file matching a pattern
    
    Args:
        root_dir (str): Root directory to search in
        file_pattern (str): File pattern to match
        
    Returns:
        str or None: Path to the latest file or None if not found
    """
    try:
        latest_file = None
        latest_time = 0
        
        for dirpath, _, filenames in os.walk(root_dir):
            for filename in filenames:
                if filename.startswith(file_pattern.replace('*', '')):
                    file_path = os.path.join(dirpath, filename)
                    file_time = os.path.getmtime(file_path)
                    
                    if file_time > latest_time:
                        latest_time = file_time
                        latest_file = file_path
        
        return latest_file
    except Exception as e:
        write_log(f"Error finding latest file: {str(e)}", "RED")
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
        root_dir2 = os.path.join(USER_PROFILE)
        
        # Use concurrent futures to search for files in parallel
        with concurrent.futures.ThreadPoolExecutor() as executor:
            # Start concurrent file search jobs
            write_log("Starting concurrent file search operations...", "YELLOW")
            
            alm_file_future = executor.submit(find_latest_file, root_dir, FILE_PATTERNS['gsn'])
            data_file_future = executor.submit(find_latest_file, root_dir2, FILE_PATTERNS['er'])
            sharepoint_exists_future = executor.submit(os.path.exists, SYNCED_FILE_PATH)
            
            # Wait for all file search jobs to complete
            write_log("Waiting for file search operations to complete...", "YELLOW")
            
            excel_file_path = alm_file_future.result()
            data_file_path = data_file_future.result()
            sharepoint_exists = sharepoint_exists_future.result()
        
        # Validate required files exist
        if not excel_file_path:
            write_log(f"No files found matching pattern '{FILE_PATTERNS['gsn']}' in {root_dir}", "RED")
            return
        
        if not data_file_path:
            write_log(f"No files found matching pattern '{FILE_PATTERNS['er']}' in {root_dir2}", "RED")
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