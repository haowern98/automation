"""
Active Directory data processing for SharePoint Automation
Uses PowerShell script for AD operations via a batch file
"""
import os
import json
import subprocess
import time
from src.utils.logger import write_log
from src.config import AD_SEARCH, AD_RESULTS_FILE, AD_COMPARISON_FILE

def process_ad_data(ldap_filter=None, search_base=None):
    """
    Process AD data by calling PowerShell script through a batch file
    
    Args:
        ldap_filter (str, optional): LDAP filter for the query
        search_base (str, optional): Base DN for the search
        
    Returns:
        list: List of computer names
    """
    # Use defaults from config if not provided
    if ldap_filter is None:
        ldap_filter = AD_SEARCH['ldap_filter']
    if search_base is None:
        search_base = AD_SEARCH['search_base']
    
    output_file = AD_RESULTS_FILE
    
    write_log(f"Starting AD data processing using PowerShell with LDAP filter", "YELLOW")
    start_time = time.time()
    
    try:
        # FIXED: Calculate path to project root, then to batch file
        # From src/processors/ad_processor.py -> project root -> run_ad_processor.bat
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
        batch_file = os.path.join(project_root, "run_ad_processor.bat")
        
        write_log(f"Looking for batch file at: {batch_file}", "CYAN")
        
        # Check if batch file exists
        if not os.path.exists(batch_file):
            write_log(f"Batch file not found at: {batch_file}", "RED")
            write_log("Please ensure run_ad_processor.bat is in the project root directory", "RED")
            return []
        
        write_log(f"Executing batch file: {batch_file}", "CYAN")
        result = subprocess.run([batch_file], capture_output=True, text=True, cwd=project_root)
        
        # Log output from batch/PowerShell
        write_log("PowerShell script output:", "CYAN")
        for line in result.stdout.splitlines():
            if line.strip():  # Skip empty lines
                write_log(f"PowerShell: {line}", "CYAN")
        
        # Check if there are any errors in stderr
        if result.stderr:
            write_log("PowerShell script errors:", "RED")
            for line in result.stderr.splitlines():
                if line.strip():
                    write_log(f"PowerShell Error: {line}", "RED")
        
        # Check if the JSON file was created
        if os.path.exists(output_file):
            # Read JSON file
            try:
                with open(output_file, 'r', encoding='utf-8-sig') as f:
                    file_content = f.read()
                    
                    # Check if file has content
                    if not file_content.strip():
                        write_log(f"AD results file is empty: {output_file}", "RED")
                        return []
                    
                    # Parse JSON content
                    write_log(f"Parsing AD results from: {output_file}", "CYAN")
                    ad_computers = json.loads(file_content)
                    
                    # Ensure we have a list
                    if not isinstance(ad_computers, list):
                        write_log(f"AD results is not a list: {type(ad_computers)}", "RED")
                        # Try to convert to list if possible
                        if isinstance(ad_computers, str):
                            ad_computers = [ad_computers]
                        else:
                            try:
                                ad_computers = list(ad_computers)
                            except:
                                ad_computers = []
                    
                    # Ensure all items are strings
                    ad_computers = [str(item) for item in ad_computers if item]
                    
                    end_time = time.time()
                    duration = end_time - start_time
                    write_log(f"AD data processing complete - {len(ad_computers)} entries found in {duration:.2f} seconds", "GREEN")
                    
                    # Display the hostnames if found
                    if ad_computers:
                        write_log("\n==============================================", "CYAN")
                        write_log(f"AD HOSTNAMES: {len(ad_computers)}", "YELLOW")
                        write_log("==============================================", "CYAN")
                        
                        # Sort and display hostnames in columns
                        sorted_hostnames = sorted(ad_computers)
                        
                        # Display a sample of hostnames (5 to 10)
                        sample_size = min(10, len(sorted_hostnames))
                        write_log(f"Sample of {sample_size} AD hostnames:", "CYAN")
                        for i in range(sample_size):
                            write_log(f"  {i+1}. {sorted_hostnames[i]}", "WHITE")
                        
                        if len(sorted_hostnames) > sample_size:
                            write_log(f"  ... and {len(sorted_hostnames) - sample_size} more", "WHITE")
                        
                        write_log("==============================================", "CYAN")
                    else:
                        write_log("No AD hostnames found!", "YELLOW")
                    
                    return ad_computers
            except Exception as e:
                write_log(f"Error parsing AD results file: {str(e)}", "RED")
                if os.path.exists(output_file):
                    with open(output_file, 'r') as f:
                        content = f.read()
                        write_log(f"File content: {content[:500]}", "RED")
                return []
        else:
            write_log(f"AD results file not found: {output_file}", "RED")
            write_log("Checking current directory:", "YELLOW")
            current_dir = os.getcwd()
            write_log(f"Current directory: {current_dir}", "YELLOW")
            files = os.listdir(current_dir)
            write_log(f"Files in directory: {files}", "YELLOW")
            return []
            
    except Exception as e:
        write_log(f"Error processing AD data: {str(e)}", "RED")
        import traceback
        write_log(traceback.format_exc(), "RED")
        write_log("Please check if PowerShell and the AD module are available.", "RED")
        return []

# Rest of the file remains the same...
def compare_gsn_with_ad(gsn_entries, ad_entries):
    """
    Compare GSN and AD data sets
    
    Args:
        gsn_entries (list): List of GSN entries
        ad_entries (list): List of AD entries
        
    Returns:
        dict: Dictionary containing comparison results
    """
    write_log("\n=========================================", "YELLOW")
    write_log("COMPARING GSN AND AD ENTRIES", "YELLOW")
    write_log("=========================================", "YELLOW")
    
    # Ensure inputs are lists of strings
    gsn_entries = [str(item) for item in gsn_entries if item]
    ad_entries = [str(item) for item in ad_entries if item]
    
    # Log input details for debugging
    write_log(f"GSN Entries: {len(gsn_entries)}", "CYAN")
    write_log(f"AD Entries: {len(ad_entries)}", "CYAN")
    
    # If AD entries is empty, try to load from file
    if not ad_entries:
        ad_results_file = AD_RESULTS_FILE
        if os.path.exists(ad_results_file):
            write_log(f"AD entries is empty, trying to load from file: {ad_results_file}", "YELLOW")
            try:
                with open(ad_results_file, 'r', encoding='utf-8-sig') as f:
                    ad_entries = json.load(f)
                write_log(f"Loaded {len(ad_entries)} AD entries from file", "GREEN")
            except Exception as e:
                write_log(f"Error loading AD entries from file: {str(e)}", "RED")
    
    # Find entries in GSN but not in AD
    missing_in_ad = [item for item in gsn_entries if item not in ad_entries]
    
    # Find entries in AD but not in GSN
    missing_in_gsn = [item for item in ad_entries if item not in gsn_entries]
    
    # Report GSN entries not in AD
    if missing_in_ad:
        write_log("\nIn GSN but not in AD:", "MAGENTA")
        display_count = min(len(missing_in_ad), 10)
        for item in sorted(missing_in_ad)[:display_count]:
            write_log(f"  {item}", "MAGENTA")
        if len(missing_in_ad) > display_count:
            write_log(f"  ... and {len(missing_in_ad) - display_count} more", "MAGENTA")
    else:
        write_log("\nNo entries in GSN that are not in AD.", "GREEN")
    
    # Report AD entries not in GSN
    if missing_in_gsn:
        write_log("\nIn AD but not in GSN:", "CYAN")
        display_count = min(len(missing_in_gsn), 10)
        for item in sorted(missing_in_gsn)[:display_count]:
            write_log(f"  {item}", "CYAN")
        if len(missing_in_gsn) > display_count:
            write_log(f"  ... and {len(missing_in_gsn) - display_count} more", "CYAN")
    else:
        write_log("\nNo entries in AD that are not in GSN.", "GREEN")
    
    # Create summary of comparison results
    write_log("\nComparison Summary:", "YELLOW")
    write_log(f"- Total GSN entries: {len(gsn_entries)}", "WHITE")
    write_log(f"- Total AD entries: {len(ad_entries)}", "WHITE")
    write_log(f"- GSN entries not in AD: {len(missing_in_ad)}", "MAGENTA")
    write_log(f"- AD entries not in GSN: {len(missing_in_gsn)}", "CYAN")
    write_log("=========================================", "YELLOW")
    
    # Define the output file path
    output_file_path = AD_COMPARISON_FILE
    
    # Save results to JSON file
    result = {
        "MissingInAD": missing_in_ad,
        "MissingInGSN": missing_in_gsn
    }
    
    # Ensure the data directory exists
    os.makedirs(os.path.dirname(output_file_path), exist_ok=True)
    
    # Write the comparison results to file
    try:
        with open(output_file_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=4)
        write_log(f"Comparison results saved to: {output_file_path}", "CYAN")
    except Exception as e:
        write_log(f"Error saving comparison results: {str(e)}", "RED")
    
    return result

# Optional test function for direct testing
def test_ad_processor():
    """Test the AD processor directly"""
    write_log("Testing AD processor...", "YELLOW")
    
    # Process AD data
    ad_computers = process_ad_data()
    
    # Create mock GSN data for testing (using the same AD data)
    gsn_entries = ad_computers.copy() if ad_computers else []
    
    # Compare the data
    compare_gsn_with_ad(gsn_entries, ad_computers)
    
    return ad_computers

# Run the test function if this module is executed directly
if __name__ == "__main__":
    test_ad_processor()