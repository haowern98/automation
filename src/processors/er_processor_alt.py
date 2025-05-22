"""
Alternative ER data processing for SharePoint Automation using openpyxl
"""
import time
import openpyxl
import pandas as pd
from src.utils.logger import write_log

def process_er_data_alt(data_file_path):
    """
    Process ER data from Excel file using openpyxl (no COM automation)
    
    Args:
        data_file_path (str): Path to the Excel file containing ER data
        
    Returns:
        dict: Dictionary containing filtered hostnames and serial numbers
    """
    write_log(f"Starting ER data processing (alternative method) from: {data_file_path}", "YELLOW")
    start_time = time.time()
    
    # Initialize lists to hold the filtered data
    filtered_er_hostnames = []
    filtered_hostnames2 = []
    er_serial_number = []
    
    try:
        # First try with pandas (faster)
        write_log("Attempting to read Excel file with pandas...", "CYAN")
        try:
            # Read the Excel file
            df = pd.read_excel(data_file_path)
            write_log(f"Successfully read Excel file with pandas. Found {len(df)} rows.", "GREEN")
            
            # Log column names for debugging
            write_log(f"Columns in the file: {list(df.columns)}", "CYAN")
            
            # Find relevant columns
            hostname_col = None
            status_col = None
            sn_col = None
            
            # Try to find the hostname column (could be named differently)
            for col in df.columns:
                col_lower = str(col).lower()
                if 'host' in col_lower or 'computer' in col_lower or 'name' in col_lower:
                    hostname_col = col
                    write_log(f"Found hostname column: '{hostname_col}'", "GREEN")
                    break
            
            # Try to find the status column
            for col in df.columns:
                col_lower = str(col).lower()
                if 'status' in col_lower or 'state' in col_lower or 'days' in col_lower:
                    status_col = col
                    write_log(f"Found status column: '{status_col}'", "GREEN")
                    break
            
            # Try to find the serial number column
            for col in df.columns:
                col_lower = str(col).lower()
                if 'serial' in col_lower or 'sn' in col_lower or 'number' in col_lower:
                    sn_col = col
                    write_log(f"Found SN column: '{sn_col}'", "GREEN")
                    break
            
            # If we couldn't find the columns, use default indices
            if not hostname_col:
                write_log("Hostname column not found, attempting to use index 10 (K)", "YELLOW")
                if len(df.columns) > 10:
                    hostname_col = df.columns[10]  # Column K (index 10)
                    write_log(f"Using column '{hostname_col}' for hostnames", "YELLOW")
            
            if not status_col:
                write_log("Status column not found, attempting to use index 36 (AK)", "YELLOW")
                if len(df.columns) > 36:
                    status_col = df.columns[36]  # Column AK (index 36)
                    write_log(f"Using column '{status_col}' for status", "YELLOW")
            
            if not sn_col:
                write_log("SN column not found, attempting to use index 14 (O)", "YELLOW")
                if len(df.columns) > 14:
                    sn_col = df.columns[14]  # Column O (index 14)
                    write_log(f"Using column '{sn_col}' for serial numbers", "YELLOW")
            
            # Process the data if we found the necessary columns
            if hostname_col:
                write_log("Processing data with pandas...", "CYAN")
                
                # Convert columns to string to ensure startswith works
                df[hostname_col] = df[hostname_col].astype(str)
                
                # Filter hostnames starting with specific prefixes
                filtered_df = df[
                    df[hostname_col].str.startswith('SGASC') | 
                    df[hostname_col].str.startswith('SGESC') | 
                    df[hostname_col].str.startswith('SGSC') | 
                    df[hostname_col].str.startswith('SGWSC') | 
                    df[hostname_col].str.startswith('SGXSC')
                ]
                
                write_log(f"Found {len(filtered_df)} entries matching the hostname criteria", "GREEN")
                
                # Extract hostnames
                filtered_er_hostnames = filtered_df[hostname_col].tolist()
                
                # Apply additional filtering if we have the status column
                if status_col and sn_col:
                    # Filter for "Between 31 and 60 days"
                    filtered_df2 = filtered_df[filtered_df[status_col] == "Between 31 and 60 days"]
                    
                    # Extract hostnames and serial numbers
                    filtered_hostnames2 = filtered_df2[hostname_col].tolist()
                    er_serial_number = filtered_df2[sn_col].tolist()
                    
                    write_log(f"Found {len(filtered_hostnames2)} entries with status 'Between 31 and 60 days'", "GREEN")
                
                # Calculate duration
                end_time = time.time()
                duration = end_time - start_time
                
                write_log(f"ER data processing complete - {len(filtered_er_hostnames)} entries found in {duration:.2f} seconds", "GREEN")
                write_log(f"- Filtered entries (31-60 days): {len(filtered_hostnames2)}", "CYAN")
                
                return {
                    "FilteredERHostnames": filtered_er_hostnames,
                    "FilteredHostnames2": filtered_hostnames2,
                    "ErSerialNumber": er_serial_number
                }
            else:
                write_log("Could not find all required columns using pandas", "YELLOW")
                raise Exception("Required columns not found")
                
        except Exception as pd_error:
            write_log(f"Pandas method failed: {str(pd_error)}", "RED")
            write_log("Falling back to openpyxl method...", "YELLOW")
        
        # Fallback to openpyxl if pandas fails
        write_log("Loading workbook with openpyxl...", "CYAN")
        workbook = openpyxl.load_workbook(data_file_path, read_only=True, data_only=True)
        
        # Get the active worksheet or first sheet
        sheet = workbook.active
        
        write_log(f"Successfully loaded workbook with openpyxl. Active sheet: {sheet.title}", "GREEN")
        
        # Define the column indices (convert to 0-based for openpyxl)
        hostname_column_index = 11  # Column K
        status_column_index = 37    # Column AK
        er_sn_column_index = 15     # Column O
        
        # Log some info for debugging
        first_row = list(sheet.rows)[3]  # Get the 4th row (index 3)
        write_log(f"Row 4 sample data: {[cell.value for cell in first_row]}", "CYAN")
        
        row_index = 4  # Starting from row 4 (0-based index would be 3)
        max_empty_rows = 10
        empty_row_count = 0
        
        # Process the data
        for row_idx, row in enumerate(sheet.iter_rows(min_row=4), start=4):
            try:
                # Get cell values with proper indexing adjustment
                cell_value = str(row[hostname_column_index-1].value or "")
                status_value = str(row[status_column_index-1].value or "")
                sn_value = str(row[er_sn_column_index-1].value or "")
                
                if not cell_value:
                    empty_row_count += 1
                    if empty_row_count >= max_empty_rows:
                        write_log(f"Detected {max_empty_rows} consecutive empty rows, stopping processing", "YELLOW")
                        break
                    continue
                
                empty_row_count = 0  # Reset counter when we find a value
                
                # Log progress every 100 rows
                if row_idx % 100 == 0:
                    write_log(f"Processed {row_idx} rows...", "CYAN")
                
                # Filter to include hostnames starting with specific prefixes
                if (cell_value.startswith("SGASC") or 
                    cell_value.startswith("SGESC") or 
                    cell_value.startswith("SGSC") or 
                    cell_value.startswith("SGWSC") or 
                    cell_value.startswith("SGXSC")):
                    
                    filtered_er_hostnames.append(cell_value)
                    
                    # Second round of filtering: Check if the status is "Between 31 and 60 days"
                    if status_value == "Between 31 and 60 days":
                        filtered_hostnames2.append(cell_value)
                        er_serial_number.append(sn_value)
                
            except Exception as row_error:
                write_log(f"Error processing row {row_idx}: {str(row_error)}", "RED")
        
        # Close the workbook
        workbook.close()
        
        # Calculate duration
        end_time = time.time()
        duration = end_time - start_time
        
        write_log(f"ER data processing complete - {len(filtered_er_hostnames)} entries found in {duration:.2f} seconds", "GREEN")
        write_log(f"- Filtered entries (31-60 days): {len(filtered_hostnames2)}", "CYAN")
        
        return {
            "FilteredERHostnames": filtered_er_hostnames,
            "FilteredHostnames2": filtered_hostnames2,
            "ErSerialNumber": er_serial_number
        }
        
    except Exception as e:
        import traceback
        write_log(f"Error processing ER data with alternative method: {str(e)}", "RED")
        write_log(traceback.format_exc(), "RED")
        return {"FilteredERHostnames": [], "FilteredHostnames2": [], "ErSerialNumber": []}