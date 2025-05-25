"""
ER data processing for SharePoint Automation
"""
import time
from src.utils.logger import write_log
from src.utils.excel_functions import ExcelApplication

def process_er_data(data_file_path):
    """
    Process ER data from Excel file
    
    Args:
        data_file_path (str): Path to the Excel file containing ER data
        
    Returns:
        dict: Dictionary containing filtered hostnames and serial numbers
    """
    write_log(f"Starting ER data processing from: {data_file_path}", "YELLOW")
    start_time = time.time()
    
    # Initialize Excel application
    excel_app = ExcelApplication()
    
    # Initialize lists to hold the filtered data
    filtered_er_hostnames = []
    filtered_hostnames2 = []
    er_serial_number = []
    
    try:
        # Open workbook
        if not excel_app.open_workbook(data_file_path):
            return {"FilteredERHostnames": [], "FilteredHostnames2": [], "ErSerialNumber": []}
        
        # Get the first worksheet
        worksheet = excel_app.get_worksheet()
        if not worksheet:
            return {"FilteredERHostnames": [], "FilteredHostnames2": [], "ErSerialNumber": []}
        
        # Define the column indices
        hostname_column_index = 11  # Column K
        status_column_index = 37    # Column AK
        er_sn_column_index = 15     # Column O
        
        row_index = 4  # Starting from row 4
        
        # Process the data
        while True:
            cell_value = worksheet.Cells(row_index, hostname_column_index).Text
            status_value = worksheet.Cells(row_index, status_column_index).Text
            sn_value = worksheet.Cells(row_index, er_sn_column_index).Text
            
            if not cell_value:
                break
            
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
            
            row_index += 1
        
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
        write_log(f"Error processing ER data: {str(e)}", "RED")
        return {"FilteredERHostnames": [], "FilteredHostnames2": [], "ErSerialNumber": []}
        
    finally:
        # Clean up
        excel_app.close()