"""
GSN data processing for SharePoint Automation
"""
import time
from utils.logger import write_log
from utils.excel_functions import ExcelApplication

def process_gsn_data(excel_file_path):
    """
    Process GSN data from Excel file
    
    Args:
        excel_file_path (str): Path to the Excel file containing GSN data
        
    Returns:
        list: List of extracted values
    """
    write_log(f"Starting GSN data processing from: {excel_file_path}", "YELLOW")
    start_time = time.time()
    
    # Initialize Excel application
    excel_app = ExcelApplication()
    extracted_values = []
    
    try:
        # Open workbook
        if not excel_app.open_workbook(excel_file_path):
            return []
        
        # Get the first worksheet
        worksheet = excel_app.get_worksheet()
        if not worksheet:
            return []
        
        # Define the column to extract (Column A)
        column_to_extract = 1
        row_index = 2  # Starting from row 2
        
        # Extract values
        while True:
            cell_value = worksheet.Cells(row_index, column_to_extract).Text
            if not cell_value:
                break
                
            extracted_values.append(cell_value)
            row_index += 1
        
        # Calculate duration
        end_time = time.time()
        duration = end_time - start_time
        
        write_log(f"GSN data processing complete - {len(extracted_values)} entries found in {duration:.2f} seconds", "GREEN")
        return extracted_values
        
    except Exception as e:
        write_log(f"Error processing GSN data: {str(e)}", "RED")
        return []
        
    finally:
        # Clean up
        excel_app.close()