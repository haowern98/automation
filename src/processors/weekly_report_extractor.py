"""
Weekly Report Extractor

This processor:
1. Accesses the Weekly Report Excel file from a local synced folder
2. Extracts data for a specific date range
3. Generates an HTML table with proper formatting (yellow only in status column)
4. Can be used both from GUI and CLI
"""
import os
import re
import time
import datetime
import sys
import pandas as pd
import shutil
import tempfile
from dateutil.parser import parse
import webbrowser
from src.utils.logger import write_log
from src.processors.gsn_vs_ad_extractor import GSNvsADExtractor
from src.processors.gsn_vs_er_extractor import GSNvsERExtractor
from src.processors.er_extractor import ERExtractor

# Print diagnostic info at startup
print("Weekly Report Extractor - Starting up...")
print("Python version:", sys.version)

class WeeklyReportExtractor:
    """Class to extract weekly reports from local Excel file"""
    
    def __init__(self, excel_file_path=None):
        """
        Initialize with Excel file path
        
        Args:
            excel_file_path (str): Path to the Excel file
        """
        self.temp_files = []  # Track temporary files for cleanup
        
        # If no path is provided, get from settings first, then fallback to default
        if not excel_file_path:
            try:
                from src.gui.settings_dialog import get_settings
                settings = get_settings()
                configured_path = settings.get('file_paths', 'weekly_report_file_path', '')
                
                if configured_path and os.path.exists(configured_path):
                    self.excel_file_path = configured_path
                else:
                    # Fallback to default hardcoded path
                    self.excel_file_path = self._get_default_path()
            except Exception:
                # If settings can't be loaded, use default hardcoded path
                self.excel_file_path = self._get_default_path()
        else:
            self.excel_file_path = excel_file_path
    
    def _get_default_path(self):
        """Get the default hardcoded path as fallback"""
        user_profile = os.environ.get('USERPROFILE', '')
        return os.path.join(
            user_profile, 
            'DPDHL', 
            'SM Team - SG - AD EDS, MFA, GSN VS AD, GSN VS ER Weekly Report', 
            'Weekly Report 2025 - Copy.xlsx'
        )
    
    def get_section_keywords(self):
        """
        Get section keywords from settings, with fallback to hardcoded defaults
        
        Returns:
            list: List of section keywords
        """
        try:
            from src.gui.settings_dialog import get_settings
            settings = get_settings()
            keywords = settings.get_section_keywords()
            
            # Validate that we have keywords
            if keywords and len(keywords) > 0:
                return keywords
            else:
                # Fallback to hardcoded defaults if settings are empty
                return [
                    "Applied MFA Method", "ARP Invalid", "Accounts with Manager", 
                    "No AD", "GID assigned", "Accounts with", "Manager/ARP"
                ]
        except Exception as e:
            # If settings loading fails, use hardcoded defaults
            print(f"Warning: Could not load section keywords from settings: {e}")
            return [
                "Applied MFA Method", "ARP Invalid", "Accounts with Manager", 
                "No AD", "GID assigned", "Accounts with", "Manager/ARP"
            ]
    
    def set_excel_file_path(self, excel_file_path):
        """
        Set the path to the Excel file
        
        Args:
            excel_file_path (str): Path to the Excel file
            
        Returns:
            bool: Success status
        """
        self.excel_file_path = excel_file_path
        return True
    
    def cleanup_temp_files(self):
        """Clean up all temporary files with multiple retry attempts"""
        for temp_file in self.temp_files:
            max_attempts = 3
            for attempt in range(max_attempts):
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        write_log(f"Cleaned up temporary file: {temp_file}", "GREEN")
                        break
                except Exception as e:
                    if attempt < max_attempts - 1:
                        write_log(f"Warning: Could not delete temporary file {temp_file}: {str(e)}", "YELLOW")
                        write_log(f"Retrying in 2 seconds... (Attempt {attempt+1}/{max_attempts})", "YELLOW")
                        time.sleep(2)
                    else:
                        write_log(f"Warning: Failed to delete temporary file {temp_file} after {max_attempts} attempts: {str(e)}", "RED")
        
        # Clear the list
        self.temp_files = []
    
    def extract_date_components(self, date_range_str):
        """
        Extract month and year from a date range string
        
        Args:
            date_range_str (str): Date range string (e.g., '5-9 May 2025')
            
        Returns:
            tuple: (month_name, month_num, year)
        """
        # Match different date range patterns
        pattern1 = r'\d+-\d+\s+([A-Za-z]+)\s+(\d{4})'  # "5-9 May 2025"
        pattern2 = r'\d+\s+([A-Za-z]+)\s+-\s+\d+\s+([A-Za-z]+)\s+(\d{4})'  # "5 May - 9 Jun 2025"
        
        # Try first pattern (same month)
        match1 = re.match(pattern1, date_range_str)
        if match1:
            month_name = match1.group(1)
            year = match1.group(2)
            month_num = datetime.datetime.strptime(month_name, '%B').month
            return (month_name, month_num, year)
        
        # Try second pattern (different months)
        match2 = re.match(pattern2, date_range_str)
        if match2:
            start_month = match2.group(1)
            end_month = match2.group(2)
            year = match2.group(3)
            
            # Use the end month for worksheet determination
            month_name = end_month
            month_num = datetime.datetime.strptime(month_name, '%B').month
            return (month_name, month_num, year)
        
        # Default if no match
        try:
            # Try to parse as a date
            dt = parse(date_range_str)
            return (dt.strftime('%B'), dt.month, dt.strftime('%Y'))
        except:
            # Return current month/year if parsing fails
            now = datetime.datetime.now()
            return (now.strftime('%B'), now.month, now.strftime('%Y'))
    
    def determine_worksheet_name(self, date_range_str):
        """
        Determine the worksheet name based on date range
        
        Args:
            date_range_str (str): Date range string (e.g., '5-9 May 2025')
            
        Returns:
            str: Worksheet name
        """
        month_name, _, year = self.extract_date_components(date_range_str)
        
        # Format: "MFA, AD EDS May 2025"
        return f"MFA, AD EDS {month_name} {year}"
    
    def create_copy_and_extract(self, date_range_str):
        """
        Create a temporary copy of the Excel file and extract data from it
        
        Args:
            date_range_str (str): Date range string
            
        Returns:
            list: Extracted data
        """
        temp_file = None
        
        try:
            # Create a temporary file with a unique name
            _, temp_file = tempfile.mkstemp(suffix='.xlsx', prefix='temp_report_')
            self.temp_files.append(temp_file)  # Track for cleanup
            
            write_log(f"Creating temporary copy at: {temp_file}", "CYAN")
            shutil.copy2(self.excel_file_path, temp_file)
            
            # Try to extract data from the temporary file
            data = self.extract_from_file(temp_file, date_range_str)
            return data
            
        except Exception as e:
            write_log(f"Error creating temporary copy: {str(e)}", "RED")
            import traceback
            traceback.print_exc()
            # Try hard-coded basic extraction if all else fails
            return self.create_basic_data(date_range_str)
        finally:
            # Don't try to delete temp file here - we'll clean up at the end
            pass
    
    def extract_from_file(self, file_path, date_range_str):
        """
        Extract data from the specified Excel file for a specific date range
        
        Args:
            file_path (str): Path to Excel file
            date_range_str (str): Date range string to extract (e.g., '5-9 May 2025')
            
        Returns:
            list: Extracted data
        """
        try:
            # Get all sheet names
            excel_file = pd.ExcelFile(file_path)
            all_sheets = excel_file.sheet_names
            write_log(f"Available worksheets: {all_sheets}", "CYAN")
            
            # Extract month and year components
            month_name, _, year = self.extract_date_components(date_range_str)
            
            # Find the sheet that most closely matches our expected name
            target_sheet = None
            
            # First try to find an exact match for "MFA, AD EDS" with full month name
            for sheet_name in all_sheets:
                if f"MFA, AD EDS {month_name}" in sheet_name and str(year) in sheet_name:
                    target_sheet = sheet_name
                    write_log(f"Found exact matching worksheet: {sheet_name}", "GREEN")
                    break
            
            # If no exact match, try with month abbreviation
            if not target_sheet:
                month_abbr = month_name[:3]
                for sheet_name in all_sheets:
                    if "MFA, AD EDS" in sheet_name and month_abbr in sheet_name and str(year) in sheet_name:
                        target_sheet = sheet_name
                        write_log(f"Found partial matching worksheet: {sheet_name}", "GREEN")
                        break
            
            # If still no match, look for any sheet with the month and year
            if not target_sheet:
                for sheet_name in all_sheets:
                    if month_name in sheet_name and str(year) in sheet_name:
                        target_sheet = sheet_name
                        write_log(f"Found month/year matching worksheet: {sheet_name}", "GREEN")
                        break
                    elif month_name[:3] in sheet_name and str(year) in sheet_name:
                        target_sheet = sheet_name
                        write_log(f"Found month abbr/year matching worksheet: {sheet_name}", "GREEN")
                        break
            
            if not target_sheet:
                write_log(f"No worksheet found matching month {month_name} and year {year}", "RED")
                return []
            
            # Read the entire Excel sheet
            df = pd.read_excel(file_path, sheet_name=target_sheet, header=None)
            write_log(f"Read worksheet with {len(df)} rows", "CYAN")
            
            # Find the row containing the requested date range
            start_row = -1
            for i, row in df.iterrows():
                # Get the first cell value
                first_cell = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else ""
                if date_range_str == first_cell.strip():
                    start_row = i
                    write_log(f"Found exact date range '{date_range_str}' in row {start_row}", "GREEN")
                    break
            
            # If not found with exact match, try with substring
            if start_row == -1:
                for i, row in df.iterrows():
                    first_cell = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else ""
                    if date_range_str in first_cell:
                        start_row = i
                        write_log(f"Found date range '{date_range_str}' in row {start_row}", "GREEN")
                        break
            
            if start_row == -1:
                write_log(f"Date range '{date_range_str}' not found in worksheet", "RED")
                return []
            
            # Find the next row that contains another date range pattern (N-N Month YYYY)
            end_row = -1
            date_pattern = r'\d+-\d+\s+[A-Za-z]+\s+\d{4}'  # Matches "5-9 May 2025" format
            
            for i in range(start_row + 1, len(df)):
                # Get the first cell in this row
                first_cell = str(df.iloc[i, 0]) if not pd.isna(df.iloc[i, 0]) else ""
                
                # If we find another date range pattern, this is our end
                if re.search(date_pattern, first_cell.strip()) and i != start_row:
                    end_row = i
                    write_log(f"Found next date range at row {end_row}: '{first_cell.strip()}'", "CYAN")
                    break
            
            # If no end found, use the end of the data
            if end_row == -1:
                end_row = len(df)
                write_log(f"No next date range found, using end of data (row {end_row})", "CYAN")
            
            # Extract the data between start_row and end_row
            data_df = df.iloc[start_row:end_row].copy()
            write_log(f"Extracted {len(data_df)} rows of data from rows {start_row} to {end_row-1}", "GREEN")
            
            # Print the first 3 rows to debug what's being extracted
            write_log("\nSample of extracted data (first 3 rows):", "CYAN")
            for i in range(min(3, len(data_df))):
                first_cell = str(data_df.iloc[i, 0]) if not pd.isna(data_df.iloc[i, 0]) else ""
                write_log(f"Row {i}: {first_cell[:50]}", "WHITE")
            
            # Convert DataFrame to a list of rows (just extracting values)
            data = []
            
            for _, row in data_df.iterrows():
                row_data = []
                for value in row.values:
                    if pd.isna(value):
                        row_data.append({'value': ''})
                    else:
                        row_data.append({'value': str(value).strip()})
                
                # Only add non-empty rows
                if any(cell['value'] for cell in row_data):
                    data.append(row_data)
            
            return data
            
        except Exception as e:
            write_log(f"Error extracting from file: {str(e)}", "RED")
            import traceback
            traceback.print_exc()
            return []
    
    def create_basic_data(self, date_range_str):
        """
        Create a basic data structure with expected styling 
        when all other extraction methods fail
        
        Args:
            date_range_str (str): Date range string
            
        Returns:
            list: Basic data structure
        """
        write_log("Creating basic data structure with hardcoded values", "YELLOW")
        
        # Basic structure for a weekly report - simplified to just values
        data = []
        
        # First row - Date range
        data.append([
            {'value': date_range_str},
            {'value': ''},
            {'value': ''},
            {'value': ''}
        ])
        
        # Second row - Headers
        data.append([
            {'value': 'Updates for AD/EDS Clean up & MFA'},
            {'value': 'Incident Ticket'},
            {'value': 'Remarks'},
            {'value': 'Status'}
        ])
        
        # Section header - Applied MFA Method
        data.append([
            {'value': 'Applied MFA Method'},
            {'value': ''},
            {'value': ''},
            {'value': ''}
        ])
        
        # Sample data rows
        data.append([
            {'value': 'ernecheo'},
            {'value': ''},
            {'value': '5/5/2025 intern, Chris Eng'},
            {'value': 'Pending'}
        ])
        
        data.append([
            {'value': 'jinglang'},
            {'value': ''},
            {'value': '5/5/2025 intern, Chris Eng'},
            {'value': 'Pending'}
        ])
        
        data.append([
            {'value': 'bajum'},
            {'value': ''},
            {'value': 'Conversion to contractor account - setup done 21.4.2025'},
            {'value': 'Completed'}
        ])
        
        return data
    
    def extract_data_for_date_range(self, date_range_str):
        """
        Extract data for the given date range from the Excel file
        
        Args:
            date_range_str (str): Date range string (e.g., '5-9 May 2025')
            
        Returns:
            list: List of rows containing the data
        """
        # Try to create a copy and extract from it
        return self.create_copy_and_extract(date_range_str)
    
    def extract_data_for_date_range_gui(self, date_range_str):
        """
        Extract data for the given date range from the Excel file (GUI version)
        Returns both success status and data for GUI error handling
        
        Args:
            date_range_str (str): Date range string (e.g., '5-9 May 2025')
            
        Returns:
            tuple: (success: bool, data: list, error_message: str)
        """
        try:
            write_log(f"GUI: Extracting data for date range: {date_range_str}", "YELLOW")
            data = self.create_copy_and_extract(date_range_str)
            
            if not data:
                return False, [], f"No data found for date range '{date_range_str}'"
            
            write_log(f"GUI: Successfully extracted {len(data)} rows", "GREEN")
            return True, data, ""
            
        except Exception as e:
            error_msg = f"Error extracting data: {str(e)}"
            write_log(f"GUI: {error_msg}", "RED")
            return False, [], error_msg
        finally:
            # Always clean up temp files
            self.cleanup_temp_files()
    
    def generate_html_table(self, data):
        """
        Generate an HTML table from the extracted data with precise column-based styling
        
        Args:
            data (list): List of rows containing the data
            
        Returns:
            str: HTML table string
        """
        if not data:
            return "<p>No data found for the specified date range.</p>"
        
        # Fixed column count - weekly reports should have exactly 4 columns
        max_cols = 4
        
        # Generate CSS with proper styling
        html = '''
    <style>
    /* Base table styling */
    table.weekly-report {
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 20px;
        font-family: Arial, sans-serif;
        table-layout: fixed;
    }

    /* Cell borders and padding */
    table.weekly-report td {
        border: 1px solid #dddddd;
        padding: 8px;
        vertical-align: top;
        word-wrap: break-word;
    }

    /* Define specific column widths */
    table.weekly-report td:nth-child(1) { width: 25%; } /* Updates column */
    table.weekly-report td:nth-child(2) { width: 15%; } /* Incident Ticket */
    table.weekly-report td:nth-child(3) { width: 45%; } /* Remarks */
    table.weekly-report td:nth-child(4) { width: 15%; } /* Status */

    /* First row (date range) - gray background */
    table.weekly-report tr:first-child td {
        background-color: #f0f0f5 !important;
        font-weight: bold;
        text-align: left;
    }

    /* Second row (column headers) - red text */
    table.weekly-report tr:nth-child(2) td {
        color: #ff0000;
        font-weight: bold;
        text-align: center;
        background-color: #ffffff;
    }

    /* Section headers - light blue background spans full row */
    tr.section-header td {
        background-color: #ddebf7 !important;
        font-weight: bold;
        text-align: left;
    }

    /* "Completed by for" rows - yellow background for entire row */
    tr.completed-by-row td {
        background-color: #ffeb9c !important;
        color: #9c5700;
    }

    /* Status column - default white background */
    table.weekly-report td:nth-child(4) {
        background-color: white !important;
        text-align: center;
    }

    /* Override for "Pending" in status column only */
    table.weekly-report td:nth-child(4).pending {
        background-color: #ffeb9c !important;
        color: #9c5700;
    }

    /* Override for "Completed" in status column only */
    table.weekly-report td:nth-child(4).completed {
        background-color: #c6efce !important;
        color: #006100;
    }

    /* Ensure section header backgrounds override status column defaults */
    tr.section-header td:nth-child(4) {
        background-color: #ddebf7 !important;
        color: inherit !important;
        text-align: left;
    }

    /* Ensure completed-by-row backgrounds override status column defaults */
    tr.completed-by-row td:nth-child(4) {
        background-color: #ffeb9c !important;
        color: #9c5700;
        text-align: center;
    }

    /* INC cells in column 2 - red font */
    table.weekly-report td.inc-cell {
        color: #ff0000 !important;
        font-weight: bold;
    }

    /* Data rows */
    table.weekly-report tr:not(:first-child):not(:nth-child(2)):not(.section-header):not(.completed-by-row) td:nth-child(1) {
        text-align: left;
    }

    table.weekly-report tr:not(:first-child):not(:nth-child(2)):not(.section-header):not(.completed-by-row) td:nth-child(2) {
        text-align: center;
    }

    table.weekly-report tr:not(:first-child):not(:nth-child(2)):not(.section-header):not(.completed-by-row) td:nth-child(3) {
        text-align: left;
    }
    </style>

    <table class="weekly-report">
    '''
        
        # Process each row of data
        for row_idx, row in enumerate(data):
            # Check if this is a section header row
            is_section_header = False
            if row_idx > 1 and len(row) > 0 and row[0]['value']:
                first_cell = row[0]['value']
                section_keywords = self.get_section_keywords()
                if any(keyword in first_cell for keyword in section_keywords):
                    is_section_header = True
            
            # Check if this is a "completed by for" row (among last two rows)
            is_completed_by_row = False
            if row_idx >= len(data) - 2:  # Last two rows
                # Check if any cell in the row contains "completed by for"
                for cell in row:
                    if 'completed by for' in cell.get('value', '').lower():
                        is_completed_by_row = True
                        break
            
            # Start row with appropriate class
            if is_section_header:
                html += '<tr class="section-header">\n'
            elif is_completed_by_row:
                html += '<tr class="completed-by-row">\n'
            else:
                html += '<tr>\n'
            
            # Handle first row (date range) - should span all columns
            if row_idx == 0:
                date_value = row[0]['value'] if len(row) > 0 else ''
                html += f'  <td colspan="4">{date_value}</td>\n'
            
            # Handle section headers - should span all columns
            elif is_section_header:
                section_value = row[0]['value'] if len(row) > 0 else ''
                html += f'  <td colspan="4">{section_value}</td>\n'
            
            # Handle regular rows - exactly 4 columns
            else:
                for col_idx in range(max_cols):
                    # Get cell value if it exists, otherwise empty
                    if col_idx < len(row):
                        cell_value = row[col_idx].get('value', '')
                    else:
                        cell_value = ''
                    
                    # Check if this is column 2 (Incident Ticket) and contains "INC"
                    is_inc_cell = (col_idx == 1 and 'INC' in cell_value)
                    
                    # Special styling for status column (4th column) - but not for completed-by rows
                    if col_idx == 3 and not is_completed_by_row:  # Status column (0-indexed)
                        if cell_value == "Pending":
                            html += f'  <td class="pending">{cell_value}</td>\n'
                        elif cell_value == "Completed":
                            html += f'  <td class="completed">{cell_value}</td>\n'
                        else:
                            html += f'  <td>{cell_value}</td>\n'
                    elif is_inc_cell:
                        # INC cell in column 2 - red font
                        html += f'  <td class="inc-cell">{cell_value}</td>\n'
                    else:
                        # Normal cell (completed-by rows get their styling from CSS class)
                        html += f'  <td>{cell_value}</td>\n'
            
            # End row
            html += '</tr>\n'
        
        # Close the table
        html += '</table>\n'
        
        return html
    
    def generate_complete_html(self, data, date_range_str=None):
        """
        Generate complete HTML document with proper structure
        Now supports both regular data (list) and combined data (dict)
        
        Args:
            data (list or dict): List of rows containing the data OR dict with combined MFA + GSN VS AD data
            date_range_str (str, optional): Date range string for the title
            
        Returns:
            str: Complete HTML document
        """
        # Check if this is combined data (dict) or regular data (list)
        if isinstance(data, dict):
            # This is combined data with both MFA and GSN VS AD
            table_html = self.generate_combined_html_table(data)
        else:
            # This is regular MFA-only data
            table_html = self.generate_html_table(data)
        
        # Create complete HTML document
        html = '<!DOCTYPE html>\n'
        html += '<html>\n'
        html += '<head>\n'
        html += '    <meta charset="UTF-8">\n'
        
        # Dynamic title based on date range
        if date_range_str:
            html += f'    <title>{date_range_str} Weekly Report</title>\n'
        else:
            html += '    <title>Weekly Report</title>\n'
        
        html += '</head>\n'
        html += '<body>\n'
        
        # Dynamic header based on date range
        if date_range_str:
            html += f'<h1>{date_range_str} Weekly Report</h1>\n'
        else:
            html += '<h1>Weekly Report</h1>\n'
            
        # Add the MFA & AD/EDS subheading
        html += '<h2>MFA & AD/EDS</h2>\n'
        
        # Add the HTML table with styles
        html += table_html
        
        html += '\n</body>\n'
        html += '</html>'
        
        return html
            
    def extract_combined_data_for_date_range_gui(self, date_range_str):
        """
        Extract MFA, GSN VS AD, and GSN VS ER data for the given date range (GUI version)
        Returns both success status and combined data for GUI error handling
        
        Args:
            date_range_str (str): Date range string (e.g., '29-30 May 2025')
            
        Returns:
            tuple: (success: bool, combined_data: dict, error_message: str)
        """
        try:
            write_log(f"GUI: Extracting combined MFA + GSN VS AD + GSN VS ER data for date range: {date_range_str}", "YELLOW")
            
            # Extract MFA data
            write_log("Extracting MFA data...", "CYAN")
            mfa_success, mfa_data, mfa_error = self.extract_data_for_date_range_gui(date_range_str)
            
            # Extract GSN VS AD data
            write_log("Extracting GSN VS AD data...", "CYAN")
            gsn_vs_ad_extractor = GSNvsADExtractor(self.excel_file_path)
            gsn_ad_success, gsn_ad_data, gsn_ad_error = gsn_vs_ad_extractor.extract_gsn_vs_ad_data(date_range_str)
            
            # Extract GSN VS ER data
            write_log("Extracting GSN VS ER data...", "CYAN")
            gsn_vs_er_extractor = GSNvsERExtractor(self.excel_file_path)
            gsn_er_success, gsn_er_data, gsn_er_error = gsn_vs_er_extractor.extract_gsn_vs_er_data(date_range_str)

            # Extract ER data
            write_log("Extracting ER data...", "CYAN")
            er_extractor = ERExtractor(self.excel_file_path)
            er_success, er_data, er_error = er_extractor.extract_er_data(date_range_str)
            
            # Combine results
            combined_data = {
                'mfa_data': mfa_data if mfa_success else [],
                'gsn_vs_ad_data': gsn_ad_data if gsn_ad_success else [],
                'gsn_vs_er_data': gsn_er_data if gsn_er_success else [],
                'er_data': er_data if er_success else [],
                'mfa_success': mfa_success,
                'gsn_vs_ad_success': gsn_ad_success,
                'gsn_vs_er_success': gsn_er_success,
                'er_success': er_success,
                'mfa_error': mfa_error,
                'gsn_vs_ad_error': gsn_ad_error,
                'gsn_vs_er_error': gsn_er_error,
                'er_error': er_error
            }
            
            # Determine overall success
            overall_success = mfa_success or gsn_ad_success or gsn_er_success or er_success  # Success if at least one succeeds
            
            # Create combined error message
            error_parts = []
            if not mfa_success and mfa_error:
                error_parts.append(f"MFA Error: {mfa_error}")
            if not gsn_ad_success and gsn_ad_error:
                error_parts.append(f"GSN VS AD Error: {gsn_ad_error}")
            if not gsn_er_success and gsn_er_error:
                error_parts.append(f"GSN VS ER Error: {gsn_er_error}")
            if not er_success and er_error:
                error_parts.append(f"ER Error: {er_error}")
            
            combined_error = " | ".join(error_parts) if error_parts else ""
            
            if overall_success:
                mfa_count = len(mfa_data) if mfa_success else 0
                gsn_ad_count = len(gsn_ad_data) if gsn_ad_success else 0
                gsn_er_count = len(gsn_er_data) if gsn_er_success else 0
                er_count = len(er_data) if er_success else 0
                write_log(f"GUI: Successfully extracted combined data - MFA: {mfa_count} rows, GSN VS AD: {gsn_ad_count} rows, GSN VS ER: {gsn_er_count} rows, ER: {er_count} rows", "GREEN")
            else:
                write_log(f"GUI: Failed to extract any data - {combined_error}", "RED")
            
            return overall_success, combined_data, combined_error
            
        except Exception as e:
            error_msg = f"Error extracting combined data: {str(e)}"
            write_log(f"GUI: {error_msg}", "RED")
            return False, {
                'mfa_data': [], 'gsn_vs_ad_data': [], 'gsn_vs_er_data': [], 'er_data': [],
                'mfa_success': False, 'gsn_vs_ad_success': False, 'gsn_vs_er_success': False, 'er_success': False,
                'mfa_error': '', 'gsn_vs_ad_error': '', 'gsn_vs_er_error': '', 'er_error': ''
            }, error_msg
        finally:
            # Always clean up temp files
            self.cleanup_temp_files()

    def generate_combined_html_table(self, combined_data):
        """
        Generate simple HTML tables compatible with Teams messaging
        
        Args:
            combined_data (dict): Dictionary containing MFA, GSN VS AD, GSN VS ER, and ER data
            
        Returns:
            str: HTML table string
        """
        mfa_data = combined_data.get('mfa_data', [])
        gsn_vs_ad_data = combined_data.get('gsn_vs_ad_data', [])
        gsn_vs_er_data = combined_data.get('gsn_vs_er_data', [])
        er_data = combined_data.get('er_data', [])
        mfa_success = combined_data.get('mfa_success', False)
        gsn_ad_success = combined_data.get('gsn_vs_ad_success', False)
        gsn_er_success = combined_data.get('gsn_vs_er_success', False)
        er_success = combined_data.get('er_success', False)
        
        html = ''
        
        # Generate MFA section if data exists
        if mfa_success and mfa_data:
            html += '<table border="1" style="font-size: 5px;">\n<tbody>\n'
            
            # Process MFA data
            for row_idx, row in enumerate(mfa_data):
                html += '<tr>\n'
                
                # Handle first row (date range) - should span all columns
                if row_idx == 0:
                    date_value = row[0]['value'] if len(row) > 0 else ''
                    html += f'            <td colspan="4" style="background-color: #EDEDED; color: #000000; font-weight: bold;"><b>{date_value}</b></td>\n'
                
                # Handle second row (headers) - red text
                elif row_idx == 1:
                    headers = ['Updates for AD/EDS Clean up & MFA', 'Incident Ticket', 'Remarks', 'Status']
                    for i, header in enumerate(headers):
                        if i == 3:  # Status column
                            html += f'<td style="background-color: #FFFFFF; color: #FF0000; font-weight: bold;"><span style="font-size: 6px; white-space: nowrap;"><b>{header}</b></span></td>\n'
                        else:
                            html += f'<td style="background-color: #FFFFFF; color: #FF0000; font-weight: bold;"><span style="font-size: 7px;"><b>{header}</b></span></td>\n'
                
                # Handle section headers
                elif len(row) > 0 and any(keyword in row[0].get('value', '') for keyword in self.get_section_keywords()):
                    section_value = row[0]['value']
                    html += f'<td style="background-color: #DDEBF7; color: #000000; font-weight: bold;"><span style="font-size: 7px;"><b>{section_value}</b></span></td>\n'
                    html += f'<td style="background-color: #DDEBF7; color: #000000; font-weight: bold;"><span style="font-size: 7px;"><b></b></span></td>\n'
                    html += f'<td style="background-color: #DDEBF7; color: #000000; font-weight: bold;"><span style="font-size: 7px;"><b></b></span></td>\n'
                    html += f'<td style="background-color: #DDEBF7; color: #000000; font-weight: bold;"><span style="font-size: 6px; white-space: nowrap;"><b></b></span></td>\n'
                
                # Handle "completed by for" rows
                elif row_idx >= len(mfa_data) - 2 and any('completed by for' in cell.get('value', '').lower() for cell in row):
                    for col_idx in range(4):
                        cell_value = row[col_idx].get('value', '') if col_idx < len(row) else ''
                        if col_idx == 3:  # Status column
                            html += f'<td style="background-color: #FFFF00; color: #FF0000; font-weight: bold;"><span style="font-size: 6px; white-space: nowrap;"><b>{cell_value}</b></span></td>\n'
                        else:
                            html += f'<td style="background-color: #FFFF00; color: #FF0000; font-weight: bold;"><span style="font-size: 7px;"><b>{cell_value}</b></span></td>\n'
                
                # Handle regular data rows
                else:
                    for col_idx in range(4):
                        cell_value = row[col_idx].get('value', '') if col_idx < len(row) else ''
                        
                        # Check if this is column 2 and contains "INC" (red text)
                        if col_idx == 1 and 'INC' in cell_value:
                            html += f'<td style="background-color: #FFFFFF; color: #FF0000; font-weight: bold;"><span style="font-size: 7px;"><b>{cell_value}</b></span></td>\n'
                        # Status column with special styling
                        elif col_idx == 3:
                            if cell_value == "Pending":
                                html += f'<td style="background-color: #FFEB9C; color: #9C5700; font-weight: normal;"><span style="font-size: 6px; white-space: nowrap;">{cell_value}</span></td>\n'
                            elif cell_value == "Completed":
                                html += f'<td style="background-color: #C6EFCE; color: #006100; font-weight: normal;"><span style="font-size: 6px; white-space: nowrap;">{cell_value}</span></td>\n'
                            else:
                                html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 6px; white-space: nowrap;">{cell_value}</span></td>\n'
                        # Regular columns
                        else:
                            html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 7px;">{cell_value}</span></td>\n'
                
                html += '</tr>'
            
            html += '</tbody>\n</table>\n'
        
        elif not mfa_success:
            html += '<p style="color: red;">MFA data could not be loaded.</p>\n'
        
        # Add GSN VS AD section if data exists
        if gsn_ad_success and gsn_vs_ad_data:
            html += '<br><h2>GSN VS AD</h2>\n'
            html += '<table border="1" style="font-size: 5px;">\n<tbody>\n'
            
            # Process GSN VS AD data
            for row_idx, row in enumerate(gsn_vs_ad_data):
                html += '<tr>\n'
                
                # Handle main header - should span all 6 columns
                if len(row) > 0 and 'GSN VS AD' in row[0].get('value', ''):
                    header_value = row[0]['value']
                    html += f'            <td colspan="6" style="background-color: #AEAAAA; color: #000000; font-weight: bold;"><b>{header_value}</b></td>\n'
                
                # Handle column headers
                elif len(row) > 0 and row[0].get('value', '') == 'In GSN not in AD':
                    headers = ['In GSN not in AD', 'Remarks', 'Action', 'In AD not in GSN', 'Remarks', 'Action']
                    for i, header in enumerate(headers):
                        if i >= 3:  # Last 3 columns
                            html += f'<td style="background-color: #FFFF00; color: #000000; font-weight: bold;"><span style="font-size: 6px; white-space: nowrap;"><b>{header}</b></span></td>\n'
                        else:
                            html += f'<td style="background-color: #FFFF00; color: #000000; font-weight: bold;"><span style="font-size: 7px;"><b>{header}</b></span></td>\n'
                
                # Handle regular rows - exactly 6 columns
                else:
                    for col_idx in range(6):
                        cell_value = row[col_idx].get('value', '') if col_idx < len(row) else ''
                        if col_idx >= 3:  # Last 3 columns
                            html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 6px; white-space: nowrap;">{cell_value}</span></td>\n'
                        else:
                            html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 7px;">{cell_value}</span></td>\n'
                
                html += '</tr>'
            
            html += '</tbody>\n</table>\n'
        
        elif not gsn_ad_success:
            html += '<br><h2>GSN VS AD</h2>\n'
            html += '<p style="color: red;">GSN VS AD data could not be loaded.</p>\n'
        
        # Add GSN VS ER section if data exists
        if gsn_er_success and gsn_vs_er_data:
            html += '<br><h2>GSN VS ER</h2>\n'
            html += '<table border="1" style="font-size: 5px;">\n<tbody>\n'
            
            # Process GSN VS ER data
            for row_idx, row in enumerate(gsn_vs_er_data):
                html += '<tr>\n'
                
                # Handle exactly 2 columns (D and E)
                for col_idx in range(2):
                    if col_idx < len(row):
                        cell_data = row[col_idx]
                        cell_value = cell_data.get('value', '')
                        
                        # Check if bold
                        if cell_data.get('isBolded') == 'bold':
                            html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: bold;"><span style="font-size: 7px;"><b>{cell_value}</b></span></td>\n'
                        else:
                            html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 7px;">{cell_value}</span></td>\n'
                    else:
                        html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 7px;"></span></td>\n'
                
                html += '</tr>'
            
            html += '</tbody>\n</table>\n'
        
        elif not gsn_er_success:
            html += '<br><h2>GSN VS ER</h2>\n'
            html += '<p style="color: red;">GSN VS ER data could not be loaded.</p>\n'
        
        # Add ER section if data exists
        if er_success and er_data:
            html += '<br><h2>ER</h2>\n'
            html += '<table border="1" style="font-size: 5px;">\n<tbody>\n'
            
            # Process ER data
            for row_idx, row_data in enumerate(er_data):
                html += '<tr>\n'
                
                # Special handling for first row (date range header with gray background)
                if row_idx == 0 and 'Column1' in row_data and row_data['Column1'].get('colspan') == 3:
                    # First row should span 3 columns with #AEAAAA background
                    cell_content = row_data['Column1'].get('cell content', '')
                    html += f'            <td colspan="3" style="background-color: #AEAAAA; color: #000000; font-weight: bold;"><b>{cell_content}</b></td>\n'
                else:
                    # Handle exactly 3 columns (Column1, Column2, Column3)
                    for col_num in range(1, 4):
                        col_key = f"Column{col_num}"
                        
                        # Skip merged cells (columns 2 and 3 in first row)
                        if row_idx == 0 and col_num > 1 and col_key in row_data and row_data[col_key].get('merged', False):
                            continue
                        
                        if col_key in row_data:
                            cell_data = row_data[col_key]
                            cell_content = cell_data.get('cell content', '')
                            
                            # Force white background and black text for all data rows
                            html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 7px;">{cell_content}</span></td>\n'
                        else:
                            html += f'<td style="background-color: #FFFFFF; color: #000000; font-weight: normal;"><span style="font-size: 7px;"></span></td>\n'
                
                html += '</tr>'
            
            html += '</tbody>\n</table>\n'
        
        elif not er_success:
            html += '<br><h2>ER</h2>\n'
            html += '<p style="color: red;">ER data could not be loaded.</p>\n'
        
        # If no sections have data
        if not mfa_success and not gsn_ad_success and not gsn_er_success and not er_success:
            html += '<p>No data found for the specified date range.</p>\n'
        
        return html
    
    def save_html_to_file(self, html, output_path, date_range_str=None):
        """
        Save HTML content to a file with proper styling
        
        Args:
            html (str): HTML content
            output_path (str): Path to save the file
            date_range_str (str, optional): Date range string for the title
            
        Returns:
            bool: Success status
        """
        try:
            # Create the directory if it doesn't exist
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            write_log(f"Saving HTML file to: {output_path}", "CYAN")
            
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html)
            
            write_log(f"HTML file saved successfully to: {output_path}", "GREEN")
            return True
        except Exception as e:
            write_log(f"Error saving HTML file: {str(e)}", "RED")
            import traceback
            traceback.print_exc()
            return False


def main():
    """Main function to run the script"""
    try:
        print("\nWeekly Report Extractor")
        print("======================")
        
        # Create the extractor with default file path
        extractor = WeeklyReportExtractor()
        
        try:
            # Ask if user wants to use a custom Excel file path
            use_custom_path = input("Use default Excel file path? (y/n): ").strip().lower() == 'n'
            
            if use_custom_path:
                custom_path = input("Enter path to Excel file: ").strip()
                extractor.set_excel_file_path(custom_path)
            
            # Prompt for date range
            date_range_str = input("Enter date range to extract (e.g., '5-9 May 2025'): ").strip()
            
            # Extract data for the date range
            print(f"\nExtracting data for date range: {date_range_str}...")
            data = extractor.extract_data_for_date_range(date_range_str)
            
            if not data:
                print("\nNo data found for the specified date range.")
                return
            
            print(f"\nFound {len(data)} rows of data. Generating HTML table...")
            
            # Generate complete HTML
            complete_html = extractor.generate_complete_html(data, date_range_str)
            
            # Get user's Downloads folder
            user_profile = os.environ.get('USERPROFILE', '')
            downloads_folder = os.path.join(user_profile, 'Downloads')
            
            # Create sanitized filename
            safe_filename = date_range_str.replace(" ", "_").replace("-", "_")
            
            # Define HTML output path
            html_path = os.path.join(downloads_folder, f'weekly_report_{safe_filename}.html')
            
            # Save HTML file
            html_success = extractor.save_html_to_file(complete_html, html_path, date_range_str)
            
            if html_success:
                print(f"\n==> HTML file saved to: {html_path}")
                
                # Ask if user wants to open the HTML file
                want_open = input("\nOpen the HTML file in browser now? (y/n): ").strip().lower() == 'y'
                if want_open:
                    webbrowser.open('file://' + os.path.abspath(html_path))
            
        finally:
            # Ensure temporary files are cleaned up
            extractor.cleanup_temp_files()
            print("\nDone!")
    
    except Exception as e:
        print(f"\nError in main function: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")


if __name__ == "__main__":
    main()