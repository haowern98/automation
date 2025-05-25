"""
Data comparison functionality for SharePoint Automation
"""
import datetime
import time
from src.utils.logger import write_log
from src.utils.excel_functions import ExcelApplication

def compare_data_sets(gsn_entries, er_entries):
    """
    Compare GSN and ER data sets and return the differences
    
    Args:
        gsn_entries (list): List of GSN entries
        er_entries (list): List of ER entries
        
    Returns:
        dict: Dictionary containing comparison results
    """
    write_log("\n=========================================", "YELLOW")
    write_log("COMPARING GSN AND ER ENTRIES", "YELLOW")
    write_log("=========================================", "YELLOW")
    
    # Find entries in GSN but not in ER
    missing_in_er = [item for item in gsn_entries if item not in er_entries]
    
    # Find entries in ER but not in GSN
    missing_in_gsn = [item for item in er_entries if item not in gsn_entries]
    
    # Report GSN entries not in ER
    if missing_in_er:
        write_log("\nIn GSN but not in ER:", "MAGENTA")
        for item in sorted(missing_in_er):
            write_log(f"  {item}", "MAGENTA")
    else:
        write_log("\nNo entries in GSN that are not in ER.", "GREEN")
    
    # Report ER entries not in GSN
    if missing_in_gsn:
        write_log("\nIn ER but not in GSN:", "CYAN")
        for item in sorted(missing_in_gsn):
            write_log(f"  {item}", "CYAN")
    else:
        write_log("\nNo entries in ER that are not in GSN.", "GREEN")
    
    # Create summary of comparison results
    write_log("\nComparison Summary:", "YELLOW")
    write_log(f"- Total GSN entries: {len(gsn_entries)}", "WHITE")
    write_log(f"- Total ER entries: {len(er_entries)}", "WHITE")
    write_log(f"- GSN entries not in ER: {len(missing_in_er)}", "MAGENTA")
    write_log(f"- ER entries not in GSN: {len(missing_in_gsn)}", "CYAN")
    write_log("=========================================", "YELLOW")
    
    return {
        "MissingInER": missing_in_er,
        "MissingInGSN": missing_in_gsn
    }

def format_date_range(date_range, full_month_name=False):
    """
    Format a date range object
    
    Args:
        date_range: DateRangeResult object
        full_month_name (bool): Whether to use full month names
        
    Returns:
        str: Formatted date range string
    """
    if not date_range:
        return ""
    
    format_string = "%B" if full_month_name else "%b"
    
    start_date = date_range.start_date
    end_date = date_range.end_date
    
    if start_date.month == end_date.month and start_date.year == end_date.year:
        # Same month format: "15-17 Apr 2025" or "15-17 April 2025"
        month_format = start_date.strftime(format_string)
        return f"{start_date.day}-{end_date.day} {month_format} {start_date.year}"
    else:
        # Different month format: "15 Apr - 17 May 2025" or "15 April - 17 May 2025"
        start_month_format = start_date.strftime(format_string)
        end_month_format = end_date.strftime(format_string)
        return f"{start_date.day} {start_month_format} - {end_date.day} {end_month_format} {end_date.year}"

class ExcelUpdater:
    """Class to update Excel files with comparison data"""
    
    def __init__(self, file_path):
        """
        Initialize Excel updater
        
        Args:
            file_path (str): Path to the Excel file
        """
        self.file_path = file_path
        self.excel_app = ExcelApplication()
        self.workbook = None
        
    def analyze_excel_file(self, gsn_entries, er_entries, ad_entries, date_range, 
                           missing_in_er, missing_in_gsn, filtered_hostnames2, er_serial_number):
        """
        Update Excel file with comparison data
        
        Args:
            gsn_entries (list): GSN entries
            er_entries (list): ER entries
            ad_entries (list): AD entries
            date_range: DateRangeResult object
            missing_in_er (list): GSN entries not in ER
            missing_in_gsn (list): ER entries not in GSN
            filtered_hostnames2 (list): Filtered hostnames (31-60 days)
            er_serial_number (list): ER serial numbers
            
        Returns:
            bool: Success status
        """
        if not self.excel_app.open_workbook(self.file_path):
            write_log(f"Failed to open Excel file: {self.file_path}", "RED")
            return False
        
        self.workbook = self.excel_app.workbook
            
        try:
            # Display worksheet information
            self._display_worksheet_info()
            
            # Update GSN vs ER worksheet
            write_log("\nUpdating GSN vs ER worksheet...", "YELLOW")
            self._update_gsner_worksheet(gsn_entries, er_entries, missing_in_er, missing_in_gsn, date_range)
            
            # Update ER NO LOGON worksheet
            write_log("\nUpdating ER NO LOGON worksheet...", "YELLOW")
            self._update_er_nologon_worksheet(date_range, filtered_hostnames2, er_serial_number)
            
            # Update GSN VS AD worksheet
            write_log("\nUpdating GSN VS AD worksheet...", "YELLOW")
            self._update_gsnvsad_worksheet(date_range, gsn_entries, ad_entries)
            
            # Save the workbook
            write_log("\nSaving changes to file: " + self.file_path, "YELLOW")
            self.excel_app.save()
            
            write_log("Excel file updates completed successfully", "GREEN")
            return True
            
        except Exception as e:
            write_log(f"Error analyzing Excel file: {str(e)}", "RED")
            import traceback
            write_log(traceback.format_exc(), "RED")
            return False
            
        finally:
            # Close the workbook
            self.excel_app.close()
    
    def _find_available_worksheet_name(self, base_name):
        """
        Find an available worksheet name by checking if base_name exists,
        and if so, append (copy) or (copy 2), (copy 3), etc.
        
        Args:
            base_name (str): The desired base worksheet name
            
        Returns:
            str: Available worksheet name
        """
        if not self.workbook:
            write_log("No workbook is open", "RED")
            return base_name
        
        try:
            # Check if base name is available
            if not self._worksheet_exists(base_name):
                write_log(f"Worksheet name '{base_name}' is available", "GREEN")
                return base_name
            
            write_log(f"Worksheet '{base_name}' already exists, finding alternative name...", "YELLOW")
            
            # Try with (copy) suffix
            copy_name = f"{base_name} (copy)"
            if not self._worksheet_exists(copy_name):
                write_log(f"Using worksheet name '{copy_name}'", "GREEN")
                return copy_name
            
            # Try with numbered copies: (copy 2), (copy 3), etc.
            copy_number = 2
            while copy_number <= 100:  # Reasonable limit to prevent infinite loop
                numbered_copy_name = f"{base_name} (copy {copy_number})"
                if not self._worksheet_exists(numbered_copy_name):
                    write_log(f"Using worksheet name '{numbered_copy_name}'", "GREEN")
                    return numbered_copy_name
                copy_number += 1
            
            # If we reach here, use a timestamp-based name as fallback
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            fallback_name = f"{base_name} ({timestamp})"
            write_log(f"Using timestamp-based fallback name '{fallback_name}'", "YELLOW")
            return fallback_name
            
        except Exception as e:
            write_log(f"Error finding available worksheet name: {str(e)}", "RED")
            # Return base name as fallback
            return base_name

    def _worksheet_exists(self, worksheet_name):
        """
        Check if a worksheet with the given name exists in the workbook
        
        Args:
            worksheet_name (str): Name of the worksheet to check
            
        Returns:
            bool: True if worksheet exists, False otherwise
        """
        try:
            # Try to access the worksheet
            worksheet = self.workbook.Worksheets(worksheet_name)
            return True
        except:
            # Worksheet doesn't exist
            return False

    def _create_new_worksheet(self, worksheet_name):
        """
        Create a new worksheet with the specified name
        
        Args:
            worksheet_name (str): Name for the new worksheet
            
        Returns:
            worksheet: Excel worksheet object or None if creation failed
        """
        if not self.workbook:
            write_log("No workbook is open", "RED")
            return None
        
        # Clean up any existing default sheets first
        self._cleanup_default_sheets()
        
        # Validate and clean the worksheet name
        clean_name = self._clean_worksheet_name(worksheet_name)
        write_log(f"Creating new worksheet '{clean_name}'...", "YELLOW")
        
        max_retries = 3
        worksheet = None
        
        for attempt in range(max_retries):
            try:
                # Create the worksheet
                worksheet = self.workbook.Worksheets.Add()
                original_name = worksheet.Name  # Store the default name
                
                # Try to rename it immediately
                try:
                    worksheet.Name = clean_name
                    write_log(f"Successfully created and named worksheet '{clean_name}'", "GREEN")
                    return worksheet
                    
                except Exception as name_error:
                    write_log(f"Error setting worksheet name '{clean_name}': {str(name_error)}", "YELLOW")
                    
                    # Delete the failed worksheet immediately
                    try:
                        write_log(f"Deleting failed worksheet '{original_name}'", "YELLOW")
                        worksheet.Delete()
                    except Exception as del_error:
                        write_log(f"Could not delete failed worksheet: {str(del_error)}", "RED")
                    
                    worksheet = None
                    
                    # Try with a simplified fallback name on last attempt
                    if attempt == max_retries - 1:
                        try:
                            fallback_name = f"Report_{datetime.datetime.now().strftime('%H%M%S')}"
                            worksheet = self.workbook.Worksheets.Add()
                            worksheet.Name = fallback_name
                            write_log(f"Used simple fallback name '{fallback_name}'", "YELLOW")
                            return worksheet
                        except Exception as fallback_error:
                            write_log(f"Even simple fallback failed: {str(fallback_error)}", "RED")
                            if worksheet:
                                try:
                                    worksheet.Delete()
                                except:
                                    pass
                            return None
                        
            except Exception as e:
                write_log(f"Error creating worksheet (attempt {attempt+1}/{max_retries}): {str(e)}", "YELLOW")
                if worksheet:
                    try:
                        worksheet.Delete()
                    except:
                        pass
                    worksheet = None
                
                if attempt < max_retries - 1:
                    write_log(f"Retrying in 1 second...", "YELLOW")
                    time.sleep(1)
        
        write_log(f"Failed to create worksheet '{clean_name}' after all attempts", "RED")
        return None
    
    def _cleanup_default_sheets(self):
        """Remove any default sheets (Sheet1, Sheet2, etc.) that might exist"""
        try:
            # Get list of worksheets
            sheets_to_delete = []
            
            for i in range(1, self.workbook.Worksheets.Count + 1):
                worksheet = self.workbook.Worksheets(i)
                sheet_name = worksheet.Name
                
                # Check if it's a default sheet name and is empty
                if (sheet_name.startswith("Sheet") and 
                    sheet_name[5:].isdigit() and 
                    self._is_worksheet_empty(worksheet)):
                    sheets_to_delete.append(worksheet)
            
            # Delete the empty default sheets
            for sheet in sheets_to_delete:
                try:
                    write_log(f"Cleaning up empty default sheet: {sheet.Name}", "CYAN")
                    sheet.Delete()
                except Exception as del_error:
                    write_log(f"Could not delete sheet {sheet.Name}: {str(del_error)}", "YELLOW")
                    
        except Exception as e:
            write_log(f"Error during sheet cleanup: {str(e)}", "YELLOW")
    
    def _is_worksheet_empty(self, worksheet):
        """Check if a worksheet is empty (has no data)"""
        try:
            used_range = worksheet.UsedRange
            if used_range is None:
                return True
            
            # Check if the used range is just one cell with no value
            if (used_range.Rows.Count == 1 and 
                used_range.Columns.Count == 1 and 
                not used_range.Value):
                return True
                
            return False
            
        except:
            # If we can't determine, assume it's not empty to be safe
            return False
    
    def _clean_worksheet_name(self, name):
        """
        Clean worksheet name to ensure it's valid for Excel
        
        Args:
            name (str): Original worksheet name
            
        Returns:
            str: Cleaned worksheet name
        """
        if not name:
            return "Sheet"
        
        # Excel worksheet name restrictions:
        # - Max 31 characters
        # - Cannot contain: \ / ? * [ ] :
        # - Cannot be empty
        # - Cannot be "History" (reserved)
        
        # Remove invalid characters
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        clean_name = name
        for char in invalid_chars:
            clean_name = clean_name.replace(char, '_')
        
        # Trim to 31 characters
        if len(clean_name) > 31:
            clean_name = clean_name[:31]
        
        # Ensure it's not empty
        if not clean_name.strip():
            clean_name = "Sheet"
        
        # Avoid reserved name
        if clean_name.lower() == "history":
            clean_name = "Report_History"
        
        return clean_name.strip()

    def _generate_unique_suffix(self):
        """Generate a unique suffix for table names to avoid conflicts"""
        import datetime
        return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    def _set_cell_value_safely(self, worksheet, row, col, value, max_retries=3):
        """Safely set cell value with retry mechanism and error handling"""
        for attempt in range(max_retries):
            try:
                # Convert value to string to avoid type issues
                str_value = str(value) if value is not None else ""
                
                # Set the cell value
                cell = worksheet.Cells(row, col)
                cell.Value = str_value
                return True
                
            except Exception as e:
                if attempt < max_retries - 1:
                    write_log(f"Error setting cell ({row}, {col}) attempt {attempt+1}: {str(e)}", "YELLOW")
                    time.sleep(0.1)  # Short delay before retry
                else:
                    write_log(f"Failed to set cell ({row}, {col}) value '{value}' after {max_retries} attempts: {str(e)}", "RED")
                    return False

    def _highlight_matching_cells_safely(self, worksheet, sorted_gsn_values, sorted_er_values, max_row):
        """Safely highlight matching cells with error handling"""
        try:
            # Colors for highlighting - pink for matches, white for non-matches
            pink_color = 0xC1B6FF  # Decimal equivalent of RGB(244, 204, 204)
            white_color = 0xFFFFFF  # Decimal equivalent of RGB(255, 255, 255)
            
            # Process column A (GSN) - highlight cells that match with column B
            write_log(f"Highlighting {len(sorted_gsn_values)} GSN entries...", "CYAN")
            for i in range(len(sorted_gsn_values)):
                try:
                    value_a = sorted_gsn_values[i]
                    cell_a = worksheet.Cells(i + 2, 1)  # +2 because we start at row 2 (after header)
                    
                    if value_a in sorted_er_values:
                        cell_a.Interior.Color = pink_color
                    else:
                        cell_a.Interior.Color = white_color
                        
                except Exception as highlight_error:
                    write_log(f"Warning: Could not highlight GSN cell {i+2}: {str(highlight_error)}", "YELLOW")
            
            # Process column B (ER) - highlight cells that match with column A
            write_log(f"Highlighting {len(sorted_er_values)} ER entries...", "CYAN")
            for j in range(len(sorted_er_values)):
                try:
                    value_b = sorted_er_values[j]
                    cell_b = worksheet.Cells(j + 2, 2)  # +2 because we start at row 2 (after header)
                    
                    if value_b in sorted_gsn_values:
                        cell_b.Interior.Color = pink_color
                    else:
                        cell_b.Interior.Color = white_color
                        
                except Exception as highlight_error:
                    write_log(f"Warning: Could not highlight ER cell {j+2}: {str(highlight_error)}", "YELLOW")
            
            write_log("Cell highlighting completed", "GREEN")
            
        except Exception as e:
            write_log(f"Error in cell highlighting: {str(e)}", "RED")
            write_log("Continuing without cell highlighting...", "YELLOW")

    def _date_range_exists_in_worksheet(self, worksheet, date_range_text):
        """
        Check if a date range already exists in the worksheet
        
        Args:
            worksheet: Excel worksheet object
            date_range_text (str): Date range text to search for
            
        Returns:
            bool: True if date range exists, False otherwise
        """
        try:
            # Search the first column for the date range text
            last_used_range = worksheet.UsedRange
            if not last_used_range:
                return False
                
            last_row = last_used_range.Row + last_used_range.Rows.Count - 1
            
            for i in range(1, last_row + 1):
                try:
                    cell_value = worksheet.Cells(i, 1).Text
                    if str(cell_value).strip() == date_range_text:
                        return True
                except:
                    continue
                    
            return False
            
        except Exception as e:
            write_log(f"Error checking if date range exists: {str(e)}", "YELLOW")
            return False
            
    def _display_worksheet_info(self):
        """Display information about worksheets in the workbook"""
        if not self.excel_app.workbook:
            return
            
        write_log("\nWorksheets in the Excel file:", "GREEN")
        write_log("-------------------------------------------", "GREEN")
        
        visible_count = 0
        hidden_count = 0
        very_hidden_count = 0
        
        # Get worksheets collection
        for i in range(1, self.excel_app.workbook.Worksheets.Count + 1):
            worksheet = self.excel_app.workbook.Worksheets(i)
            visibility_value = worksheet.Visible
            
            if visibility_value == -1:
                visibility = "Visible"
                color = "WHITE"
                visible_count += 1
            elif visibility_value == 0:
                visibility = "Hidden"
                color = "YELLOW"
                hidden_count += 1
            else:
                visibility = "Very Hidden"
                color = "RED"
                very_hidden_count += 1
                
            write_log(f"Sheet {i}: {worksheet.Name} [{visibility}]", color)
            
        # Display summary
        total_count = visible_count + hidden_count + very_hidden_count
        write_log("\nSummary:", "CYAN")
        write_log(f"Total worksheets: {total_count}", "WHITE")
        write_log(f"Visible worksheets: {visible_count}", "WHITE")
        write_log(f"Hidden worksheets: {hidden_count}", "YELLOW")
        write_log(f"Very hidden worksheets: {very_hidden_count}", "RED")
    
    def _update_gsner_worksheet(self, gsn_entries, er_entries, missing_in_er, missing_in_gsn, date_range):
        """Update the GSN vs ER worksheet with auto-copy naming for existing worksheets"""
        # Format the date range for worksheet name
        date_range_formatted = format_date_range(date_range)
        
        # Set the base worksheet name
        base_worksheet_name = "GSN VS ER"
        if date_range_formatted:
            base_worksheet_name = f"GSN VS ER {date_range_formatted}"
            
        # Find an available worksheet name (adds (copy) if needed)
        worksheet_name = self._find_available_worksheet_name(base_worksheet_name)
        write_log(f"Using worksheet name: {worksheet_name}", "YELLOW")
        
        # Create the new worksheet
        worksheet = self._create_new_worksheet(worksheet_name)
        if not worksheet:
            return False
        
        try:
            # Set the header for the table
            self._set_cell_value_safely(worksheet, 1, 1, "GSN")   # Column A header
            self._set_cell_value_safely(worksheet, 1, 2, "ER")    # Column B header
            
            # Format headers
            try:
                worksheet.Cells(1, 1).Font.Bold = True
                worksheet.Cells(1, 2).Font.Bold = True
                worksheet.Cells(1, 1).HorizontalAlignment = -4108  # Center
                worksheet.Cells(1, 2).HorizontalAlignment = -4108  # Center
            except Exception as format_error:
                write_log(f"Warning: Could not format headers: {str(format_error)}", "YELLOW")
            
            # Load the extracted values into the first column (Column A - GSN)
            sorted_gsn_values = sorted(gsn_entries)
            write_log(f"Writing {len(sorted_gsn_values)} GSN entries to column A...", "CYAN")
            
            row_index = 2  # Start filling from row 2
            for i, value in enumerate(sorted_gsn_values):
                try:
                    self._set_cell_value_safely(worksheet, row_index, 1, value)  # Column A (GSN)
                    row_index += 1
                    
                    # Progress indicator for large datasets
                    if i > 0 and i % 100 == 0:
                        write_log(f"  Written {i} GSN entries...", "CYAN")
                        
                except Exception as cell_error:
                    write_log(f"Error writing GSN entry '{value}' to row {row_index}: {str(cell_error)}", "RED")
                    row_index += 1  # Continue with next row
            
            # Load the filtered hostnames into the second column (Column B - ER)
            sorted_er_values = sorted(er_entries)
            write_log(f"Writing {len(sorted_er_values)} ER entries to column B...", "CYAN")
            
            row_index = 2  # Reset to start filling from row 2
            for i, value in enumerate(sorted_er_values):
                try:
                    self._set_cell_value_safely(worksheet, row_index, 2, value)  # Column B (ER)
                    row_index += 1
                    
                    # Progress indicator for large datasets
                    if i > 0 and i % 100 == 0:
                        write_log(f"  Written {i} ER entries...", "CYAN")
                        
                except Exception as cell_error:
                    write_log(f"Error writing ER entry '{value}' to row {row_index}: {str(cell_error)}", "RED")
                    row_index += 1  # Continue with next row
            
            # Determine the max row count between both primary columns
            max_row = max(len(sorted_gsn_values), len(sorted_er_values)) + 1  # +1 for the header row
            
            # Create a table from the range for GSN and ER with error handling
            try:
                table_range = worksheet.Range(f"A1:B{max_row}")
                
                # Create a new table for GSN and ER
                new_table = worksheet.ListObjects.Add(1, table_range, None, 1)
                new_table.Name = f"GSN_ER_Table_{self._generate_unique_suffix()}"  # Unique table name
                new_table.TableStyle = "TableStyleLight15"
                new_table.ShowTableStyleRowStripes = False
                
                write_log("Created new GSN_ER_Table successfully", "GREEN")
                
            except Exception as table_error:
                write_log(f"Warning: Could not create table: {str(table_error)}", "YELLOW")
                write_log("Continuing without table formatting...", "YELLOW")
            
            # Highlight matching cells
            write_log("Highlighting matching cells between GSN and ER columns...", "YELLOW")
            self._highlight_matching_cells_safely(worksheet, sorted_gsn_values, sorted_er_values, max_row)
            
            # Set the tab color to gold with error handling
            try:
                gold_color = 0xABE5F3 
                worksheet.Tab.Color = gold_color
            except Exception as color_error:
                write_log(f"Warning: Could not set tab color: {str(color_error)}", "YELLOW")
            
            # === "In GSN but not in ER" Section ===
            row_after_missing_in_er = self._add_comparison_section(
                worksheet, missing_in_er, "In GSN but not in ER", "Remarks", 4, 1)
            
            # === "In ER but not in GSN" Section ===
            row_after_missing_in_gsn = self._add_comparison_section(
                worksheet, missing_in_gsn, "In ER but not in GSN", "Remarks", 4, row_after_missing_in_er + 1)
            
            # === Count Summary Section ===
            self._add_count_summary(
                worksheet, ["ER", "GSN"], [len(er_entries), len(gsn_entries)], 4, row_after_missing_in_gsn + 1)
            
            # Auto-fit the columns with error handling
            try:
                worksheet.Columns.AutoFit()
            except Exception as autofit_error:
                write_log(f"Warning: Could not auto-fit columns: {str(autofit_error)}", "YELLOW")
            
            write_log("GSN vs ER worksheet updated successfully!", "GREEN")
            return True
            
        except Exception as e:
            write_log(f"Error updating GSN vs ER worksheet: {str(e)}", "RED")
            import traceback
            write_log(traceback.format_exc(), "RED")
            return False
            
    def _update_er_nologon_worksheet(self, date_range, filtered_hostnames2, er_serial_number):
        """Update the ER NO LOGON DETAILS worksheet - adds entries even if date range exists"""
        if not date_range or not date_range.year:
            write_log("No year information available in date range object. Skipping year-specific worksheet update.", "YELLOW")
            return False
                
        # Format date range
        date_range_formatted = format_date_range(date_range)
        
        year_worksheet_name = f"ER {date_range.year}"
        write_log(f"Looking for worksheet: '{year_worksheet_name}'", "CYAN")
        
        # Find the worksheet - if it doesn't exist, create it
        year_worksheet = None
        try:
            year_worksheet = self.excel_app.workbook.Worksheets(year_worksheet_name)
            write_log(f"Found existing worksheet '{year_worksheet_name}'", "GREEN")
                    
        except:
            write_log(f"Worksheet '{year_worksheet_name}' not found, creating new one...", "YELLOW")
            year_worksheet = self._create_new_worksheet(year_worksheet_name)
            if not year_worksheet:
                return False
        
        # Check if this date range already exists in the worksheet
        if self._date_range_exists_in_worksheet(year_worksheet, date_range_formatted):
            write_log(f"Date range '{date_range_formatted}' already exists in worksheet. Adding new entry anyway...", "YELLOW")
        else:
            write_log(f"Date range '{date_range_formatted}' is new. Adding entry...", "GREEN")
                
        # Find the last used row in the worksheet
        try:
            last_used_range = year_worksheet.UsedRange
            if last_used_range:
                last_row = last_used_range.Row + last_used_range.Rows.Count - 1
            else:
                last_row = 0
            
            write_log(f"Last used row in worksheet: Row {last_row}", "CYAN")
            
            # Start at the row after the last used row
            start_row = last_row + 1
            
            # Add date range header with merged cells
            self._add_merged_header(
                year_worksheet, date_range_formatted, start_row, 1, 3, bg_color=0xAAAAAE)
            
            # Process the filtered hostnames data (31-60 days)
            if not filtered_hostnames2:
                # If no data, display "NIL"
                nil_row = start_row + 1
                nil_range = year_worksheet.Range(f"A{nil_row}:C{nil_row}")
                nil_range.Merge()
                year_worksheet.Cells(nil_row, 1).Value = "NIL"
                year_worksheet.Cells(nil_row, 1).HorizontalAlignment = -4108  # Center
                
                # Add borders to the merged NIL cell - all sides
                self._add_borders_to_range(nil_range)
                
                write_log("No devices found with login between 31-60 days. Added 'NIL' entry.", "MAGENTA")
            else:
                # Add the hostname and serial number data
                write_log(f"Adding {len(filtered_hostnames2)} devices with login between 31-60 days.", "CYAN")
                
                for i in range(len(filtered_hostnames2)):
                    current_row = start_row + 1 + i
                    hostname = filtered_hostnames2[i]
                    serial_num = er_serial_number[i]
                    
                    # Set values in columns A and B
                    year_worksheet.Cells(current_row, 1).Value = hostname  # Column A - Hostname
                    year_worksheet.Cells(current_row, 2).Value = serial_num  # Column B - Serial Number
                    
                    # Add borders to each cell
                    for col in range(1, 4):
                        self._add_borders_to_cell(year_worksheet, current_row, col)
            
            # Auto-fit the columns
            year_worksheet.Columns.AutoFit()
            
            write_log(f"Successfully updated worksheet with ER NO LOGON DETAILS", "GREEN")
            return True
            
        except Exception as e:
            write_log(f"Error while updating year worksheet: {str(e)}", "RED")
            return False
            
    def _update_gsnvsad_worksheet(self, date_range, gsn_entries, ad_entries):
        """Update the GSN VS AD worksheet with copy naming logic"""
        if not date_range or not date_range.year:
            write_log("No year information available in date range object. Skipping GSN VS AD year worksheet update.", "YELLOW")
            return False
                    
        year_value = date_range.year
        base_target_worksheet_name = f"GSN VS AD {year_value}"
                
        write_log(f"Looking for worksheet: '{base_target_worksheet_name}'", "CYAN")
                
        # Find the worksheet - if it doesn't exist, create it
        gsn_ad_year_worksheet = None
        try:
            gsn_ad_year_worksheet = self.excel_app.workbook.Worksheets(base_target_worksheet_name)
            write_log(f"Found existing worksheet '{base_target_worksheet_name}'", "GREEN")
            
            # Check if this date range data might conflict
            date_range_text = format_date_range(date_range, full_month_name=True)
            if self._date_range_exists_in_worksheet(gsn_ad_year_worksheet, date_range_text):
                write_log(f"Date range '{date_range_text}' already exists. Creating copy worksheet...", "YELLOW")
                
                # Find available name and create new worksheet
                available_name = self._find_available_worksheet_name(base_target_worksheet_name)
                gsn_ad_year_worksheet = self._create_new_worksheet(available_name)
                if not gsn_ad_year_worksheet:
                    return False
                    
        except:
            write_log(f"Worksheet '{base_target_worksheet_name}' not found, creating new one...", "YELLOW")
            gsn_ad_year_worksheet = self._create_new_worksheet(base_target_worksheet_name)
            if not gsn_ad_year_worksheet:
                return False
                    
        try:
            # Find the last used row in the worksheet
            last_used_range = gsn_ad_year_worksheet.UsedRange
            if last_used_range:
                last_row = last_used_range.Row + last_used_range.Rows.Count - 1
            else:
                last_row = 0
                    
            write_log(f"Last used row in worksheet: Row {last_row}", "CYAN")
                    
            # Format date range with full month name
            date_range_text = format_date_range(date_range, full_month_name=True)
                    
            # Extract date parts to determine if month header is needed
            date_parts = date_range_text.split(" ")
                    
            # Extract the start date
            start_date = 0
            if "-" in date_parts[0]:
                # Format like "15-17 April 2025"
                start_date = int(date_parts[0].split("-")[0])
            else:
                # Format like "15 April - 17 May 2025"
                start_date = int(date_parts[0])
                    
            # Find the appropriate month name
            month_name = date_parts[1].upper() if len(date_parts) > 1 else ""
                    
            # Determine if month header should be shown (when start date is less than 5)
            show_month_header = start_date < 5
                    
            # Month header row position
            month_row = last_row + 2
                    
            # Add month header if needed
            if show_month_header:
                self._add_merged_header(
                    gsn_ad_year_worksheet, month_name, month_row, 1, 6, 
                    font_size=12, font_color=0xFF7B00)
                        
                # Set position for date range header
                start_row = month_row + 1
            else:
                # No month header, use month row for date range
                start_row = month_row
                    
            # Add date range header
            self._add_merged_header(
                gsn_ad_year_worksheet, f"{date_range_text} GSN VS AD", 
                start_row, 1, 6, bg_color=0xAAAAAE, border_weight=2)
                    
            # Add column headers
            second_row = start_row + 1
            headers = ["In GSN not in AD", "Remarks", "Action", "In AD not in GSN", "Remarks", "Action"]
                    
            for i, header in enumerate(headers):
                cell = gsn_ad_year_worksheet.Cells(second_row, i + 1)
                cell.Value = header
                cell.Font.Bold = True
                cell.HorizontalAlignment = -4108  # Center
                cell.Interior.Color = 65535  # Yellow
                cell.Borders.Weight = 2
                    
            # Compare GSN and AD datasets directly within this method
            # Normalize both lists to ensure consistent comparison
            gsn_normalized = [str(item).strip() for item in gsn_entries if item]
            ad_normalized = [str(item).strip() for item in ad_entries if item]
                    
            # Find entries in GSN but not in AD
            missing_in_ad = [item for item in gsn_normalized if item not in ad_normalized]
                    
            # Find entries in AD but not in GSN
            missing_in_gsn = [item for item in ad_normalized if item not in gsn_normalized]
                    
            # Sort both lists for consistent output
            missing_in_ad.sort()
            missing_in_gsn.sort()
                    
            # Get the max length of the two arrays
            max_length = max(len(missing_in_ad), len(missing_in_gsn))
                    
            # Add the data rows
            write_log(f"Starting to add data rows. Max length: {max_length}", "CYAN")
            write_log(f"Missing in AD count: {len(missing_in_ad)}", "CYAN")
            write_log(f"Missing in GSN count: {len(missing_in_gsn)}", "CYAN")
                    
            for i in range(max_length):
                current_row = second_row + 1 + i
                        
                # Set value in column A (In GSN not in AD)
                if i < len(missing_in_ad):
                    value = missing_in_ad[i]
                    gsn_ad_year_worksheet.Cells(current_row, 1).NumberFormat = "@"  # Force text format
                    gsn_ad_year_worksheet.Cells(current_row, 1).Value = str(value)
                        
                # Set value in column D (In AD not in GSN)
                if i < len(missing_in_gsn):
                    value = missing_in_gsn[i]
                    gsn_ad_year_worksheet.Cells(current_row, 4).NumberFormat = "@"  # Force text format
                    gsn_ad_year_worksheet.Cells(current_row, 4).Value = str(value)
                        
                # Add borders to all cells in the row
                for col in range(1, 7):
                    gsn_ad_year_worksheet.Cells(current_row, col).Borders.Weight = 2
            
            # Auto-fit the columns
            gsn_ad_year_worksheet.Columns.AutoFit()
                    
            write_log(f"Successfully updated worksheet with GSN VS AD comparison data", "GREEN")
            write_log(f"- In GSN not in AD entries: {len(missing_in_ad)}", "MAGENTA")
            write_log(f"- In AD not in GSN entries: {len(missing_in_gsn)}", "CYAN")
                    
            return True
                    
        except Exception as e:
            write_log(f"Error while updating GSN VS AD year worksheet: {str(e)}", "RED")
            import traceback
            write_log(traceback.format_exc(), "RED")
            return False
            
    def _add_comparison_section(self, worksheet, data, header_text, remarks_header, start_col, start_row):
        """Add a comparison section to a worksheet"""
        # Add headers
        worksheet.Cells(start_row, start_col).Value = header_text
        worksheet.Cells(start_row, start_col + 1).Value = remarks_header
        
        # Format headers
        worksheet.Cells(start_row, start_col).Font.Bold = True
        worksheet.Cells(start_row, start_col + 1).Font.Bold = True
        
        # Add borders to headers
        worksheet.Cells(start_row, start_col).Borders.Weight = 2
        worksheet.Cells(start_row, start_col + 1).Borders.Weight = 2
        
        # Sort data
        sorted_data = sorted(data)
        
        if not sorted_data:
            # If no entries, display "NIL"
            nil_range = worksheet.Range(
                f"{chr(64 + start_col)}{start_row + 1}:{chr(64 + start_col + 1)}{start_row + 1}")
            nil_range.Merge()
            worksheet.Cells(start_row + 1, start_col).Value = "NIL"
            worksheet.Cells(start_row + 1, start_col).HorizontalAlignment = -4108  # Center
            nil_range.Borders.Weight = 2
            return start_row + 2  # Next row after the NIL row
        else:
            # Add data
            row_index = start_row + 1
            for value in sorted_data:
                worksheet.Cells(row_index, start_col).Value = value
                
                # Add borders
                worksheet.Cells(row_index, start_col).Borders.Weight = 2
                worksheet.Cells(row_index, start_col + 1).Borders.Weight = 2
                
                row_index += 1
            return row_index  # Next row after the last entry
            
    def _add_count_summary(self, worksheet, labels, values, start_col, start_row):
        """Add a count summary section to a worksheet"""
        for i in range(len(labels)):
            # Add label
            worksheet.Cells(start_row + i, start_col).Value = labels[i]
            
            # Format label
            worksheet.Cells(start_row + i, start_col).Font.Bold = True
            worksheet.Cells(start_row + i, start_col).HorizontalAlignment = -4152  # Right
            
            # Add value
            worksheet.Cells(start_row + i, start_col + 1).Value = values[i]
            
            # Add borders
            worksheet.Cells(start_row + i, start_col).Borders.Weight = 2
            worksheet.Cells(start_row + i, start_col + 1).Borders.Weight = 2
            
    def _add_merged_header(self, worksheet, header_text, start_row, start_col, end_col, 
                        bg_color=-1, border_weight=-1, font_size=-1, font_color=-1):
        """Add a merged header to a worksheet"""
        # Set header text
        worksheet.Cells(start_row, start_col).Value = header_text
        worksheet.Cells(start_row, start_col).Font.Bold = True
        
        if font_size > 0:
            worksheet.Cells(start_row, start_col).Font.Size = font_size
            
        if font_color >= 0:
            worksheet.Cells(start_row, start_col).Font.Color = font_color
            
        try:
            # Try to clear any existing merges
            try:
                merge_range = worksheet.Range(
                    f"{chr(64 + start_col)}{start_row}:{chr(64 + end_col)}{start_row}")
                merge_range.MergeCells = False
            except:
                pass
                
            # Merge cells
            first_cell = worksheet.Cells(start_row, start_col)
            last_cell = worksheet.Cells(start_row, end_col)
            merge_range = worksheet.Range(first_cell, last_cell)
            merge_range.MergeCells = True
            
            # Apply formatting
            worksheet.Cells(start_row, start_col).HorizontalAlignment = -4108  # Center
            
            if bg_color >= 0:
                merge_range.Interior.Color = bg_color
                
            if border_weight > 0:
                merge_range.Borders.Weight = border_weight
                
            return True
            
        except Exception as e:
            write_log(f"Error merging cells: {str(e)}", "RED")
            
            # Fallback - just style the first cell without merging
            worksheet.Cells(start_row, start_col).HorizontalAlignment = -4108  # Center
            
            if bg_color >= 0:
                worksheet.Cells(start_row, start_col).Interior.Color = bg_color
                
            if border_weight > 0:
                worksheet.Cells(start_row, start_col).Borders.Weight = border_weight
                
            return False
            
    def _add_borders_to_range(self, range_obj, line_style=1):
        """Add borders to a range"""
        xl_edge_left = 7
        xl_edge_top = 8
        xl_edge_bottom = 9
        xl_edge_right = 10
        
        range_obj.Borders.Item(xl_edge_left).LineStyle = line_style
        range_obj.Borders.Item(xl_edge_top).LineStyle = line_style
        range_obj.Borders.Item(xl_edge_bottom).LineStyle = line_style
        range_obj.Borders.Item(xl_edge_right).LineStyle = line_style
        
    def _add_borders_to_cell(self, worksheet, row, col, line_style=1):
        """Add borders to a cell"""
        xl_edge_left = 7
        xl_edge_top = 8
        xl_edge_bottom = 9
        xl_edge_right = 10
        
        cell = worksheet.Cells(row, col)
        cell.Borders.Item(xl_edge_left).LineStyle = line_style
        cell.Borders.Item(xl_edge_top).LineStyle = line_style
        cell.Borders.Item(xl_edge_bottom).LineStyle = line_style
        cell.Borders.Item(xl_edge_right).LineStyle = line_style