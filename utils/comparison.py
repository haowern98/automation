"""
Data comparison functionality for SharePoint Automation
"""
import datetime
from utils.logger import write_log
from utils.excel_functions import ExcelApplication

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
            return False
            
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
        """Update the GSN vs ER worksheet"""
        # Format the date range for worksheet name
        date_range_formatted = format_date_range(date_range)
        
        # Set the worksheet name
        worksheet_name = "GSN VS ER"
        if date_range_formatted:
            worksheet_name = f"GSN VS ER {date_range_formatted}"
            
        write_log(f"Working with worksheet: {worksheet_name}", "YELLOW")
        
        # Find or create worksheet
        worksheet = self.excel_app.find_or_create_worksheet(worksheet_name)
        if not worksheet:
            return False
            
        # Set the header for the table
        worksheet.Cells(1, 1).Value = "GSN"   # Column A header
        worksheet.Cells(1, 2).Value = "ER"    # Column B header
        
        # Format headers
        worksheet.Cells(1, 1).Font.Bold = True
        worksheet.Cells(1, 2).Font.Bold = True
        worksheet.Cells(1, 1).HorizontalAlignment = -4108  # Center
        worksheet.Cells(1, 2).HorizontalAlignment = -4108  # Center
        
        # Load the extracted values into the first column (Column A - GSN)
        # Sort values alphabetically first
        sorted_gsn_values = sorted(gsn_entries)
        row_index = 2  # Start filling from row 2
        for value in sorted_gsn_values:
            worksheet.Cells(row_index, 1).Value = value  # Column A (GSN)
            row_index += 1
        
        # Load the filtered hostnames into the second column (Column B - ER)
        # Sort values alphabetically first
        sorted_er_values = sorted(er_entries)
        row_index = 2  # Reset to start filling from row 2
        for value in sorted_er_values:
            worksheet.Cells(row_index, 2).Value = value  # Column B (ER)
            row_index += 1
        
        # Determine the max row count between both primary columns
        max_row = max(len(sorted_gsn_values), len(sorted_er_values)) + 1  # +1 for the header row
        
        # Create a table from the range for GSN and ER
        table_range = worksheet.Range(f"A1:B{max_row}")
        
        # Delete existing table if it exists
        for i in range(1, worksheet.ListObjects.Count + 1):
            table = worksheet.ListObjects(i)
            if table.Name == "GSN_ER_Table":
                table.Delete()
                break
        
        # Create a new table for GSN and ER
        new_table = worksheet.ListObjects.Add(1, table_range, None, 1)
        new_table.Name = "GSN_ER_Table"
        new_table.TableStyle = "TableStyleLight15"
        new_table.ShowTableStyleRowStripes = False
        
        # Highlight matching cells with pink color
        write_log("Highlighting matching cells between GSN and ER columns...", "YELLOW")
        
        # Colors for highlighting - pink for matches, white for non-matches
        pink_color = 0xC1B6FF  # Decimal equivalent of RGB(244, 204, 204)
        white_color = 0xFFFFFF  # Decimal equivalent of RGB(255, 255, 255)
        
        # Get the column ranges (excluding header)
        column_a_range = worksheet.Range(f"A2:A{max_row}")
        column_b_range = worksheet.Range(f"B2:B{max_row}")
        
        # Process column A (GSN) - highlight cells that match with column B
        for i in range(len(sorted_gsn_values)):
            value_a = sorted_gsn_values[i]
            cell_a = column_a_range.Cells(i + 1, 1)  # +1 because range is 1-based
            
            if value_a in sorted_er_values:
                cell_a.Interior.Color = pink_color
            else:
                cell_a.Interior.Color = white_color
        
        # Process column B (ER) - highlight cells that match with column A
        for j in range(len(sorted_er_values)):
            value_b = sorted_er_values[j]
            cell_b = column_b_range.Cells(j + 1, 1)  # +1 because range is 1-based
            
            if value_b in sorted_gsn_values:
                cell_b.Interior.Color = pink_color
            else:
                cell_b.Interior.Color = white_color
        
        # Set the tab color to gold (#F3E5AB)
        gold_color = 0xABE5F3 
        worksheet.Tab.Color = gold_color
        
        # === "In GSN but not in ER" Section ===
        row_after_missing_in_er = self._add_comparison_section(
            worksheet, missing_in_er, "In GSN but not in ER", "Remarks", 4, 1)
        
        # === "In ER but not in GSN" Section ===
        row_after_missing_in_gsn = self._add_comparison_section(
            worksheet, missing_in_gsn, "In ER but not in GSN", "Remarks", 4, row_after_missing_in_er + 1)
        
        # === Count Summary Section ===
        self._add_count_summary(
            worksheet, ["ER", "GSN"], [len(er_entries), len(gsn_entries)], 4, row_after_missing_in_gsn + 1)
        
        # Auto-fit the columns
        worksheet.Columns.AutoFit()
        
        write_log("GSN vs ER worksheet updated successfully!", "GREEN")
        return True
        
    def _update_er_nologon_worksheet(self, date_range, filtered_hostnames2, er_serial_number):
        """Update the ER NO LOGON DETAILS worksheet"""
        if not date_range or not date_range.year:
            write_log("No year information available in date range object. Skipping year-specific worksheet update.", "YELLOW")
            return False
            
        # Format date range
        date_range_formatted = format_date_range(date_range)
        
        year_worksheet_name = f"ER {date_range.year}"
        write_log(f"Looking for worksheet: '{year_worksheet_name}'", "CYAN")
        
        # Find the worksheet
        year_worksheet = None
        try:
            year_worksheet = self.excel_app.workbook.Worksheets(year_worksheet_name)
            write_log(f"Found worksheet '{year_worksheet_name}'", "GREEN")
        except:
            write_log(f"Worksheet '{year_worksheet_name}' not found in the workbook", "YELLOW")
            return False
            
        # Find the last used row in the worksheet
        try:
            last_used_range = year_worksheet.UsedRange
            last_row = last_used_range.Row + last_used_range.Rows.Count - 1
            
            write_log(f"Last used row in worksheet '{year_worksheet_name}': Row {last_row}", "CYAN")
            
            # Check if date range already exists in the worksheet
            is_date_range_present = False
            for i in range(1, last_row + 1):
                cell_value = year_worksheet.Cells(i, 1).Text
                if cell_value == date_range_formatted:
                    is_date_range_present = True
                    write_log(f"Date range '{date_range_formatted}' is already present in the worksheet. Skipping update.", "YELLOW")
                    break
                    
            # If date range is already present, skip update
            if is_date_range_present:
                return True
                
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
            
            write_log(f"Successfully updated worksheet '{year_worksheet_name}' with ER NO LOGON DETAILS", "GREEN")
            return True
            
        except Exception as e:
            write_log(f"Error while updating year worksheet: {str(e)}", "RED")
            return False
            
    def _update_gsnvsad_worksheet(self, date_range, gsn_entries, ad_entries):
        """Update the GSN VS AD worksheet"""
        if not date_range or not date_range.year:
            write_log("No year information available in date range object. Skipping GSN VS AD year worksheet update.", "YELLOW")
            return False
                
        year_value = date_range.year
        target_worksheet_name = f"GSN VS AD {year_value}"
            
        write_log(f"Looking for worksheet: '{target_worksheet_name}'", "CYAN")
            
        # Find the worksheet
        try:
            gsn_ad_year_worksheet = self.excel_app.workbook.Worksheets(target_worksheet_name)
            write_log(f"Found worksheet '{target_worksheet_name}'", "GREEN")
        except:
            write_log(f"Worksheet '{target_worksheet_name}' not found in the workbook", "YELLOW")
            return False
                
        try:
            # Find the last used row in the worksheet
            last_used_range = gsn_ad_year_worksheet.UsedRange
            last_row = last_used_range.Row + last_used_range.Rows.Count - 1
                
            write_log(f"Last used row in worksheet '{target_worksheet_name}': Row {last_row}", "CYAN")
                
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
                
            write_log(f"Successfully updated worksheet '{target_worksheet_name}' with GSN VS AD comparison data", "GREEN")
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