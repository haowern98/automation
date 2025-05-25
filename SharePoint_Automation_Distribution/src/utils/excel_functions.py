"""
Excel handling utilities for SharePoint Automation
"""
import os
import time
import win32com.client
import pythoncom
import gc
from src.utils.logger import write_log

class ExcelApplication:
    """Class to handle Excel application connections and operations"""
    
    def __init__(self, visible=False):
        """
        Initialize Excel application
        
        Args:
            visible (bool): Whether Excel should be visible
        """
        # Initialize COM in this thread
        pythoncom.CoInitialize()
        
        self.excel = None
        self.workbook = None
        self.worksheet = None
        
        try:
            self.excel = win32com.client.Dispatch("Excel.Application")
            self.excel.Visible = visible
            self.excel.DisplayAlerts = False
            
            # Ensure Excel is ready
            self.ensure_excel_ready()
        except Exception as e:
            write_log(f"Error initializing Excel application: {str(e)}", "RED")
            self.close()
    
    def ensure_excel_ready(self, max_retries=3, retry_delay=2):
        """
        Ensure Excel application is ready for automation
        
        Args:
            max_retries (int): Maximum number of retry attempts
            retry_delay (int): Delay between retries in seconds
            
        Returns:
            bool: Success status
        """
        attempt = 0
        while attempt < max_retries:
            try:
                # Try to access a simple property to check if Excel is responsive
                version = self.excel.Version
                write_log(f"Excel is ready (Version: {version})", "GREEN")
                return True
            except Exception as e:
                attempt += 1
                write_log(f"Excel not ready (attempt {attempt}/{max_retries}): {str(e)}", "YELLOW")
                if attempt < max_retries:
                    write_log(f"Waiting {retry_delay} seconds before retrying...", "YELLOW")
                    time.sleep(retry_delay)
                else:
                    write_log(f"Excel not ready after {max_retries} attempts", "RED")
                    # Try to recreate the Excel application
                    try:
                        self.excel = None
                        pythoncom.CoUninitialize()
                        pythoncom.CoInitialize()
                        self.excel = win32com.client.Dispatch("Excel.Application")
                        self.excel.Visible = False
                        self.excel.DisplayAlerts = False
                        version = self.excel.Version
                        write_log(f"Excel recreated successfully (Version: {version})", "GREEN")
                        return True
                    except Exception as reinit_error:
                        write_log(f"Failed to recreate Excel: {str(reinit_error)}", "RED")
                        return False
    
    def open_workbook(self, file_path, max_retries=3, retry_delay=2):
        """
        Open an Excel workbook with retry mechanism
        
        Args:
            file_path (str): Path to the Excel file
            max_retries (int): Maximum number of retry attempts
            retry_delay (int): Delay between retries in seconds
            
        Returns:
            bool: Success status
        """
        if not os.path.exists(file_path):
            write_log(f"Excel file not found: {file_path}", "RED")
            return False
        
        attempt = 0
        while attempt < max_retries:
            try:
                self.workbook = self.excel.Workbooks.Open(file_path)
                return True
            except Exception as e:
                attempt += 1
                write_log(f"Error opening workbook (attempt {attempt}/{max_retries}): {str(e)}", "YELLOW")
                if attempt < max_retries:
                    write_log(f"Waiting {retry_delay} seconds before retrying...", "YELLOW")
                    time.sleep(retry_delay)
                else:
                    write_log(f"Failed to open workbook after {max_retries} attempts: {str(e)}", "RED")
                    return False
    
    def get_worksheet(self, sheet_name=None, sheet_index=1, max_retries=3, retry_delay=2):
        """
        Get a worksheet by name or index with retry
        
        Args:
            sheet_name (str, optional): Name of the worksheet
            sheet_index (int, optional): Index of the worksheet (1-based)
            max_retries (int): Maximum number of retry attempts
            retry_delay (int): Delay between retries in seconds
            
        Returns:
            worksheet: Excel worksheet object or None
        """
        if not self.workbook:
            write_log("No workbook is open", "RED")
            return None
        
        attempt = 0
        while attempt < max_retries:
            try:
                if sheet_name:
                    self.worksheet = self.workbook.Worksheets(sheet_name)
                else:
                    self.worksheet = self.workbook.Worksheets(sheet_index)
                return self.worksheet
            except Exception as e:
                attempt += 1
                write_log(f"Error getting worksheet (attempt {attempt}/{max_retries}): {str(e)}", "YELLOW")
                if attempt < max_retries:
                    write_log(f"Waiting {retry_delay} seconds before retrying...", "YELLOW")
                    time.sleep(retry_delay)
                else:
                    write_log(f"Failed to get worksheet after {max_retries} attempts: {str(e)}", "RED")
                    return None
    
    def find_or_create_worksheet(self, sheet_name, max_retries=3, retry_delay=2):
        """
        Find an existing worksheet or create a new one with retry
        
        Args:
            sheet_name (str): Name of the worksheet
            max_retries (int): Maximum number of retry attempts
            retry_delay (int): Delay between retries in seconds
            
        Returns:
            worksheet: Excel worksheet object
        """
        if not self.workbook:
            write_log("No workbook is open", "RED")
            return None
            
        attempt = 0
        while attempt < max_retries:
            try:
                # Try to get the worksheet
                try:
                    worksheet = self.workbook.Worksheets(sheet_name)
                    write_log(f"Worksheet '{sheet_name}' already exists. Updating it...", "YELLOW")
                    
                    # Clear existing content
                    worksheet.UsedRange.Clear()
                except:
                    # Create new worksheet
                    write_log(f"Creating new worksheet '{sheet_name}'...", "YELLOW")
                    worksheet = self.workbook.Worksheets.Add()
                    worksheet.Name = sheet_name
                    
                self.worksheet = worksheet
                return worksheet
            except Exception as e:
                attempt += 1
                write_log(f"Error finding/creating worksheet (attempt {attempt}/{max_retries}): {str(e)}", "YELLOW")
                if attempt < max_retries:
                    write_log(f"Waiting {retry_delay} seconds before retrying...", "YELLOW")
                    time.sleep(retry_delay)
                else:
                    write_log(f"Failed to find/create worksheet after {max_retries} attempts: {str(e)}", "RED")
                    return None
    
    def save(self, max_retries=3, retry_delay=2):
        """
        Save the workbook with retry
        
        Args:
            max_retries (int): Maximum number of retry attempts
            retry_delay (int): Delay between retries in seconds
            
        Returns:
            bool: Success status
        """
        if not self.workbook:
            write_log("No workbook to save", "YELLOW")
            return False
            
        attempt = 0
        while attempt < max_retries:
            try:
                self.workbook.Save()
                write_log("Workbook saved successfully", "GREEN")
                return True
            except Exception as e:
                attempt += 1
                write_log(f"Error saving workbook (attempt {attempt}/{max_retries}): {str(e)}", "YELLOW")
                if attempt < max_retries:
                    write_log(f"Waiting {retry_delay} seconds before retrying...", "YELLOW")
                    time.sleep(retry_delay)
                else:
                    write_log(f"Failed to save workbook after {max_retries} attempts: {str(e)}", "RED")
                    return False
    
    def close(self, save_changes=False):
        """
        Close the workbook and release COM objects
        
        Args:
            save_changes (bool): Whether to save changes
        """
        try:
            if self.worksheet:
                self.worksheet = None
                
            if self.workbook:
                try:
                    self.workbook.Close(SaveChanges=save_changes)
                except Exception as wb_error:
                    write_log(f"Error closing workbook: {str(wb_error)}", "YELLOW")
                self.workbook = None
                
            if self.excel:
                try:
                    self.excel.Quit()
                except Exception as excel_error:
                    write_log(f"Error quitting Excel: {str(excel_error)}", "YELLOW")
                self.excel = None
        except Exception as e:
            write_log(f"Error closing Excel objects: {str(e)}", "RED")
        finally:
            # Force garbage collection to ensure COM objects are released
            gc.collect()
            
            # Release COM objects
            try:
                pythoncom.CoUninitialize()
            except:
                pass