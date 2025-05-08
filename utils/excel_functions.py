"""
Excel handling utilities for SharePoint Automation
"""
import os
import win32com.client
import pythoncom
from utils.logger import write_log

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
        
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = visible
        self.excel.DisplayAlerts = False
        self.workbook = None
        self.worksheet = None
    
    def open_workbook(self, file_path):
        """
        Open an Excel workbook
        
        Args:
            file_path (str): Path to the Excel file
            
        Returns:
            bool: Success status
        """
        if not os.path.exists(file_path):
            write_log(f"Excel file not found: {file_path}", "RED")
            return False
        
        try:
            self.workbook = self.excel.Workbooks.Open(file_path)
            return True
        except Exception as e:
            write_log(f"Error opening workbook: {str(e)}", "RED")
            return False
    
    def get_worksheet(self, sheet_name=None, sheet_index=1):
        """
        Get a worksheet by name or index
        
        Args:
            sheet_name (str, optional): Name of the worksheet
            sheet_index (int, optional): Index of the worksheet (1-based)
            
        Returns:
            worksheet: Excel worksheet object or None
        """
        if not self.workbook:
            write_log("No workbook is open", "RED")
            return None
        
        try:
            if sheet_name:
                self.worksheet = self.workbook.Worksheets(sheet_name)
            else:
                self.worksheet = self.workbook.Worksheets(sheet_index)
            return self.worksheet
        except Exception as e:
            write_log(f"Error getting worksheet: {str(e)}", "RED")
            return None
    
    def find_or_create_worksheet(self, sheet_name):
        """
        Find an existing worksheet or create a new one
        
        Args:
            sheet_name (str): Name of the worksheet
            
        Returns:
            worksheet: Excel worksheet object
        """
        if not self.workbook:
            write_log("No workbook is open", "RED")
            return None
            
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
            write_log(f"Error finding/creating worksheet: {str(e)}", "RED")
            return None
    
    def save(self):
        """Save the workbook"""
        if self.workbook:
            try:
                self.workbook.Save()
                return True
            except Exception as e:
                write_log(f"Error saving workbook: {str(e)}", "RED")
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
                self.workbook.Close(SaveChanges=save_changes)
                self.workbook = None
                
            if self.excel:
                self.excel.Quit()
                self.excel = None
        except Exception as e:
            write_log(f"Error closing Excel objects: {str(e)}", "RED")
        finally:
            # Release COM objects
            try:
                pythoncom.CoUninitialize()
            except:
                pass