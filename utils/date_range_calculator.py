"""
Date range calculator utility for SharePoint Automation
Handles automatic date calculation with timeout for user input
"""
import datetime
import threading
import time
import calendar
from utils.logger import write_log
from gui.date_selector import show_date_range_selection, DateRangeResult

def get_date_range_with_timeout(timeout_seconds=120):
    """
    Show date selector dialog with timeout. If user doesn't provide input within
    the timeout period, automatically calculate date range based on current week.
    
    Args:
        timeout_seconds (int): Timeout in seconds
        
    Returns:
        tuple: (primary_date_range, secondary_date_range) where secondary_date_range 
               is None if dates are in same month
    """
    # Variable to store user input date range
    user_date_range = [None]
    dialog_closed = [False]
    
    # Function to run when dialog is shown
    def show_dialog():
        user_date_range[0] = show_date_range_selection()
        dialog_closed[0] = True
    
    # Start the dialog in a separate thread
    dialog_thread = threading.Thread(target=show_dialog)
    dialog_thread.daemon = True
    dialog_thread.start()
    
    # Wait for user input or timeout
    start_time = datetime.datetime.now()
    while (datetime.datetime.now() - start_time).total_seconds() < timeout_seconds and not dialog_closed[0]:
        time.sleep(0.1)
    
    # If user provided input, use it
    if user_date_range[0] is not None:
        write_log("User selected date range", "GREEN")
        return user_date_range[0], None
    
    # If dialog is still open, force close it 
    if not dialog_closed[0]:
        write_log("Timeout reached, automatically calculating date range", "YELLOW")
        # In a more complete implementation, you would close the dialog here
        # But since we're using a separate thread, we'll just let it run and ignore its result
    
    # Calculate date range based on current week
    # today = datetime.date.today()
    today = datetime.date(2025, 4, 18)  # Year, Month, Day
    
    # If today is not Friday, find the next Friday
    while today.weekday() != 4:  # 4 is Friday (0-based, Monday is 0)
        today += datetime.timedelta(days=1)
    
    # Find the Monday of this week
    monday = today - datetime.timedelta(days=4)  # Go back 4 days from Friday to get Monday
    
    write_log(f"Automatically calculated date range: Monday={monday}, Friday={today}", "CYAN")
    
    # Check if Monday and Friday are in the same month
    if monday.month == today.month:
        # Same month case
        date_range = DateRangeResult(
            start_date=monday,
            end_date=today,
            formatted_date=f"{monday.day}-{today.day} {today.strftime('%B')} {today.year}"
        )
        date_range.year = str(today.year)  # Make sure to set the year property
        return date_range, None
    else:
        # Different month case - create two date ranges
        # Last day of the Monday's month
        last_day_of_month = calendar.monthrange(monday.year, monday.month)[1]
        month_end = datetime.date(monday.year, monday.month, last_day_of_month)
        
        # First day of Friday's month
        month_start = datetime.date(today.year, today.month, 1)
        
        # Create first date range (Monday to end of month)
        first_range = DateRangeResult(
            start_date=monday,
            end_date=month_end,
            formatted_date=f"{monday.day}-{month_end.day} {monday.strftime('%B')} {monday.year}"
        )
        first_range.year = str(monday.year)  # Set the year property
        
        # Create second date range (1st of month to Friday)
        second_range = DateRangeResult(
            start_date=month_start,
            end_date=today,
            formatted_date=f"{month_start.day}-{today.day} {today.strftime('%B')} {today.year}"
        )
        second_range.year = str(today.year)  # Set the year property
        
        return first_range, second_range

def process_with_date_range(date_range):
    """
    Utility function that processes automation with the given date range
    This is a placeholder for where you'd implement the processing logic
    or call the main program's functions with the date range
    
    Args:
        date_range (DateRangeResult): The date range to process
    """
    write_log(f"Processing with date range: {date_range.date_range_formatted}", "GREEN")
    # In a real implementation, you would call the main program's functions here
    # or implement the processing logic
    
    # Example:
    # from main import process_data
    # process_data(date_range)