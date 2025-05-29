"""
Terminal Control Utilities

This module provides functions to show/hide the terminal window on Windows.
"""
import sys
import os

def hide_terminal():
    """Hide the terminal/console window (Windows only)"""
    if sys.platform == "win32":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            SW_HIDE = 0
            
            hWnd = kernel32.GetConsoleWindow()
            if hWnd:
                user32.ShowWindow(hWnd, SW_HIDE)
                return True
        except Exception as e:
            print(f"Could not hide terminal: {str(e)}")
            return False
    return False

def show_terminal():
    """Show the terminal/console window (Windows only)"""
    if sys.platform == "win32":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            SW_SHOW = 5
            
            hWnd = kernel32.GetConsoleWindow()
            if hWnd:
                user32.ShowWindow(hWnd, SW_SHOW)
                return True
        except Exception as e:
            print(f"Could not show terminal: {str(e)}")
            return False
    return False

def apply_terminal_setting():
    """Apply terminal visibility based on current settings"""
    try:
        from src.gui.settings_dialog import get_settings
        settings = get_settings()
        show_terminal_setting = settings.get('general', 'show_terminal', False)
        
        if show_terminal_setting:
            return show_terminal()
        else:
            return hide_terminal()
    except Exception as e:
        print(f"Could not apply terminal setting: {str(e)}")
        return False