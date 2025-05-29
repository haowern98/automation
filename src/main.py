#!/usr/bin/env python3
"""
SharePoint Automation - Main Script with Smart Mode Detection
"""
import os
import sys
import argparse
import psutil
from datetime import datetime

# Add the parent directory to the Python path so we can import src modules
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from PyQt5.QtWidgets import QApplication
from src.utils.logger import write_log
from src.config import DATA_DIR

print("DEBUG: All imports successful, about to define smart mode detection")

def detect_execution_mode():
    """
    Automatically detect whether to run in manual or auto mode based on execution context
    
    Detection Rules:
    1. If --manual flag: Force manual mode
    2. If --auto flag: Force auto mode
    3. If run by double-clicking EXE: Manual mode (parent is explorer.exe)
    4. If run from GUI environment interactively: Manual mode
    5. If run from command line/task scheduler: Auto mode
    6. If no console attached: Manual mode (likely GUI launch)
    
    Returns:
        tuple: (is_manual_mode: bool, detection_reason: str)
    """
    detection_reason = ""
    
    try:
        # Check command line arguments first (highest priority)
        if len(sys.argv) > 1:
            for arg in sys.argv[1:]:
                if arg == '--manual':
                    return True, "Command line flag --manual"
                elif arg == '--auto':
                    return False, "Command line flag --auto"
        
        # Method 1: Check parent process
        try:
            current_process = psutil.Process()
            parent_process = current_process.parent()
            parent_name = parent_process.name().lower()
            
            # If parent is Windows Explorer, user double-clicked the EXE
            if 'explorer' in parent_name:
                return True, f"Double-clicked from Explorer (parent: {parent_name})"
            
            # If parent is cmd.exe, powershell, or task scheduler
            command_line_parents = ['cmd.exe', 'powershell.exe', 'conhost.exe', 'svchost.exe', 'taskeng.exe', 'taskhostw.exe']
            if any(parent in parent_name for parent in command_line_parents):
                return False, f"Command line/Task Scheduler execution (parent: {parent_name})"
                
        except Exception as e:
            detection_reason += f"Parent process check failed: {str(e)}; "
        
        # Method 2: Check if console is attached
        try:
            # Try to get console window handle
            import ctypes
            from ctypes import wintypes
            
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            
            # Get console window
            console_window = kernel32.GetConsoleWindow()
            
            if console_window == 0:
                # No console window - likely GUI launch
                return True, "No console window attached (GUI launch)"
            
            # Check if console window is visible
            is_visible = user32.IsWindowVisible(console_window)
            if not is_visible:
                return False, "Hidden console window (automated execution)"
                
        except Exception as e:
            detection_reason += f"Console check failed: {str(e)}; "
        
        # Method 3: Check standard input/output
        try:
            # If stdin is not a TTY, likely automated
            if not sys.stdin.isatty():
                return False, "stdin is not a TTY (automated execution)"
            
            # If stdout is not a TTY, likely redirected/automated
            if not sys.stdout.isatty():
                return False, "stdout is not a TTY (redirected/automated)"
                
        except Exception as e:
            detection_reason += f"TTY check failed: {str(e)}; "
        
        # Method 4: Check environment variables for automation
        automation_indicators = [
            'JENKINS_URL', 'BUILD_NUMBER',  # Jenkins
            'GITHUB_ACTIONS', 'CI',         # GitHub Actions
            'TF_BUILD',                     # Azure DevOps
            'SCHEDULED_TASK_NAME',          # Windows Task Scheduler
            'SYSTEM'                        # System account
        ]
        
        for indicator in automation_indicators:
            if indicator in os.environ:
                return False, f"Automation environment detected ({indicator})"
        
        # Method 5: Check current working directory
        try:
            cwd = os.getcwd().lower()
            # If running from system directories, likely automated
            system_dirs = ['system32', 'windows', 'program files']
            if any(sys_dir in cwd for sys_dir in system_dirs):
                return False, f"Running from system directory: {cwd}"
        except Exception as e:
            detection_reason += f"CWD check failed: {str(e)}; "
        
        # Method 6: Check session type (interactive vs service)
        try:
            import ctypes
            from ctypes import wintypes
            
            # Get current session ID
            kernel32 = ctypes.windll.kernel32
            session_id = kernel32.GetCurrentProcessId()
            
            # Try to determine if running in an interactive session
            user32 = ctypes.windll.user32
            desktop = user32.GetThreadDesktop(kernel32.GetCurrentThreadId())
            
            if desktop == 0:
                return False, "No desktop session (service/automated execution)"
        except Exception as e:
            detection_reason += f"Session check failed: {str(e)}; "
        
        # Default fallback: If we can't determine, assume manual mode for safety
        # This ensures the user gets a GUI if detection fails
        return True, f"Default fallback to manual mode (detection uncertain: {detection_reason.rstrip('; ')})"
        
    except Exception as e:
        # If all detection fails, default to manual mode for user safety
        return True, f"Detection failed, defaulting to manual: {str(e)}"

def main():
    """Main function to run the SharePoint automation with smart mode detection"""
    print("DEBUG: Entered main function")
    
    # Record the start time
    start_time = datetime.now()
    write_log("Starting SharePoint Automation Script with Smart Mode Detection", "YELLOW")

    # Parse command line arguments
    parser = argparse.ArgumentParser(description='SharePoint Automation Script')
    parser.add_argument('--manual', action='store_true', 
                        help='Force manual mode with GUI interface')
    parser.add_argument('--auto', action='store_true',
                        help='Force automatic mode without GUI prompts')
    parser.add_argument('--debug', action='store_true',
                        help='Run in debug mode with additional logging')
    args = parser.parse_args()
    
    # Detect execution mode
    detected_manual_mode, detection_reason = detect_execution_mode()
    
    # Override detection if explicit flags are provided
    if args.manual and args.auto:
        write_log("Both --manual and --auto flags provided. Using --manual.", "YELLOW")
        manual_mode = True
        final_reason = "Command line override: --manual flag"
    elif args.manual:
        manual_mode = True
        final_reason = "Command line override: --manual flag"
    elif args.auto:
        manual_mode = False
        final_reason = "Command line override: --auto flag"
    else:
        manual_mode = detected_manual_mode
        final_reason = f"Auto-detected: {detection_reason}"
    
    # Log the mode decision
    mode_text = "MANUAL" if manual_mode else "AUTO"
    write_log(f"Execution Mode: {mode_text}", "GREEN")
    write_log(f"Detection Reason: {final_reason}", "CYAN")
    
    print(f"DEBUG: Mode detected - manual: {manual_mode}, debug: {args.debug}")
    print(f"DEBUG: Detection reason: {final_reason}")
    
    # Ensure data directory exists
    os.makedirs(DATA_DIR, exist_ok=True)
    
    # Create QApplication for GUI components (even in auto mode for dialogs)
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    
    # In auto mode, hide the application from taskbar unless debug mode
    if not manual_mode and not args.debug:
        app.setQuitOnLastWindowClosed(True)
    
    print("DEBUG: About to call automation function")
    
    try:
        from src.utils.app_controller import run_sharepoint_automation_with_loading
        success = run_sharepoint_automation_with_loading(manual_mode, args.debug)
        
        if success:
            write_log("SharePoint automation completed successfully", "GREEN")
        else:
            write_log("SharePoint automation was cancelled or failed", "YELLOW")
        
        # In auto mode, provide final status
        if not manual_mode:
            status = "SUCCESS" if success else "FAILED"
            write_log(f"AUTOMATION STATUS: {status}", "GREEN" if success else "RED")
        
    except Exception as e:
        print(f"DEBUG: Exception in main: {str(e)}")
        write_log(f"An error occurred: {str(e)}", "RED")
        import traceback
        traceback.print_exc()
        
        # In auto mode, ensure we exit with proper code
        if not manual_mode:
            write_log("AUTOMATION STATUS: ERROR", "RED")
            sys.exit(1)
        
    finally:
        # Record the end time
        end_time = datetime.now()
        duration = end_time - start_time
        write_log(f"Total script execution time: {duration.total_seconds()} seconds", "CYAN")
        
        # In auto mode, don't wait for user input
        if not manual_mode:
            write_log("Auto mode completed, exiting...", "YELLOW")
        else:
            # In manual mode, app will handle its own event loop through GUI

            print("DEBUG: main() function defined")

if __name__ == "__main__":
    print("DEBUG: About to call main()")
    main()
    print("DEBUG: main() completed")