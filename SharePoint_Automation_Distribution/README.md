# SharePoint Automation Distribution Package

This package contains both manual and automated modes for SharePoint Automation.

## Files Included

### Manual Mode (GUI)
- **SharePointAutomation-Manual.exe** - Double-click to run in manual mode with GUI

### Automated Mode (Task Scheduler)
- **run_setup_script.bat** - Run this first to set up Python environment
- **run_sharepoint_automation.bat** - Use this in Task Scheduler for automated runs
- **run_ad_processor.bat** - Helper script for Active Directory processing
- **src/** - Source code directory (required for batch file execution)

### Configuration & Data
- **settings.json** - Configuration file (automatically created/updated)
- **data/** - Data storage directory
- **ADProcessor.txt** - PowerShell script for AD operations

## Setup Instructions

### For Manual Mode:
1. Simply double-click `SharePointAutomation-Manual.exe`
2. No additional setup required - all dependencies are included

### For Automated Mode (Task Scheduler):
1. First run `run_setup_script.bat` to set up Python environment
2. Configure Task Scheduler to run `run_sharepoint_automation.bat`
3. Ensure the task runs from this directory

## Task Scheduler Configuration

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., weekly on Friday)
4. Action: Start a program
5. Program: `run_sharepoint_automation.bat`
6. Start in: Path to this directory

## Settings Configuration

Both modes use the same `settings.json` file for configuration:
- File search directories
- File name patterns  
- Auto mode timeout
- Other preferences

Use the manual mode to easily configure settings through the GUI.

## Troubleshooting

### Manual Mode:
- If EXE doesn't start, check antivirus software
- Settings are stored in the same directory as the EXE

### Automated Mode:
- Ensure Python virtual environment is set up (run setup script)
- Check that all source files are present
- Verify Task Scheduler is running from correct directory
- Check Windows Event Viewer for error logs

## Support

For issues or questions, check the application logs and ensure:
1. All required files are present
2. Settings are properly configured
3. Network access is available for SharePoint sync
4. Excel and PowerShell are available on the system
