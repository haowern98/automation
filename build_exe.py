#!/usr/bin/env python3
"""
Enhanced Build script to create EXE file for SharePoint Automation Manual Mode
Keeps automated mode as batch file for task scheduler
"""
import os
import sys
import subprocess
import shutil
import json
from pathlib import Path

def install_pyinstaller():
    """Install PyInstaller if not already installed"""
    try:
        import PyInstaller
        print("✓ PyInstaller is already installed")
        return True
    except ImportError:
        print("📦 Installing PyInstaller...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✓ PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to install PyInstaller: {e}")
            return False

def check_files():
    """Check if required files exist"""
    required_files = [
        "src/main.py",
        "src/config.py", 
        "src/utils/app_controller.py",
        "src/gui/tabbed_app.py",
        "ADProcessor.txt"
    ]
    
    missing_files = []
    for file_path in required_files:
        if not os.path.exists(file_path):
            missing_files.append(file_path)
    
    if missing_files:
        print("❌ Missing required files:")
        for file in missing_files:
            print(f"   - {file}")
        return False
    
    # Check optional files
    if not os.path.exists("settings.json"):
        print("⚠️  settings.json not found - will create default")
        create_default_settings()
    
    if not os.path.exists("data"):
        print("⚠️  data directory not found - creating it")
        os.makedirs("data", exist_ok=True)
    
    return True

def create_default_settings():
    """Create default settings.json if it doesn't exist"""
    user_profile = os.environ.get('USERPROFILE', '')
    downloads_path = os.path.join(user_profile, 'Downloads')
    
    default_settings = {
        "file_paths": {
            "gsn_search_directory": downloads_path,
            "er_search_directory": downloads_path,
            "gsn_file_pattern": "alm_hardware",
            "er_file_pattern": "data"
        },
        "general": {
            "auto_mode_timeout": "30",
            "show_terminal": False
        }
    }
    
    with open("settings.json", "w") as f:
        json.dump(default_settings, f, indent=4)
    print("✓ Created default settings.json")

def create_version_file():
    """Create a version file for the EXE"""
    version_content = '''# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx
VSVersionInfo(
  ffi=FixedFileInfo(
# filevers and prodvers should be always a tuple with four items: (1, 2, 3, 4)
# Set not needed items to zero 0.
filevers=(1,0,0,0),
prodvers=(1,0,0,0),
# Contains a bitmask that specifies the valid bits 'flags'r
mask=0x3f,
# Contains a bitmask that specifies the Boolean attributes of the file.
flags=0x0,
# The operating system for which this file was designed.
# 0x4 - NT and there is no need to change it.
OS=0x4,
# The general type of file.
# 0x1 - the file is an application.
fileType=0x1,
# The function of the file.
# 0x0 - the function is not defined for this fileType
subtype=0x0,
# Creation date and time stamp.
date=(0, 0)
),
  kids=[
StringFileInfo(
  [
  StringTable(
    u'040904B0',
    [StringStruct(u'CompanyName', u'Your Company'),
    StringStruct(u'FileDescription', u'SharePoint Automation Tool - Manual Mode'),
    StringStruct(u'FileVersion', u'1.0.0'),
    StringStruct(u'InternalName', u'SharePointAutomation'),
    StringStruct(u'LegalCopyright', u'Copyright (C) 2025'),
    StringStruct(u'OriginalFilename', u'SharePointAutomation-Manual.exe'),
    StringStruct(u'ProductName', u'SharePoint Automation'),
    StringStruct(u'ProductVersion', u'1.0.0')])
  ]), 
VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)'''
    
    with open('version_info.txt', 'w') as f:
        f.write(version_content)
    
    print("✓ Created version_info.txt")

def build_manual_exe():
    """Build the manual mode EXE"""
    print("🔨 Building SharePoint Automation (Manual Mode) EXE...")
    
    # PyInstaller command for manual mode
    cmd = [
        "pyinstaller",
        "--onefile",  # Single executable file
        "--windowed",  # No console window (GUI only)
        "--name", "SharePointAutomation-Manual",
        "--add-data", "settings.json;.",  # Include settings file
        "--add-data", "data;data",  # Include data directory
        "--add-data", "ADProcessor.txt;.",  # Include AD processor
        "--add-data", "run_ad_processor.bat;.",  # Include AD processor batch file
        "--version-file", "version_info.txt",  # Include version info
        # Hidden imports for all required modules
        "--hidden-import", "PyQt5.QtCore",
        "--hidden-import", "PyQt5.QtGui", 
        "--hidden-import", "PyQt5.QtWidgets",
        "--hidden-import", "PyQt5.QtWebEngineWidgets",
        "--hidden-import", "PyQt5.QtWebEngineCore", 
        "--hidden-import", "PyQt5.QtWebEngineProcess",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "openpyxl",
        "--hidden-import", "pandas",
        "--hidden-import", "psutil",
        "--hidden-import", "ldap3",
        "--hidden-import", "pyad",
        "--hidden-import", "winshell",
        "--hidden-import", "xlwings",
        "--hidden-import", "dateutil",
        "--hidden-import", "dateutil.parser",
        # Project-specific imports
        "--hidden-import", "src",
        "--hidden-import", "src.main",
        "--hidden-import", "src.config",
        "--hidden-import", "src.gui",
        "--hidden-import", "src.gui.date_selector",
        "--hidden-import", "src.gui.settings_dialog",
        "--hidden-import", "src.gui.tabbed_app",
        "--hidden-import", "src.gui.utils",
        "--hidden-import", "src.processors",
        "--hidden-import", "src.processors.gsn_processor",
        "--hidden-import", "src.processors.er_processor", 
        "--hidden-import", "src.processors.er_processor_alt",
        "--hidden-import", "src.processors.ad_processor",
        "--hidden-import", "src.utils",
        "--hidden-import", "src.utils.logger",
        "--hidden-import", "src.utils.excel_functions",
        "--hidden-import", "src.utils.comparison",
        "--hidden-import", "src.utils.app_controller",
        # Exclude unnecessary modules to reduce size
        "--exclude-module", "tkinter",
        "--exclude-module", "matplotlib", 
        "--exclude-module", "IPython",
        "--exclude-module", "jupyter",
        "--exclude-module", "notebook",
        "--exclude-module", "scipy",
        "--exclude-module", "numpy.testing",
        # Entry point with manual mode argument
        "manual_launcher.py"
    ]
    
    # Add icon if it exists
    if os.path.exists("icon.ico"):
        cmd.extend(["--icon", "icon.ico"])
        print("  ✓ Using custom icon: icon.ico")
    
    try:
        subprocess.check_call(cmd)
        print("✓ Manual mode EXE built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build manual EXE: {e}")
        return False

def create_manual_launcher():
    """Create a launcher script specifically for manual mode"""
    launcher_content = '''#!/usr/bin/env python3
"""
Manual Mode Launcher for SharePoint Automation
This launcher ensures the application starts in manual mode
"""
import sys
import os

# Add current directory to path so we can import src modules
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Import and run the main application in manual mode
if __name__ == "__main__":
    # Force manual mode by setting sys.argv
    sys.argv = [sys.argv[0], "--manual"]
    
    # Import and run main
    from src.main import main
    main()
'''
    
    with open("manual_launcher.py", "w") as f:
        f.write(launcher_content)
    
    print("✓ Created manual_launcher.py")

def create_distribution_package():
    """Create a distribution package with both manual EXE and auto batch files"""
    dist_dir = "SharePoint_Automation_Distribution"
    
    # Remove existing distribution directory
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir)
    
    # Create distribution directory
    os.makedirs(dist_dir, exist_ok=True)
    
    # Copy the manual EXE
    if os.path.exists("dist/SharePointAutomation-Manual.exe"):
        shutil.copy2("dist/SharePointAutomation-Manual.exe", 
                    os.path.join(dist_dir, "SharePointAutomation-Manual.exe"))
        print(f"✓ Copied manual EXE to {dist_dir}")
    
    # Copy batch files for automated mode
    batch_files = [
        "run_setup_script.bat",
        "run_sharepoint_automation.bat", 
        "run_ad_processor.bat"
    ]
    
    for batch_file in batch_files:
        if os.path.exists(batch_file):
            shutil.copy2(batch_file, os.path.join(dist_dir, batch_file))
            print(f"✓ Copied {batch_file} to {dist_dir}")
    
    # Copy essential files
    essential_files = [
        "ADProcessor.txt",
        "settings.json"
    ]
    
    for file in essential_files:
        if os.path.exists(file):
            shutil.copy2(file, os.path.join(dist_dir, file))
            print(f"✓ Copied {file} to {dist_dir}")
    
    # Copy data directory
    if os.path.exists("data"):
        shutil.copytree("data", os.path.join(dist_dir, "data"))
        print(f"✓ Copied data directory to {dist_dir}")
    
    # Copy source code for batch file usage
    if os.path.exists("src"):
        shutil.copytree("src", os.path.join(dist_dir, "src"))
        print(f"✓ Copied src directory to {dist_dir}")
    
    # Copy setup.py and requirements if they exist
    if os.path.exists("setup.py"):
        shutil.copy2("setup.py", os.path.join(dist_dir, "setup.py"))
        print(f"✓ Copied setup.py to {dist_dir}")
    
    # Create README for distribution
    create_distribution_readme(dist_dir)
    
    print(f"\n📦 Distribution package created in: {dist_dir}")

def create_distribution_readme(dist_dir):
    """Create a README file for the distribution"""
    readme_content = """# SharePoint Automation Distribution Package

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
"""
    
    readme_path = os.path.join(dist_dir, "README.md")
    with open(readme_path, "w") as f:
        f.write(readme_content)
    
    print(f"✓ Created README.md in {dist_dir}")

def main():
    """Main build function"""
    print("🚀 SharePoint Automation Manual EXE Builder")
    print("=" * 60)
    print("This will create an EXE for manual mode while keeping")
    print("automated mode as batch files for task scheduler.")
    print("=" * 60)
    
    # Check if we're in the project root
    if not os.path.exists("src"):
        print("❌ Please run this script from the project root directory")
        print("   (The directory containing the 'src' folder)")
        return
    
    # Check if we're in a virtual environment
    if hasattr(sys, 'real_prefix') or (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
        print("✓ Running in virtual environment")
    else:
        print("⚠️  Warning: Not running in virtual environment")
        print("   Consider activating your virtual environment first")
        response = input("   Continue anyway? (y/N): ")
        if response.lower() != 'y':
            return
    
    # Check required files
    if not check_files():
        print("❌ Missing required files. Please check your project structure.")
        return
    
    # Install PyInstaller
    if not install_pyinstaller():
        print("❌ Failed to install PyInstaller. Exiting.")
        return
    
    # Clean previous builds
    if os.path.exists('dist'):
        print("🧹 Cleaning previous builds...")
        shutil.rmtree('dist')
    if os.path.exists('build'):
        shutil.rmtree('build')
    
    # Create launcher and version file
    create_manual_launcher()
    create_version_file()
    
    print("\n🔨 Building Manual Mode EXE...")
    
    # Build manual EXE
    manual_success = build_manual_exe()
    
    # Clean up build artifacts
    cleanup_files = ['build', 'manual_launcher.py', 'version_info.txt']
    for item in cleanup_files:
        if os.path.exists(item):
            if os.path.isdir(item):
                shutil.rmtree(item)
            else:
                os.remove(item)
    
    # Clean up spec file
    if os.path.exists('SharePointAutomation-Manual.spec'):
        os.remove('SharePointAutomation-Manual.spec')
    
    print("\n" + "=" * 60)
    if manual_success:
        print("🎉 Manual EXE build completed successfully!")
        
        # Create distribution package
        print("\n📦 Creating distribution package...")
        create_distribution_package()
        
        print(f"\n📁 Generated files:")
        if os.path.exists('dist/SharePointAutomation-Manual.exe'):
            size_manual = os.path.getsize('dist/SharePointAutomation-Manual.exe') / (1024*1024)
            print(f"   - dist/SharePointAutomation-Manual.exe ({size_manual:.1f} MB)")
        
        print(f"\n📋 Usage:")
        print(f"   Manual Mode:")
        print(f"   - Double-click SharePointAutomation-Manual.exe")
        print(f"   - GUI interface with settings and date selection")
        print(f"   ")
        print(f"   Automated Mode (Task Scheduler):")
        print(f"   - Use run_sharepoint_automation.bat")
        print(f"   - Requires Python environment setup")
        print(f"   - Configure via Task Scheduler")
        
        print(f"\n📦 Distribution:")
        print(f"   - Complete package available in: SharePoint_Automation_Distribution/")
        print(f"   - Copy entire folder to target machines")
        print(f"   - Manual EXE requires no additional setup")
        print(f"   - Automated mode requires Python (use setup script)")
        
    else:
        print("❌ Build failed. Check the error messages above.")
        print(f"\n🔧 Troubleshooting tips:")
        print(f"   - Make sure all dependencies are installed")
        print(f"   - Check that you're in the project root directory")
        print(f"   - Try running: pip install -r requirements.txt")
        print(f"   - Ensure PyQt5 and win32com are properly installed")

if __name__ == "__main__":
    main()