#!/usr/bin/env python3
"""
Enhanced Build script to create SINGLE EXE file for SharePoint Automation
Supports both manual and automatic modes with smart detection
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
        "ADProcessor.txt",
        "single_launcher.py"  # New requirement
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
            "er_file_pattern": "data",
            "weekly_report_file_path": os.path.join(
                user_profile,
                'DPDHL',
                'SM Team - SG - AD EDS, MFA, GSN VS AD, GSN VS ER Weekly Report',
                'Weekly Report 2025.xlsx'
            )
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
    StringStruct(u'FileDescription', u'SharePoint Automation Tool - Universal Mode'),
    StringStruct(u'FileVersion', u'1.0.0'),
    StringStruct(u'InternalName', u'SharePointAutomation'),
    StringStruct(u'LegalCopyright', u'Copyright (C) 2025'),
    StringStruct(u'OriginalFilename', u'SharePointAutomation.exe'),
    StringStruct(u'ProductName', u'SharePoint Automation'),
    StringStruct(u'ProductVersion', u'1.0.0')])
  ]), 
VarFileInfo([VarStruct(u'Translation', [1033, 1200])])
  ]
)'''
    
    with open('version_info.txt', 'w') as f:
        f.write(version_content)
    
    print("✓ Created version_info.txt")

def build_single_exe():
    """Build the single universal EXE"""
    print("🔨 Building SharePoint Automation (Universal Mode) EXE...")
    
    # PyInstaller command for single universal EXE
    cmd = [
        "pyinstaller",
        "--onefile",  # Single executable file
        "--windowed",  # No console window by default (smart detection will handle this)
        "--name", "SharePointAutomation",  # Simple name
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
        "--hidden-import", "src.processors.weekly_report_extractor",
        "--hidden-import", "src.processors.gsn_vs_ad_extractor",
        "--hidden-import", "src.processors.gsn_vs_er_extractor",
        "--hidden-import", "src.processors.er_extractor",
        "--hidden-import", "src.utils",
        "--hidden-import", "src.utils.logger",
        "--hidden-import", "src.utils.excel_functions",
        "--hidden-import", "src.utils.comparison",
        "--hidden-import", "src.utils.app_controller",
        "--hidden-import", "src.gui.loading_screen",
        # Exclude unnecessary modules to reduce size
        "--exclude-module", "tkinter",
        "--exclude-module", "matplotlib", 
        "--exclude-module", "IPython",
        "--exclude-module", "jupyter",
        "--exclude-module", "notebook",
        "--exclude-module", "scipy",
        "--exclude-module", "numpy.testing",
        # Entry point - the single launcher
        "single_launcher.py"
    ]
    
    # Add icon if it exists
    if os.path.exists("icon.ico"):
        cmd.extend(["--icon", "icon.ico"])
        print("  ✓ Using custom icon: icon.ico")
    
    try:
        subprocess.check_call(cmd)
        print("✓ Universal EXE built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build universal EXE: {e}")
        return False

def create_distribution_package():
    """Create a distribution package with the single EXE and batch files"""
    dist_dir = "SharePoint_Automation_Distribution"
    
    # Remove existing distribution directory
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir)
    
    # Create distribution directory
    os.makedirs(dist_dir, exist_ok=True)
    
    # Copy the universal EXE
    if os.path.exists("dist/SharePointAutomation.exe"):
        shutil.copy2("dist/SharePointAutomation.exe", 
                    os.path.join(dist_dir, "SharePointAutomation.exe"))
        print(f"✓ Copied universal EXE to {dist_dir}")
    
    # Copy batch files for automated mode (still useful for task scheduler)
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
    
    # Copy source code for batch file usage (fallback)
    if os.path.exists("src"):
        shutil.copytree("src", os.path.join(dist_dir, "src"))
        print(f"✓ Copied src directory to {dist_dir}")
    
    # Copy setup.py if it exists
    if os.path.exists("setup.py"):
        shutil.copy2("setup.py", os.path.join(dist_dir, "setup.py"))
        print(f"✓ Copied setup.py to {dist_dir}")
    
    # Create README for distribution
    create_distribution_readme(dist_dir)
    
    print(f"\n📦 Distribution package created in: {dist_dir}")

def create_distribution_readme(dist_dir):
    """Create a README file for the distribution"""
    readme_content = """# SharePoint Automation Distribution Package

This package contains the **Universal SharePoint Automation EXE** that automatically detects how to run.

## Files Included

### Universal EXE (Recommended)
- **SharePointAutomation.exe** - Universal EXE that automatically detects execution mode:
  - **Double-click**: Runs in manual mode with GUI
  - **Task Scheduler**: Runs in auto mode without GUI
  - **Command line**: Supports `--manual` and `--auto` flags

### Legacy Batch Files (Optional)
- **run_setup_script.bat** - Set up Python environment (only needed for batch mode)
- **run_sharepoint_automation.bat** - Legacy batch execution
- **run_ad_processor.bat** - Helper script for AD operations
- **src/** - Source code directory (for batch file usage)

### Configuration & Data
- **settings.json** - Configuration file (automatically created/updated)
- **data/** - Data storage directory
- **ADProcessor.txt** - PowerShell script for AD operations

## Usage Instructions

### Method 1: Universal EXE (Recommended)

**For Manual Use:**
1. Double-click `SharePointAutomation.exe`
2. GUI will appear with date selection and settings
3. No additional setup required

**For Task Scheduler:**
1. Create scheduled task
2. Set program: `SharePointAutomation.exe`
3. Set arguments: `--auto` (optional, will auto-detect)
4. Set working directory to this folder
5. EXE will run automatically without GUI

**For Command Line:**
```cmd
SharePointAutomation.exe                    # Auto-detects mode
SharePointAutomation.exe --manual           # Force manual mode
SharePointAutomation.exe --auto             # Force auto mode
SharePointAutomation.exe --debug            # Enable debug logging
```

### Method 2: Legacy Batch Files

If you prefer the old method:
1. Run `run_setup_script.bat` first to set up Python
2. Use `run_sharepoint_automation.bat` for Task Scheduler

## Smart Mode Detection

The EXE automatically detects:
- **Manual Mode**: Double-click, GUI environment, interactive use
- **Auto Mode**: Task Scheduler, command line, service execution, CI/CD

## Task Scheduler Setup

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., weekly on Friday)
4. Action: Start a program
5. Program: `SharePointAutomation.exe`
6. Arguments: `--auto` (optional)
7. Start in: Path to this directory

## Settings Configuration

Use manual mode (double-click EXE) to easily configure:
- File search directories
- File name patterns  
- Auto mode timeout
- Other preferences

Settings are saved in `settings.json` and used by both modes.

## Troubleshooting

### Universal EXE Issues:
- If mode detection fails, use explicit flags: `--manual` or `--auto`
- Check antivirus software if EXE doesn't start
- Ensure all files are in the same directory

### Task Scheduler Issues:
- Verify working directory is set correctly
- Check Windows Event Viewer for error logs
- Ensure network access for SharePoint sync
- Verify Excel and PowerShell are available

### Mode Detection Issues:
- Use `--debug` flag to see detection reasoning
- Check console output for detection details
- Use explicit `--manual` or `--auto` flags to override

## Support

The Universal EXE provides the best experience:
- ✅ No Python setup required
- ✅ Automatic mode detection
- ✅ Works for both manual and scheduled use
- ✅ Single file distribution
- ✅ Backward compatible with all existing functionality
"""
    
    readme_path = os.path.join(dist_dir, "README.md")
    with open(readme_path, "w") as f:
        f.write(readme_content)
    
    print(f"✓ Created README.md in {dist_dir}")

def main():
    """Main build function"""
    print("🚀 SharePoint Automation Universal EXE Builder")
    print("=" * 60)
    print("This will create ONE EXE that works for both manual and")
    print("automated modes with smart detection.")
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
    
    # Create version file
    create_version_file()
    
    print("\n🔨 Building Universal EXE...")
    
    # Build single EXE
    success = build_single_exe()
    
    # Clean up build artifacts
    cleanup_files = ['build', 'version_info.txt']
    for item in cleanup_files:
        if os.path.exists(item):
            if os.path.isdir(item):
                shutil.rmtree(item)
            else:
                os.remove(item)
    
    # Clean up spec file
    if os.path.exists('SharePointAutomation.spec'):
        os.remove('SharePointAutomation.spec')
    
    print("\n" + "=" * 60)
    if success:
        print("🎉 Universal EXE build completed successfully!")
        
        # Create distribution package
        print("\n📦 Creating distribution package...")
        create_distribution_package()
        
        print(f"\n📁 Generated files:")
        if os.path.exists('dist/SharePointAutomation.exe'):
            size = os.path.getsize('dist/SharePointAutomation.exe') / (1024*1024)
            print(f"   - dist/SharePointAutomation.exe ({size:.1f} MB)")
        
        print(f"\n📋 Usage:")
        print(f"   Manual Mode (GUI):")
        print(f"   - Double-click SharePointAutomation.exe")
        print(f"   - Or run: SharePointAutomation.exe --manual")
        print(f"   ")
        print(f"   Auto Mode (Task Scheduler):")
        print(f"   - SharePointAutomation.exe --auto")
        print(f"   - Or just SharePointAutomation.exe (auto-detects)")
        print(f"   ")
        print(f"   Smart Detection:")
        print(f"   - Double-click → Manual mode")
        print(f"   - Task Scheduler → Auto mode")
        print(f"   - Command line → Auto mode")
        
        print(f"\n📦 Distribution:")
        print(f"   - Complete package: SharePoint_Automation_Distribution/")
        print(f"   - Copy entire folder to target machines")
        print(f"   - Universal EXE requires no additional setup")
        print(f"   - Works for both manual and automated use")
        
    else:
        print("❌ Build failed. Check the error messages above.")
        print(f"\n🔧 Troubleshooting tips:")
        print(f"   - Make sure all dependencies are installed")
        print(f"   - Check that you're in the project root directory")
        print(f"   - Ensure single_launcher.py exists")
        print(f"   - Try running: pip install -r requirements.txt")

if __name__ == "__main__":
    main()