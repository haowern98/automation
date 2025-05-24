#!/usr/bin/env python3
"""
Build script to create EXE files for SharePoint Automation
"""
import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_pyinstaller():
    """Install PyInstaller if not already installed"""
    try:
        import PyInstaller
        print("PyInstaller is already installed")
        return True
    except ImportError:
        print("Installing PyInstaller...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("PyInstaller installed successfully")
            return True
        except subprocess.CalledProcessError as e:
            print(f"Failed to install PyInstaller: {e}")
            return False

def check_files():
    """Check if required files exist"""
    required_files = [
        "src/main.py",
        "src/config.py",
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
    import json
    default_settings = {
        "file_paths": {
            "gsn_search_directory": "",
            "er_search_directory": "",
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

def build_manual_exe():
    """Build the manual mode EXE"""
    print("Building SharePoint Automation (Manual Mode) EXE...")
    
    # PyInstaller command for manual mode
    cmd = [
        "pyinstaller",
        "--onefile",  # Single executable file
        "--windowed",  # No console window (GUI only)
        "--name", "SharePointAutomation-Manual",
        "--add-data", "settings.json;.",  # Include settings file
        "--add-data", "data;data",  # Include data directory
        "--add-data", "ADProcessor.txt;.",  # Include AD processor
        "--hidden-import", "PyQt5.QtCore",
        "--hidden-import", "PyQt5.QtGui", 
        "--hidden-import", "PyQt5.QtWidgets",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "openpyxl",
        "--hidden-import", "pandas",
        "--hidden-import", "psutil",
        "--hidden-import", "src.gui.date_selector",
        "--hidden-import", "src.gui.settings_dialog",
        "--hidden-import", "src.gui.tabbed_app",
        "--hidden-import", "src.processors.gsn_processor",
        "--hidden-import", "src.processors.er_processor", 
        "--hidden-import", "src.processors.ad_processor",
        "--hidden-import", "src.utils.logger",
        "--hidden-import", "src.utils.excel_functions",
        "--hidden-import", "src.utils.comparison",
        "--hidden-import", "src.utils.app_controller",
        "--exclude-module", "tkinter",  # Reduce size
        "--exclude-module", "matplotlib",  # Reduce size
        "--exclude-module", "IPython",  # Reduce size
        "--exclude-module", "jupyter",  # Reduce size
        "src/main.py"
    ]
    
    # Add icon if it exists
    if os.path.exists("icon.ico"):
        cmd.extend(["--icon", "icon.ico"])
        print("  Using custom icon: icon.ico")
    
    try:
        subprocess.check_call(cmd)
        print("✓ Manual mode EXE built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build manual EXE: {e}")
        return False

def build_auto_exe():
    """Build the auto mode EXE"""
    print("Building SharePoint Automation (Auto Mode) EXE...")
    
    # PyInstaller command for auto mode
    cmd = [
        "pyinstaller",
        "--onefile",  # Single executable file
        "--console",  # Show console for auto mode
        "--name", "SharePointAutomation-Auto",
        "--add-data", "settings.json;.",  # Include settings file
        "--add-data", "data;data",  # Include data directory
        "--add-data", "ADProcessor.txt;.",  # Include AD processor
        "--hidden-import", "PyQt5.QtCore",
        "--hidden-import", "PyQt5.QtGui", 
        "--hidden-import", "PyQt5.QtWidgets",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "openpyxl",
        "--hidden-import", "pandas",
        "--hidden-import", "psutil",
        "--hidden-import", "src.gui.date_selector",
        "--hidden-import", "src.gui.settings_dialog",
        "--hidden-import", "src.gui.tabbed_app",
        "--hidden-import", "src.processors.gsn_processor",
        "--hidden-import", "src.processors.er_processor", 
        "--hidden-import", "src.processors.ad_processor",
        "--hidden-import", "src.utils.logger",
        "--hidden-import", "src.utils.excel_functions",
        "--hidden-import", "src.utils.comparison",
        "--hidden-import", "src.utils.app_controller",
        "--exclude-module", "tkinter",  # Reduce size
        "--exclude-module", "matplotlib",  # Reduce size
        "--exclude-module", "IPython",  # Reduce size
        "--exclude-module", "jupyter",  # Reduce size
        "src/main.py"
    ]
    
    # Add icon if it exists
    if os.path.exists("icon.ico"):
        cmd.extend(["--icon", "icon.ico"])
        print("  Using custom icon: icon.ico")
    
    try:
        subprocess.check_call(cmd)
        print("✓ Auto mode EXE built successfully!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build auto EXE: {e}")
        return False

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
    StringStruct(u'FileDescription', u'SharePoint Automation Tool'),
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

def main():
    """Main build function"""
    print("🚀 SharePoint Automation EXE Builder")
    print("=" * 50)
    
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
    
    # Build EXEs
    print("\n🔨 Building EXE files...")
    
    manual_success = build_manual_exe()
    print()  # Add spacing
    auto_success = build_auto_exe()
    
    # Clean up build artifacts
    if os.path.exists('build'):
        shutil.rmtree('build')
    
    # Clean up spec files
    for spec_file in ['SharePointAutomation-Manual.spec', 'SharePointAutomation-Auto.spec']:
        if os.path.exists(spec_file):
            os.remove(spec_file)
    
    print("\n" + "=" * 50)
    if manual_success and auto_success:
        print("🎉 Build completed successfully!")
        print("\n📁 Generated files:")
        if os.path.exists('dist/SharePointAutomation-Manual.exe'):
            size_manual = os.path.getsize('dist/SharePointAutomation-Manual.exe') / (1024*1024)
            print(f"   - dist/SharePointAutomation-Manual.exe ({size_manual:.1f} MB) - GUI mode")
        if os.path.exists('dist/SharePointAutomation-Auto.exe'):
            size_auto = os.path.getsize('dist/SharePointAutomation-Auto.exe') / (1024*1024)
            print(f"   - dist/SharePointAutomation-Auto.exe ({size_auto:.1f} MB) - Console mode")
        
        print("\n📋 Usage:")
        print("   - Double-click SharePointAutomation-Manual.exe for manual mode")
        print("   - Run SharePointAutomation-Auto.exe from command line for auto mode")
        print("   - Or schedule SharePointAutomation-Auto.exe in Windows Task Scheduler")
        
        print("\n📦 Distribution:")
        print("   - Copy the EXE files to target machines")
        print("   - No Python installation required on target machines")
        print("   - Settings will be created automatically on first run")
        
    else:
        print("❌ Build failed. Check the error messages above.")
        print("\n🔧 Troubleshooting tips:")
        print("   - Make sure all dependencies are installed")
        print("   - Check that you're in the project root directory")
        print("   - Try running: pip install -r requirements.txt")

if __name__ == "__main__":
    main()