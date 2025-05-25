@echo off
setlocal enabledelayedexpansion

echo ===============================================
echo SharePoint Automation - Setup Script (Manual Python)
echo ===============================================
echo.

set VENV_NAME=venv-py312

:CHECK_PYTHON
echo Checking for Python 3.12...
py -3.12 --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo Python 3.12 is installed.
    set PYTHON_CMD=py -3.12
    goto CREATE_VENV
)

python --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    for /f "tokens=2" %%i in ('python --version 2^>^&1') do set INSTALLED_VERSION=%%i
    echo Found Python !INSTALLED_VERSION! installed.
    set PYTHON_CMD=python
    goto CREATE_VENV
) else (
    echo Python is not installed or not in PATH.
    echo Please install Python 3.12 from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    goto END
)

:CREATE_VENV
echo.
echo Creating virtual environment (%VENV_NAME%)...
if exist %VENV_NAME% (
    echo Virtual environment already exists.
    choice /M "Do you want to recreate it"
    if !ERRORLEVEL! EQU 2 goto ACTIVATE_VENV
    echo Removing existing virtual environment...
    rmdir /S /Q %VENV_NAME%
)

%PYTHON_CMD% -m venv %VENV_NAME%
if %ERRORLEVEL% NEQ 0 (
    echo Failed to create virtual environment.
    echo Installing virtualenv and trying again...
    %PYTHON_CMD% -m pip install virtualenv
    %PYTHON_CMD% -m virtualenv %VENV_NAME%
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to create virtual environment.
        echo Please check your Python installation.
        goto END
    )
)

:ACTIVATE_VENV
echo.
echo Activating virtual environment...
call %VENV_NAME%\Scripts\activate
if %ERRORLEVEL% NEQ 0 (
    echo Failed to activate virtual environment.
    goto END
)

:INSTALL_DEPENDENCIES
echo.
echo Installing/upgrading pip...
python -m pip install --upgrade pip

echo.
echo Installing setuptools (required for setup.py)...
python -m pip install setuptools wheel

echo.
echo Installing dependencies using pip...
python -m pip install -e .
if %ERRORLEVEL% NEQ 0 (
    echo Failed to install dependencies.
    echo Please check the setup.py file and your internet connection.
    goto END
)

echo.
echo ===============================================
echo Setup completed successfully!
echo.
echo The virtual environment '%VENV_NAME%' has been created and all dependencies installed.
echo.
echo You can now run the application using:
echo   run_sharepoint_automation.bat - For normal mode
echo   run_manual.bat - For manual mode
echo ===============================================

:END
endlocal
pause