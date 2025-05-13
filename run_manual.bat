@echo off
echo Starting SharePoint Automation in Manual Mode...
echo Date: %date% Time: %time%
echo.

rem Set working directory to the script's location
cd /d "%~dp0"

rem Activate Python virtual environment
if exist venv-py312\Scripts\activate (
    call venv-py312\Scripts\activate
) else (
    echo WARNING: Virtual environment not found, using system Python.
)

rem Run the Python program in manual mode
python main.py --manual

echo.
echo SharePoint Automation completed.
pause