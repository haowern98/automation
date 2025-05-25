@echo off
echo Starting SharePoint Automation...
echo.

rem Activate Python virtual environment
call venv-py312\Scripts\activate

rem Run the Python program
python -m src.main

echo.
echo SharePoint Automation completed.
pause