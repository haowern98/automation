@echo off
echo Running AD Processor...
echo.

set TEMP_PS_FILE=%TEMP%\ad_processor_temp.ps1
set AD_PROCESSOR_TXT=%~dp0ADProcessor.txt

REM Check if text file exists
if not exist "%AD_PROCESSOR_TXT%" (
    echo Error: Cannot find AD processor file at:
    echo %AD_PROCESSOR_TXT%
    exit /b 1
)

REM Create a temporary PS1 file with the Write-Log function and ADProcessor code
(
echo # Function to write log messages
echo function Write-Log {
echo     param (
echo         [string]$Message,
echo         [string]$Color = "White"
echo     ^)
echo     $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
echo     Write-Host "[$timestamp] $Message" -ForegroundColor $Color
echo }
echo.
type "%AD_PROCESSOR_TXT%"
echo.
echo # Process AD data and save to JSON
echo $adComputers = Process-ADData -OutputFilePath "ad_results.json"
echo Write-Log "AD processing completed with $($adComputers.Count) computers" -Color Green
) > "%TEMP_PS_FILE%"

REM Execute the temporary PS1 file
powershell -ExecutionPolicy Bypass -File "%TEMP_PS_FILE%"

REM Check if the JSON file was created
if exist "ad_results.json" (
    echo.
    echo AD results file created: ad_results.json
) else (
    echo.
    echo WARNING: AD results file not created!
    exit /b 1
)

REM Clean up
del "%TEMP_PS_FILE%" 2>nul

echo.
echo AD Processing completed.