@echo off
echo Starting Adobe Creative Cloud Removal Tool...
echo.
echo This will completely remove Adobe Creative Cloud from your system.
echo Please ensure you have backed up any important work.
echo.
pause

REM Change to the script directory
cd /d "%~dp0"

REM Run PowerShell script with execution policy bypass
powershell.exe -ExecutionPolicy Bypass -File "Remove-Adobe.ps1"

echo.
echo Script completed. Press any key to exit.
pause >nul