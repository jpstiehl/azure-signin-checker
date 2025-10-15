@echo off
REM Azure AD Sign-in Checker Launcher
REM This batch file launches the PowerShell GUI script with proper execution policy

echo Starting Azure AD Sign-in Checker...
echo.

REM Change to script directory
cd /d "%~dp0"

REM Run PowerShell script with bypass execution policy
powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "Check-UserSignIns-GUI.ps1"

REM Pause if there were any errors
if errorlevel 1 (
    echo.
    echo Script encountered an error. Press any key to close...
    pause >nul
)