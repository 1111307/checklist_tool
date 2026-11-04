@echo off
title Config Check Tool - Windows XP/7

:: Check admin privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] Not running as Administrator. Some checks may be inaccurate.
    echo Please right-click run.bat and select "Run as administrator".
    echo.
    pause
)

:: Change to script directory
cd /d "%~dp0"

:: Create output directory
if not exist output mkdir output

echo Starting configuration check...
echo.

:: Run VBScript via cscript (console mode)
cscript //NoLogo check_xp7.vbs

echo.
echo Press any key to exit...
pause >nul
