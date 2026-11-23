@echo off
chcp 65001 >nul
title DaData Parser - Administrations

echo ===============================================
echo  DaData Parser - Administrations (Windows)
echo  Branch arena/363bbf0c-gov
echo ===============================================
echo.

:: Check Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found! Install Python 3.10+ from https://python.org
    echo Don't forget to check "Add Python to PATH" during install
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
pip install openpyxl requests --quiet
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies
    echo Try: pip install openpyxl requests --user
    pause
    exit /b 1
)

echo [2/3] Starting GUI...
echo If window didn't open, run manually: python DaData_GUI.py
echo.

:: Start GUI
python DaData_GUI.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Program exited with error code %errorlevel%
    pause
)

echo.
echo [3/3] Done. Results in program folder.
pause
