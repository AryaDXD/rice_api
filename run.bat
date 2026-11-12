@echo off
setlocal enabledelayedexpansion

title Student Management API Tester ^& Bulk Upload

:: Change directory to this script's location
cd /d "%~dp0"

echo ========================================================
echo    Student Management API Tester ^& Bulk Upload
echo ========================================================
echo.

:: Detect Python executable
set "PYTHON_CMD="

:: 1. Check local virtual environment (.venv or venv)
if exist ".venv\Scripts\python.exe" (
    set "PYTHON_CMD=.venv\Scripts\python.exe"
) else if exist "venv\Scripts\python.exe" (
    set "PYTHON_CMD=venv\Scripts\python.exe"
)

:: 2. Check Python Launcher for Windows (py)
if "%PYTHON_CMD%"=="" (
    where py >nul 2>&1
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=py -3"
    )
)

:: 3. Check python command in PATH
if "%PYTHON_CMD%"=="" (
    where python >nul 2>&1
    if !errorlevel! equ 0 (
        set "PYTHON_CMD=python"
    )
)

:: If Python is not installed or not in PATH
if "%PYTHON_CMD%"=="" (
    echo [ERROR] Python was not found on this system.
    echo.
    echo Please install Python 3.9 or higher from:
    echo    https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, be sure to check:
    echo    [x] Add python.exe to PATH
    echo.
    pause
    exit /b 1
)

echo Using Python: %PYTHON_CMD%
echo.
echo Starting Student API Tester...
echo Access the application in your browser at:
echo    http://127.0.0.1:8000
echo.
echo Press Ctrl+C in this window at any time to stop the server.
echo.

:: Automatically launch browser after 2 seconds in the background
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://127.0.0.1:8000"

:: Start the application
%PYTHON_CMD% app.py

:: Keep window open if the application exits unexpectedly
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Application exited with code %errorlevel%.
    pause
)

