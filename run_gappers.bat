@echo off
REM ============================================================
REM  Stock Gap Analysis - Windows launcher
REM  Double-click this file, or run it from cmd with arguments:
REM    run_gappers.bat --ticker TSLA --start 2026-01-01 --end 2026-06-01
REM  With no arguments, it runs with the script's defaults (AAPL).
REM ============================================================

setlocal

REM Go to the folder this .bat file lives in, regardless of where it's run from
cd /d "%~dp0"

REM Create a virtual environment on first run if it doesn't exist yet
if not exist venv (
    echo [setup] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo [error] Could not create venv. Is Python installed and on PATH?
        pause
        exit /b 1
    )
)

call venv\Scripts\activate.bat

echo [setup] Installing/checking dependencies...
pip install -q -r requirements.txt
if errorlevel 1 (
    echo [error] Failed to install dependencies.
    pause
    exit /b 1
)

echo [run] Starting gap analysis...
python gappers.py %*

echo.
echo [done] Finished. Press any key to close.
pause >nul