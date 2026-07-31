@echo off
REM ============================================================
REM  Stock Gap Analysis - Windows launcher
REM
REM  Run with your own arguments (skips the prompts below):
REM    run_gappers.bat --ticker TSLA --start 2026-01-01 --end 2026-06-01
REM    run_gappers.bat --ticker AAPL,TSLA,NVDA
REM
REM  Or just double-click with no arguments - it will ask you
REM  interactively, including for multiple tickers at once.
REM ============================================================

setlocal enabledelayedexpansion

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

REM If arguments were passed (dragged onto the file, or run from cmd), use them as-is
if not "%~1"=="" (
    echo [run] Starting gap analysis with provided arguments...
    python gappers.py %*
    goto :end
)

REM ---- No arguments: ask interactively ----
echo.
echo === Stock Gap Analysis - interactive setup ===
echo (press Enter to accept the default shown in brackets)
echo.

set "TICKERS="
set /p TICKERS="Ticker(s), space or comma separated [AAPL]: "
if "%TICKERS%"=="" set "TICKERS=AAPL"
REM normalize spaces to commas so "AAPL TSLA NVDA" becomes "AAPL,TSLA,NVDA"
set "TICKERS=%TICKERS: =,%"

set "START_DATE="
set /p START_DATE="Start date YYYY-MM-DD [2026-01-01]: "
if "%START_DATE%"=="" set "START_DATE=2026-01-01"

set "END_DATE="
set /p END_DATE="End date YYYY-MM-DD [2026-05-15]: "
if "%END_DATE%"=="" set "END_DATE=2026-05-15"

set "THRESHOLD="
set /p THRESHOLD="Gap threshold, decimal [0.015]: "
if "%THRESHOLD%"=="" set "THRESHOLD=0.015"

set "DIRECTION="
set /p DIRECTION="Gap direction: up / down / both [up]: "
if "%DIRECTION%"=="" set "DIRECTION=up"

echo.
echo [run] Analyzing: %TICKERS%  (%START_DATE% to %END_DATE%, threshold %THRESHOLD%, direction %DIRECTION%)
python gappers.py --ticker %TICKERS% --start %START_DATE% --end %END_DATE% --threshold %THRESHOLD% --direction %DIRECTION%

:end
echo.
echo [done] Finished. Press any key to close.
pause >nul