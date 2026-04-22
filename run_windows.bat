@echo off
REM =====================================================================
REM  QBD -> QuickBooks Online Migration  (Windows launcher)
REM
REM  Run this AFTER QuickBooks Desktop is open with the company file you
REM  want to migrate. The first run creates a virtualenv and installs
REM  dependencies; subsequent runs just launch the app.
REM =====================================================================

setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo =====================================================
echo   QBD -^> QuickBooks Online Migration Connector
echo =====================================================
echo.

REM --- 1. Check Python is installed -----------------------------------
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python is not on PATH.
    echo         Install Python 3.9+ from https://www.python.org/downloads/
    echo         During install, tick "Add python.exe to PATH".
    pause
    exit /b 1
)

REM --- 2. Create venv on first run ------------------------------------
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment in .venv ...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtualenv.
        pause
        exit /b 1
    )
)

REM --- 3. Install / update dependencies -------------------------------
echo Installing dependencies...
call .venv\Scripts\python.exe -m pip install -q --upgrade pip
call .venv\Scripts\python.exe -m pip install -q -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Dependency install failed.
    pause
    exit /b 1
)

REM --- 4. Warn if pywin32 / QB SDK seems missing ----------------------
call .venv\Scripts\python.exe -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo [WARNING] pywin32 did not import. Direct QBD connect may fail.
    echo           Try:  .venv\Scripts\python.exe -m pip install pywin32
)

REM --- 5. Launch --------------------------------------------------------
echo.
echo Starting web UI at http://localhost:5050
echo Press Ctrl+C in this window to stop the server.
echo.
echo ------------------------------------------------------

set PORT=5050
call .venv\Scripts\python.exe app.py

endlocal
