@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating Python environment...
    py -3 -m venv .venv
    if errorlevel 1 (
        echo Could not create Python environment. Install Python 3.11 or newer.
        pause
        exit /b 1
    )
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Python dependency installation failed.
        pause
        exit /b 1
    )
    ".venv\Scripts\python.exe" -m playwright install chromium
    if errorlevel 1 (
        echo Chromium installation failed.
        pause
        exit /b 1
    )
)

if exist "data\customers.xlsx" (
    set "INPUT_FILE=data\customers.xlsx"
) else (
    set "INPUT_FILE=data\customers.csv"
)

if not exist "%INPUT_FILE%" (
    copy /Y "data\customers.example.csv" "data\customers.csv" >nul
    set "INPUT_FILE=data\customers.csv"
    echo Created data\customers.csv. Fill in customer records and run again.
    pause
    exit /b 0
)

if not exist ".playwright-ready" (
    echo Installing Chromium browser runtime...
    ".venv\Scripts\python.exe" -m playwright install chromium
    if errorlevel 1 (
        echo Chromium installation failed.
        pause
        exit /b 1
    )
    echo ready>.playwright-ready
)

echo Starting GUI. Tasks run in formal execution mode after confirmation.
".venv\Scripts\python.exe" gui.py
pause
