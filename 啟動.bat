@echo off
cd /d "%~dp0"

python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Download: https://www.python.org/downloads/
    pause
    exit /b 1
)

python -c "import win32print, PIL, qrcode, openpyxl" > nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    python -m pip install pywin32 qrcode[pil] pillow openpyxl --quiet
    if errorlevel 1 (
        echo [ERROR] Installation failed. Check internet connection.
        pause
        exit /b 1
    )
    echo Packages installed.
)

python run.py
if errorlevel 1 pause
