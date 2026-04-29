@echo off
TITLE NexusCRM Server
echo =======================================
echo     Starting NexusCRM Server...
echo =======================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python 3.8+ from python.org and check "Add to PATH/Environment variables".
    pause
    exit /b
)

:: Ensure a Python virtual environment exists
if not exist "venv" (
    echo [INFO] Creating Python virtual environment...
    python -m venv venv
)

:: Activate the environment and install dependencies quietly
call venv\Scripts\activate.bat
if not exist "venv\Lib\site-packages\flask" (
    echo [INFO] Installing required dependencies...
    pip install -q -r requirements.txt
)

:: Check if a custom Excel file was passed as drag-and-drop
if "%~1"=="" (
    echo  Data Source: CRM_Sample_Data_Template.xlsx ^(Default^)
    echo  URL:         http://localhost:5000
    echo.
    echo Press Ctrl+C to stop the application.
    python app.py
) else (
    echo  Data Source: %~1
    echo  URL:         http://localhost:5000
    echo.
    echo Press Ctrl+C to stop the application.
    python app.py "%~1"
)

pause
