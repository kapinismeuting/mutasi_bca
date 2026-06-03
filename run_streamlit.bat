@echo off
REM BCA Statement Converter - Streamlit Launcher (Windows)
REM Simple script to start the Streamlit application

setlocal enabledelayedexpansion

echo.
echo ============================================
echo    BCA Statement Converter - Streamlit
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.7+
    echo Download from: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [OK] Python found: 
python --version

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate.bat
echo [OK] Virtual environment activated

REM Install requirements
echo Checking dependencies...
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements.txt
echo [OK] Dependencies installed

REM Load .env if exists
if exist ".env" (
    echo [OK] Loading environment variables from .env
    for /f "tokens=*" %%A in (.env) do (
        if not "%%A"=="" if not "!%%A:~0,1!"=="#" (
            set "%%A"
        )
    )
)

echo.
echo Configuration:
echo   PDF Folder:    %PDF_FOLDER%
echo   Output Folder: %OUTPUT_FOLDER%
echo   Log Level:     %LOG_LEVEL%
echo.

echo Starting Streamlit application...
echo If browser doesn't open, visit: http://localhost:8501
echo.

streamlit run streamlit_app.py

REM Deactivate on exit
call venv\Scripts\deactivate.bat

pause
