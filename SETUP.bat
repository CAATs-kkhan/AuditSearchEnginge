@echo off
echo ============================================
echo   Audit Search Engine - Setup Script
echo ============================================
echo.

echo Step 1: Checking Python installation...
python --version
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo.
echo Step 2: Creating virtual environment...
python -m venv venv
if %errorlevel% neq 0 (
    echo ERROR: Could not create virtual environment
    pause
    exit /b 1
)

echo.
echo Step 3: Activating virtual environment...
call venv\Scripts\activate.bat

echo.
echo Step 4: Upgrading pip...
python -m pip install --upgrade pip

echo.
echo Step 5: Installing required packages...
echo This may take a few minutes...
pip install -r requirements.txt

echo.
echo ============================================
echo   Setup Complete!
echo ============================================
echo.
echo To run the application:
echo   1. Double-click RUN_APP.bat
echo   OR
echo   2. Open Command Prompt in this folder and run:
echo      venv\Scripts\activate
echo      streamlit run app\search_app.py
echo.
echo Place your documents in the "documents" folder.
echo.
pause
