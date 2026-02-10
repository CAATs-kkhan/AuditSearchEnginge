@echo off
echo ============================================
echo   Starting Audit Search Engine...
echo ============================================
echo.

cd /d "%~dp0"

REM Check if virtual environment exists
if not exist "venv\Scripts\python.exe" (
    echo ERROR: Virtual environment not found!
    echo.
    echo Please run SETUP.bat first to install the application.
    echo.
    pause
    exit /b 1
)

echo Activating virtual environment...
call venv\Scripts\activate.bat

echo.
echo Starting the application...
echo.
echo The application will open in your web browser.
echo If browser doesn't open, go to: http://localhost:8501
echo.
echo To stop the application, press Ctrl+C in this window.
echo.

REM Use full path to streamlit
venv\Scripts\streamlit.exe run app\search_app.py

pause
