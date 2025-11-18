@echo off
cls

echo ==========================================
echo Excel to PDF Converter - Web Application
echo ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo X Python is not installed. Please install Python 3 first.
    pause
    exit /b 1
)

echo [OK] Python found

REM Install dependencies
echo.
echo [*] Installing dependencies...
pip install -q flask pandas reportlab openpyxl

if errorlevel 1 (
    echo X Failed to install dependencies
    pause
    exit /b 1
)

echo [OK] Dependencies installed successfully

REM Start the server
echo.
echo [*] Starting web server...
echo.
echo ==========================================
echo [OK] Server is running!
echo ==========================================
echo.
echo Open your browser and go to:
echo.
echo    http://localhost:5000
echo.
echo ==========================================
echo.
echo Press Ctrl+C to stop the server
echo.

REM Run the Flask app
python app.py
