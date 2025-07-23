@echo off
echo OptimoRoute Sorter - Installation and Setup
echo ==========================================
echo.

echo Checking Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://python.org
    pause
    exit /b 1
)

echo Python found. Installing dependencies...
pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies
    echo Please check your internet connection and try again
    pause
    exit /b 1
)

echo.
echo Installation completed successfully!
echo Starting OptimoRoute Sorter...
echo.

python optimoroute_sorter_app.py

if errorlevel 1 (
    echo.
    echo ERROR: Failed to start application
    echo Please check the error messages above
    pause
    exit /b 1
)
