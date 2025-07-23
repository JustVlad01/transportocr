@echo off
echo Starting OptimoRoute Sorter...
python optimoroute_sorter_app.py
if errorlevel 1 (
    echo.
    echo ERROR: Failed to start application
    echo Make sure you have run install_and_run.bat first
    pause
)
