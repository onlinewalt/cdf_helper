@echo off
rem CDF Helper launcher - double-click to start the web UI
rem main.py starts the server and opens the browser automatically
cd /d "%~dp0"
python main.py
if errorlevel 1 (
    echo.
    echo Failed to start. Please install Python and run: pip install -r requirements.txt
    pause
)