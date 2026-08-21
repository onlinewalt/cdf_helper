@echo off
REM ============================================================
REM IIS + wfastcgi setup script for CDF Helper
REM ============================================================
REM Run this script once on the Windows Server AFTER installing Python.
REM It will:
REM   1. Install project dependencies (requirements.txt)
REM   2. Install wfastcgi
REM   3. Enable wfastcgi (registers FastCGI with IIS)
REM   4. Auto-detect Python path and update web.config
REM ============================================================

cd /d "%~dp0"

echo.
echo [1/4] Installing project dependencies...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies. Ensure Python is in PATH and pip is available.
    pause
    exit /b 1
)

echo.
echo [2/4] Installing wfastcgi...
python -m pip install wfastcgi
if errorlevel 1 (
    echo ERROR: Failed to install wfastcgi.
    pause
    exit /b 1
)

echo.
echo [3/4] Enabling wfastcgi (registers with IIS FastCGI module)...
python -m wfastcgi-enable
if errorlevel 1 (
    echo WARNING: wfastcgi-enable failed. You may need to run this script as Administrator
    echo and ensure IIS + CGI + FastCGI features are installed.
    echo.
    echo Manual commands if needed:
    echo   python -m wfastcgi-enable
    echo.
    echo Then manually update web.config scriptProcessor with the path returned above.
    pause
    exit /b 1
)

echo.
echo [4/4] Auto-detecting Python path for web.config...
python -c "
import sys, os, re

python_exe = sys.executable
wfastcgi_path = os.path.join(sys.prefix, 'lib', 'site-packages', 'wfastcgi.py')

# If not found in site-packages, try to find it
if not os.path.exists(wfastcgi_path):
    import wfastcgi
    wfastcgi_path = wfastcgi.__file__

script_processor = f'{python_exe}|{wfastcgi_path}'
print(f'  Python: {python_exe}')
print(f'  wfastcgi: {wfastcgi_path}')
print(f'  scriptProcessor: {script_processor}')

# Update web.config
with open('web.config', 'r', encoding='utf-8') as f:
    config = f.read()

# Replace the scriptProcessor value
config = re.sub(
    r'scriptProcessor=\"[^\"]*\"',
    f'scriptProcessor=\"{script_processor}\"',
    config
)

with open('web.config', 'w', encoding='utf-8') as f:
    f.write(config)

print('  web.config updated successfully.')
"

if errorlevel 1 (
    echo WARNING: Could not auto-update web.config.
    echo Please manually edit web.config and set the correct scriptProcessor path.
    echo The format should be: C:\path\to\python.exe|C:\path\to\wfastcgi.py
    pause
    exit /b 1
)

echo.
echo ================================================================
echo Setup complete! Next steps:
echo.
echo 1. Create an IIS Site / Application pointing to this folder
echo 2. Set the Application Pool to "No Managed Code"
echo 3. Ensure IIS has read+execute permissions on this folder
echo 4. Ensure IIS has write permissions on the 'uploads/', 'generated/',
echo    'logs/', and 'translate_cache.json' paths
echo 5. Browse to http://your-server/
echo ================================================================
echo.
pause
