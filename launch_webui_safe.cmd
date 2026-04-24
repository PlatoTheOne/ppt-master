@echo off
setlocal

cd /d "%~dp0"
set "URL=http://127.0.0.1:8765/"

where python >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=py -3"
    ) else (
        echo.
        echo [PPT Master] Python was not found.
        echo Install Python or add it to PATH, then try again.
        echo.
        pause
        exit /b 1
    )
)

echo.
echo [PPT Master] Starting local Web UI...
echo [PPT Master] Browser address: %URL%
echo [PPT Master] Keep this window open while using the app.
echo.

call %PYTHON_CMD% -c "import sys, urllib.request; urllib.request.urlopen('%URL%', timeout=2); sys.exit(0)" >nul 2>nul
if %errorlevel%==0 (
    echo [PPT Master] Web UI is already running.
    start "" "%URL%"
    exit /b 0
)

start "" cmd /c "timeout /t 3 /nobreak >nul && start "" %URL%"
call %PYTHON_CMD% run_webui.py
