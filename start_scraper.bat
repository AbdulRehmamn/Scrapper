@echo off
setlocal
cd /d "%~dp0"
title ScrapeFlow Pro

where py >nul 2>&1
if %errorlevel%==0 (
  set "PYTHON_CMD=py -3"
) else (
  set "PYTHON_CMD=python"
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating the local Python environment...
  %PYTHON_CMD% -m venv .venv
  if errorlevel 1 goto :error
)

echo Installing or checking required packages...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto :error

echo Installing or checking the Chromium browser engine...
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 echo WARNING: Chromium setup failed. Direct HTTP scraping will still work.

echo.
echo Starting ScrapeFlow Pro at http://127.0.0.1:8765
".venv\Scripts\python.exe" app.py
goto :end

:error
echo.
echo Setup could not finish. Confirm that Python 3.10 or newer and internet access are available.
pause

:end
endlocal
