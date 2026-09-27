@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Run start_scraper.bat once before using this installer.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m playwright install chromium
if errorlevel 1 (
  echo Chromium installation failed.
) else (
  echo Chromium is ready for JavaScript-rendered pages.
)
pause
endlocal
