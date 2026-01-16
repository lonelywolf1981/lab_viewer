@echo off
setlocal EnableExtensions
cd /d "%~dp0" || (echo FAILED: cd to script folder & pause & exit /b 1)

set "LOG=%CD%\run.log"
echo ===== Run started %date% %time% ===== > "%LOG%"

if not exist ".venv\Scripts\python.exe" (
  echo Venv not found. Run Setup_Once.cmd first.
  echo Venv not found. >> "%LOG%"
  echo Log: "%LOG%"
  pause
  exit /b 1
)

call ".venv\Scripts\activate.bat"


python server.py >> "%LOG%" 2>&1

echo.
echo Server stopped or crashed.
echo Log: "%LOG%"
pause
