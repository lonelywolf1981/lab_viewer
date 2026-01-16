@echo off
setlocal EnableExtensions
cd /d "%~dp0" || (echo FAILED: cd to script folder & pause & exit /b 1)

set "LOG=%CD%\setup.log"
echo ===== Setup started %date% %time% ===== > "%LOG%"

REM Prefer Python Launcher (py). If not present, fallback to python.
where py >nul 2>nul
if errorlevel 1 (
  set "PY=python"
) else (
  set "PY=py"
)

%PY% -V >> "%LOG%" 2>&1
if errorlevel 1 (
  echo Python not found. Install Python and retry. >> "%LOG%"
  echo Python not found. Install Python and retry.
  echo Log: "%LOG%"
  pause
  exit /b 1
)

echo [1/3] Creating virtual environment... >> "%LOG%"
%PY% -m venv .venv >> "%LOG%" 2>&1
if errorlevel 1 (
  echo FAILED to create venv. >> "%LOG%"
  echo FAILED to create venv. Log: "%LOG%"
  pause
  exit /b 1
)

echo [2/3] Installing dependencies... >> "%LOG%"
call ".venv\Scripts\activate.bat" >> "%LOG%" 2>&1
if errorlevel 1 (
  echo FAILED to activate venv. >> "%LOG%"
  echo FAILED to activate venv. Log: "%LOG%"
  pause
  exit /b 1
)

python -m pip install --upgrade pip setuptools wheel >> "%LOG%" 2>&1
if errorlevel 1 (
  echo FAILED to upgrade pip. >> "%LOG%"
  echo FAILED to upgrade pip. Log: "%LOG%"
  pause
  exit /b 1
)

python -m pip install -r requirements.txt >> "%LOG%" 2>&1
if errorlevel 1 (
  echo FAILED to install requirements. >> "%LOG%"
  echo FAILED to install requirements. Log: "%LOG%"
  pause
  exit /b 1
)

echo [3/3] DONE. >> "%LOG%"
echo.
echo Setup complete.
echo Log: "%LOG%"
pause
