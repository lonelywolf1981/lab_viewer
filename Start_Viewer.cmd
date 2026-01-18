@echo off
setlocal EnableExtensions

cd /d "%~dp0" || (echo FAILED: cannot cd to script folder & pause & exit /b 1)

set "LOG=%CD%\run.log"
>"%LOG%" echo ===== Run started %date% %time% =====

set "VENV_PY=%CD%\.venv\Scripts\python.exe"

if not exist "%VENV_PY%" (
  echo .venv not found. Running Setup_Once...
  echo .venv not found. Running Setup_Once...>>"%LOG%"
  call "Setup_Once.cmd" >>"%LOG%" 2>&1
)

if not exist "%VENV_PY%" (
  echo ERROR: .venv still not found. See setup.log
  echo ERROR: .venv still not found.>>"%LOG%"
  pause
  exit /b 1
)

rem Quick sanity: Flask must import
"%VENV_PY%" -c "import flask" >>"%LOG%" 2>&1
if errorlevel 1 (
  echo Flask not importable. Running Setup_Once again...
  echo Flask not importable.>>"%LOG%"
  call "Setup_Once.cmd" >>"%LOG%" 2>&1
)

echo [run] Python:>>"%LOG%"
"%VENV_PY%" -V >>"%LOG%" 2>&1

echo [run] starting server.py>>"%LOG%"


rem Open browser after a short delay (works on Win7+)
start "" /b cmd /c "ping 127.0.0.1 -n 2 >nul & start http://127.0.0.1:8787"

"%VENV_PY%" server.py >>"%LOG%" 2>&1

echo.
echo Server stopped or crashed.
echo Log: "%LOG%"
pause
exit /b 0
