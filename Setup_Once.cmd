@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ============================================================
rem LeMuRe Viewer - Setup Once (portable, adaptive)
rem - Creates .venv using host Python (py -3 preferred)
rem - Installs Flask/Werkzeug compatible with host Python
rem - Filters Flask/Werkzeug out of requirements.txt (if present)
rem - Adds sitecustomize.py shim when pkgutil.get_loader/find_loader are missing (Py 3.14+)
rem - Verifies Flask can initialize (Flask('server'))
rem ============================================================

cd /d "%~dp0" || (echo FAILED: cannot cd to script folder & pause & exit /b 1)

set "LOG=%CD%\setup.log"
>"%LOG%" echo ===== Setup started %date% %time% =====

call :log App folder: "%CD%"
call :log Log file  : "%LOG%"

rem Pick python command
set "PY=python"
where py >nul 2>nul && set "PY=py -3"

rem Detect host python version
set "PYVER_FULL="
for /f "tokens=1,2" %%A in ('%PY% -V 2^>^&1') do (
  if /i "%%A"=="Python" set "PYVER_FULL=%%B"
)

if not defined PYVER_FULL (
  call :log ERROR: Python not found in PATH.
  call :log Install Python 3.8.x for Windows 7, or Python 3.9+ for modern Windows.
  pause
  exit /b 1
)

call :log Host Python: %PYVER_FULL%

rem Parse major/minor
set "PY_MAJOR="
set "PY_MINOR="
for /f "tokens=1,2 delims=." %%a in ("%PYVER_FULL%") do (
  set "PY_MAJOR=%%a"
  set "PY_MINOR=%%b"
)

if not defined PY_MAJOR goto :badver
if not defined PY_MINOR goto :badver

rem Require Python >= 3.8
if %PY_MAJOR% LSS 3 goto :need_py38
if %PY_MAJOR% EQU 3 if %PY_MINOR% LSS 8 goto :need_py38

rem Choose Flask/Werkzeug pins
set "FLASK_SPEC="
set "WERKZEUG_SPEC="

if %PY_MAJOR% EQU 3 if %PY_MINOR% EQU 8 (
  rem Windows 7-friendly
  set "FLASK_SPEC=Flask>=3.0,<3.1"
  set "WERKZEUG_SPEC=Werkzeug>=3.0,<3.1"
) else (
  rem Modern
  set "FLASK_SPEC=Flask>=3.1,<4"
  set "WERKZEUG_SPEC=Werkzeug>=3.1,<4"
)

rem Escape < and > for logging (CMD redirection operators)
set "FLASK_SPEC_LOG=%FLASK_SPEC:>=^>%"
set "FLASK_SPEC_LOG=%FLASK_SPEC_LOG:<=^<%"
set "WERKZEUG_SPEC_LOG=%WERKZEUG_SPEC:>=^>%"
set "WERKZEUG_SPEC_LOG=%WERKZEUG_SPEC_LOG:<=^<%"

call :log Pins:
call :log   %FLASK_SPEC_LOG%
call :log   %WERKZEUG_SPEC_LOG%

rem Recreate venv to avoid mixed deps
if exist ".venv" (
  call :log Removing existing .venv ...
  rmdir /s /q ".venv" >>"%LOG%" 2>&1
)

call :log [1/5] Creating virtual environment...
%PY% -m venv .venv >>"%LOG%" 2>&1
if errorlevel 1 (
  call :log FAILED: could not create .venv
  pause
  exit /b 1
)

set "VENV_PY=%CD%\.venv\Scripts\python.exe"

call :log [2/5] Upgrading pip tooling...
"%VENV_PY%" -m pip install -U pip setuptools wheel >>"%LOG%" 2>&1
if errorlevel 1 (
  call :log FAILED: pip tooling upgrade failed
  pause
  exit /b 1
)

call :log [3/5] Installing remaining dependencies (requirements.txt, if exists)...
if exist "requirements.txt" (
  rem Create filtered requirements to prevent downgrading Flask/Werkzeug
  "%VENV_PY%" -c "import re, pathlib; p=pathlib.Path('requirements.txt'); out=pathlib.Path('requirements._filtered.txt'); txt=p.read_text(encoding='utf-8',errors='ignore').splitlines(); r=re.compile(r'^\s*(flask|werkzeug)\s*($|[<>=!~])', re.I); kept=[ln for ln in txt if ln.strip() and not r.match(ln) and not ln.lstrip().startswith('#')]; out.write_text('\n'.join(kept)+('\n' if kept else ''), encoding='utf-8')" >>"%LOG%" 2>&1
  "%VENV_PY%" -m pip install --disable-pip-version-check -r "requirements._filtered.txt" >>"%LOG%" 2>&1
  if errorlevel 1 call :log WARN: requirements.txt install had errors (continuing)
) else (
  call :log requirements.txt not found, skipping.
)

call :log [4/5] Installing/Upgrading Flask & Werkzeug...
"%VENV_PY%" -m pip install -U --disable-pip-version-check "%FLASK_SPEC%" "%WERKZEUG_SPEC%" >>"%LOG%" 2>&1
if errorlevel 1 (
  call :log FAILED: Flask/Werkzeug install failed
  pause
  exit /b 1
)

rem If pkgutil.get_loader is missing (Python 3.14+), create sitecustomize.py shim.
"%VENV_PY%" -c "import pkgutil,sys; sys.exit(0 if hasattr(pkgutil,'get_loader') else 1)" >nul 2>&1
if errorlevel 1 call :make_sitecustomize

call :log [5/5] Verifying environment...
"%VENV_PY%" -c "import sys; from importlib.metadata import version; import flask, werkzeug; print('OK: python',sys.version.split()[0]); print('OK: flask',version('flask')); print('OK: werkzeug',version('werkzeug')); from flask import Flask; Flask('server'); print('OK: Flask(server) created')" >>"%LOG%" 2>&1
if errorlevel 1 (
  call :log FAILED: Verification failed. See setup.log.
  pause
  exit /b 1
)

call :log ===== Setup DONE %date% %time% =====
call :log Setup complete.
pause
exit /b 0

:make_sitecustomize
if exist "sitecustomize.py" exit /b 0
call :log [fix] Creating sitecustomize.py (pkgutil.get_loader compat)...
"%VENV_PY%" -c "import io; io.open('sitecustomize.py','w',encoding='utf-8').write('# -*- coding: utf-8 -*-\n# Auto-generated for LeMuRe Viewer\nimport pkgutil, types, importlib.util\n\nif not hasattr(pkgutil, \"get_loader\"):\n    def get_loader(module_or_name):\n        if module_or_name is None:\n            return None\n        if isinstance(module_or_name, types.ModuleType):\n            return getattr(module_or_name, \"__loader__\", None)\n        spec = importlib.util.find_spec(module_or_name)\n        return None if spec is None else spec.loader\n    pkgutil.get_loader = get_loader\n\nif not hasattr(pkgutil, \"find_loader\"):\n    pkgutil.find_loader = pkgutil.get_loader\n')" >>"%LOG%" 2>&1
exit /b 0

:badver
call :log ERROR: Could not parse Python version: "%PYVER_FULL%"
pause
exit /b 1

:need_py38
call :log ERROR: Python 3.8+ is required. Detected: %PYVER_FULL%
pause
exit /b 1

:log
echo %*
>>"%LOG%" echo %*
exit /b 0
