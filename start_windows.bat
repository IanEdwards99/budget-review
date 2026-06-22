@echo off
setlocal

cd /d "%~dp0"

where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
  set "PYTHON_CMD=py -3"
) else (
  where python >nul 2>nul
  if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python"
  ) else (
    echo Python is not installed or not available on PATH.
    echo Install Python 3.11 or newer from https://www.python.org/downloads/
    pause
    exit /b 1
  )
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating local Python environment...
  %PYTHON_CMD% -m venv .venv
  if %ERRORLEVEL% NEQ 0 (
    echo Failed to create Python virtual environment.
    pause
    exit /b 1
  )
)

echo Installing/updating dependencies...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
  echo Failed to install dependencies.
  pause
  exit /b 1
)

echo Starting Budget Review App...
start "" http://127.0.0.1:5057
".venv\Scripts\python.exe" app.py

endlocal

