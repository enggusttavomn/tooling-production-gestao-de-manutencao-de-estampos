@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe"
set "APP_ENTRY=%SCRIPT_DIR%app.py"

echo Starting server...
echo Using: "%PYTHON_EXE%"

if not exist "%PYTHON_EXE%" (
  echo.
  echo ERROR: Python venv not found.
  echo Expected: "%PYTHON_EXE%"
  echo.
  echo Tip: Create the venv or update this script.
  pause
  exit /b 1
)

if not exist "%APP_ENTRY%" (
  echo.
  echo ERROR: app.py not found.
  echo Expected: "%APP_ENTRY%"
  pause
  exit /b 1
)

"%PYTHON_EXE%" "%APP_ENTRY%"

echo.
echo Server stopped or failed to start.
pause
