@echo off
title Test certificates
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo   Python is not installed on this computer.
  echo   Install it from python.org, then double-click Start.bat again.
  echo.
  pause
  exit /b 1
)

python -c "import flask, openpyxl, xlrd" >nul 2>nul
if errorlevel 1 (
  echo   Installing what the app needs, one time only...
  python -m pip install --quiet -r requirements.txt
)

python -m app.web
pause
