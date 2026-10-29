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

python -c "import flask, openpyxl, xlrd, docx, reportlab" >nul 2>nul
if errorlevel 1 (
  echo   Installing what the app needs, one time only...
  python -m pip install --quiet -r requirements.txt
)

rem Close any copy of the app that is still open (an older one would keep showing old screens)
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -match 'app\.web' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>nul

python -m app.web
pause
