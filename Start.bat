@echo off
title Test certificates
cd /d "%~dp0"

rem A portable Python in this folder (made by tools\make_portable.py) needs nothing on the PC
if exist "python\python.exe" (
  set PY="%~dp0python\python.exe" -s
  goto run
)

where python >nul 2>nul
if errorlevel 1 (
  echo.
  echo   Python is not installed on this computer.
  echo   Install it from python.org, then double-click Start.bat again.
  echo.
  pause
  exit /b 1
)
set PY=python

%PY% -c "import flask, openpyxl, xlrd, docx, reportlab" >nul 2>nul
if errorlevel 1 (
  echo   Installing what the app needs, one time only...
  %PY% -m pip install --quiet -r requirements.txt
)

:run
rem Close any copy of the app that is still open (an older one would keep showing old screens)
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -match 'app\.web' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>nul

%PY% -m app.web
pause
