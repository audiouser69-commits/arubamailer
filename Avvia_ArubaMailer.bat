@echo off
title ArubaMailer
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY (
  echo Python non e' installato.
  echo Si apre la pagina di download: installa Python 3 e spunta "Add python.exe to PATH".
  start https://www.python.org/downloads/windows/
  pause
  exit /b 1
)
%PY% -c "import openpyxl" 2>nul || (
  echo Installo il componente per leggere Excel ^(solo la prima volta^)...
  %PY% -m pip install --user openpyxl
)
%PY% ArubaMailer.pyw
if errorlevel 1 pause
