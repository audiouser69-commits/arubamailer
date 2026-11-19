@echo off
title Creazione ArubaMailer.exe
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py
if not defined PY set PY=python
%PY% -m pip install --user --upgrade pyinstaller openpyxl python-docx pillow
%PY% -m PyInstaller --noconfirm --onefile --windowed --name ArubaMailer --hidden-import docx_html --collect-all docx ArubaMailer.pyw
echo.
echo Fatto: il programma e' in  dist\ArubaMailer.exe  (puoi copiarlo dove vuoi).
pause
