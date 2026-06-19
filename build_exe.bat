@echo off
chcp 65001 >nul
cd /d "%~dp0"

call upd\Scripts\activate.bat

pyinstaller --onefile --windowed --name xlsx2xml gui.py
pyinstaller --onefile --name xlsx2xml_cli generate_xml.py

echo.
echo Exe-файлы в папке dist\
pause
