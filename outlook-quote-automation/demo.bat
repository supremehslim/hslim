@echo off
chcp 65001 >nul
cd /d "%~dp0"
.venv\Scripts\python.exe -m quote_automation demo
if errorlevel 1 goto end
start "" "%~dp0runtime\demo\데모_메일.html"
start "" "%~dp0runtime\demo\데모_견적서.xlsx"
:end
pause
