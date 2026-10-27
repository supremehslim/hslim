@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -3 -m venv .venv
if errorlevel 1 goto fail
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto fail
if not exist .env copy .env.example .env >nul
if not exist config.json copy config.example.json config.json >nul
echo 설치 완료. README.md의 M365 앱 등록 및 API 설정을 진행하세요.
notepad .env
notepad config.json
pause
exit /b 0
:fail
echo 설치 실패. Python 3.11 이상 설치 및 인터넷 연결을 확인하세요.
pause
exit /b 1
