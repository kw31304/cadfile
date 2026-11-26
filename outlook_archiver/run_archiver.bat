@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Virtual environment not found: .venv\Scripts\python.exe
    exit /b 1
)

".venv\Scripts\python.exe" "outlook_archiver.py" %*
exit /b %ERRORLEVEL%
