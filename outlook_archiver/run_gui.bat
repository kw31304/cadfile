@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
    echo [ERROR] Virtual environment not found: .venv\Scripts\pythonw.exe
    exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" "outlook_archiver_gui.py"
exit /b 0
