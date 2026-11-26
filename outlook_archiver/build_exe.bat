@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

if not exist ".venv\Scripts\pyinstaller.exe" (
    echo [ERROR] PyInstaller not found in .venv
    exit /b 1
)

".venv\Scripts\pyinstaller.exe" ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "OutlookArchiver_V1.0" ^
    --version-file "version_info.txt" ^
    --distpath "." ^
    --workpath "build" ^
    --paths "." ^
    --hidden-import "outlook_archiver" ^
    --hidden-import "win32com.client" ^
    --hidden-import "win32timezone" ^
    --hidden-import "pythoncom" ^
    --hidden-import "pywintypes" ^
    --exclude-module "PySide6.QtWebEngineCore" ^
    --exclude-module "PySide6.QtWebEngineWidgets" ^
    --exclude-module "PySide6.QtWebEngineQuick" ^
    --exclude-module "PySide6.QtPdf" ^
    --exclude-module "PySide6.QtPdfWidgets" ^
    --exclude-module "PySide6.QtMultimedia" ^
    --exclude-module "PySide6.QtMultimediaWidgets" ^
    --exclude-module "PySide6.QtQml" ^
    --exclude-module "PySide6.QtQuick" ^
    --exclude-module "PySide6.QtQuickWidgets" ^
    --exclude-module "PySide6.Qt3DCore" ^
    --exclude-module "PySide6.Qt3DRender" ^
    "outlook_archiver_gui.py"

exit /b %ERRORLEVEL%
