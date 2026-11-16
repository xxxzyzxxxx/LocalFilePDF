@echo off
cd /d "%~dp0"
if exist "dist\LocalFilePDF\LocalFilePDF.exe" (
    start "" "dist\LocalFilePDF\LocalFilePDF.exe"
    exit /b
)
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "app.py"
    exit /b
)
echo Portable program or Python environment not found. Please see README.md.
pause
