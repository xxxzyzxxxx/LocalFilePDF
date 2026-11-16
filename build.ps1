$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
if (-not (Test-Path -LiteralPath '.venv\Scripts\python.exe')) {
    python -m venv .venv
    if ($LASTEXITCODE -ne 0) { throw 'Python venv creation failed' }
}
& '.\.venv\Scripts\python.exe' -m pip install -r requirements-dev.txt
if ($LASTEXITCODE -ne 0) { throw 'Dependency install failed' }
& '.\.venv\Scripts\python.exe' -m PyInstaller --noconfirm --clean --windowed --onedir --name LocalFilePDF --hidden-import win32timezone --hidden-import pythoncom --hidden-import pywintypes --collect-submodules win32com --exclude-module PySide6.QtWebEngineCore --exclude-module PySide6.QtWebEngineWidgets --exclude-module PySide6.QtQml --exclude-module PySide6.QtQuick app.py
if ($LASTEXITCODE -ne 0) { throw 'Packaging failed' }
Copy-Item -LiteralPath 'README.md','TEST_RESULTS.md','THIRD_PARTY_NOTICES.md' -Destination 'dist\LocalFilePDF'
Copy-Item -LiteralPath 'requirements-lock.txt' -Destination 'dist\LocalFilePDF'
Copy-Item -LiteralPath 'licenses' -Destination 'dist\LocalFilePDF\licenses' -Recurse -Force
Write-Host 'Ready: dist\LocalFilePDF\LocalFilePDF.exe'
