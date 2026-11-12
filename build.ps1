# Builds dist\HyroxRadar.exe, a single file that runs without Python.
# Usage (PowerShell, from the project folder):  .\build.ps1
$ErrorActionPreference = "Stop"
$python = ".venv/Scripts/python.exe"

if (-not (Test-Path $python)) {
    uv venv .venv
}
uv pip install --python $python -r requirements-dev.txt

& $python -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name HyroxRadar `
    --exclude-module pytest --exclude-module IPython --exclude-module PyQt5 `
    --exclude-module PyQt6 --exclude-module PySide6 --exclude-module scipy `
    app.py
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

$exe = Get-Item "dist/HyroxRadar.exe"
"Built {0} ({1:N0} MB)" -f $exe.FullName, ($exe.Length / 1MB)
