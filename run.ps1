# Starts the grade automation. Run from PowerShell in this folder:
#   powershell -ExecutionPolicy Bypass -File .\run.ps1
#
# (Run setup.ps1 once first.)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$venvPy = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPy)) {
    Write-Host "Not set up yet. Run this first:" -ForegroundColor Yellow
    Write-Host "    powershell -ExecutionPolicy Bypass -File .\setup.ps1"
    Read-Host "Press ENTER to close"
    exit 1
}

& $venvPy poc.py
