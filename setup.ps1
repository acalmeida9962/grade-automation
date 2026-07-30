# One-time setup for Windows. Run from PowerShell in this folder:
#   powershell -ExecutionPolicy Bypass -File .\setup.ps1
#
# Creates a local Python virtual environment, installs the dependencies, and
# downloads the Chromium browser that Playwright drives. Safe to re-run.

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

Write-Host "=== Grade automation setup ===" -ForegroundColor Cyan

# 1. Find a Python launcher (python.exe or the 'py' launcher).
$py = $null
foreach ($cand in @("python", "py")) {
    if (Get-Command $cand -ErrorAction SilentlyContinue) { $py = $cand; break }
}
if (-not $py) {
    Write-Host ""
    Write-Host "Python was not found." -ForegroundColor Red
    Write-Host "Install it first (see README, 'Step 0'), e.g.:"
    Write-Host "    winget install -e --id Python.Python.3.12"
    Write-Host "Then open a NEW PowerShell window and run this script again."
    Read-Host "Press ENTER to close"
    exit 1
}
Write-Host "Using Python launcher: $py"
& $py --version

# 2. Create the virtual environment.
if (-not (Test-Path ".\.venv")) {
    Write-Host "Creating virtual environment (.venv)..."
    & $py -m venv .venv
} else {
    Write-Host "Virtual environment already exists (.venv)."
}

$venvPy = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

# 3. Install dependencies.
Write-Host "Upgrading pip..."
& $venvPy -m pip install --upgrade pip
Write-Host "Installing Python packages (playwright, openpyxl)..."
& $venvPy -m pip install -r requirements.txt

# 4. Download the Chromium browser Playwright will drive (~150 MB).
Write-Host "Downloading Chromium for Playwright (this can take a few minutes)..."
& $venvPy -m playwright install chromium

Write-Host ""
Write-Host "Setup complete. Start the tool with:  .\run.ps1" -ForegroundColor Green
Read-Host "Press ENTER to close"
