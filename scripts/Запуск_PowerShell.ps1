# DaData Parser - PowerShell launcher (UTF-8)
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host " DaData Parser - Administrations (Windows)" -ForegroundColor White
Write-Host " Branch arena/363bbf0c-gov" -ForegroundColor Gray
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

# Check Python
try {
    $pyVersion = python --version 2>&1
    Write-Host "[1/3] Python found: $pyVersion" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Python not found! Install from https://python.org" -ForegroundColor Red
    Write-Host "Check 'Add Python to PATH'"
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "[2/3] Installing dependencies..." -ForegroundColor Yellow
pip install openpyxl requests --quiet
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Failed to install deps, trying --user" -ForegroundColor Red
    pip install openpyxl requests --user --quiet
}

Write-Host "[3/3] Starting GUI..." -ForegroundColor Yellow
Write-Host "If window didn't open, run: python DaData_Administracii_GUI.py"
Write-Host ""

python DaData_Administracii_GUI.py

if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] Program exited with code $LASTEXITCODE" -ForegroundColor Red
}

Write-Host ""
Write-Host "Done. Results in program folder." -ForegroundColor Green
Read-Host "Press Enter to exit"
