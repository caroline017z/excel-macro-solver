$ErrorActionPreference = "Stop"
$repo = $PSScriptRoot
$wb   = "$env:USERPROFILE\OneDrive - 38 Degrees North\Desktop\38DN-IL_US Solar_PricingModel_Test - Copy.xlsm"

if (-not (Test-Path $wb)) {
    Write-Host "Workbook not found at: $wb" -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit 1
}

Set-Location $repo
python solve_via_macro.py $wb --timeout 5400
Read-Host "Done. Press Enter to close"
