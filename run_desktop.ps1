$ErrorActionPreference = "Stop"
$repo = $PSScriptRoot
$wb   = "$env:USERPROFILE\Box\2. Deal Flow\Novel Energy Solutions\Pricing Model\38DN-IL_Novel Energy Solutions_lease financing_PricingModel_100% Commercial_2026.04.15.xlsm"

if (-not (Test-Path $wb)) {
    Write-Host "Workbook not found at: $wb" -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit 1
}

Set-Location $repo
python -u solve_via_macro.py $wb --timeout 5400
Read-Host "Done. Press Enter to close"
