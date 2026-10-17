# Taxi Software - Vendor Masterfile Live Sync (PowerShell)
Set-Location $PSScriptRoot

Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "    TAXI SOFTWARE - LIVE VENDOR MASTERFILE SYNCHRONIZER" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Directory: $(Get-Location)" -ForegroundColor Green
Write-Host "Monitoring masterfile entries and auto-syncing all vendor files..." -ForegroundColor Green
Write-Host ""
Write-Host "Press Ctrl+C in this window anytime to stop the synchronizer." -ForegroundColor Yellow
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

python -m app.cli.sync_vendors --watch

