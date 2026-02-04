# Script untuk deploy ke Testing URL yang tetap (terpisah dari production)
$TestingDeploymentId = "AKfycbwUb6MY7ILw_0j3ChAEQODE4GX0vbVrTFNVhiFF3f847n_298E2M0GmiGcQn-VeWfbb"
$Date = Get-Date -Format "yyyy-MM-dd HH:mm"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  TESTING DEPLOYMENT" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nPushing changes to Apps Script..."
clasp push

Write-Host "`nDeploying to testing URL..."
clasp deploy --deploymentId $TestingDeploymentId --description "Testing: $Date"

$TestingUrl = "https://script.google.com/macros/s/$TestingDeploymentId/exec"

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "  DEPLOYMENT SUCCESS!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "`nTesting URL (stable):" -ForegroundColor Yellow
Write-Host $TestingUrl -ForegroundColor White

# Copy to clipboard
$TestingUrl | Set-Clipboard
Write-Host "`n[URL copied to clipboard]" -ForegroundColor Cyan

Write-Host ""
Pause
