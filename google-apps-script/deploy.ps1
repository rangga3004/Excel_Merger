# Script untuk deploy ke Production URL yang tetap
$DeploymentId = "AKfycbxLMEU6JFZPWLEmp7KM5x9zXnbaaYUc1A1rPK_wStoJ3_L1PfFa0NQSNRW1hH_-rce3"
$Date = Get-Date -Format "yyyy-MM-dd HH:mm"

Write-Host "Pushing changes to Apps Script..."
clasp push

Write-Host "Deploying to stable ID..."
clasp deploy --deploymentId $DeploymentId --description "Update $Date"

Write-Host "Done! URL remains the same."
Pause
