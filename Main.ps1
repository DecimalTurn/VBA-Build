Write-Host "Current directory: $(pwd)"
Write-Host "Changing zip file name and location"
. ./Zip-It.ps1
Write-Host "========================="
Write-Host "Closing Office applications"
. ./Close-Office.ps1
Write-Host "========================="
Write-Host "Enabling VBOM (Visual Basic for Applications Object Model)"
. ./Enable-VBOM.ps1
Write-Host "========================="
Write-Host "Importing VBA code into Office document"
. ./Build-VBA.ps1