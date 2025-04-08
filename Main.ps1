# Get the source directory from command line argument or use default "src"
param(
    [string]$SourceDir = "src"
)

Write-Host "Create Zip file and rename it to Office document target"
. "$PSScriptRoot\scripts\Zip-It.ps1"
Write-Host "========================="

Write-Host "Closing Office applications"
. "$PSScriptRoot\scripts\Close-Office.ps1"
Write-Host "========================="

Write-Host "Enabling VBOM (Visual Basic for Applications Object Model)"
. "$PSScriptRoot\scripts\Enable-VBOM.ps1"
Write-Host "========================="

Write-Host "Importing VBA code into Office document"
. "$PSScriptRoot\scripts\Build-VBA.ps1"
