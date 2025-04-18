# Get the source directory from command line argument or use default "src"
param(
    [string]$SourceDir = "src"
)

Write-Host "Changing zip file name and location"
. "$PSScriptRoot\Zip-It.ps1"
Write-Host "========================="

Write-Host "Closing Office applications"
. "$PSScriptRoot\Close-Office.ps1"
Write-Host "========================="

Write-Host "Enabling VBOM (Visual Basic for Applications Object Model)"
. "$PSScriptRoot\Enable-VBOM.ps1"
Write-Host "========================="

Write-Host "Importing VBA code into Office document"
. "$PSScriptRoot\Build-VBA.ps1"
