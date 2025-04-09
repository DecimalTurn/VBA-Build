# Get the source directory from command line argument or use default "src"
param(
    [string]$SourceDir = "src"
)

# Read name of the folders under src into an array
$folders = Get-ChildItem -Path "$PSScriptRoot\src" -Directory | Select-Object -ExpandProperty Name
Write-Host "Folders in src: $folders"

# Check if the folders array is empty
if ($folders.Count -eq 0) {
    Write-Host "No folders found in src. Exiting script."
    exit 1
}

$officeApps = @()

function Get-OfficeApp {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FileExtension
    )

    switch -Regex ($FileExtension.ToLower()) {
        '^(xlsb|xlsm||xltm|xlam)$' { return "Excel" }
        '^(docm|dotm)$' { return "Word" }
        '^(pptm|potm)$' { return "PowerPoint" }
        '^(accdb|accda)$' { return "Access" }
        default { return $null }
    }
}

foreach ($folder in $folders) {
    $FileExtension = $folder.Substring($folder.LastIndexOf('.') + 1)
    $app = Get-OfficeApp -FileExtension $FileExtension
    
    if ($app) {
        if ($officeApps -notcontains $app) {
            $officeApps += $app
        }
    } else {
        Write-Host "Unknown file extension: $FileExtension. Skipping..."
        continue
    }
}

foreach $app in $officeApps {
    Write-Host "Enabling VBOM for $app"
    . "$PSScriptRoot\scripts\Enable-VBOM.ps1" $app
    Write-Host "========================="
}

Write-Host "Open and close Office applications"
. "$PSScriptRoot\scripts\Open-Close-Office.ps1" $officeApps
Write-Host "========================="

foreach $folder in $folders {

    $app = Get-OfficeApp -FileExtension $folder.Substring($folder.LastIndexOf('.') + 1)
    if ($app) {
        Write-Host "Creating $app document"
        . "$PSScriptRoot\scripts\Create-$app.ps1" $folder
        Write-Host "========================="
    } else {
        Write-Host "Unknown file extension: $folder. Skipping..."
        continue
    }

    if ($app -neq "Access") {
        Write-Host "Create Zip file and rename it to Office document target"
        . "$PSScriptRoot\scripts\Zip-It.ps1" $folder
        Write-Host "========================="
    }

    Write-Host "Importing VBA code into Office document" 
    . "$PSScriptRoot\scripts\Build-VBA.ps1" $folder

}