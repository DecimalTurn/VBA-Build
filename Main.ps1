# Get the source directory from command line argument or use default "src"
param(
    [string]$SourceDir = "src"
)

Write-Host "Current directory: $(pwd)"
Write-Host "Using source directory: $SourceDir"

# Read name of the folders under the specified source directory into an array
$folders = Get-ChildItem -Path "$PSScriptRoot/$SourceDir" -Directory | Select-Object -ExpandProperty Name
Write-Host "Folders in ${SourceDir}: $folders"

# Check if the folders array is empty
if ($folders.Count -eq 0) {
    Write-Host "No folders found in ${SourceDir}. Exiting script."
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

# Create a list of Office applications that are needed based on the file extensions of the folders
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


# We need to open and close the Office applications before we can enable VBOM
Write-Host "Open and close Office applications"
. "$PSScriptRoot/scripts/Open-Close-Office.ps1" $officeApps
Write-Host "========================="

# Enable VBOM for each Office application
foreach ($app in $officeApps) {
    Write-Host "Enabling VBOM for $app"
    . "$PSScriptRoot/scripts/Enable-VBOM.ps1" $app
    Write-Host "========================="
}

foreach ($folder in $folders) {
    $app = Get-OfficeApp -FileExtension $folder.Substring($folder.LastIndexOf('.') + 1)

    if ($app -ne "Access") {
        Write-Host "Create Zip file and rename it to Office document target"
        . "$PSScriptRoot/scripts/Zip-It.ps1" "${SourceDir}/${folder}"
        Write-Host "========================="
    }

    Write-Host "Importing VBA code into Office document" 
    . "$PSScriptRoot/scripts/Build-VBA.ps1" "${SourceDir}/${folder}" "$app"
}