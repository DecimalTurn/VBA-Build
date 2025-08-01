# Get the source directory from command line argument or use default "src"
param(
    [string]$SourceDir = "src",
    [string]$TestFramework = "none", # Default to "none" if not specified
    [string]$OfficeAppDetection = "automatic" # Default to "automatic" if not specified
)

Write-Host "Current directory: $(pwd)"
Write-Host "Using source directory: $SourceDir"

# Read name of the folders under the specified source directory into an array
$CurrentWorkingDir = Get-Location
$folders = Get-ChildItem -Path "$CurrentWorkingDir/$SourceDir" -Directory | Select-Object -ExpandProperty Name
Write-Host "Folders in ${SourceDir}: $folders"

# Check if the folders array is empty
if ($folders.Count -eq 0) {
    Write-Host "No folders found in ${SourceDir}. Exiting script."
    exit 1
}

$officeApps = @()
$processedFolders = 0
$successfulBuilds = 0
$accessFolders = @()
$hasAccessDatabase = $false

function Get-OfficeApp {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FileExtension
    )

    switch -Regex ($FileExtension.ToLower()) {
        '^(xlsb|xlsm||xltm|xlam)$' { return "Excel" }
        '^(docm|dotm)$' { return "Word" }
        '^(pptm|potm|ppam)$' { return "PowerPoint" }
        '^(accdb|accda|accde)$' { return "Access" }
        default { return $null }
    }
}

if ($OfficeAppDetection -ieq "automatic") {

    Write-Host "Automatic detection of Office applications based on file extensions"

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
    
} else {
    # We parse the OfficeApp parameter to get the name of the Office application
    $officeApps = $OfficeAppDetection -split ","
    $officeApps = $officeApps | ForEach-Object { $_.Trim() }
    $officeApps = $officeApps | Where-Object { $_ -in @("Excel", "Word", "PowerPoint", "Access") }
    if ($officeApps.Count -eq 0) {
        Write-Host "No valid Office applications specified. Exiting script."
        exit 1
    }
}


# We need to open and close the Office applications before we can enable VBOM
Write-Host "Open and close Office applications"
. "$PSScriptRoot/scripts/Open-Close-Office.ps1" $officeApps
Write-Host "========================="


if ($TestFramework -ieq "rubberduck") {
    Write-Host "Install Rubberduck"
    . "$PSScriptRoot/scripts/Install-Rubberduck-VBA.ps1"
    Write-Host "========================="
} else {
    Write-Host "Test framework is not Rubberduck. Skipping installation."
}

# Enable VBOM for each Office application
Write-Host "Enabling VBOM for Office applications"
foreach ($app in $officeApps) {
    Write-Host "Enabling VBOM for $app"
    . "$PSScriptRoot/scripts/Enable-VBOM.ps1" $app
    Write-Host "========================="
}

# To get better screenshots we need to minimize the "Administrator" CMD window
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/scripts/utils/Minimize.ps1"


# Import scripts
. "$PSScriptRoot/scripts/Tests-Rubberduck-VBA.ps1" # Import the Rubberduck testing script
. "$PSScriptRoot/scripts/Clean-Up.ps1" # Import the Clean-Up.ps1 script

Minimize-Window "Administrator: C:\actions"
Write-Host "========================="

foreach ($folder in $folders) {

    Write-Host "▶️ Processing folder: $folder"
    $processedFolders++

    $fileExtension = $folder.Substring($folder.LastIndexOf('.') + 1)

    if ($OfficeAppDetection -ieq "automatic") {
        $app = Get-OfficeApp -FileExtension $fileExtension
    } elseif ($officeApps.Count -eq 1) {
        # Note that when an array has only one element, PowerShell will treat it as a single value
        $app = $officeApps
    } elseif ($officeApps.Count -gt 1) {
        Write-Host "Multiple Office applications specified. Please specify only one."
        exit 1
    } else {
        Write-Host "No valid Office applications specified. Exiting script."
        exit 1
    }

    Write-Host "Office application: $app"

    if ($app -eq "Access") {
        Write-Host "Access database detected. Adding to Access folders list..."
        $accessFolders += "${SourceDir}/${folder}"
        $hasAccessDatabase = $true
        Write-Host "Access is not supported in the main build process. Skipping build but tracking for separate processing..."
        continue
    }

    $ext = "zip"
    Write-Host "Create Zip file and rename it to Office document target"
    . "$PSScriptRoot/scripts/Zip-It.ps1" "${SourceDir}/${folder}"

    Write-Host "Copy and rename the file to the correct name"
    . "$PSScriptRoot/scripts/Rename-It.ps1" "${SourceDir}/${folder}" "$ext"

    Write-Host "Importing VBA code into Office document" 
    . "$PSScriptRoot/scripts/Build-VBA.ps1" "${SourceDir}/${folder}" "$app"
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Build-VBA.ps1 failed with exit code $LASTEXITCODE"
        exit $LASTEXITCODE
    } else {
        $successfulBuilds++
    }
  
    if ($TestFramework -ieq "rubberduck" -and $fileExtension -ne "ppam") {
        Write-Host "Running tests with Rubberduck"
        $rubberduckTestResult = Test-WithRubberduck -officeApp $officeApp
        if (-not $rubberduckTestResult) {
            Write-Host "Rubberduck tests were not completed successfully, but continuing with the script..."
        }
    } else {
        if ($fileExtension -eq "ppam") {
            Write-Host "Skipping tests for PowerPoint add-in (.ppam) files since Rubberduck can't run tests on them directly."
        } else {
            Write-Host "Test framework is not Rubberduck. Skipping tests."
        }
    }

    Write-Host "Cleaning up"
    CleanUp-OfficeApp -officeApp $officeApp

    Write-Host "========================="
}

# Output variables for GitHub Actions
Write-Host "Setting GitHub Actions outputs..."

# Write to GITHUB_OUTPUT file using the current recommended method
"processed-folders=$processedFolders" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
"successful-builds=$successfulBuilds" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
"office-apps=$($officeApps -join '|||')" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
"access-folders=$($accessFolders -join '|||')" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8
"has-access-database=$hasAccessDatabase" | Out-File -FilePath $env:GITHUB_OUTPUT -Append -Encoding utf8

Write-Host "Build process completed successfully!"
Write-Host "Processed folders: $processedFolders"
Write-Host "Successful builds: $successfulBuilds"
Write-Host "Office apps used: $($officeApps -join ' ||| ')"
Write-Host "Access folders found: $($accessFolders -join ' ||| ')"
Write-Host "Has Access database: $hasAccessDatabase"