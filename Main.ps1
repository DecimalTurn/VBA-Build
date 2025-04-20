# Get the source directory from command line argument or use default "src"

param(
    [string]$SourceDir = "src"
)

# Start the main timer
$mainTimer = [System.Diagnostics.Stopwatch]::StartNew()
$stepTimer = [System.Diagnostics.Stopwatch]::StartNew()

Set-Variable -Name MainScriptDir -Value "$PSScriptRoot" -Option Constant

# Import utility functions
. "$MainScriptDir/scripts/utils/Invoke.ps1" # Function to invoke a script with a timeout
. "$MainScriptDir/scripts/utils/Minimize.ps1" # To get better screenshots we need to minimize the "Administrator" CMD window
. "$MainScriptDir/scripts/utils/Path.ps1"
. "$MainScriptDir/scripts/utils/Screenshot.ps1"
. "$MainScriptDir/scripts/utils/TimedMessage.ps1"

Write-TimedMessage "Current directory: $(pwd)" -StartNewStep
# This directory contains the directories associated with the office documents
$mainSourceDir = NormalizeDirPath ((Resolve-Path $SourceDir).Path)
Write-Host "Main Source Dir: $mainSourceDir"

# Check if the source dir was resolved correctly
if (-not (Test-Path $mainSourceDir)) {
    Write-Host "🔴 Error: Source directory not found: $mainSourceDir"
    exit 1
}

# Read name of the folders under the specified source directory into an array
$fileDirNames = Get-ChildItem -Path "$mainSourceDir" -Directory | Select-Object -ExpandProperty Name
Write-TimedMessage "Folders in ${mainSourceDir}: $fileDirNames" -StartNewStep

# Check if the folders array is empty
if ($fileDirNames.Count -eq 0) {
    Write-TimedMessage "No folders found in ${mainSourceDir}. Exiting script." -StartNewStep
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
Write-TimedMessage "Identifying required Office applications" -StartNewStep
foreach ($fileDirName in $fileDirNames) {
    $fileExt = GetFileExtension $fileDirName
    $app = Get-OfficeApp -FileExtension $fileExt
    
    if ($app) {
        if ($officeApps -notcontains $app) {
            $officeApps += $app
        }
    } else {
        Write-TimedMessage "Unknown file extension: $fileExt. Skipping..."
        continue
    }
}
Write-TimedMessage "Required Office applications: $officeApps"

# We need to open and close the Office applications before we can enable VBOM
Write-TimedMessage "Open and close Office applications" -StartNewStep
. "$PSScriptRoot/scripts/Open-Close-Office.ps1" $officeApps
Write-TimedMessage "Completed opening and closing Office applications"
Take-Screenshot -OutputPath "${mainSourceDir}screenshots/AllOfficeAppClosed_{{timestamp}}.png"
Write-Host "========================="

# Enable VBOM for each Office application
foreach ($app in $officeApps) {
    Write-TimedMessage "Enabling VBOM for $app" -StartNewStep
    . "$PSScriptRoot/scripts/Enable-VBOM.ps1" $app
    Write-TimedMessage "VBOM enabled for $app"
    Write-Host "========================="
}


Write-TimedMessage "Minimizing Administrator window" -StartNewStep
Minimize-Window "Administrator: C:\actions"
Write-TimedMessage "Window minimized"

foreach ($fileDirName in $fileDirNames) {
    Write-Host "========================="
    # Reset the source directory to the original value

    Write-TimedMessage "Processing folder: $fileDirName" -StartNewStep
    $fileExt = $folder.Substring($folder.LastIndexOf('.') + 1)
    Write-Host "File extension: $fileExt"
    $app = Get-OfficeApp -FileExtension $fileExt

    $ext = ""
    if ($app -ne "Access") {
        $ext = "zip"
        Write-TimedMessage "Creating Zip file and renaming to Office document target with path ${mainSourceDir}${fileDirName}" -StartNewStep
        & "$PSScriptRoot/scripts/Zip-It.ps1" "${mainSourceDir}${fileDirName}"
        Write-TimedMessage "Zip file created"
    }
    else {
        $ext = "accdb"
        Write-TimedMessage "Copying directory and content to Skeleton directory" -StartNewStep
        Copy-Item -Path "${mainSourceDir}${fileDirName}/DBSource" -Destination "${mainSourceDir}${fileDirName}/Skeleton" -Recurse -Force
        Write-TimedMessage "Directory copied"
    }

    Write-TimedMessage "Copying and renaming file to correct name" -StartNewStep
    & "$PSScriptRoot/scripts/Rename-It.ps1" "${mainSourceDir}${folder}" "$ext"
    Write-TimedMessage "File renamed"

    Write-TimedMessage "Importing VBA code into Office document" -StartNewStep
    # Replace the direct Build-VBA call with the timeout version
    $buildVbaScriptPath = "$PSScriptRoot/scripts/Build-VBA.ps1"
    $success = Invoke-ScriptWithTimeout -ScriptPath $buildVbaScriptPath -Arguments @("${mainSourceDir}${folder}", "$app") -TimeoutSeconds 30

    Write-Host "mainSourceDir: ${mainSourceDir}"

    if (-not $success) {
        Write-TimedMessage "🔴 Build-VBA.ps1 execution timed out or failed for ${folder}. Continuing with next file..."
        
        $screenshotDir = ${mainSourceDir} + "screenshots/"
        if (-not (Test-Path $screenshotDir)) {
            New-Item -ItemType Directory -Path $screenshotDir -Force | Out-Null
            Write-Host "Created screenshot directory: $screenshotDir"
        }

        Take-Screenshot -OutputPath "${screenshotDir}${app}_{{timestamp}}.png"
    }

    Write-TimedMessage "🟢 Completed processing folder: $folder"
}