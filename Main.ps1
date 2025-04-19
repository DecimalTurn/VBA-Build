# Get the source directory from command line argument or use default "src"

param(
    [string]$SourceDir = "src"
)

# Import utility functions
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/scripts/utils/Minimize.ps1" # To get better screenshots we need to minimize the "Administrator" CMD window
. "$scriptPath/scripts/utils/TimedMessage.ps1"
. "$scriptPath/scripts/utils/Invoke.ps1" # Function to invoke a script with a timeout
. "$scriptPath/scripts/utils/Screenshot.ps1"

# Start the main timer
$mainTimer = [System.Diagnostics.Stopwatch]::StartNew()
$stepTimer = [System.Diagnostics.Stopwatch]::StartNew()

Write-TimedMessage "Current directory: $(pwd)" -StartNewStep
Write-TimedMessage "Using source directory: $SourceDir"

# Read name of the folders under the specified source directory into an array
$folders = Get-ChildItem -Path "$PSScriptRoot/$SourceDir" -Directory | Select-Object -ExpandProperty Name
Write-TimedMessage "Folders in ${SourceDir}: $folders" -StartNewStep

# Check if the folders array is empty
if ($folders.Count -eq 0) {
    Write-TimedMessage "No folders found in ${SourceDir}. Exiting script." -StartNewStep
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
foreach ($folder in $folders) {
    $FileExtension = $folder.Substring($folder.LastIndexOf('.') + 1)
    $app = Get-OfficeApp -FileExtension $FileExtension
    
    if ($app) {
        if ($officeApps -notcontains $app) {
            $officeApps += $app
        }
    } else {
        Write-TimedMessage "Unknown file extension: $FileExtension. Skipping..."
        continue
    }
}
Write-TimedMessage "Required Office applications: $officeApps"

# We need to open and close the Office applications before we can enable VBOM
Write-TimedMessage "Open and close Office applications" -StartNewStep
. "$PSScriptRoot/scripts/Open-Close-Office.ps1" $officeApps
Write-TimedMessage "Completed opening and closing Office applications"
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



foreach ($folder in $folders) {
    Write-TimedMessage "Processing folder: $folder" -StartNewStep
    $FileExtension = $folder.Substring($folder.LastIndexOf('.') + 1)
    $app = Get-OfficeApp -FileExtension $FileExtension

    $ext = ""
    if ($app -ne "Access") {
        $ext = "zip"
        Write-TimedMessage "Creating Zip file and renaming to Office document target" -StartNewStep
        . "$PSScriptRoot/scripts/Zip-It.ps1" "${SourceDir}/${folder}"
        Write-TimedMessage "Zip file created"
    }
    else {
        $ext = "accdb"
        Write-TimedMessage "Copying folder and content to Skeleton folder" -StartNewStep
        Copy-Item -Path "${SourceDir}/${folder}/DBSource" -Destination "${SourceDir}/${folder}/Skeleton" -Recurse -Force
        Write-TimedMessage "Folder copied"
    }

    Write-TimedMessage "Copying and renaming file to correct name" -StartNewStep
    . "$PSScriptRoot/scripts/Rename-It.ps1" "${SourceDir}/${folder}" "$ext"
    Write-TimedMessage "File renamed"

    Write-TimedMessage "Importing VBA code into Office document" -StartNewStep
    # Replace the direct Build-VBA call with the timeout version
    $buildVbaScriptPath = "$PSScriptRoot/scripts/Build-VBA.ps1"
    $success = Invoke-ScriptWithTimeout -ScriptPath $buildVbaScriptPath -Arguments @("${SourceDir}/${folder}", "$app") -TimeoutSeconds 300

    if (-not $success) {
        Write-TimedMessage "🔴 Build-VBA.ps1 execution timed out or failed for ${folder}. Continuing with next file..."
        
        $screenshotDir = ${SourceDir} + "screenshots/"
        if (-not (Test-Path $screenshotDir)) {
            New-Item -ItemType Directory -Path $screenshotDir -Force | Out-Null
            Write-Host "Created screenshot directory: $screenshotDir"
        }

        Take-Screenshot -OutputPath "${screenshotDir}${app}_{{timestamp}}.png"
    }

    Write-TimedMessage "Completed processing folder: $folder"
    Write-Host "========================="
}