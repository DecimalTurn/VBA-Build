# Get the source directory from command line argument or use default "src"
param(
    [string]$SourceDir = "src",
    [string]$TestFramework = "none", # Default to "none" if not specified
    [string]$OfficeAppDetection = "automatic", # Default to "automatic" if not specified
    [string]$AccessVcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/tags/v5.0.1", # msaccess-vcs-addin release URL for Access builds
    [string]$AccessVcsSha = "", # Expected SHA256 digest of the add-in asset (empty = skip verification)
    [string]$AccessVcsConfig = "", # Application config file for Access builds (empty = no prepare step)
    [string]$AccessVcsCompile = "" # Force compiling all Access builds to ACCDE (empty = compile only folders named *.accde)
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

# ---------------------------------------------------------------------------
# Access builds via the vendored msaccess-vcs-build
#
# Access folders are detected during the main loop below, collected into
# $accessFolders, and built afterwards by Invoke-AccessBuilds. Building uses
# ONLY committed files under vendors/msaccess-vcs-build (never the submodule);
# that tree is refreshed by scripts/Bundle-Vendors.ps1.
# ---------------------------------------------------------------------------

$VendorBuildDir = Join-Path $PSScriptRoot 'vendors/msaccess-vcs-build'

function Assert-VendoredVcsBuild {
    $required = @(
        'Build.ps1',
        'scripts/Build-Accdb.ps1',
        'scripts/Install-msaccess-vcs.ps1',
        'scripts/Get-GitHubHeaders.ps1'
    )
    $missing = $required | Where-Object { -not (Test-Path (Join-Path $VendorBuildDir $_)) }
    if ($missing) {
        Write-Error @"
Vendored msaccess-vcs-build files are missing from '$VendorBuildDir':
  $($missing -join "`n  ")

Generate them by running (from the repo root, with the submodule initialized):
  git submodule update --init --recursive
  pwsh -File scripts/Bundle-Vendors.ps1
"@
        exit 1
    }
}

function Assert-VcsAddinDigest {
    param(
        [string]$ReleaseUrl,
        [string]$ExpectedSha
    )

    if ([string]::IsNullOrEmpty($ExpectedSha)) {
        Write-Host "No access-vcs-sha provided - skipping add-in digest verification."
        return
    }

    . "$VendorBuildDir/scripts/Get-GitHubHeaders.ps1"
    Write-Host "Verifying add-in digest against: $ReleaseUrl"

    $release = Invoke-RestMethod -Uri $ReleaseUrl -Headers (Get-GitHubHeaders)

    # Select the SAME asset the installer downloads (Version*.zip), not
    # .assets[0], so we verify exactly the file that will be installed.
    $asset = $release.assets | Where-Object { $_.name -like 'Version*.zip' } |
        Select-Object -First 1
    if (-not $asset -or -not $asset.digest) {
        throw "No sha256 digest found for the Version*.zip asset. Update the pin."
    }

    $actual   = ($asset.digest -replace '^sha256:', '').ToLower()
    $expected = $ExpectedSha.Trim().ToLower()
    Write-Host "Expected SHA256: $expected"
    Write-Host "Actual SHA256:   $actual"
    if ($actual -ne $expected) {
        throw "SHA256 digest mismatch for '$($asset.name)'. Expected $expected, got $actual."
    }
    Write-Host "SHA256 digest verification passed."
}

function Invoke-AccessBuilds {
    param([string[]]$AccessFolders)

    Assert-VendoredVcsBuild

    $ErrorActionPreference = 'Stop'

    # --- digest-verify the add-in release ONCE (auth via GH_TOKEN) --------
    Assert-VcsAddinDigest -ReleaseUrl $AccessVcsUrl -ExpectedSha $AccessVcsSha

    # --- install the add-in ONCE into the APPDATA default location ---------
    # Install-msaccess-vcs.ps1 expands to $TargetDir\MSAccessVCS. TargetDir =
    # $env:APPDATA puts it exactly where Build-Accdb.ps1 looks when vcsUrl is
    # empty, so per-folder Build.ps1 calls with -vcsUrl "" reuse it without
    # re-downloading. Run from a temp dir so msaccess-vcs.zip doesn't pollute
    # the workspace.
    $tempWorkDir = Join-Path $env:TEMP "vba-build-$PID"
    New-Item -ItemType Directory -Force -Path $tempWorkDir | Out-Null
    Push-Location $tempWorkDir
    try {
        $install = & "$VendorBuildDir/scripts/Install-msaccess-vcs.ps1" `
            -vcsUrl $AccessVcsUrl -TargetDir $env:APPDATA -SetTrustedLocation $true
    }
    finally {
        Pop-Location
    }
    if (-not $install -or -not $install.AddInPath) {
        throw "Failed to install the msaccess-vcs add-in (see output above)."
    }
    Write-Host "msaccess-vcs add-in installed: $($install.AddInPath)"

    # --- per folder: fresh child pwsh running the vendored Build.ps1 -------
    # Phase 1 uses a separate pwsh per folder for isolation while this new flow
    # is validated. Phase 2 (follow-up) drops the child pwsh and calls Build.ps1
    # in-process for speed.
    foreach ($folder in $AccessFolders) {
        # Replicate the old bash parse-step: per-folder compile flag.
        $compile = if (-not [string]::IsNullOrEmpty($AccessVcsCompile)) { $AccessVcsCompile }
                   elseif ($folder -match '\.accde(\.src)?$')             { 'true' }
                   else                                                   { 'false' }

        Write-Host "Building Access database: $folder (compile=$compile)"
        & pwsh -NoProfile -File "$VendorBuildDir/Build.ps1" `
            -SourceDir      $folder `
            -TargetDir      "$SourceDir/out" `
            -Compile        $compile `
            -AppConfigFile  $AccessVcsConfig `
            -vcsUrl         ""
        if ($LASTEXITCODE -ne 0) {
            throw "Access build failed for $folder (exit code $LASTEXITCODE)."
        }
    }
}

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

function Get-OfficeAppFromFolder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FolderName,
        [string]$FolderPath = ""
    )

    # An Access source is any msaccess-vcs export. msaccess-vcs writes a
    # vcs-options.json marker at the export root regardless of the folder naming
    # convention used (<Name>.accdb, <Name>.accdb.src, or a custom ExportFolder), so
    # detect it by content rather than relying on the folder name.
    if ($FolderPath -ne "" -and (Test-Path -LiteralPath (Join-Path $FolderPath 'vcs-options.json'))) {
        return "Access"
    }

    $FileExtension = $FolderName.Substring($FolderName.LastIndexOf('.') + 1)
    return Get-OfficeApp -FileExtension $FileExtension
}

if ($OfficeAppDetection -ieq "automatic") {

    Write-Host "Automatic detection of Office applications based on file extensions"

    # Create a list of Office applications that are needed based on the source folders
    foreach ($folder in $folders) {
        $app = Get-OfficeAppFromFolder -FolderName $folder -FolderPath (Join-Path $SourceDir $folder)

        if ($app) {
            if ($officeApps -notcontains $app) {
                $officeApps += $app
            }
        } else {
            Write-Host "Unknown file extension: $folder. Skipping..."
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

# VBA runtime setup (Office installation, app preparation and security settings)
# is handled by the setup-vba action before this script runs.
if ($TestFramework -ieq "rubberduck") {
    Write-Host "Install Rubberduck"
    . "$PSScriptRoot/scripts/Install-Rubberduck-VBA.ps1"
    Write-Host "========================="
} else {
    Write-Host "Test framework is not Rubberduck. Skipping installation."
}

# To get better screenshots we need to minimize the "Administrator" CMD window
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/scripts/utils/Minimize.ps1"


# Import scripts
. "$PSScriptRoot/scripts/Tests-Rubberduck-VBA.ps1" # Import the Rubberduck testing script
. "$PSScriptRoot/scripts/Clean-Up.ps1" # Import the Clean-Up.ps1 script

Minimize-Window "Administrator: C:\actions"
Minimize-Window "C:\ProgramData\GitHub\HostedComputeAgent\hosted-compute-agent"


Write-Host "========================="

foreach ($folder in $folders) {

    Write-Host "▶️ Processing folder: $folder"
    $processedFolders++

    $fileExtension = $folder.Substring($folder.LastIndexOf('.') + 1)

    if ($OfficeAppDetection -ieq "automatic") {
        $app = Get-OfficeAppFromFolder -FolderName $folder -FolderPath (Join-Path $SourceDir $folder)
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

# Build any detected Access databases using the vendored msaccess-vcs-build.
if ($hasAccessDatabase) {
    Invoke-AccessBuilds -AccessFolders $accessFolders
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