param(
    [string]$SourceDir = '', # empty => use parameter SourceFile, don't use msaccess-vcs
    [string]$SourceFile = '', # empty => use msaccess-vcs with SourceDir else: only compile file $SourceFile
    [string]$TargetDir = '', # Folder for output file, default (empty): current folder 
    [string]$Compile = 'false', # Default to "false" if not specified
    [string]$AppConfigFile = '', # Default "" => don't change database properties etc.
    [string]$RunAccUnitTests = 'false', # path to msaccess-vcs add-in, empty = don't use msaccess-vcs
    [string]$vcsUrl = 'https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest', # empty = don't install msacess-vcs
    [string]$SetTrustedLocation = 'true', # set trusted location for current folder
    [string]$FileName = '' # empty .. use file name from vcs options
)


# Prepare parameters

if ([string]::IsNullOrEmpty($SourceDir) -and [string]::IsNullOrEmpty($SourceFile)) {
    Write-Error "SourceDir or SourceFile must be specified"
    exit 1  
}

[bool]$CompileBool = $false
if ($Compile -and $Compile.ToLower() -eq "true") {
    $CompileBool = $true
}

[bool]$SetTrustedLocationBool = $false
if ($SetTrustedLocation -and $SetTrustedLocation.ToLower() -eq "true") {
    $SetTrustedLocationBool = $true
}

[bool]$RunAccUnitTestBool = $false
if ($RunAccUnitTests -and $RunAccUnitTests.ToLower() -eq "true") {
    $RunAccUnitTestBool = $true
}

if ([string]::IsNullOrEmpty($SourceDir) ) {
    $vcsUrl = "" # don't install msaccess-vcs if SourceDir is not specified
}

if (-not [string]::IsNullOrEmpty($TargetDir)) {
    if (-not ([System.IO.Path]::IsPathRooted($TargetDir))) {
        $TargetDir = Join-Path -Path (Get-Location) -ChildPath $TargetDir.TrimStart('\','/','.')
    }
}
else {
    $TargetDir = (Get-Location)
}


# Prepare environment

$curDir = $(Get-Location)
$tempTrustedLocationName = "VCS-build-folder_"  + (Get-Date -Format "yyyyMMddHHmmss")
if ($SetTrustedLocationBool)
{
    Write-Host "Set trusted location: $curDir"
    & "$PSScriptRoot/scripts/Set-TrustedLocation.ps1" "$tempTrustedLocationName" "$curDir"
    Write-Host "-----"
}

[string]$vcsAddInPath = ""
if ($vcsUrl -gt "") {
	Write-Host "Install msaccess-vcs"
    $vcsTargetDir = $curDir.Path
	$vcsInstallData = & "$PSScriptRoot/scripts/Install-msaccess-vcs.ps1" -vcsUrl "${vcsUrl}" -TargetDir "$vcsTargetDir" -SetTrustedLocation $false
    
    if (-not $vcsInstallData.Success) {
        Write-Error "Failed to install msaccess-vcs add-in"
        exit 1
    }
    
    $vcsAddInPath = $vcsInstallData.AddInPath
	Write-Host "-----"
}


# Build accdb file
if (-not ([string]::IsNullOrEmpty($SourceDir))) {
    Write-Host "Build accdb - TargetDir: $TargetDir"
    $accdbPath = & "$PSScriptRoot/scripts/Build-Accdb.ps1" -SourceDir $SourceDir -TargetDir "${TargetDir}" -VcsAddInPath $vcsAddInPath -FileName "${FileName}"
    $accdbPath = "$accdbPath" # simple join if array
    $accdbPath = $accdbPath.Trim()
    if (([string]::IsNullOrEmpty($accdbPath)) -or -not (Test-Path $accdbPath)) {
        Write-Error "Failed to create accdb file"
        exit 1
    }
    Write-Host "Built file: $accdbPath"
} 
else { # use SourceFile
    if (-not ([System.IO.Path]::IsPathRooted($SourceFile))) {
        $SourceFile = Join-Path -Path (Get-Location) -ChildPath $SourceFile.TrimStart('\','/','.')
    }
    Write-Host "Use SourceFile: $SourceFile"
    if (-not (Test-Path $SourceFile)) {
        Write-Error "Source file not found: $SourceFile"
        exit 1
    }
    $accdbPath = $SourceFile
}   
Write-Host "-----"


$accFilePath = $accdbPath

# Prepare the application. Only module/reference removal requires an editable .accdb, so it
# runs before compiling. Procedures and database properties are applied after compiling,
# directly on the .accde, matching the last release and leaving the source .accdb unaltered.
# When not compiling there is no .accde, so every step runs on the .accdb.
if ($AppConfigFile -gt "") {
    $preStage = if ($CompileBool) { "PreCompile" } else { "All" }
    Write-Host "Prepare application ($preStage) from config file: $AppConfigFile"
    & "$PSScriptRoot/scripts/Prepare-Application.ps1" -AccessFile "$accdbPath" -ConfigFile "$AppConfigFile" -Stage $preStage
    Write-Host "-----"
}

if ($CompileBool) {
    Write-Host "compile accdb"
    $compileResult = & "$PSScriptRoot/scripts/Compile-Accdb.ps1" -SourceFile "$accdbPath"
    # Write-Host "accdb: $($result.AccdbPath)"
    # Write-Host "accde: $($result.AccdePath)"
    if (-not $compileResult.Success) {
        Write-Error "Failed to create ACCDE file"
        exit 1
    }
    $accFilePath = $compileResult.AccdePath
    Write-Host "-----"

    # Run procedures and apply database properties to the compiled .accde.
    if ($AppConfigFile -gt "") {
        Write-Host "Prepare application (PostCompile) from config file: $AppConfigFile"
        & "$PSScriptRoot/scripts/Prepare-Application.ps1" -AccessFile "$accFilePath" -ConfigFile "$AppConfigFile" -Stage PostCompile
        Write-Host "-----"
    }
}

# Run AccUnit tests
if ($RunAccUnitTestBool) {

    Write-Host "Run AccUnit tests"

    if (-not (Test-Path $accdbPath)) {
        Write-Error "Accdb file not found: $accdbPath"
        exit 1
    }
    $testAccdbPath = [System.IO.Path]::GetFileName($accdbPath)
    $testAccdbPath = Join-Path -Path (Get-Location) -ChildPath $testAccdbPath
    if (-not (Test-Path $testAccdbPath)) {
        Copy-Item -Path $accdbPath -Destination $testAccdbPath -Force
    }

    $testResult = & "$PSScriptRoot/scripts/Run-AccUnit-Tests.ps1" -AccdbPath "$testAccdbPath"
#copy test log to TargetDir
    $testLogFile = $testResult.LogFile
    if ($TargetDir -ne (Get-Location) -and (Test-Path $testLogFile)) {
        $targetTestLogFile = Join-Path -Path $TargetDir -ChildPath ([System.IO.Path]::GetFileName($testLogFile))
        Copy-Item -Path $testLogFile -Destination $targetTestLogFile -Force
        Write-Host "Test log file copied to: $targetTestLogFile"
    }

    if (-not $testResult.Success) {
        Write-Host "Tests failed" -ForegroundColor Red
        if (Test-Path $testLogFile) {
            Get-Content $testLogFile | Where-Object { $_ -match "\t(Failed|Error)\t" } | ForEach-Object { Write-Host $_ }
        }
        exit 1
    }
    Write-Host "-----"
}

# clean up
if ($SetTrustedLocationBool)
{
    Write-Host "Remove trusted location: $curDir"
    & "$PSScriptRoot/scripts/Remove-TrustedLocation.ps1" "$tempTrustedLocationName" 
    Write-Host "-----"
}

exit 0
