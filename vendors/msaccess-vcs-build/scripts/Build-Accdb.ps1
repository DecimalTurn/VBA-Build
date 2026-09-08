#
# build accdb from source
#
param(
    [string]$SourceDir,
	[string]$TargetDir = "",
	[string]$FileName = "", # empty = name from vcs options
    [string]$VcsAddInPath = "" # empty = use default path (installed version)
   
)

# Check if the script is running under a Windows service account (SYSTEM, NETWORK SERVICE, LOCAL SERVICE)
$serviceAccounts = @('SYSTEM', 'NETWORK SERVICE', 'LOCAL SERVICE')
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
if ($serviceAccounts | Where-Object { $currentUser -match $_ }) {
    Write-Warning "Warning: This script is running under a Windows service account ($currentUser). Microsoft Access should not be executed as a service!"
}
else {
    Write-Host "Running script as user: $currentUser"
}

[string]$tempFileName = "VcsBuildTempApp"
[string]$accdbFileName = $tempFileName
if ($FileName -gt "") {
    $accdbFileName = $FileName
}

$curDir = $(Get-Location)
$accdbPath = "$curDir\$accdbFileName.accdb"

# open/create access file
$access = New-Object -ComObject Access.Application
$access.Visible = $true
if (-not (Test-Path $accdbPath)) {    
    $access.NewCurrentDatabase($accdbPath)
} 
else {
	$access.OpenCurrentDatabase($accdbPath)
}




[string]$addInProcessPath = ""
if ($VcsAddInPath -gt "") {
    $addInProcessPath = [System.IO.Path]::ChangeExtension($VcsAddInPath, "").TrimEnd('.')   
}
else {
    $appdata = $env:APPDATA
    $addInFolder = Join-Path $appdata "MSAccessVCS"
    $addInProcessPath = Join-Path $addInFolder "Version Control"
}

$addInPattern = "$addInProcessPath.accd[ae]"

if (-not (Test-Path $addInPattern)) {
    Write-Host "msaccess-vcs add-in not found: $addInPattern"
    Write-Host "Please install msaccess-vcs add-in first."
    exit 1
}

if (
    -not ([System.IO.Path]::IsPathRooted($SourceDir)) -or
    ($SourceDir -match "^[\\\/]") # "\source" or "/source"
) {
    $SourceDir = Join-Path -Path (Get-Location) -ChildPath $SourceDir.TrimStart('\','/','.')
}

Write-Host "Add-in path: $addInProcessPath"
Write-Host "Current path: $curDir"
Write-Host "Source: $SourceDir"
Write-Host "TargetDir: $TargetDir"
Write-Host ""

Write-Host "Start msaccess-vcs build " -NoNewline
$access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
Write-Host "." -NoNewline
$null = $access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnBuild", [ref] "$SourceDir")

# VCS Build close tempApp and reopen new accdb => check 2x for Forms.Count
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while (($access.Forms.Count -gt 0) -and ($stopwatch.Elapsed.TotalSeconds -lt 30)) {
    Start-Sleep -Seconds 2
    Write-Host "." -NoNewline
}
$stopwatch.Stop()
Start-Sleep -Seconds 3
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while (($access.Forms.Count -gt 0) -and ($stopwatch.Elapsed.TotalSeconds -lt 30)) {
    Start-Sleep -Seconds 2
    Write-Host "." -NoNewline
}
$stopwatch.Stop()
Write-Host " completed"

$builtFileName = $access.CurrentProject.Name
$builtFilePath = $access.CurrentProject.FullName

Start-Sleep -Seconds 1
Write-Host "Close Access " -NoNewline
$access.Quit(2)
Write-Host "." -NoNewline
Start-Sleep -Seconds 1
Write-Host "." -NoNewline
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
Write-Host "." -NoNewline
[GC]::WaitForPendingFinalizers()

# Kill any lingering Access processes - see issue #1 (support for msaccess-vcs-addin v5), thanks DecimalTurn
Get-Process -Name "MSACCESS" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 3

Write-Host " completed"
Write-Host ""

if ( ($builtFileName -gt "") -and ($builtFileName -ne "$tempFileName.accdb") ) {
	Write-Host "Built: $builtFileName ($builtFilePath)"
} else {
	Write-Host "Build failed"
    if ([string]::IsNullOrEmpty($builtFileName)) {
        Write-Host "   (builtFileName is empty)"
    }
    else {
        Write-Host "   $builtFileName"
    }
    if ([string]::IsNullOrEmpty($builtFilePath)) {
        Write-Host "   (builtFilePath is empty)"
    } 
    else {
	    Write-Host "  $builtFilePath"
    }
	exit 1
}


# Wait for a file to become openable with an exclusive lock, or time out.
# $access.Quit() can return before the MSACCESS process has fully exited and
# released its lock on the built database, which makes Copy-Item fail
# intermittently with "being used by another process".
function Wait-FileUnlocked {
    param(
        [string]$Path,
        [int]$TimeoutSeconds = 60
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Close()
            $stream.Dispose()
            return $true
        }
        catch {
            Start-Sleep -Milliseconds 500
        }
    }
    return $false
}

# copy file to TargetDir
if ([string]::IsNullOrEmpty($FileName)) {
    $FileName = $builtFileName
}

$targetFilePath = $builtFilePath
$builtFilePathDir = [System.IO.Path]::GetDirectoryName($builtFilePath)
if (($TargetDir -gt "") -and ($TargetDir -ne  $builtFilePathDir) ) {
	Write-Host "Copy accdb to $TargetDir"
	New-Item -Path $TargetDir -ItemType Directory -Force | Out-Null

    # Make sure the built database is no longer locked before copying it. If it is
    # still locked after the grace window, stop any lingering MSACCESS process as a
    # last resort (the CI runner should not have unrelated Access instances open).
    if (-not (Wait-FileUnlocked -Path $builtFilePath -TimeoutSeconds 60)) {
        Write-Warning "Built file still locked after 60s. Stopping lingering MSACCESS process(es)..."
        Get-Process -Name MSACCESS -ErrorAction SilentlyContinue | ForEach-Object {
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction Stop
                Write-Host "Stopped MSACCESS pid $($_.Id)"
            }
            catch {
                Write-Warning "Could not stop MSACCESS pid $($_.Id): $($_.Exception.Message)"
            }
        }
        $null = Wait-FileUnlocked -Path $builtFilePath -TimeoutSeconds 10
    }

    $copyAttempts = 0
    while ($true) {
        $copyAttempts++
        try {
            Copy-Item -Path $builtFilePath -Destination "$TargetDir\$FileName" -Force -ErrorAction Stop
            break
        }
        catch {
            if ($copyAttempts -ge 5) { throw }
            Write-Warning "Copy attempt $copyAttempts failed: $($_.Exception.Message). Retrying..."
            Start-Sleep -Seconds 2
        }
    }
	Write-Host ""
	$targetFilePath = "$TargetDir\$FileName"
} elseif ($FileName -ne $builtFileName) {
	Rename-Item -Path ".\$builtFileName" -NewName $FileName -Force
}

$tempFilePath = Join-Path -Path $curDir -ChildPath ([System.IO.Path]::ChangeExtension($tempFileName, "accdb"))
if (Test-Path $tempFilePath) {
    Remove-Item -Path $tempFilePath -Force  
}

return "$targetFilePath"
