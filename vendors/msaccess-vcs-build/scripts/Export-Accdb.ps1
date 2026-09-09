#
# export from accdb
#
param(
    [string]$FileName, 
    [bool]$ForceExportAll = $false,
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


# Check/build full file path
$curDir = $(Get-Location)
$accdbPath = $FileName
if (-not ([System.IO.Path]::IsPathRooted($accdbPath))) {
    $accdbPath = Join-Path -Path $curDir -ChildPath $FileName
    $accdbPath = [System.IO.Path]::GetFullPath($accdbPath)
}

if (-not (Test-Path $accdbPath)) {    
    Write-Error "Error: The specified Access database file does not exist: $accdbPath"
    Exit 1  
} 


# check msaccess-vcs add-in
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

Write-Host "Add-in: $addInProcessPath"
Write-Host "file: $accdbPath"


## open access file

# open access
$access = New-Object -ComObject Access.Application
$access.Visible = $true

# Define Win32 API for keyboard events
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
    public const int VK_SHIFT = 0x10;
    public const int KEYEVENTF_KEYDOWN = 0x0000;
    public const int KEYEVENTF_KEYUP = 0x0002;
}
"@

# Hold down shift key
[Win32]::keybd_event([Win32]::VK_SHIFT, 0, [Win32]::KEYEVENTF_KEYDOWN, [UIntPtr]::Zero)
Start-Sleep -Milliseconds 100

# open access file
$access.OpenCurrentDatabase($accdbPath)

Start-Sleep -Milliseconds 100
# Release shift key
[Win32]::keybd_event([Win32]::VK_SHIFT, 0, [Win32]::KEYEVENTF_KEYUP, [UIntPtr]::Zero)


## export from accdb
Write-Host "Start msaccess-vcs export " -NoNewline
$access.Run("$addInProcessPath.SetInteractionMode", [ref] 1)
Write-Host "." -NoNewline

## TODO: read return value from Export function (if error level is implemented)
if ($ForceExportAll) {
    $null = $access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnExport", [ref] "True")
} else {
    $null = $access.Run("$addInProcessPath.HandleRibbonCommand", [ref] "btnExport")
}       
Write-Host "." -NoNewline
Start-Sleep -Seconds 1
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
while (($access.Forms.Count -gt 0) -and ($stopwatch.Elapsed.TotalSeconds -lt 10)) {
    Start-Sleep -Seconds 2
    Write-Host "." -NoNewline
}
$stopwatch.Stop()
if ($access.Forms.Count -gt 0)
{
    Write-Host "  failed"
    #don't close access to allow user to see error message
    Exit 1 
}
else
{
    Write-Host " completed"
}

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
Write-Host " completed"
Write-Host ""

Exit 0