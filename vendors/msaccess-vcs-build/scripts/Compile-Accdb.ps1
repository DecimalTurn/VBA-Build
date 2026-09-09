#
# create accde
#
param(
    [string]$SourceFile,
    [string]$DestFile = "" # empty .. use SourceFile with accde extension
)

[string]$SourceFilePath = $SourceFile
[string]$DestFilePath = ""
if (
    -not ([System.IO.Path]::IsPathRooted($SourceFilePath)) -or
    ($SourceFilePath -match "^[\\\/]") # "\source" or "/source"
) {
    $SourceFilePath = Join-Path -Path (Get-Location) -ChildPath $SourceFilePath.TrimStart('\','/','.')
}

if ($DestFile -gt "") {
	if (
		-not ([System.IO.Path]::IsPathRooted($DestFile)) -or
		($DestFile -match "^[\\\/]") # "\source" or "/source"
	) {
		$DestFilePath = Join-Path -Path (Get-Location) -ChildPath $DestFile.TrimStart('\','/','.')
	}
	else {
		$DestFilePath = $DestFile
	}
} else {
	$DestFilePath = [System.IO.Path]::ChangeExtension($SourceFilePath, "accde")
}

Write-Host "accdb: $SourceFilePath"
Write-Host "accde: $DestFilePath"

if (Test-Path $DestFilePath) {
    Remove-Item $DestFilePath -Force
}


$access = New-Object -ComObject Access.Application
# $access.Visible = $true

$accessType = $access.GetType()
$result = $accessType.InvokeMember(
    "SysCmd",
    "InvokeMethod",
    $null,
    $access,
    @(603, $SourceFilePath, $DestFilePath)
)

$timeout = 10
$success = $false

for ($i = 0; $i -lt $timeout; $i++) {
    if (Test-Path $DestFilePath) {
        $success = $true
        break
    }
    Start-Sleep -Seconds 1
}

if ($success) {
    Write-Host "accde successfully created."
} else {
    Write-Error "Error: accde file was not created."
}

$access.Quit(2)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

$result = [PSCustomObject]@{
    AccdbPath = "$SourceFilePath"
    AccdePath = "$DestFilePath"
    Success   = $success
}

return $result