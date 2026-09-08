#
# install AccUnitLoader.accda
#
param(
    [string]$AccUnitUrl = "https://api.github.com/repos/AccessCodeLib/AccUnit/releases/latest",
    [string]$TargetRootDir = "", # empty = use current directory
    [bool]$SetTrustedLocation = $true # set trusted location for add-in folder
)

Write-Host "Download url: $AccUnitUrl"

$TargetOfficeApp = "Access"
. "$PSScriptRoot/Get-GitHubHeaders.ps1"
$headers = Get-GitHubHeaders
$release = Invoke-RestMethod -Uri $AccUnitUrl -Headers $headers

# zip url
$asset = $release.assets | Where-Object { $_.name -like "AccUnitLoader*$TargetOfficeApp.zip" } | Select-Object -First 1
$zipUrl = $asset.browser_download_url

# save as
$zipFile = "AccUnitLoader.zip"

# download file
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile

Write-Host "zip file downloaded from $zipUrl to $zipFile"

# extrat to local folder (don't use original MSAccessVCS folder)
if ([string]::IsNullOrEmpty($TargetRootDir)) {
    $TargetRootDir = (Get-Location).Path
}
else {
    if (-not ([System.IO.Path]::IsPathRooted($TargetRootDir))) {
        $TargetRootDir = Join-Path -Path (Get-Location) -ChildPath $TargetRootDir
        if ($TargetRootDir -match '[\\/][.][\\/]')
        {
            $TargetRootDir = $TargetRootDir -replace '[\\/][.][\\/]', '\'
        }
    }
}
$addInFolder = Join-Path -Path $TargetRootDir -ChildPath "AccUnit"
Expand-Archive -Path $zipFile -DestinationPath $addInFolder -Force

$addInFileName = "AccUnitLoader.accda"
[string]$addInPath = Join-Path $addInFolder $addInFileName

# check if add-in file exists
if (-not (Test-Path $addInPath)) {
    Write-Error "Add-in file not found: $addInPath"
    exit 1
}

#rename file to accdb (compile it to accda later)
[string]$accdbPath = [System.IO.Path]::ChangeExtension($addInPath, "accdb")

if (Test-Path $accdbPath) {
    Remove-Item -Path $accdbPath -Force
}
Rename-Item -Path $addInPath -NewName $accdbPath

$access = New-Object -ComObject Access.Application
$access.Visible = $true

$access.OpenCurrentDatabase($accdbPath)
$access.Visible = $true
$access.Run("CheckAccUnitTypeLibFile")
$access.Run("ExportAccUnitFiles")
$access.CloseCurrentDatabase()

$accessType = $access.GetType()
$result = $accessType.InvokeMember(
    "SysCmd",
    "InvokeMethod",
    $null,
    $access,
    @(603, $accdbPath, $addInPath)
)




$access.Quit(2)
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($access)
Remove-Variable access
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "AccUnitLoader installed: $addInPath"

if ($SetTrustedLocation)
{
    Write-Host "Set trusted location: $addInFolder"
    & "$PSScriptRoot/Set-TrustedLocation.ps1" "AccUnit-add-in-folder" "$addInFolder"
}

$result = [PSCustomObject]@{
    AddInPath = "$addInPath"
    AddInFolder = "$addInFolder"
    AddInFileName = "$addInFileName"
    Success = $true
}
return $result