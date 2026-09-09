#
# install msaccess-vcs
#
param(
    [string]$vcsUrl = "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest",
    [string]$TargetDir = "", # empty = use current directory
    [bool]$SetTrustedLocation = $true # set trusted location for add-in folder
)
# URL
# $vcsUrl = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/latest"
# $vcsUrl = "https://api.github.com/repos/josef-poetzl/msaccess-vcs-addin/releases/latest"

Write-Host "Download url: $vcsUrl"

. "$PSScriptRoot/Get-GitHubHeaders.ps1"
$headers = Get-GitHubHeaders
$release = Invoke-RestMethod -Uri $vcsUrl -Headers $headers

# zip url
$asset = $release.assets | Where-Object { $_.name -like "Version*.zip" } | Select-Object -First 1
$zipUrl = $asset.browser_download_url

# save as
$zipFile = "msaccess-vcs.zip"

# download file
Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile

Write-Host "zip file downloaded from $zipUrl to $zipFile"

# extrat to local folder (don't use original MSAccessVCS folder)
if ([string]::IsNullOrEmpty($TargetDir)) {
    $TargetDir = (Get-Location).Path
}
else {
    if (-not ([System.IO.Path]::IsPathRooted($TargetDir))) {
        $TargetDir = Join-Path -Path (Get-Location) -ChildPath $TargetDir.TrimStart('\','/','.')
    }
}
$addInFolder = Join-Path -Path $TargetDir -ChildPath "MSAccessVCS"
Expand-Archive -Path $zipFile -DestinationPath $addInFolder -Force

$addInFileName = "Version Control.accda"
$addInPath = Join-Path $addInFolder $addInFileName

Write-Host "msaccess-vcs installed: $addInPath"

if ($SetTrustedLocation)
{
    Write-Host "Set trusted location: $addInFolder"
    & "$PSScriptRoot/Set-TrustedLocation.ps1" "VCS-add-in-folder" "$addInFolder"
}

$result = [PSCustomObject]@{
    AddInPath = "$addInPath"
    AddInFolder = "$addInFolder"
    AddInFileName = "$addInFileName"
    Success = $true
}
return $result