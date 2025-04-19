# This script uses 7-Zip to compress files and directory in the ${docDirPath}/XMLsource directory into a zip file.


$dirPath = $args[0]
if (-not $dirPath) {
    Write-Host "Error: No dirPath specified. Usage: Zip-It.ps1 <dirPath>"
    exit 1
}
Write-Host "dirPath: $dirPath"

# Import utility functions
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Path.ps1"

$currentDir = (Get-Location).Path + "/"
Write-Host "Current directory: $currentDir"
$docDirPath = GetAbsPath -path $dirPath -basePath $currentDir
Write-Host "Using docDirPath: $docDirPath"

$docDirPath = NormalizeDirPath $docDirPath

$fileName = GetDirName $docDirPath
$fileNameNoExt = $fileName.Substring(0, $fileName.LastIndexOf('.'))
$fileExtension = $fileName.Substring($fileName.LastIndexOf('.') + 1)

Write-Host "Staring the compression process..."

# Define the source directory and the output zip file
$xmlDir = $docDirPath + "XMLsource/"
$outputDir = $docDirPath + "Skeleton/"
$outputZipFile = $outputDir + "$fileNameNoExt.zip"

# Path to the 7-Zip executable
$sevenZipPath = "7z"  # Assumes 7-Zip is in the system PATH. Adjust if necessary.

# Check if the source directory exists
if (-not (Test-Path $xmlDir)) {
    Write-Host "Source dir not found: $xmlDir"
    ls
    exit 1
}

# Ensure the destination directory exists
Write-Host "Output directory: $outputDir"

if (-not (Test-Path $outputDir)) {
    Write-Host "Creating output directory: $outputDir"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Change the working directory to the source folder
Write-Host "Changing directory to $xmlDir"
Set-Location -Path $xmlDir
Write-Host "Current directory after change: $(Get-Location)"

# Compress the files and directory using 7-Zip
Write-Host "Compressing files in $xmlDir to $outputDir..."
& $sevenZipPath a -tzip "$outputZipFile" "*" | Out-Null

# Check if the compression was successful using $LASTEXITCODE
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error: Compression failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

# Restore the original working directory
Set-Location -Path $currentDir
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error: Failed to restore directory to $currentDir"
    exit $LASTEXITCODE
}

if (-not (Test-Path $outputZipFile)) {
    Write-Host "Error: Zip file not found after compression: $outputZipFile"
    exit 1
}

Write-Host "Compression completed successfully. Zip file created at: $outputDir"
