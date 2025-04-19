# This script uses 7-Zip to compress files and directory in the ${docDirPath}/XMLsource directory into a zip file.


$docDirPath = $args[0]
if (-not $docDirPath) {
    Write-Host "Error: No docDirPath specified. Usage: Zip-It.ps1 <docDirPath>"
    exit 1
}

$currentDir = (Get-Location).Path + "/"
Write-Host "Current directory: $currentDir"
$docDirPath = GetAbsPath -path $docDirPath -basePath $currentDir
Write-Host "Using docDirPath: $docDirPath"

$docDirPath = NormalizeDirPath $docDirPath

$fileName = GetDirName $docDirPath
$fileNameNoExt = $fileName.Substring(0, $fileName.LastIndexOf('.'))
$fileExtension = $fileName.Substring($fileName.LastIndexOf('.') + 1)

Write-Host "Staring the compression process..."

# Define the source directory and the output zip file
$xmlDir = $docDirPath + "XMLsource/"
$outputZipFile = $docDirPath + "Skeleton/$fileNameNoExt.zip"

# Path to the 7-Zip executable
$sevenZipPath = "7z"  # Assumes 7-Zip is in the system PATH. Adjust if necessary.

# Check if the source directory exists
if (-not (Test-Path $xmlDir)) {
    Write-Host "Source dir not found: $xmlDir"
    ls
    exit 1
}

# Ensure the destination directory exists
$outputDir = Split-Path -Path $outputZipFile
Write-Host "Output directory: $outputDir"

if (-not (Test-Path $outputDir)) {
    Write-Host "Creating output directory: $outputDir"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Ensure the destination directory exists
$outputDir = Split-Path -Path $outputZipFile
if (-not (Test-Path $outputDir)) {
    Write-Host "Creating output directory: $outputDir"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# Change the working directory to the source folder
Write-Host "Changing directory to $outputDir..."
Set-Location -Path $outputDir
Write-Host "Current directory after change: $(Get-Location)"

# Compress the files and directory using 7-Zip
Write-Host "Compressing files in $xmlDir to $outputDir..."
& $sevenZipPath a -tzip "${outputDir}${fileNameNoExt}.zip" "*" | Out-Null

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

Write-Host "Compression completed successfully. Zip file created at: $outputDir"
