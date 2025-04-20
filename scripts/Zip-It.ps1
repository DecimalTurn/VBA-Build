# This script uses 7-Zip to compress files and folders in the ${folderName}/XMLsource directory into a zip file.

# Read the name of the folder from the argument passed to the script
$folderName = $args[0]
if (-not $folderName) {
    Write-Host "Error: No folder name specified. Usage: Zip-It.ps1 <FolderName>"
    exit 1
}

$sourceDir = $folderName.Substring(0, $folderName.LastIndexOf('/'))

$filNameWithExtension = $folderName.Substring($folderName.LastIndexOf('/') + 1)
$fileName = $filNameWithExtension.Substring(0, $filNameWithExtension.LastIndexOf('.'))
$fileExtension = $filNameWithExtension.Substring($filNameWithExtension.LastIndexOf('.') + 1)

Write-Host "Staring the compression process..."

$currentDir = Get-Location
Write-Host "Current directory: $currentDir"

# Define the source folder and the output zip file
$sourceFolder = Join-Path -Path $currentDir -ChildPath "$folderName/XMLsource/"
$outputZipFile = Join-Path -Path $currentDir -ChildPath "$folderName/Skeleton/$fileName.zip"

# Path to the 7-Zip executable
$sevenZipPath = "7z"  # Assumes 7-Zip is in the system PATH. Adjust if necessary.

# Check if the source folder exists
if (-not (Test-Path $sourceFolder)) {
    Write-Host "Source folder not found: $sourceFolder"
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

if (-not (Test-Path $sourceFolder)) {
    Write-Host "Source folder not found: $sourceFolder"
    exit 1
}

# Ensure the destination directory exists
$outputDir = Split-Path -Path $outputZipFile
if (-not (Test-Path $outputDir)) {
    Write-Host "Creating output directory: $outputDir"
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$absoluteSourceFolder = Resolve-Path -Path $sourceFolder
if (-not (Test-Path $absoluteSourceFolder)) {
    Write-Host "Error: Source folder not found: $absoluteSourceFolder"
    exit 1
}

$absoluteDestinationFolder = Resolve-Path -Path $outputDir

# Change the working directory to the source folder
Write-Host "Changing directory to $absoluteSourceFolder..."
cd $absoluteSourceFolder
# if ($LASTEXITCODE -ne 0) {
#     Write-Host "Error: Failed to change directory to $absoluteSourceFolder"
#     exit $LASTEXITCODE
# }

Write-Host "Current directory after change: $(Get-Location)"

# Compress the files and folders using 7-Zip
Write-Host "Compressing files in $sourceFolder to $absoluteDestinationFolder..."
& $sevenZipPath a -tzip "$absoluteDestinationFolder/$fileName.zip" "*" | Out-Null

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

Write-Host "Compression completed successfully. Zip file created at: $absoluteDestinationFolder"
