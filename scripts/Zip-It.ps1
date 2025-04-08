# This script uses 7-Zip to compress files and folders in the src/XMLsource directory into a zip file named skeleton.zip.

Write-Host "Staring the compression process..."

$currentDir = Get-Location
Write-Host "Current directory: $currentDir"

# Define the source folder and the output zip file
$sourceFolder = Join-Path -Path $currentDir -ChildPath "src/skeleton.xlsm/XMLsource/"
$outputZipFile = Join-Path -Path $currentDir -ChildPath "src/skeleton.xlsm/XMLoutput/skeleton.zip"

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
& $sevenZipPath a -tzip "$absoluteDestinationFolder/skeleton.zip" "*" | Out-Null

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


# Create a copy of the zip file in the src/skeleton.xlsm/XMLoutput folder at the /src level
$copySource = "src/skeleton.xlsm/XMLoutput/skeleton.zip"
$renameDestination = "./skeleton.xlsm"

# Delete the destination file if it exists
if (Test-Path $renameDestination) {
    Write-Host "Deleting existing file: $renameDestination"
    Remove-Item -Path $renameDestination -Force
}

# Copy and rename the file in one step
Write-Host "Copying and renaming $copySource to $renameDestination..."
Copy-Item -Path $copySource -Destination $renameDestination -Force

# Verify if the file exists after the copy
if (-not (Test-Path $renameDestination)) {
    Write-Host "Error: File not found after copy: $renameDestination"
    exit 1
}

Write-Host "File successfully copied and renamed to: $renameDestination"