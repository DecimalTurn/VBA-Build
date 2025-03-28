# This is a a powershell script that used 7zip to compress files and folders
# The files and folders to compress are located in the src/XMLSource folder
# It contains files such as src/skeleton.xlsm/XMLsource/[Content_Types].xml and folders such as src/skeleton.xlsm/XMLsource/_rels
# The script will compress the files and folders into a zip file located in the src/skeleton.xlsm/XMLsource folder
# The zip file will be named skeleton.zip
# The script will use 7zip to compress the files and folders
# The script will use the 7zip command line interface to compress the files and folders

# This script uses 7-Zip to compress files and folders in the src/XMLSource directory into a zip file named skeleton.zip.

Write-Host "Staring the compression process..."

# Define the source folder and the output zip file
$sourceFolder = "src/skeleton.xlsm/XMLSource"
$outputZipFile = "src/skeleton.xlsm/XMLOutput/skeleton.zip"

# Path to the 7-Zip executable
$sevenZipPath = "7z"  # Assumes 7-Zip is in the system PATH. Adjust if necessary.

# Check if the source folder exists
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

# Compress the files and folders using 7-Zip
Write-Host "Compressing files in $sourceFolder to $outputZipFile..."
& $sevenZipPath a -tzip "$outputZipFile" "$sourceFolder/*" | Out-Null

# Check if the compression was successful using $LASTEXITCODE
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error: Compression failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

Write-Host "Compression completed successfully. Zip file created at: $outputZipFile"

# Create a copy of the zip file in the src/skeleton.xlsm/XMLOutput folder at the /src level
$copySource = "src/skeleton.xlsm/XMLOutput/skeleton.zip"
$copyDestination = "src"

# Just perform the copy
Write-Host "Copying $copySource to $copyDestination..."
Copy-Item -Path $copySource -Destination $copyDestination -Force
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error: Copy failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

# Rename src/skeleton.zip to src/skeleton.xlsm
$renameSource = "src/skeleton.zip"
$renameDestination = "src/skeleton.xlsm"
# Just perform the rename
Write-Host "Renaming $renameSource to $renameDestination..."
Rename-Item -Path $renameSource -NewName $renameDestination -Force
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error: Rename failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}
Write-Host "Renaming completed successfully. Zip file renamed to: $renameDestination"