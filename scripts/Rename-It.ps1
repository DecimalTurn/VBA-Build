# This scripts simply copies the file from the source to the destination folder
# and renames it to the correct file name based on the folder name.

# Read the name of the folder from the argument passed to the script
$folderName = $args[0]
if (-not $folderName) {
    Write-Host "Error: No folder name specified. Usage: Rename-It.ps1 <FolderName>"
    exit 1
}

$ext = $args[1]
if (-not $ext) {
    Write-Host "Error: No file extension specified. Usage: Rename-It.ps1 <FolderName> <FileExtension>"
    exit 1
}

$sourceDir = $folderName.Substring(0, $folderName.LastIndexOf('/'))

$filNameWithExtension = $folderName.Substring($folderName.LastIndexOf('/') + 1)
$fileName = $filNameWithExtension.Substring(0, $filNameWithExtension.LastIndexOf('.'))
$fileExtension = $filNameWithExtension.Substring($filNameWithExtension.LastIndexOf('.') + 1)

# Since we can't create an .xlsb file from source code directly, we need to create a .xlsm file and then save it as .xlsb
# We will use the xlsb.xlsm file extension in that case
if ($fileExtension -eq "xlsb") {
    $fileExtension = "xlsb.xlsm"
}

# Since we can't edit the .ppam file directly, we will use the .pptm file extension
if ($fileExtension -eq "ppam") {
    $fileExtension = "ppam.pptm"
}

# Create a copy of the zip/document file in the $folderName/Skeleton folder at the top level
$copySource = "$folderName/Skeleton/$fileName.$ext"
$renameDestinationFolder = "$sourceDir/out"
$renameDestinationFilePath = "$renameDestinationFolder/$fileName.$fileExtension"

# Create rename destination folder if it doesn't exist
if (-not (Test-Path $renameDestinationFolder)) {
    Write-Host "Creating destination folder: $renameDestinationFolder"
    New-Item -ItemType Directory -Path $renameDestinationFolder -Force | Out-Null
}

# Delete the destination file if it exists
if (Test-Path $renameDestinationFilePath) {
    Write-Host "Deleting existing file: $renameDestinationFilePath"
    Remove-Item -Path $renameDestinationFilePath -Force
}

# Copy and rename the file in one step
Write-Host "Copying and renaming $copySource to $renameDestinationFilePath"
Copy-Item -Path $copySource -Destination $renameDestinationFilePath -Force

# Verify if the file exists after the copy
if (-not (Test-Path $renameDestinationFilePath)) {
    Write-Host "Error: File not found after copy: $renameDestinationFilePath"
    exit 1
}

Write-Host "File successfully copied and renamed to: $renameDestinationFilePath"