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

# Import utility functions
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Path.ps1"

$rootSrcDir = DirUp (NormalizeDirPath $folderName)

$currentDir = (Get-Location).Path + "/"
$srcDir = GetAbsPath -path $folderName -basePath $currentDir

$fileName = GetDirName $srcDir
$fileNameNoExt = $fileName.Substring(0, $fileName.LastIndexOf('.'))
$fileExtension = $fileName.Substring($fileName.LastIndexOf('.') + 1)

# Create a copy of the zip/document file in the $folderName/Skeleton folder at the top level
$copySource = "$folderName/Skeleton/$fileNameNoExt.$ext"
$renameDestinationFolder = $rootSrcDir + "out/"
$renameDestinationFilePath = "$renameDestinationFolder$fileNameNoExt.$fileExtension"

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