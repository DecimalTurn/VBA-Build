# Summary:
# This PowerShell script automates the process of importing VBA code into an Office document.
# It retrieves the current working directory, constructs the path to the Office file,
# and imports .bas, .frm and .cls files from a specified folder into the document and saves it.


# Load utiliies
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Screenshot.ps1"

# Args
$folderName = $args[0]
$officeAppName = $args[1]

if (-not $folderName) {
    Write-Host "🔴 Error: No folder name specified. Usage: Build-VBA.ps1 <FolderName>"
    exit 1
}

if (-not $officeAppName) {
    Write-Host "🔴 Error: No Office application specified. Usage: Build-VBA.ps1 <FolderName> <officeAppName>"
    exit 1
}

# Import utility functions
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Path.ps1"

$currentDir = (Get-Location).Path + "/"
$srcDir = GetAbsPath -path $folderName -basePath $currentDir

$fileName = GetDirName $srcDir
$fileNameNoExt = $fileName.Substring(0, $fileName.LastIndexOf('.'))
$fileExtension = $fileName.Substring($fileName.LastIndexOf('.') + 1)

$outputDir = (DirUp $srcDir) + "out/"
$outputFilePath = $outputDir + $fileName

if ($outputFilePath.EndsWith(".xlsb")) {
    $outputFilePath = $outputFilePath -replace "\.xlsb$", ".xlsb.xlsm"
}

# Make sure the output file already exists
if (-not (Test-Path $outputFilePath)) {
    Write-Host "🔴 Error: Output file not found: $outputFilePath"
    exit 1
}

$screenshotDir = (DirUp $outputDir) + "screenshots/"
if (-not (Test-Path $screenshotDir)) {
    New-Item -ItemType Directory -Path $screenshotDir -Force | Out-Null
    Write-Host "Created screenshot directory: $screenshotDir"
}

# Allows to double-ckeck if the VBOM is enabled
# HKCU:\Software\Microsoft\Office\16.0\Common\TrustCenter
# Just check the registry entry
function Test-VBOMAccess {
    param (
        [string]$officeAppName
    )

    # Check if the VBOM is enabled
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\TrustCenter"
    if (-not (Test-Path $regPath)) {
        Write-Host "Warning: Registry path not found: $regPath"
        Write-Host "Please enable Access to the VBA project object model in Excel Trust Center settings."
        return $false
    }
    
    # Check if the AccessVBOM property is set to 1
    $accessVBOM = Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($null -eq $accessVBOM) {
        Write-Host "Warning: AccessVBOM property not found. Please enable Access to the VBA project object model in Excel Trust Center settings."
        return $false
    } elseif ($accessVBOM.AccessVBOM -ne 1) {
        Write-Host "Warning: AccessVBOM is not enabled. Please enable Access to the VBA project object model in Excel Trust Center settings."
        return $false
    } elseif ($accessVBOM.AccessVBOM -eq 1) {
        Write-Host "AccessVBOM is enabled. Proceeding with import..."
        return $true
    }
    
    # This line should not be reached, but adding as a fallback
    return $false
}

# Check if VBOM access is enabled before attempting imports
if (-not (Test-VBOMAccess -officeAppName $officeAppName)) {
    Write-Host "🔴 Error: VBOM access is not enabled. Please enable it in the Trust Center settings."
    exit 1
}

# Create the application instance
$officeApp = New-Object -ComObject "$officeAppName.Application"

# Make app visible (uncomment if needed)
$officeApp.Visible = $true

# Check if the application instance was created successfully
if ($null -eq $officeApp) {
    Write-Host "🔴 Error: Failed to create COM object for $officeApp"
    exit 1
}

# Open the document
if ($officeAppName -eq "Excel") {
    $doc = $officeApp.Workbooks.Open($outputFilePath)
} elseif ($officeAppName -eq "Word") {
    $doc = $officeApp.Documents.Open($outputFilePath)
} elseif ($officeAppName -eq "PowerPoint") {
    $doc = $officeApp.Presentations.Open($outputFilePath)
} else {
    Write-Host "🔴 Error: Unsupported Office application: $officeAppName"
    exit 1
}

# Check if the document was opened successfully
if ($null -eq $doc) {
    Write-Host "🔴 Error: Failed to open the document: $outputFilePath"
    Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
    exit 1
} else {
    Write-Host "Document opened successfully: $outputFilePath"
}


# Define the module folder path

$moduleFolder = GetAbsPath -path "$folderName/Modules" -basePath $currentDir
Write-Host "Module folder path: $moduleFolder"

# Define the class modules folder path
$classModulesFolder = GetAbsPath -path "$folderName/Class Modules" -basePath $currentDir
Write-Host "Class Modules folder path: $classModulesFolder"

# Define Microsoft Excel Objects folder path
$excelObjectsFolder = GetAbsPath -path "$folderName/Microsoft Excel Objects" -basePath $currentDir
Write-Host "Microsoft Excel Objects folder path: $excelObjectsFolder"

# Define Microsoft Word Objects folder path
$wordObjectsFolder = GetAbsPath -path "$folderName/Microsoft Word Objects" -basePath $currentDir
Write-Host "Microsoft Word Objects folder path: $wordObjectsFolder"

# Define the forms folder path
$formsFolder = GetAbsPath -path "$folderName/Forms" -basePath $currentDir
Write-Host "Forms folder path: $formsFolder"

#Check if the module folder does not exist create an empty one
if (-not (Test-Path $moduleFolder)) {
    Write-Host "Module folder not found: $moduleFolder"
    New-Item -ItemType Directory -Path $moduleFolder -Force | Out-Null
    Write-Host "Created module folder: $moduleFolder"
}

#Check if the class modules folder does not exist create an empty one
if (-not (Test-Path $classModulesFolder)) {
    Write-Host "Class Modules folder not found: $classModulesFolder"
    New-Item -ItemType Directory -Path $classModulesFolder -Force | Out-Null
    Write-Host "Created class modules folder: $classModulesFolder"
}

#Check if the forms folder does not exist create an empty one
if (-not (Test-Path $formsFolder)) {
    Write-Host "Forms folder not found: $formsFolder"
    New-Item -ItemType Directory -Path $formsFolder -Force | Out-Null
    Write-Host "Created forms folder: $formsFolder"
}

#Check if the Microsoft Excel Objects folder does not exist create an empty one (only for Excel)
if ($officeAppName -eq "Excel" -and (-not (Test-Path $excelObjectsFolder))) {
    Write-Host "Microsoft Excel Objects folder not found: $excelObjectsFolder"
    New-Item -ItemType Directory -Path $excelObjectsFolder -Force | Out-Null
    Write-Host "Created Microsoft Excel Objects folder: $excelObjectsFolder"
}

#Check if the Microsoft Word Objects folder does not exist create an empty one (only for Word)
if ($officeAppName -eq "Word" -and (-not (Test-Path $wordObjectsFolder))) {
    Write-Host "Microsoft Word Objects folder not found: $wordObjectsFolder"
    New-Item -ItemType Directory -Path $wordObjectsFolder -Force | Out-Null
    Write-Host "Created Microsoft Word Objects folder: $wordObjectsFolder"
}

# Get VBProject once and reuse it for all imports
$vbProject = $null
try {
    $vbProject = $doc.VBProject
    # Check if the VBProject is accessible
    if ($null -eq $vbProject) {
        Write-Host "VBProject is not accessible. Attempting to re-open the application..."
        $officeApp.Quit()
        Start-Sleep -Seconds 2
        $officeApp = New-Object -ComObject "$officeAppName.Application"
        $officeApp.Visible = $true
        
        # Re-open the document based on application type
        if ($officeAppName -eq "Excel") {
            $doc = $officeApp.Workbooks.Open($outputFilePath)
        } elseif ($officeAppName -eq "Word") {
            $doc = $officeApp.Documents.Open($outputFilePath)
        } elseif ($officeAppName -eq "PowerPoint") {
            $doc = $officeApp.Presentations.Open($outputFilePath)
        } else {
            Write-Host "🔴 Error: Unsupported Office application: $officeAppName"
            exit 1
        }
        
        $vbProject = $doc.VBProject

        if ($null -eq $vbProject) {
            Write-Host "🔴 Error: VBProject is still not accessible after re-opening the application."
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
            exit 1
        } else {
            Write-Host "VBProject is now accessible after re-opening the application."
        }
    }
} catch {
    Write-Host "🔴 Error accessing VB Project: $($_.Exception.Message)"
    Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
    exit 1
}
    
# Check if we have Excel-specific objects to import first (when we're building for Excel)
if ($officeAppName -eq "Excel" -and (Test-Path $excelObjectsFolder)) {
    Write-Host "Importing Excel-specific objects from: $excelObjectsFolder"
    
    # Find ThisWorkbook.wbk.cls in the Excel Objects folder
    $excelObjectsFiles = Get-ChildItem -Path $excelObjectsFolder -Filter *.cls -ErrorAction SilentlyContinue
    $thisWorkbookFile = $excelObjectsFiles | Where-Object { $_.Name -eq "ThisWorkbook.wbk.cls" }
    
    $wbkFileCount = 0
    if ($null -ne $thisWorkbookFile) { $wbkFileCount = 1 }
    Write-Host "Found $wbkFileCount .wbk.cls files to import"

    if ($null -ne $thisWorkbookFile) {
        # Find the ThisWorkbook component in the VBA project
        $thisWorkbookComponent = $null
        foreach ($component in $vbProject.VBComponents) {
            if ($component.Name -eq "ThisWorkbook") {
                $thisWorkbookComponent = $component
                Write-Host "Found ThisWorkbook component in VBA project"
                break
            }
        }
        
        if ($null -eq $thisWorkbookComponent) {
            Write-Host "Error: Could not find ThisWorkbook component in VBA project"
            # Throw an error 
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
            exit 1
        }

        # Get the code from the ThisWorkbook file
        $rawFileContent_wb = Get-Content -Path $thisWorkbookFile.FullName -Raw
        
        # Process raw code to remove headers and metadata for ThisWorkbook
        $lines_wb = $rawFileContent_wb -split [System.Environment]::NewLine
        Write-Host "Processing ThisWorkbook code with $($lines_wb.Count) lines"
        $processedLinesList_wb = New-Object System.Collections.Generic.List[string]
        $insideBeginEndBlock_wb = $false
        $metadataHeaderProcessed_wb = $false # Flag to indicate metadata section is passed

        foreach ($line_iter_wb in $lines_wb) {
            if ($metadataHeaderProcessed_wb) {
                $processedLinesList_wb.Add($line_iter_wb)
                continue
            }

            $trimmedLine_wb = $line_iter_wb.Trim()
            if ($trimmedLine_wb -eq "BEGIN") { $insideBeginEndBlock_wb = $true; continue }
            if ($insideBeginEndBlock_wb -and $trimmedLine_wb -eq "END") { $insideBeginEndBlock_wb = $false; continue }
            if ($insideBeginEndBlock_wb) { continue }
            if ($trimmedLine_wb -match "^VERSION\s") { continue }
            if ($trimmedLine_wb -match "^Attribute\sVB_") { continue }

            # If none of the above, we're past the metadata header
            $metadataHeaderProcessed_wb = $true
            $processedLinesList_wb.Add($line_iter_wb) # Add this first non-metadata line
        }
        $processedVbaCodeString_wb = $processedLinesList_wb -join [System.Environment]::NewLine

        try {
            # Clear existing code and import new code
            $codeModule = $thisWorkbookComponent.CodeModule
            if ($codeModule.CountOfLines -gt 0) {
                $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                Write-Host "Cleared existing code from ThisWorkbook component"
            }
            
            $codeModule.AddFromString($processedVbaCodeString_wb) # Use processed code
            Write-Host "Successfully imported code into ThisWorkbook component"
        } catch {
            Write-Host "Error importing ThisWorkbook code: $($_.Exception.Message)"
            
            # Fallback to line-by-line import
            try {
                Write-Host "Attempting line-by-line import for ThisWorkbook..."
                $processedVbaCodeArray_wb = $processedLinesList_wb.ToArray() # Use processed lines
                
                # Ensure $codeModule is available; it should be from the outer try's assignment
                if ($null -ne $codeModule) {
                    if ($codeModule.CountOfLines -gt 0) {
                        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                    }
                    
                    $lineIndex = 1
                    foreach ($line_in_fallback_wb in $processedVbaCodeArray_wb) {
                        $codeModule.InsertLines($lineIndex, $line_in_fallback_wb)
                        $lineIndex++
                    }
                    Write-Host "Successfully imported ThisWorkbook code line by line"
                } else {
                    Write-Host "Error: CodeModule for ThisWorkbook is null in fallback."
                }
            } catch {
                Write-Host "Failed line-by-line import for ThisWorkbook: $($_.Exception.Message)"
            }
        }

    }
    
    # Import Sheet objects from Excel Objects folder (only files ending with .sheet.cls)
    $sheetFiles = $excelObjectsFiles | Where-Object { $_.Name -like "*.sheet.cls" }
    
    Write-Host "Found $($sheetFiles.Count) .sheet.cls files to import"
    
    foreach ($sheetFile in $sheetFiles) {
        Write-Host "Processing Excel sheet object: $($sheetFile.Name)"
        
        # Extract the sheet name from the filename (e.g., Sheet1.sheet.cls -> Sheet1)
        $sheetName = [System.IO.Path]::GetFileNameWithoutExtension($sheetFile.Name)
        $sheetName = $sheetName -replace "\.sheet$", ""
        
        # Find the corresponding sheet component
        $sheetComponent = $null
        foreach ($component in $vbProject.VBComponents) {
            if ($component.Name -eq $sheetName) {
                $sheetComponent = $component
                Write-Host "Found sheet component: $sheetName"
                break
            }
        }

        # If the sheet component is not found, return an error
        if ($null -eq $sheetComponent) {
            Write-Host "Error: Could not find sheet component for $sheetName"
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
            exit 1
        }
        
        
        # Get the code from the sheet file
        $rawFileContent_sh = Get-Content -Path $sheetFile.FullName -Raw

        # Process raw code to remove headers and metadata for Sheet objects
        $lines_sh = $rawFileContent_sh -split [System.Environment]::NewLine
        Write-Host "Processing sheet code with $($lines_sh.Count) lines for $sheetName"
        $processedLinesList_sh = New-Object System.Collections.Generic.List[string]
        $insideBeginEndBlock_sh = $false
        $metadataHeaderProcessed_sh = $false # Flag to indicate metadata section is passed

        foreach ($line_iter_sh in $lines_sh) {
            if ($metadataHeaderProcessed_sh) {
                $processedLinesList_sh.Add($line_iter_sh)
                continue
            }

            $trimmedLine_sh = $line_iter_sh.Trim()
            if ($trimmedLine_sh -eq "BEGIN") { $insideBeginEndBlock_sh = $true; continue }
            if ($insideBeginEndBlock_sh -and $trimmedLine_sh -eq "END") { $insideBeginEndBlock_sh = $false; continue }
            if ($insideBeginEndBlock_sh) { continue }
            if ($trimmedLine_sh -match "^VERSION\s") { continue }
            if ($trimmedLine_sh -match "^Attribute\sVB_") { continue }

            # If none of the above, we're past the metadata header
            $metadataHeaderProcessed_sh = $true
            $processedLinesList_sh.Add($line_iter_sh) # Add this first non-metadata line
        }
        $processedVbaCodeString_sh = $processedLinesList_sh -join [System.Environment]::NewLine
        
        try {
            # Clear existing code and import new code
            $codeModule = $sheetComponent.CodeModule
            if ($codeModule.CountOfLines -gt 0) {
                $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                Write-Host "Cleared existing code from ${sheetName} component"
            }
            
            $codeModule.AddFromString($processedVbaCodeString_sh) # Use processed code
            Write-Host "Successfully imported code into ${sheetName} component"
        } catch {
            Write-Host "Error importing sheet code: $($_.Exception.Message)"
            
            # Fallback to line-by-line import
            try {
                Write-Host "Attempting line-by-line import for ${sheetName}..."
                $processedVbaCodeArray_sh = $processedLinesList_sh.ToArray() # Use processed lines
                
                # Ensure $codeModule is available; it should be from the outer try's assignment
                if ($null -ne $codeModule) {
                    if ($codeModule.CountOfLines -gt 0) {
                        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                    }
                    
                    $lineIndex = 1
                    foreach ($line_in_fallback_sh in $processedVbaCodeArray_sh) {
                        $codeModule.InsertLines($lineIndex, $line_in_fallback_sh)
                        $lineIndex++
                    }
                    Write-Host "Successfully imported ${sheetName} code line by line"
                } else {
                    Write-Host "Error: CodeModule for ${sheetName} is null in fallback."
                }
            } catch {
                Write-Host "Failed line-by-line import for ${sheetName}: $($_.Exception.Message)"
            }
        }

    }
}

# Check if we have Word-specific objects to import (when we're building for Word)
if ($officeAppName -eq "Word" -and (Test-Path $wordObjectsFolder)) {
    Write-Host "Importing Word-specific objects from: $wordObjectsFolder"
    
    # Find ThisDocument.doc.cls in the Word Objects folder
    $wordObjectsFiles = Get-ChildItem -Path $wordObjectsFolder -Filter *.cls -ErrorAction SilentlyContinue
    $thisDocumentFile = $wordObjectsFiles | Where-Object { $_.Name -eq "ThisDocument.doc.cls" }
    
    $docFileCount = 0
    if ($null -ne $thisDocumentFile) { $docFileCount = 1 }
    Write-Host "Found $docFileCount .doc.cls files to import"

    if ($null -ne $thisDocumentFile) {
        # Find the ThisDocument component in the VBA project
        $thisDocumentComponent = $null
        foreach ($component in $vbProject.VBComponents) {
            if ($component.Name -eq "ThisDocument") {
                $thisDocumentComponent = $component
                Write-Host "Found ThisDocument component in VBA project"
                break
            }
        }
        
        if ($null -eq $thisDocumentComponent) {
            Write-Host "Error: Could not find ThisDocument component in VBA project"
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
            exit 1
        }

        # Get the code from the ThisDocument file
        $rawFileContent_doc = Get-Content -Path $thisDocumentFile.FullName -Raw
        
        # Process raw code to remove headers and metadata for ThisDocument
        $lines_doc = $rawFileContent_doc -split [System.Environment]::NewLine
        Write-Host "Processing ThisDocument code with $($lines_doc.Count) lines"
        $processedLinesList_doc = New-Object System.Collections.Generic.List[string]
        $insideBeginEndBlock_doc = $false
        $metadataHeaderProcessed_doc = $false # Flag to indicate metadata section is passed

        foreach ($line_iter_doc in $lines_doc) {
            if ($metadataHeaderProcessed_doc) {
                $processedLinesList_doc.Add($line_iter_doc)
                continue
            }

            $trimmedLine_doc = $line_iter_doc.Trim()
            if ($trimmedLine_doc -eq "BEGIN") { $insideBeginEndBlock_doc = $true; continue }
            if ($insideBeginEndBlock_doc -and $trimmedLine_doc -eq "END") { $insideBeginEndBlock_doc = $false; continue }
            if ($insideBeginEndBlock_doc) { continue }
            if ($trimmedLine_doc -match "^VERSION\s") { continue }
            if ($trimmedLine_doc -match "^Attribute\sVB_") { continue }

            # If none of the above, we're past the metadata header
            $metadataHeaderProcessed_doc = $true
            $processedLinesList_doc.Add($line_iter_doc) # Add this first non-metadata line
        }
        $processedVbaCodeString_doc = $processedLinesList_doc -join [System.Environment]::NewLine

        try {
            # Clear existing code and import new code
            $codeModule = $thisDocumentComponent.CodeModule
            if ($codeModule.CountOfLines -gt 0) {
                $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                Write-Host "Cleared existing code from ThisDocument component"
            }
            
            $codeModule.AddFromString($processedVbaCodeString_doc) # Use processed code
            Write-Host "Successfully imported code into ThisDocument component"
        } catch {
            Write-Host "Error importing ThisDocument code: $($_.Exception.Message)"
            
            # Fallback to line-by-line import
            try {
                Write-Host "Attempting line-by-line import for ThisDocument..."
                $processedVbaCodeArray_doc = $processedLinesList_doc.ToArray() # Use processed lines
                
                # Ensure $codeModule is available; it should be from the outer try's assignment
                if ($null -ne $codeModule) {
                    if ($codeModule.CountOfLines -gt 0) {
                        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                    }
                    
                    $lineIndex = 1
                    foreach ($line_in_fallback_doc in $processedVbaCodeArray_doc) {
                        $codeModule.InsertLines($lineIndex, $line_in_fallback_doc)
                        $lineIndex++
                    }
                    Write-Host "Successfully imported ThisDocument code line by line"
                } else {
                    Write-Host "Error: CodeModule for ThisDocument is null in fallback."
                }
            } catch {
                Write-Host "Failed line-by-line import for ThisDocument: $($_.Exception.Message)"
            }
        }
    }
    
    # Look for other potential Word objects to import
    $otherWordFiles = $wordObjectsFiles | Where-Object { $_.Name -ne "ThisDocument.doc.cls" }
    
    Write-Host "Found $($otherWordFiles.Count) other Word object files to import"
    
    foreach ($wordFile in $otherWordFiles) {
        Write-Host "Processing Word object: $($wordFile.Name)"
        
        # Extract the component name from the filename (e.g., SomeObject.doc.cls -> SomeObject)
        $objectName = [System.IO.Path]::GetFileNameWithoutExtension($wordFile.Name)
        $objectName = $objectName -replace "\.doc$", ""
        
        # Try to find corresponding component if it exists
        $objectComponent = $null
        foreach ($component in $vbProject.VBComponents) {
            if ($component.Name -eq $objectName) {
                $objectComponent = $component
                Write-Host "Found Word component: $objectName"
                break
            }
        }

        # If component doesn't exist, we'll need to try importing as a regular component
        if ($null -eq $objectComponent) {
            Write-Host "Component $objectName not found in VBA project, attempting to import as a regular component"
            try {
                $vbProject.VBComponents.Import($wordFile.FullName)
                Write-Host "Successfully imported $($wordFile.Name) as a new component"
                continue
            } catch {
                Write-Host "Error importing $($wordFile.Name): $($_.Exception.Message)"
                continue
            }
        }
        
        # Process the existing component similar to ThisDocument
        $rawFileContent_obj = Get-Content -Path $wordFile.FullName -Raw
        $lines_obj = $rawFileContent_obj -split [System.Environment]::NewLine
        $processedLinesList_obj = New-Object System.Collections.Generic.List[string]
        $insideBeginEndBlock_obj = $false
        $metadataHeaderProcessed_obj = $false

        foreach ($line_iter_obj in $lines_obj) {
            if ($metadataHeaderProcessed_obj) {
                $processedLinesList_obj.Add($line_iter_obj)
                continue
            }

            $trimmedLine_obj = $line_iter_obj.Trim()
            if ($trimmedLine_obj -eq "BEGIN") { $insideBeginEndBlock_obj = $true; continue }
            if ($insideBeginEndBlock_obj -and $trimmedLine_obj -eq "END") { $insideBeginEndBlock_obj = $false; continue }
            if ($insideBeginEndBlock_obj) { continue }
            if ($trimmedLine_obj -match "^VERSION\s") { continue }
            if ($trimmedLine_obj -match "^Attribute\sVB_") { continue }

            $metadataHeaderProcessed_obj = $true
            $processedLinesList_obj.Add($line_iter_obj)
        }
        $processedVbaCodeString_obj = $processedLinesList_obj -join [System.Environment]::NewLine
        
        try {
            $codeModule = $objectComponent.CodeModule
            if ($codeModule.CountOfLines -gt 0) {
                $codeModule.DeleteLines(1, $codeModule.CountOfLines)
            }
            $codeModule.AddFromString($processedVbaCodeString_obj)
            Write-Host "Successfully imported code into $objectName component"
        } catch {
            Write-Host "Error importing code for $objectName: $($_.Exception.Message)"
            
            # Fallback to line-by-line import
            try {
                $processedVbaCodeArray_obj = $processedLinesList_obj.ToArray()
                if ($null -ne $codeModule) {
                    if ($codeModule.CountOfLines -gt 0) {
                        $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                    }
                    $lineIndex = 1
                    foreach ($line_in_fallback_obj in $processedVbaCodeArray_obj) {
                        $codeModule.InsertLines($lineIndex, $line_in_fallback_obj)
                        $lineIndex++
                    }
                    Write-Host "Successfully imported $objectName code line by line"
                }
            } catch {
                Write-Host "Failed line-by-line import for $objectName: $($_.Exception.Message)"
            }
        }
    }
}

# Import class modules (.cls files)
$clsFiles = Get-ChildItem -Path $classModulesFolder -Filter *.cls -ErrorAction SilentlyContinue
Write-Host "Found $($clsFiles.Count) .cls files to import"

# Loop through each class module file
$clsFiles | ForEach-Object {
    Write-Host "Importing class module $($_.Name)..."
    try {
        $vbProject.VBComponents.Import($_.FullName)
        Write-Host "Successfully imported class module $($_.Name)"
    } catch {
        Write-Host "Failed to import class module $($_.Name): $($_.Exception.Message)"
    }
}

# Import form modules (.frm files)
$frmFiles = Get-ChildItem -Path $formsFolder -Filter *.frm -ErrorAction SilentlyContinue
Write-Host "Found $($frmFiles.Count) .frm files to import"

# Loop through each form file
$frmFiles | ForEach-Object {
    Write-Host "Importing form $($_.Name)..."
    try {
        $vbProject.VBComponents.Import($_.FullName)
        Write-Host "Successfully imported form $($_.Name)"
    } catch {
        Write-Host "Failed to import form $($_.Name): $($_.Exception.Message)"
    }
}

# First check if there are any .bas files
$basFiles = Get-ChildItem -Path $moduleFolder -Filter *.bas
Write-Host "Found $($basFiles.Count) .bas files to import"

# Loop through each file in the module folder
$basFiles | ForEach-Object {
    Write-Host "Importing $($_.Name)..."
    try {
        $vbProject.VBComponents.Import($_.FullName)
        Write-Host "Successfully imported $($_.Name)"
    } catch {
        Write-Host "Failed to import $($_.Name): $($_.Exception.Message)"
    }
}

# Save the document
Write-Host "Saving document..."
try {
    if ($officeAppName -eq "PowerPoint" -or $officeAppName -eq "Word") {
        # For PowerPoint, use SaveAs with the same file name to force save
        $doc.SaveAs($outputFilePath)
        Write-Host "Document saved using SaveAs method"
    } elseif ($officeAppName -eq "Excel") {
        # For Excel, we need to check if the file name ends with .xlsb.xlsm
        # If so, we need to save as .xlsb
        if ($outputFilePath.EndsWith(".xlsb.xlsm")) {
            $newFilePath = $outputFilePath -replace "\.xlsb\.xlsm$", ".xlsb"
            # Replace forward slashes with backslashes
            $newFilePath = $newFilePath -replace "/", "\"
            Write-Host "Saving document as .xlsb: $newFilePath"
            $doc.SaveAs($newFilePath, 50) # 50 is the xlExcel12 file format for .xlsb
            # Delete the .xlsb.xlsm file
            Remove-Item -Path $outputFilePath -Force
            Write-Host "Document saved as .xlsb"
        } else {
            $doc.Save()
            Write-Host "Document saved successfully"
        }
    } else {
        $doc.Save()
        Write-Host "Document saved successfully"
    }
} catch {
    Write-Host "Warning: Could not save document: $($_.Exception.Message)"
    Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
    
    # Alternative approach for PowerPoint if SaveAs fails
    if ($officeAppName -eq "PowerPoint") {
        try {
            # Try saving with a temporary file name and then renaming
            $tempPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.pptm'
            Write-Host "Attempting to save to temporary location: $tempPath"
            $doc.SaveAs($tempPath)
            
            # Close the document and application
            $doc.Close()
            $officeApp.Quit()
            
            # Release COM objects
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($officeApp) | Out-Null
            
            # Wait a moment for resources to be released
            Start-Sleep -Seconds 2
            
            # Copy the temp file to the intended destination
            Copy-Item -Path $tempPath -Destination $outputFilePath -Force
            Remove-Item -Path $tempPath -Force
            
            Write-Host "Document saved using alternative method"
            
            # Skip the rest of the cleanup as we've already done it
            Write-Host "VBA import completed successfully."
            exit 0
        } catch {
            Write-Host "Error: Alternative save method also failed: $($_.Exception.Message)"
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
        }
    }
}

# Generic test
# Call the WriteToFile macro to check if the module was imported correctly
try {
    
    $vbaModule = $doc.VBProject.VBComponents.Item(1)
    if ($null -eq $vbaModule) {
        Write-Host "Error: No VBA module found in the document."
        Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
        exit 1
    }
    Write-Host "VBA module found: $($vbaModule.Name)"

    $macroName = "WriteToFile"
    Write-Host "Macro to execute: $macroName"
    Write-Host "Application state before macro execution: Type=$($officeApp.GetType().FullName)"
    $officeApp.Run($macroName)
    Write-Host "Macro finished"
    
    # Check if the file was created successfully with the correct content
    $outputFile = "$outputDir/$fileNameNoExt.txt"
    if (Test-Path $outputFile) {
        $fileContent = Get-Content -Path $outputFile
        if ($fileContent -eq "Hello, World!") {
            Write-Host "🟢 Macro executed successfully and file content is correct."
        } else {
            Write-Host "🟡 Warning: Macro executed, but file content is incorrect.: $fileContent"
        }

        # Delete the output file after checking
        Remove-Item -Path $outputFile -Force
        Write-Host "Output test file deleted successfully."

    } else {
        Write-Host "🟡 Warning: Macro executed, but output file was not created."
    }

} catch {
    Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
    Write-Host "🟡 Warning: Could not execute macro ${macroName}: $($_.Exception.Message)"
}

# Clean-Up: Release the document
try {
    if ($null -ne $doc -and $doc.GetType().IsCOMObject) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
        Write-Host "Released document COM object"
    }
} catch {
    Write-Host "Warning: Error releasing document COM object: $($_.Exception.Message)"
}
