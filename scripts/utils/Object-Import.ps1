# Summary:
# This script contains utility functions for importing VBA code into Office documents
# with special handling for application-specific object files.

# Function to strip metadata headers from VBA class files
function Parse-Lines {
    param (
        [string]$rawFileContent
    )
    
    # Process raw code to remove headers and metadata
    $lines = $rawFileContent -split [System.Environment]::NewLine
    $processedLinesList = New-Object System.Collections.Generic.List[string]
    $insideBeginEndBlock = $false
    $metadataHeaderProcessed = $false # Flag to indicate metadata section is passed

    foreach ($line in $lines) {
        if ($metadataHeaderProcessed) {
            $processedLinesList.Add($line)
            continue
        }

        $trimmedLine = $line.Trim()
        if ($trimmedLine -eq "BEGIN") { $insideBeginEndBlock = $true; continue }
        if ($insideBeginEndBlock -and $trimmedLine -eq "END") { $insideBeginEndBlock = $false; continue }
        if ($insideBeginEndBlock) { continue }
        if ($trimmedLine -match "^VERSION\s") { continue }
        if ($trimmedLine -match "^Attribute\sVB_") { continue }

        # If none of the above, we're past the metadata header
        $metadataHeaderProcessed = $true
        $processedLinesList.Add($line) # Add this first non-metadata line
    }
    
    return $processedLinesList -join [System.Environment]::NewLine
}

# Function to import code into a VBA component
function Import-CodeToComponent {
    param (
        $component,
        [string]$processedCode,
        [string]$componentName
    )
    
    try {
        # Clear existing code and import new code
        $codeModule = $component.CodeModule
        if ($codeModule.CountOfLines -gt 0) {
            $codeModule.DeleteLines(1, $codeModule.CountOfLines)
            Write-Host "Cleared existing code from $componentName component"
        }
        
        $codeModule.AddFromString($processedCode)
        Write-Host "Successfully imported code into $componentName component"
        return $true
    } catch {
        Write-Host "Error importing $componentName code: $($_.Exception.Message)"
        
        # Fallback to line-by-line import
        try {
            Write-Host "Attempting line-by-line import for $componentName..."
            $processedLines = $processedCode -split [System.Environment]::NewLine
            
            # Ensure $codeModule is available; it should be from the outer try's assignment
            if ($null -ne $codeModule) {
                if ($codeModule.CountOfLines -gt 0) {
                    $codeModule.DeleteLines(1, $codeModule.CountOfLines)
                }
                
                $lineIndex = 1
                foreach ($line in $processedLines) {
                    $codeModule.InsertLines($lineIndex, $line)
                    $lineIndex++
                }
                Write-Host "Successfully imported $componentName code line by line"
                return $true
            } else {
                Write-Host "Error: CodeModule for $componentName is null in fallback."
                return $false
            }
        } catch {
            Write-Host "Failed line-by-line import for ${componentName}: $($_.Exception.Message)"
            return $false
        }
    }
}

# Function to find a component by name in a VBA project
function Find-VbaComponent {
    param (
        $vbProject,
        [string]$componentName
    )
    
    foreach ($component in $vbProject.VBComponents) {
        if ($component.Name -eq $componentName) {
            Write-Host "Found $componentName component in VBA project"
            return $component
        }
    }
    
    Write-Host "Error: Could not find $componentName component in VBA project"
    return $null
}

# Function to import Excel-specific objects
function Import-ExcelObjects {
    param (
        $vbProject,
        [string]$excelObjectsFolder,
        [string]$screenshotDir,
        [string]$fileNameNoExt
    )
    
    if (-not (Test-Path $excelObjectsFolder)) {
        Write-Host "Excel Objects folder not found: $excelObjectsFolder"
        return
    }
    
    Write-Host "Importing Excel-specific objects from: $excelObjectsFolder"
    
    # Find ThisWorkbook.wbk.cls in the Excel Objects folder
    $excelObjectsFiles = Get-ChildItem -Path $excelObjectsFolder -Filter *.cls -ErrorAction SilentlyContinue
    $thisWorkbookFile = $excelObjectsFiles | Where-Object { $_.Name -eq "ThisWorkbook.wbk.cls" }
    
    $wbkFileCount = 0
    if ($null -ne $thisWorkbookFile) { $wbkFileCount = 1 }
    Write-Host "Found $wbkFileCount .wbk.cls files to import"

    if ($null -ne $thisWorkbookFile) {
        # Find the ThisWorkbook component in the VBA project
        $thisWorkbookComponent = Find-VbaComponent -vbProject $vbProject -componentName "ThisWorkbook"
        
        if ($null -eq $thisWorkbookComponent) {
            # Capture screenshot and exit if component not found
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
            return $false
        }

        # Get the code from the ThisWorkbook file
        $rawFileContent = Get-Content -Path $thisWorkbookFile.FullName -Raw
        
        # Process raw code to remove headers and metadata
        $processedCode = Parse-Lines -rawFileContent $rawFileContent
        Write-Host "Processing ThisWorkbook code with $($rawFileContent.Split("`n").Count) lines"
        
        # Import the code into the component
        $importSuccess = Import-CodeToComponent -component $thisWorkbookComponent -processedCode $processedCode -componentName "ThisWorkbook"
        if (-not $importSuccess) {
            return $false
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
        $sheetComponent = Find-VbaComponent -vbProject $vbProject -componentName $sheetName
        
        # If the sheet component is not found, return an error
        if ($null -eq $sheetComponent) {
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
            return $false
        }
        
        # Get the code from the sheet file
        $rawFileContent = Get-Content -Path $sheetFile.FullName -Raw
        
        # Process raw code to remove headers and metadata
        $processedCode = Parse-Lines -rawFileContent $rawFileContent
        Write-Host "Processing sheet code with $($rawFileContent.Split("`n").Count) lines for $sheetName"
        
        # Import the code into the component
        $importSuccess = Import-CodeToComponent -component $sheetComponent -processedCode $processedCode -componentName $sheetName
        if (-not $importSuccess) {
            return $false
        }
    }
    
    return $true
}

# Function to import Word-specific objects
function Import-WordObjects {
    param (
        $vbProject,
        [string]$wordObjectsFolder,
        [string]$screenshotDir,
        [string]$fileNameNoExt
    )
    
    if (-not (Test-Path $wordObjectsFolder)) {
        Write-Host "Word Objects folder not found: $wordObjectsFolder"
        return
    }
    
    Write-Host "Importing Word-specific objects from: $wordObjectsFolder"
    
    # Find ThisDocument.doc.cls in the Word Objects folder
    $wordObjectsFiles = Get-ChildItem -Path $wordObjectsFolder -Filter *.cls -ErrorAction SilentlyContinue
    $thisDocumentFile = $wordObjectsFiles | Where-Object { $_.Name -eq "ThisDocument.doc.cls" }
    
    $docFileCount = 0
    if ($null -ne $thisDocumentFile) { $docFileCount = 1 }
    Write-Host "Found $docFileCount .doc.cls files to import"

    if ($null -ne $thisDocumentFile) {
        # Find the ThisDocument component in the VBA project
        $thisDocumentComponent = Find-VbaComponent -vbProject $vbProject -componentName "ThisDocument"
        
        if ($null -eq $thisDocumentComponent) {
            # Capture screenshot and exit if component not found
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
            return $false
        }

        # Get the code from the ThisDocument file
        $rawFileContent = Get-Content -Path $thisDocumentFile.FullName -Raw
        
        # Process raw code to remove headers and metadata
        $processedCode = Parse-Lines -rawFileContent $rawFileContent
        Write-Host "Processing ThisDocument code with $($rawFileContent.Split("`n").Count) lines"
        
        # Import the code into the component
        $importSuccess = Import-CodeToComponent -component $thisDocumentComponent -processedCode $processedCode -componentName "ThisDocument"
        if (-not $importSuccess) {
            return $false
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
        $objectComponent = Find-VbaComponent -vbProject $vbProject -componentName $objectName
        
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
        
        # Get the code from the file
        $rawFileContent = Get-Content -Path $wordFile.FullName -Raw
        
        # Process raw code to remove headers and metadata
        $processedCode = Parse-Lines -rawFileContent $rawFileContent
        
        # Import the code into the component
        $importSuccess = Import-CodeToComponent -component $objectComponent -processedCode $processedCode -componentName $objectName
        if (-not $importSuccess) {
            # We'll continue with other files even if one fails
            Write-Host "Warning: Failed to import $objectName, continuing with other files"
        }
    }
    
    return $true
}

# Main function to import application-specific objects
function Import-ObjectCode {
    param (
        [string]$officeAppName,
        $vbProject,
        [string]$objectsFolder,
        [string]$screenshotDir,
        [string]$fileNameNoExt
    )
    
    Write-Host "Starting import of $officeAppName object code"
    
    switch ($officeAppName) {
        "Excel" {
            return Import-ExcelObjects -vbProject $vbProject -excelObjectsFolder $objectsFolder -screenshotDir $screenshotDir -fileNameNoExt $fileNameNoExt
        }
        "Word" {
            return Import-WordObjects -vbProject $vbProject -wordObjectsFolder $objectsFolder -screenshotDir $screenshotDir -fileNameNoExt $fileNameNoExt
        }
        default {
            Write-Host "No specific object handling for $officeAppName"
            return $true
        }
    }
}

# Export the functions so they can be used when the script is dot-sourced
Export-ModuleMember -Function Import-ObjectCode, Parse-Lines, Import-CodeToComponent, Find-VbaComponent, Import-ExcelObjects, Import-WordObjects
