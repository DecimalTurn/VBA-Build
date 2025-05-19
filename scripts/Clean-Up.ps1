# This script will make sure that the Office application is closed with all documents to avoid any issues
# when building the next VBA project.
function CleanUp-OfficeApp {
    param (
        [Parameter(Mandatory=$true)]
        $officeApp
    )

    # Close all documents based on the application type
    try {
        # Determine the application type and close all open documents
        if ($officeApp.Name -eq "Microsoft Excel") {
            # Excel - close all workbooks
            for ($i = $officeApp.Workbooks.Count; $i -ge 1; $i--) {
                $workbook = $officeApp.Workbooks.Item($i)
                $workbook.Close($false) # Close without saving changes
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            Write-Host "All Excel workbooks closed successfully"
        }
        elseif ($officeApp.Name -eq "Microsoft Word") {
            # Word - close all documents
            for ($i = $officeApp.Documents.Count; $i -ge 1; $i--) {
                $document = $officeApp.Documents.Item($i)
                $document.Close($false) # Close without saving changes
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
            }
            Write-Host "All Word documents closed successfully"
        }
        elseif ($officeApp.Name -eq "Microsoft PowerPoint") {
            # PowerPoint - close all presentations
            for ($i = $officeApp.Presentations.Count; $i -ge 1; $i--) {
                $presentation = $officeApp.Presentations.Item($i)
                $presentation.Close()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
            }
            Write-Host "All PowerPoint presentations closed successfully"
        }
        elseif ($officeApp.Name -eq "Microsoft Access") {
            # For Access, there's typically one database open
            # Close any open database objects if needed
            Write-Host "Closed Access database"
        }
        else {
            Write-Host "Unknown Office application type. Unable to close documents specifically."
        }
    } catch {
        Write-Host "Warning: Could not close documents: $($_.Exception.Message)"
    }

    # Quit the application
    try {
        $officeApp.Quit()
        Write-Host "Application closed successfully"
    } catch {
        Write-Host "Warning: Could not quit application: $($_.Exception.Message)"
    }

    # Clean up COM objects safely
    try {
        if ($null -ne $officeApp -and $officeApp.GetType().IsCOMObject) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($officeApp) | Out-Null
            Write-Host "Released application COM object"
        }
    } catch {
        Write-Host "Warning: Error releasing application COM object: $($_.Exception.Message)"
    }

    # Force garbage collection to ensure COM objects are fully released
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "Clean-up completed successfully."
}