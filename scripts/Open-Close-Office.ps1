# This script will list all the process that are part of the Office suite (and generic Office processes) and close them.
# It will also check if the processes are running and close them if they are.

# Define an array of Office-related process names
$officeProcesses = @(
    "WINWORD",      # Microsoft Word
    "EXCEL",        # Microsoft Excel
    "POWERPNT",     # Microsoft PowerPoint
    "OUTLOOK",      # Microsoft Outlook
    "ONENOTE",      # Microsoft OneNote
    "MSACCESS",     # Microsoft Access
    "MSPUB",        # Microsoft Publisher
    "TEAMS",        # Microsoft Teams
    "lync",         # Skype for Business
    "ONENOTEM",     # OneNote quick notes
    "VISIO",        # Microsoft Visio
    "OfficeClickToRun", # Office Click-to-Run
    "OfficeC2RClient"   # Office Click-to-Run Client
)

# Function to check and close Office processes
function Close-OfficeProcesses {
    Write-Host "Checking for running Office processes..." -ForegroundColor Cyan
    
    $runningProcesses = Get-Process | Where-Object { $officeProcesses -contains $_.Name }
    
    if ($runningProcesses.Count -eq 0) {
        Write-Host "No Office processes are currently running." -ForegroundColor Green
        return
    }
    
    Write-Host "Found $($runningProcesses.Count) running Office processes:" -ForegroundColor Yellow
    foreach ($process in $runningProcesses) {
        Write-Host "  - $($process.Name) (PID: $($process.Id))" -ForegroundColor Yellow
    }
   
    Write-Host "Closing Office processes..." -ForegroundColor Cyan
    foreach ($process in $runningProcesses) {
        try {
            $process | Stop-Process -Force -ErrorAction Stop
            Write-Host "  Successfully closed $($process.Name)." -ForegroundColor Green
        }
        catch {
            Write-Host "  Failed to close $($process.Name): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    # Wait for a moment to ensure processes are closed
    Start-Sleep -Seconds 15

    Write-Host "Re-checking for any remaining Office processes..." -ForegroundColor Cyan

    # Check if any processes remain running
    $remainingProcesses = Get-Process | Where-Object { $officeProcesses -contains $_.Name }
    if ($remainingProcesses.Count -gt 0) {
        Write-Host "`nWARNING: $($remainingProcesses.Count) Office processes could not be closed:" -ForegroundColor Red
        # Let's try a second time to close them
        foreach ($process in $remainingProcesses) {
            try {
                $process | Stop-Process -Force -ErrorAction Stop
                Write-Host "  Successfully closed $($process.Name)." -ForegroundColor Green
            }
            catch {
                Write-Host "  Failed to close $($process.Name): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        # Re-check for any remaining processes
        Start-Sleep -Seconds 15
        $remainingProcesses = Get-Process | Where-Object { $officeProcesses -contains $_.Name }
        if ($remainingProcesses.Count -eq 0) {
            Write-Host "All Office processes have been successfully closed." -ForegroundColor Green
        }
        else {
            Write-Host "Some Office processes are still running:" -ForegroundColor Red
                    # List any remaining processes
            foreach ($process in $remainingProcesses) {
                Write-Host "  - $($process.Name) (PID: $($process.Id))" -ForegroundColor Red
            }
        }

    }
    else {
        Write-Host "`nAll Office processes have been successfully closed." -ForegroundColor Green
    }
}


$officeApps = $args[0]
if (-not $officeApps) {
    Write-Host "No Office applications specified. Exiting script." -ForegroundColor Red
    exit 1
}

# Convert the comma-separated string to an array
$officeApps = $officeApps -split ',' | ForEach-Object { $_.Trim() }

# Open each required Office application
foreach ($app in $officeApps) {
    Write-Host "Opening $app application..." -ForegroundColor Cyan
    try {
        # Create a new COM object for the specified Office application
        $comObject = New-Object -ComObject "$app.Application"
        Write-Host "$app application opened successfully." -ForegroundColor Green
        $comObject = $null
    }
    catch {
        Write-Host "Failed to open $app application: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Wait for a moment to ensure processes are opened
Start-Sleep -Seconds 5

# Execute the function
Close-OfficeProcesses