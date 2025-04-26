# The goal of this function is to detect if an Office dialog is open and return a boolean value indicating whether it is open or not.
# This uses Windows UI Automation to detect dialog windows belonging to the Office application process.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName System.Drawing

# Add required Windows API functions
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class WindowsAPI {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool IsWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
    
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
}
"@

function Detect-OfficeDialog {
    param (
        [Parameter(Mandatory=$true)]
        [object]$officeApp
    )
   
    # First, try to get the process ID of the Office application
    try {
        $processId = $officeApp.Application.Pid
    }
    catch {
        # If .Pid is not available, try another approach to get the process
        try {
            $appTitle = $officeApp.Caption
            $process = Get-Process | Where-Object { $_.MainWindowTitle -eq $appTitle }
            $processId = $process.Id
        }
        catch {
            Write-Host "Could not determine the process ID of the Office application."
            return $false
        }
    }
    
    if ($null -eq $processId) {
        Write-Host "Could not determine the process ID of the Office application."
        return $false
    }
    
    Write-Host "Office application process ID: $processId"
    
    # List of common dialog class names and window titles
    $dialogClassNames = @(
        "#32770",          # Standard Windows dialog
        "bosa_sdm_Microsoft Excel", # Excel VBA dialog
        "bosa_sdm_Microsoft Word",  # Word VBA dialog
        "bosa_sdm_Microsoft PowerPoint", # PowerPoint VBA dialog
        "bosa_sdm_Microsoft Access", # Access VBA dialog
        "ThunderDFrame",    # Office FileDialog
        "NUIDialog"         # Modern Office dialog
    )
    
    $dialogFound = $false
    $windows = @()
    
    # Callback function for EnumWindows
    $enumWindowsCallback = [WindowsAPI+EnumWindowsProc] {
        param($hwnd, $lparam)
        
        # Check if the window is visible
        if ([WindowsAPI]::IsWindowVisible($hwnd)) {
            # Get the process ID for the window
            $windowProcessId = 0
            [void][WindowsAPI]::GetWindowThreadProcessId($hwnd, [ref]$windowProcessId)
            
            # If this window belongs to our Office process
            if ($windowProcessId -eq $processId) {
                # Get the class name
                $classNameBuilder = New-Object System.Text.StringBuilder 256
                [WindowsAPI]::GetClassName($hwnd, $classNameBuilder, 256) | Out-Null
                $className = $classNameBuilder.ToString()
                
                # Get the window title
                $titleBuilder = New-Object System.Text.StringBuilder 256
                [WindowsAPI]::GetWindowText($hwnd, $titleBuilder, 256) | Out-Null
                $title = $titleBuilder.ToString()
                
                $windows += [PSCustomObject]@{
                    Handle = $hwnd
                    ClassName = $className
                    Title = $title
                }
                
                Write-Host "Window found - Class: $className, Title: $title"
                
                # Check if this is a dialog
                foreach ($dialogClass in $dialogClassNames) {
                    if ($className -like "*$dialogClass*") {
                        $script:dialogFound = $true
                        Write-Host "Dialog detected - Class: $className, Title: $title"
                        # We could break here, but continue to collect all windows for debugging
                    }
                }
                
                # Also check for specific dialog titles
                if ($title -like "*Save As*" -or 
                    $title -like "*Open*" -or 
                    $title -like "*Warning*" -or 
                    $title -like "*Error*" -or
                    $title -like "*Message*" -or
                    $title -like "*Visual Basic*") {
                    $script:dialogFound = $true
                    Write-Host "Dialog detected by title - Class: $className, Title: $title"
                }
            }
        }
        
        # Return true to continue enumeration
        return $true
    }
    
    # Enumerate all windows
    [void][WindowsAPI]::EnumWindows($enumWindowsCallback, [IntPtr]::Zero)
    
    # Also check if a dialog is currently the foreground window
    $foregroundWindow = [WindowsAPI]::GetForegroundWindow()
    $foregroundProcessId = 0
    [void][WindowsAPI]::GetWindowThreadProcessId($foregroundWindow, [ref]$foregroundProcessId)
    
    if ($foregroundProcessId -eq $processId) {
        $classNameBuilder = New-Object System.Text.StringBuilder 256
        [WindowsAPI]::GetClassName($foregroundWindow, $classNameBuilder, 256) | Out-Null
        $className = $classNameBuilder.ToString()
        
        $titleBuilder = New-Object System.Text.StringBuilder 256
        [WindowsAPI]::GetWindowText($foregroundWindow, $titleBuilder, 256) | Out-Null
        $title = $titleBuilder.ToString()
        
        Write-Host "Foreground window - Class: $className, Title: $title"
        
        foreach ($dialogClass in $dialogClassNames) {
            if ($className -like "*$dialogClass*") {
                $dialogFound = $true
                Write-Host "Dialog detected (foreground) - Class: $className, Title: $title"
                break
            }
        }
    }
    
    # Return whether a dialog was found
    return $dialogFound
}

# Example usage
$appName = "Excel" # Change this to the desired Office application (Excel, Word, PowerPoint, Access)
$officeApp = New-Object -ComObject "$appName.Application"
$officeApp.Visible = $true # Make sure the application is visible

try {
    # Open a document
    $docPath = "C:\Users\leduc\Workbench\temp_2025\MsgBox-Test.xlsm"
    if (Test-Path $docPath) {
        $doc = $officeApp.Workbooks.Open($docPath)
        
        # Try to run a macro that opens a dialog
        try {
            $officeApp.Run("AsyncFileDialog")
            Write-Host "Ran the AsyncFileDialog macro"
        }
        catch {
            Write-Host "Could not run the macro: $_"
        }
        
        # Wait for a few seconds to allow the dialog to open
        Write-Host "Waiting for dialog to appear..."
        Start-Sleep -Seconds 5
        
        # Check for dialogs
        $dialogOpen = Detect-OfficeDialog -officeApp $officeApp
        Write-Host "Is there a dialog open in ${appName}? $dialogOpen"
    }
    else {
        Write-Host "Test file not found: $docPath"
    }
}
catch {
    Write-Host "Error: $_"
}
finally {
    # Clean up
    try {
        if ($null -ne $doc) {
            $doc.Close($false)
        }
        $officeApp.Quit()
    }
    catch {
        Write-Host "Warning during cleanup: $_"
    }
    
    # Release the COM object
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($officeApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Cleanup completed"
}