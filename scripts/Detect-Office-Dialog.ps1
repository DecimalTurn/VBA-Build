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
    
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    public static extern int GetDlgItem(IntPtr hDlg, int nIDDlgItem);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int GetWindowTextLength(IntPtr hWnd);
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
}
"@

function Detect-OfficeDialog {
    param (
        [Parameter(Mandatory=$true)]
        [object]$officeAppName,
        
        [Parameter(Mandatory=$false)]
        [string]$documentTitle = ""
    )
   
    Write-Host "Detecting Office dialog for application: $officeAppName"
    Write-Host "Looking for Office window with document: $documentTitle"
    
    $caption = "$documentTitle - $officeAppName"
    
    # Find the process ID by enumerating windows and matching the document title
    $script:processId = $null  # Changed to script scope
    $script:mainWindowHwnd = $null
    
    $findWindowCallback = [WindowsAPI+EnumWindowsProc] {
        param($hwnd, $lparam)
        
        if ([WindowsAPI]::IsWindowVisible($hwnd)) {
            $titleBuilder = New-Object System.Text.StringBuilder 256
            [WindowsAPI]::GetWindowText($hwnd, $titleBuilder, 256) | Out-Null
            $windowTitle = $titleBuilder.ToString()
            
            # Check if this window contains our document title
            if ($windowTitle -match [regex]::Escape($caption) -or 
                ($documentTitle -ne "" -and $windowTitle -match [regex]::Escape($documentTitle))) {  # Fixed to use documentTitle
                
                $windowProcessId = 0
                [void][WindowsAPI]::GetWindowThreadProcessId($hwnd, [ref]$windowProcessId)
                
                $script:processId = $windowProcessId
                $script:mainWindowHwnd = $hwnd
                
                Write-Host "Found main window - Title: $windowTitle, Process ID: $windowProcessId"
                return $false  # Stop enumeration once we find the window
            }
        }
        
        return $true  # Continue enumeration
    }
    
    # Find the main application window
    [void][WindowsAPI]::EnumWindows($findWindowCallback, [IntPtr]::Zero)
    
    # If we couldn't find the window by title, try by application caption
    if ($null -eq $script:processId) {  # Changed to script scope
        Write-Host "Couldn't find window by document title, trying by application caption..."
        return $false
    }
    
    Write-Host "Office application process ID: $($script:processId)"  # Changed to script scope
    
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
    
    $script:dialogFound = $false
    $script:dialogWindows = @()  # Array to store dialog window details
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
            if ($windowProcessId -eq $script:processId) {  # Changed to script scope
                # Get the class name
                $classNameBuilder = New-Object System.Text.StringBuilder 256
                [WindowsAPI]::GetClassName($hwnd, $classNameBuilder, 256) | Out-Null
                $className = $classNameBuilder.ToString()
                
                # TODO: Find class name for the other Office applications
                # For now, we will skip the main application window (Excel, Word, etc.)
                if ($className -eq "XLMAIN") {
                    # This is the main application window, skip it
                    return $true  # Continue enumeration
                }

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
                        
                        # Extract dialog content
                        $dialogContent = Get-DialogContent -hwnd $hwnd
                        
                        $script:dialogWindows += [PSCustomObject]@{
                            Handle = $hwnd
                            ClassName = $className
                            Title = $title
                            Type = "Class Match"
                            MessageText = $dialogContent.FullMessage
                            Buttons = ($dialogContent.ButtonTexts -join ", ")
                        }
                        Write-Host "Dialog detected - Class: $className, Title: $title, Handle: $hwnd"
                        Write-Host "  Message: $($dialogContent.FullMessage)"
                        Write-Host "  Buttons: $($dialogContent.ButtonTexts -join ", ")"
                        # Continue enumeration to collect all dialogs
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
                    $script:dialogWindows += [PSCustomObject]@{
                        Handle = $hwnd
                        ClassName = $className
                        Title = $title
                        Type = "Title Match"
                    }
                    Write-Host "Dialog detected by title - Class: $className, Title: $title, Handle: $hwnd"
                    # Continue enumeration to collect all dialogs
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
    
    if ($foregroundProcessId -eq $script:processId) {  # Changed to script scope
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
    
    # Return more detailed information
    return [PSCustomObject]@{
        DialogFound = $dialogFound
        DialogWindows = $script:dialogWindows
        AllWindows = $windows
    }
}

function Get-DialogContent {
    param (
        [Parameter(Mandatory=$true)]
        [IntPtr]$hwnd,
        
        [Parameter(Mandatory=$false)]
        [switch]$DetailedLogging = $false
    )
    
    Write-Host "-----------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Get-DialogContent: Starting analysis of window handle: $hwnd" -ForegroundColor Cyan
    
    try {
        # Try to get automation element
        Write-Host "  Attempting to access UI Automation for window handle: $hwnd"
        $automation = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
        
        if ($null -eq $automation) {
            Write-Host "  ERROR: Could not access UI Automation for window handle: $hwnd" -ForegroundColor Red
            return [PSCustomObject]@{
                MessageTexts = @()
                ButtonTexts = @()
                FullMessage = "ERROR: Could not access dialog automation"
                Success = $false
                ErrorMessage = "Failed to create automation element from handle"
            }
        }
        
        Write-Host "  Successfully created automation element. Name: $($automation.Current.Name)" -ForegroundColor Green
        
        # Get the class name of the window
        $classNameBuilder = New-Object System.Text.StringBuilder 256
        [WindowsAPI]::GetClassName($hwnd, $classNameBuilder, 256) | Out-Null
        $className = $classNameBuilder.ToString()
        Write-Host "  Window class: $className" -ForegroundColor Yellow
        
        # VBA MsgBox specific handling - try to find static text controls
        $dialogTexts = @()
        $buttonTexts = @()
        
        # First approach: Try Text control type (standard automation)
        try {
            Write-Host "  Approach 1: Looking for Text controls..."
            $condition = New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty, 
                [System.Windows.Automation.ControlType]::Text
            )
            
            $textElements = $automation.FindAll([System.Windows.Automation.TreeScope]::Descendants, $condition)
            Write-Host "  Found $($textElements.Count) text elements" -ForegroundColor Yellow
            
            if ($textElements -ne $null -and $textElements.Count -gt 0) {
                foreach ($element in $textElements) {
                    $name = $element.Current.Name
                    if (-not [string]::IsNullOrWhiteSpace($name)) {
                        $dialogTexts += $name
                        Write-Host "    Text: $name" -ForegroundColor Gray
                    }
                }
            }
        }
        catch {
            Write-Host "  ERROR: Failed to find text elements - $_" -ForegroundColor Red
        }
        
        # Second approach: Try to get all static text controls
        if ($dialogTexts.Count -eq 0) {
            try {
                Write-Host "  Approach 2: Looking for Static controls..."
                $staticCondition = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty, 
                    [System.Windows.Automation.ControlType]::Edit
                )
                
                $staticElements = $automation.FindAll([System.Windows.Automation.TreeScope]::Descendants, $staticCondition)
                Write-Host "  Found $($staticElements.Count) static elements" -ForegroundColor Yellow
                
                if ($staticElements -ne $null -and $staticElements.Count -gt 0) {
                    foreach ($element in $staticElements) {
                        $name = $element.Current.Name
                        if (-not [string]::IsNullOrWhiteSpace($name)) {
                            $dialogTexts += $name
                            Write-Host "    Static: $name" -ForegroundColor Gray
                        }
                    }
                }
            }
            catch {
                Write-Host "  ERROR: Failed to find static elements - $_" -ForegroundColor Red
            }
        }
        
        # Third approach: Look for all elements
        if ($dialogTexts.Count -eq 0) {
            try {
                Write-Host "  Approach 3: Looking for all elements..."
                $allElements = $automation.FindAll([System.Windows.Automation.TreeScope]::Descendants, 
                    [System.Windows.Automation.Condition]::TrueCondition)
                
                Write-Host "  Found $($allElements.Count) total elements" -ForegroundColor Yellow
                
                # Loop through all elements to find text-like content
                foreach ($element in $allElements) {
                    if (-not [string]::IsNullOrWhiteSpace($element.Current.Name)) {
                        Write-Host "    Element: $($element.Current.ControlType.ProgrammaticName) - '$($element.Current.Name)'" -ForegroundColor Gray
                        
                        # Only include elements that are likely to contain message text
                        if ($element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Text -or
                            $element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Edit -or
                            $element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Document -or
                            $element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Pane -or
                            $element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Group) {
                            
                            $dialogTexts += $element.Current.Name
                            Write-Host "      Added as dialog text" -ForegroundColor Green
                        }
                        
                        # Identify buttons
                        if ($element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Button) {
                            $buttonTexts += $element.Current.Name
                            Write-Host "      Added as button" -ForegroundColor Green
                        }
                    }
                }
            }
            catch {
                Write-Host "  ERROR: Failed during element enumeration - $_" -ForegroundColor Red
            }
        }
        
        # Fourth approach: Try direct Win32 API approach for VBA dialogs
        if ($dialogTexts.Count -eq 0) {
            try {
                Write-Host "  Approach 4: Using Win32 API to find child windows..." -ForegroundColor Yellow
                
                # Define a delegate for EnumChildWindows callback
                $enumChildCallback = [WindowsAPI+EnumWindowsProc] {
                    param($childHwnd, $lparam)
                    
                    # Get the class name of the child window
                    $childClassBuilder = New-Object System.Text.StringBuilder 256
                    [WindowsAPI]::GetClassName($childHwnd, $childClassBuilder, 256) | Out-Null
                    $childClass = $childClassBuilder.ToString()
                    
                    # Get the text of the child window
                    $textBuilder = New-Object System.Text.StringBuilder 1024
                    [WindowsAPI]::GetWindowText($childHwnd, $textBuilder, 1024) | Out-Null
                    $childText = $textBuilder.ToString()
                    
                    Write-Host "    Child Window - Class: $childClass, Text: $childText" -ForegroundColor Gray
                    
                    # Static text control often has the message
                    if ($childClass -eq "Static" -and -not [string]::IsNullOrWhiteSpace($childText)) {
                        $script:tempDialogTexts += $childText
                    }
                    
                    # Button controls
                    if ($childClass -eq "Button" -and -not [string]::IsNullOrWhiteSpace($childText)) {
                        $script:tempButtonTexts += $childText
                    }
                    
                    return $true  # Continue enumeration
                }
                
                # Create script-scope variables to collect results in the callback
                $script:tempDialogTexts = @()
                $script:tempButtonTexts = @()
                
                # Add EnumChildWindows to WindowsAPI if it doesn't exist
                Add-Type -TypeDefinition @"
                using System;
                using System.Runtime.InteropServices;
                
                public static class ChildWindowAPI {
                    [DllImport("user32.dll")]
                    [return: MarshalAs(UnmanagedType.Bool)]
                    public static extern bool EnumChildWindows(IntPtr hWndParent, WindowsAPI.EnumWindowsProc lpEnumFunc, IntPtr lParam);
                }
"@ -ErrorAction SilentlyContinue
                
                # Enumerate child windows
                [ChildWindowAPI]::EnumChildWindows($hwnd, $enumChildCallback, [IntPtr]::Zero)
                
                # Add results to our main collections
                if ($script:tempDialogTexts.Count -gt 0) {
                    $dialogTexts += $script:tempDialogTexts
                    Write-Host "  Found dialog text using Win32 API: $($script:tempDialogTexts -join ', ')" -ForegroundColor Green
                }
                
                if ($script:tempButtonTexts.Count -gt 0) {
                    $buttonTexts += $script:tempButtonTexts
                    Write-Host "  Found button text using Win32 API: $($script:tempButtonTexts -join ', ')" -ForegroundColor Green
                }
            }
            catch {
                Write-Host "  ERROR: Exception during Win32 API approach - $_" -ForegroundColor Red
            }
        }
        
        return [PSCustomObject]@{
            MessageTexts = $dialogTexts
            ButtonTexts = $buttonTexts
            FullMessage = ($dialogTexts -join " ")
        }
    }
    catch {
        Write-Host "  ERROR: Exception during Get-DialogContent - $_" -ForegroundColor Red
        return [PSCustomObject]@{
            MessageTexts = @()
            ButtonTexts = @()
            FullMessage = "ERROR: Exception during Get-DialogContent"
            Success = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}

# Example usage
$appName = "Excel" # Change this to the desired Office application (Excel, Word, PowerPoint, Access)
$officeApp = New-Object -ComObject "$appName.Application"
$officeApp.Visible = $true # Make sure the application is visible

try {
    # Open a document
    $docName = "MsgBox-Test.xlsm" # Change this to the desired document name
    $docPath = "C:\Users\leduc\Workbench\temp_2025\$docName" # Change this to the desired document path
    if (Test-Path $docPath) {
        $doc = $officeApp.Workbooks.Open($docPath)
        
        # Try to run a macro that opens a dialog
        try {
            #$officeApp.Run("AsyncFileDialog")
            $officeApp.Run("AsyncMsgBox") # Adjust the macro name as needed
            Write-Host "Ran the AsyncFileDialog macro"
        }
        catch {
            Write-Host "Could not run the macro: $_"
        }
        
        # Wait for a few seconds to allow the dialog to open
        Write-Host "Waiting for dialog to appear..."
        Start-Sleep -Seconds 5
        
        # Check for dialogs
        $dialogResult = Detect-OfficeDialog -officeAppName $appName -documentTitle $docName
        Write-Host "Is there a dialog open in ${appName}? $($dialogResult.DialogFound)"

        if ($dialogResult.DialogFound) {
            Write-Host "Found the following dialogs:"
            $dialogResult.DialogWindows | Format-Table -AutoSize
        }
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