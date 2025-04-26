# The goal of this function is to detect if an Office dialog is open and return a boolean value indicating whether it is open or not.
# This uses Windows UI Automation to detect dialog windows belonging to the Office application process.

# Add required assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName System.Drawing

# Try to create the types only if they don't already exist
try {
    if (-not ([System.Management.Automation.PSTypeName]'WindowsAPI').Type) {
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
    
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    public const uint BM_CLICK = 0x00F5;
    public const uint WM_LBUTTONDOWN = 0x0201;
    public const uint WM_LBUTTONUP = 0x0202;
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
}
"@ -ErrorAction Stop
    }
} 
catch {
    Write-Warning "Failed to create Win32 API types: $_"
    Write-Warning "The script will try to continue but some functionality may be limited."
}

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
        # Get the class name of the window
        $classNameBuilder = New-Object System.Text.StringBuilder 256
        [WindowsAPI]::GetClassName($hwnd, $classNameBuilder, 256) | Out-Null
        $className = $classNameBuilder.ToString()
        Write-Host "  Window class: $className" -ForegroundColor Yellow
        
        # Get the window title
        $titleBuilder = New-Object System.Text.StringBuilder 256
        [WindowsAPI]::GetWindowText($hwnd, $titleBuilder, 256) | Out-Null
        $title = $titleBuilder.ToString()
        Write-Host "  Window title: $title" -ForegroundColor Yellow
        
        # Initialize collections
        $dialogTexts = @()
        $buttonTexts = @()
        $controlsFound = $false
        
        # APPROACH 1: Use Win32 API directly (most reliable for VBA dialogs)
        try {
            Write-Host "  Approach 1: Using direct Win32 API to find child windows..." -ForegroundColor Yellow
            
            # Create script-scope variables to collect results in the callback
            $script:tempDialogTexts = @()
            $script:tempButtonTexts = @()
            
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
                
                if (-not [string]::IsNullOrWhiteSpace($childText)) {
                    Write-Host "    Child Window - Class: $childClass, Text: $childText" -ForegroundColor Gray
                    
                    # Static text control typically contains the message text
                    if ($childClass -eq "Static" -and -not [string]::IsNullOrWhiteSpace($childText)) {
                        $script:tempDialogTexts += $childText
                        Write-Host "      Added as dialog text" -ForegroundColor Green
                    }
                    
                    # Button controls
                    elseif ($childClass -eq "Button" -and -not [string]::IsNullOrWhiteSpace($childText)) {
                        $script:tempButtonTexts += $childText
                        Write-Host "      Added as button" -ForegroundColor Green
                    }
                }
                
                return $true  # Continue enumeration
            }
            
            # Enumerate child windows using alternative API
            if (-not [ChildWindowAPI]::EnumChildWindows($hwnd, $enumChildCallback, [IntPtr]::Zero)) {
                Write-Host "  Warning: EnumChildWindows returned false" -ForegroundColor Yellow
            }
            
            # Add results to our main collections
            if ($script:tempDialogTexts.Count -gt 0) {
                $dialogTexts += $script:tempDialogTexts
                $controlsFound = $true
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
        
        # APPROACH 2: Try UI Automation as a fallback
        if (-not $controlsFound) {
            try {
                Write-Host "  Approach 2: Attempting UI Automation..." -ForegroundColor Yellow
                $automation = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
                
                if ($null -eq $automation) {
                    Write-Host "  Failed to get automation element" -ForegroundColor Red
                }
                else {
                    Write-Host "  Successfully created automation element" -ForegroundColor Green
                    
                    # Get all elements
                    $allElements = $automation.FindAll(
                        [System.Windows.Automation.TreeScope]::Descendants, 
                        [System.Windows.Automation.Condition]::TrueCondition
                    )
                    
                    Write-Host "  Found $($allElements.Count) total UI elements" -ForegroundColor Yellow
                    
                    # Identify buttons first
                    foreach ($element in $allElements) {
                        if (-not [string]::IsNullOrWhiteSpace($element.Current.Name) -and 
                            $element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Button) {
                            $buttonTexts += $element.Current.Name
                            Write-Host "    Button: $($element.Current.Name)" -ForegroundColor Gray
                        }
                    }
                    
                    # Then look for text elements, avoiding button names
                    foreach ($element in $allElements) {
                        if (-not [string]::IsNullOrWhiteSpace($element.Current.Name) -and
                            ($element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Text -or
                             $element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Edit -or
                             $element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Pane)) {
                            
                            # Skip if this is a known button name
                            if ($buttonTexts -notcontains $element.Current.Name) {
                                $dialogTexts += $element.Current.Name
                                Write-Host "    Text: $($element.Current.Name)" -ForegroundColor Gray
                            }
                        }
                    }
                }
            }
            catch {
                Write-Host "  ERROR: Exception during UI Automation - $_" -ForegroundColor Red
            }
        }
        
        # APPROACH 3: If no better information is available, use the window title as fallback
        if ($dialogTexts.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($title)) {
            Write-Host "  Using window title as fallback for dialog text: $title" -ForegroundColor Yellow
            $dialogTexts += $title
        }
        
        # Process results for standard Office dialog buttons
        if ($buttonTexts.Count -eq 0) {
            # Check for standard dialog button patterns
            if ($className -eq "#32770") {
                Write-Host "  Standard dialog detected, checking for common button patterns..." -ForegroundColor Yellow
                
                # Look for standard button IDs
                $buttonIds = @(1, 2, 6, 7)  # OK, Cancel, Yes, No
                $foundButtons = $false
                
                foreach ($id in $buttonIds) {
                    $btnHwnd = [WindowsAPI]::GetDlgItem($hwnd, $id)
                    if ($btnHwnd -ne 0) {
                        $btnTextBuilder = New-Object System.Text.StringBuilder 256
                        [WindowsAPI]::GetWindowText($btnHwnd, $btnTextBuilder, 256) | Out-Null
                        $btnText = $btnTextBuilder.ToString()
                        
                        if (-not [string]::IsNullOrWhiteSpace($btnText)) {
                            $buttonTexts += $btnText
                            $foundButtons = $true
                            Write-Host "    Found standard button ID $id`: $btnText" -ForegroundColor Green
                        }
                    }
                }
                
                if (-not $foundButtons) {
                    Write-Host "    Could not find standard buttons by ID" -ForegroundColor Yellow
                    
                    # Make educated guesses based on dialog title
                    if ($title -match "Error|Warning|Alert") {
                        $buttonTexts += "OK"
                        Write-Host "    Added assumed 'OK' button based on error/warning dialog" -ForegroundColor Yellow
                    }
                }
            }
        }
        
        # Clean up potential duplicates
        $dialogTexts = $dialogTexts | Select-Object -Unique
        $buttonTexts = $buttonTexts | Select-Object -Unique
        
        # Create the full message
        $fullMessage = $dialogTexts -join " "
        
        Write-Host "  Final dialog content - Message: '$fullMessage'" -ForegroundColor Green
        Write-Host "  Final dialog content - Buttons: '$($buttonTexts -join ", ")'" -ForegroundColor Green
        
        return [PSCustomObject]@{
            MessageTexts = $dialogTexts
            ButtonTexts = $buttonTexts
            FullMessage = $fullMessage
            Success = $true
            WindowClass = $className
            WindowTitle = $title
        }
    }
    catch {
        Write-Host "  CRITICAL ERROR in Get-DialogContent: $_" -ForegroundColor Red
        return [PSCustomObject]@{
            MessageTexts = @()
            ButtonTexts = @()
            FullMessage = "ERROR: $($_.Exception.Message)"
            Success = $false
            ErrorMessage = $_.Exception.Message
        }
    }
    finally {
        Write-Host "-----------------------------------------------------" -ForegroundColor Cyan
    }
}

function Dismiss-Dialog {
    param (
        [Parameter(Mandatory=$true)]
        [IntPtr]$hwnd,
        
        [Parameter(Mandatory=$false)]
        [string]$buttonToClick = "OK",
        
        [Parameter(Mandatory=$false)]
        [int]$maxAttempts = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$delayBetweenAttempts = 500  # milliseconds
    )
    
    Write-Host "-----------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Attempting to dismiss dialog (handle: $hwnd) by clicking '$buttonToClick' button..." -ForegroundColor Cyan
    
    # Check if the window still exists
    if (-not [WindowsAPI]::IsWindow($hwnd)) {
        Write-Host "  Dialog already dismissed (window no longer exists)" -ForegroundColor Green
        return $true
    }

    # First try: Look for button by its control ID
    # Standard Windows dialog button IDs
    $buttonIds = @{
        "OK" = 1        # IDOK
        "Cancel" = 2    # IDCANCEL
        "Abort" = 3     # IDABORT
        "Retry" = 4     # IDRETRY
        "Ignore" = 5    # IDIGNORE
        "Yes" = 6       # IDYES
        "No" = 7        # IDNO
        "Close" = 8     # IDCLOSE
        "Help" = 9      # IDHELP
    }

    if ($buttonIds.ContainsKey($buttonToClick)) {
        $buttonId = $buttonIds[$buttonToClick]
    }
    else {
        Write-Host "  Button '$buttonToClick' not found in standard button" -ForegroundColor Red
        return $false
    }

    $btnHwnd = [WindowsAPI]::GetDlgItem($hwnd, $buttonId)

    if ($btnHwnd -eq 0) {
        Write-Host "  Button not found by ID, trying other methods..." -ForegroundColor Yellow

        # Try common alternative IDs
        $alternateIds = @(100, 101, 102, 1, 2)

        $foundButton = $false
        foreach ($id in $alternateIds) {
            $btnHwnd = [WindowsAPI]::GetDlgItem($hwnd, $id)
            if ($btnHwnd -ne 0) {
                $foundButton = $true
                Write-Host "  Found button by alternate ID: $id" -ForegroundColor Green
                break
            }
        }
        if (-not $foundButton) {
            Write-Host "  No button found by alternate IDs" -ForegroundColor Red
            return $false
        }

    }
    else {
        Write-Host "  Button found by ID: $buttonId" -ForegroundColor Green
    }

    try {
        for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
            if ($attempt -gt 1) {
                Write-Host "  Attempt $attempt of $maxAttempts..." -ForegroundColor Yellow
                Start-Sleep -Milliseconds $delayBetweenAttempts
            }

            $clickSuccess = $false

            # Get the text to confirm it's the right button
            $btnTextBuilder = New-Object System.Text.StringBuilder 256
            [WindowsAPI]::GetWindowText($btnHwnd, $btnTextBuilder, 256) | Out-Null
            $btnText = $btnTextBuilder.ToString()
            
            Write-Host "  Found button with ID $buttonId, text: '$btnText'" -ForegroundColor Gray
            
            # First method: use SendMessage with BM_CLICK
            [DialogAPI]::SendMessage($btnHwnd, [DialogAPI]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            Write-Host "  Clicked button with ID $buttonId using SendMessage" -ForegroundColor Green
            
            # Add a small delay to allow the click to process
            Start-Sleep -Milliseconds 100
            
            # Check if dialog was dismissed
            if (-not [WindowsAPI]::IsWindow($hwnd)) {
                Write-Host "  Dialog successfully dismissed!" -ForegroundColor Green
                return $true
            }

            # Try to click the button a second time in case the first attempt failed
            [DialogAPI]::SendMessage($btnHwnd, [DialogAPI]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            Write-Host "  Clicked button with ID $buttonId again using SendMessage" -ForegroundColor Green

            # Check if dialog was dismissed
            if (-not [WindowsAPI]::IsWindow($hwnd)) {
                Write-Host "  Dialog successfully dismissed!" -ForegroundColor Green
                return $true
            }
            
            # Second method: try PostMessage as an alternative
            Add-Type -TypeDefinition @"
            using System;
            using System.Runtime.InteropServices;
            
            public static class PostMessageAPI {
                [DllImport("user32.dll", CharSet = CharSet.Auto)]
                public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
                
                public const uint BM_CLICK = 0x00F5;
                public const uint WM_LBUTTONDOWN = 0x0201;
                public const uint WM_LBUTTONUP = 0x0202;
            }
"@ -ErrorAction SilentlyContinue
            
            [PostMessageAPI]::PostMessage($btnHwnd, [PostMessageAPI]::BM_CLICK, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            Write-Host "  Clicked button with ID $buttonId using PostMessage" -ForegroundColor Green
            
            # Add a slightly longer delay
            Start-Sleep -Milliseconds 150
            
            # Final attempt: simulate mouse clicks
            [PostMessageAPI]::PostMessage($btnHwnd, [PostMessageAPI]::WM_LBUTTONDOWN, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            Start-Sleep -Milliseconds 50
            [PostMessageAPI]::PostMessage($btnHwnd, [PostMessageAPI]::WM_LBUTTONUP, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            Write-Host "  Clicked button with ID $buttonId using simulated mouse click" -ForegroundColor Green
            
            # Mark as successful click attempt
            $clickSuccess = $true

            # Check if dialog was dismissed
            Start-Sleep -Milliseconds 200
            if (-not [WindowsAPI]::IsWindow($hwnd)) {
                Write-Host "  Dialog successfully dismissed!" -ForegroundColor Green
                return $true
            }
           
            # If we reached this point and made a click attempt but the dialog is still open,
            # try again in the next iteration
            if ($clickSuccess) {
                Write-Host "  Button was clicked but dialog remains open, will try again..." -ForegroundColor Yellow
            }
        }
        
        # If we reach here, it means we couldn't dismiss the dialog
        Write-Host "  FAILED to dismiss dialog after $maxAttempts attempts" -ForegroundColor Red
        return $false
        
    }
    catch {
        Write-Host "  CRITICAL ERROR in Dismiss-Dialog: $_" -ForegroundColor Red
        return $false
    }
    finally {
        Write-Host "-----------------------------------------------------" -ForegroundColor Cyan
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
            
            # Automatically dismiss dialogs
            foreach ($dialog in $dialogResult.DialogWindows) {
                # Decide which button to click based on the dialog type
                $buttonToClick = "OK"  # Default
                
                if ($dialog.Buttons -like "*OK*") {
                    $buttonToClick = "OK"
                }
                elseif ($dialog.Buttons -like "*Yes*") {
                    $buttonToClick = "Yes"
                }
                elseif ($dialog.Buttons -like "*No*") {
                    $buttonToClick = "No"  # You might want to choose differently
                }
                elseif ($dialog.Buttons -like "*Cancel*") {
                    $buttonToClick = "Cancel"
                }
                
                Write-Host "Attempting to dismiss dialog: $($dialog.Title) by clicking $buttonToClick"
                $dismissed = Dismiss-Dialog -hwnd $dialog.Handle -buttonToClick $buttonToClick
                
                if ($dismissed) {
                    Write-Host "Successfully dismissed dialog: $($dialog.Title)" -ForegroundColor Green
                }
                else {
                    Write-Host "Failed to dismiss dialog: $($dialog.Title)" -ForegroundColor Red
                }
            }
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