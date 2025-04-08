# Enable access to the VBA project object model and configure macro security settings

function Enable-VBOM ($App) {
  Try {
    # Check if the application registry key exists
    $AppKeyPath = "Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer"
    if (-not (Test-Path $AppKeyPath)) {
      Write-Output "Error: The registry path '$AppKeyPath' does not exist."
      return
    }

    # Retrieve the current version
    $CurVer = Get-ItemProperty -Path $AppKeyPath -ErrorAction Stop
    $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"

    # Define possible paths for AccessVBOM
    $Paths = @(
        "HKCU:\Software\Microsoft\Office\$Version\$App\Security",
        "HKLM:\Software\Microsoft\Office\$Version\$App\Security",
        "HKLM:\Software\WOW6432Node\Microsoft\Office\$Version\$App\Security",
        "HKCU:\Software\Microsoft\Office\$Version\Common\TrustCenter",
        "HKLM:\Software\Microsoft\Office\$Version\Common\TrustCenter"
    )

    # Check each path
    $Found = $false
    foreach ($Path in $Paths) {
        if (Test-Path $Path) {
            Write-Output "Found registry path: $Path"
            # Set the AccessVBOM property
            Set-ItemProperty -Path $Path -Name AccessVBOM -Value 1 -ErrorAction Stop
            Write-Output "Successfully enabled AccessVBOM at $Path."
            $Found = $true
        }
        else {
            Write-Output "Registry path not found: $Path"
        }
    }

    if (-not $Found) {
        Write-Output "Error: None of the registry paths for AccessVBOM were found."
    }

  } Catch {
    Write-Output "Failed to enable access to VBA project object model for $App."
    Write-Output "Error: $($_.Exception.Message)"
    Write-Output "StackTrace: $($_.Exception.StackTrace)"
  }
}

function Enable-AllMacros ($App) {
  Try {
    # Check if the application registry key exists
    $AppKeyPath = "Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer"
    if (-not (Test-Path $AppKeyPath)) {
      Write-Output "Error: The registry path '$AppKeyPath' does not exist."
      return
    }

    # Retrieve the current version
    $CurVer = Get-ItemProperty -Path $AppKeyPath -ErrorAction Stop
    $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"

    # Define possible security paths
    $SecurityPaths = @(
        "HKCU:\Software\Microsoft\Office\$Version\$App\Security",
        "HKLM:\Software\Microsoft\Office\$Version\$App\Security"
    )

    # Check each path
    $Found = $false
    foreach ($Path in $SecurityPaths) {
        if (Test-Path $Path) {
            Write-Output "Found security registry path: $Path"
            # Set VBAWarnings to 1 (Enable all macros)
            Set-ItemProperty -Path $Path -Name VBAWarnings -Value 1 -ErrorAction Stop
            Write-Output "Successfully set macro security level to 'Enable all macros' at $Path."
            $Found = $true
        }
        else {
            Write-Output "Security registry path not found: $Path"
        }
    }

    if (-not $Found) {
        Write-Output "Error: None of the registry paths for macro security settings were found."
    }

  } Catch {
    Write-Output "Failed to modify macro security settings for $App."
    Write-Output "Error: $($_.Exception.Message)"
    Write-Output "StackTrace: $($_.Exception.StackTrace)"
  }
}

# Get the app name from the argument passed to the script
$AppName = $args[0]

if (-not $AppName) {
    Write-Output "Error: No application name specified. Usage: Enable-VBOM.ps1 <ApplicationName>"
    exit 1
}

Write-Output "Enabling VBOM access for $AppName..."
Enable-VBOM $AppName

Write-Output "Setting macro security to enable all macros for $AppName..."
Enable-AllMacros $AppName