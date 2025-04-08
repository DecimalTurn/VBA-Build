# Enable access to the VBA project object model

function List-RegistrySubKeysRecursively ($Path) {
    if (-not (Test-Path $Path)) {
        Write-Output "Error: The registry path '$Path' does not exist."
        return
    }

    Write-Output "Subkeys and entries under '$Path':"
    Try {
        # List all subkeys
        $SubKeys = Get-ChildItem -Path $Path
        foreach ($SubKey in $SubKeys) {
            Write-Output " - Subkey: $($SubKey.PSChildName)"

            # List all registry entries (values) for the current subkey
            Try {
                $Entries = Get-ItemProperty -Path $SubKey.PSPath
                foreach ($Entry in $Entries.PSObject.Properties) {
                    Write-Output "   - Entry: $($Entry.Name) = $($Entry.Value)"
                }
            } Catch {
                Write-Output "   Error accessing entries for subkey '$($SubKey.PSChildName)': $($_.Exception.Message)"
            }

            # Recursively call the function for each subkey
            List-RegistrySubKeysRecursively $SubKey.PSPath
        }
    } Catch {
        Write-Output "Error accessing subkeys under '$Path': $($_.Exception.Message)"
    }
}

function Enable-VBOM ($App) {
  Try {
    # Step 1: Check if the application registry key exists
    $AppKeyPath = "Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer"
    if (-not (Test-Path $AppKeyPath)) {
      Write-Output "Error: The registry path '$AppKeyPath' does not exist."
      return
    }

    # Step 2: Retrieve the current version
    $CurVer = Get-ItemProperty -Path $AppKeyPath -ErrorAction Stop
    $Version = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"

    # # Step 3: Check if the Office version registry key exists
    # $OfficePath = "HKCU:\Software\Microsoft\Office"
    # if (-not (Test-Path $OfficePath)) {
    #     Write-Output "Error: The registry path '$OfficePath' does not exist."
    #     return
    # }


        # Recursively list all subkeys under the Office version key
        #List-RegistrySubKeysRecursively $OfficePath

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

Enable-VBOM "Excel"