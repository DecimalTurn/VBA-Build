param(
    [string]$AccessFile,
    [string]$ConfigFile,
    # All (default): run every step. PreCompile: module/reference removal and
    # procedures (require an editable .accdb). PostCompile: database properties
    # only (set on the compiled .accde so they do not alter the source .accdb).
    [ValidateSet('All', 'PreCompile', 'PostCompile')]
    [string]$Stage = 'All'
)

enum PropertyType {
    Text = 10
    Integer = 3
    Long = 4
    Boolean = 1
    DateTime = 8
}

function Set-DbProperty{
    param (
        [System.Object]$db,
        [string]$PropertyName,
        [int32]$PropertyType,
        [string]$PropertyValue,
        [bool]$RemoveProperty = $false  
    )

    try {
        $db.Properties[$PropertyName].Value = $PropertyValue    
    }
    catch  {
        # Handle the exception if the property does not exist
        # error 3270
        # The property does not exist, so we create it
        $errorCode = $null
        if ($_.Exception.InnerException -and $_.Exception.InnerException.ErrorCode) {
            $errorCode = $_.Exception.InnerException.ErrorCode
        } elseif ($_.Exception.HResult) {
            $errorCode = $_.Exception.HResult
        }
        $errorMsg = $_.Exception.Message

        # Prüfe auf Fehlercode oder auf typische Fehlermeldung
        if ($errorCode -eq -2146825018 -or $errorMsg -like "*Property not found*") {
            Write-Host "Property '$PropertyName' does not exist. Creating it."
            $db.Properties.Append($db.CreateProperty($PropertyName, $PropertyType, $PropertyValue))
        } else {
            Write-Error "An unexpected error occurred: $($errorCode)  $errorMsg"
            return
        }
    }
}

function Invoke-Procedure {
    [CmdletBinding()]   
    param (
        [System.Object]$access,
        [string]$ProcedureName,
        [Parameter(ValueFromRemainingArguments = $true)]
        [object[]]$Arguments
    )

    if (-not $access) {
        Write-Error "Access application object is null."
        return
    }
    if (-not $ProcedureName) {
        Write-Error "Procedure name is null or empty."
        return
    }

    $ArgCount = $Arguments.Count    

    switch ($ArgCount) {   
        0 { 
            $null = $access.Run($ProcedureName)
        }   
        1 { 
            $null = $access.Run($ProcedureName, [ref] $Arguments[0])
        }   
        2 { 
            $null = $access.Run($ProcedureName, [ref] $Arguments[0], [ref] $Arguments[1])
        }
        3 { 
            $null = $access.Run($ProcedureName, [ref] $Arguments[0], [ref] $Arguments[1], [ref] $Arguments[2])
        }
        Default {
            # raise error if more than 3 arguments are passed
            Write-Error "Procedure '$ProcedureName' expects at most 3 arguments, but $ArgCount were provided."
            return
        }
    }

}

function Remove-VbaModules {
    param (
        [System.Object]$vbProject,
        [string[]]$Patterns
    )

    if (-not $Patterns -or $Patterns.Count -eq 0) {
        return
    }

    # VBComponent types safe to remove: 1 = standard module, 2 = class module,
    # 3 = MSForm. Type 100 is a document module (a form or report code-behind)
    # and must never be removed this way, so a pattern that matches a
    # form/report name is skipped rather than deleted.
    $removableTypes = @(1, 2, 3)

    # Collect matching names first to avoid modifying the collection while iterating
    $componentsToRemove = @()
    foreach ($component in $vbProject.VBComponents) {
        foreach ($pattern in $Patterns) {
            if ($component.Name -like $pattern) {
                if ($removableTypes -contains $component.Type) {
                    $componentsToRemove += $component.Name
                }
                else {
                    Write-Host "Skipping '$($component.Name)' (matched '$pattern' but is a document module, type $($component.Type) - a form/report code-behind)"
                }
                break
            }
        }
    }

    foreach ($name in $componentsToRemove) {
        try {
            $component = $vbProject.VBComponents.Item($name)
            $vbProject.VBComponents.Remove($component)
            Write-Host "Removed module '$name'"
        }
        catch {
            Write-Host "Warning: Could not remove module '$name': $($_.Exception.Message)"
        }
    }

    if ($componentsToRemove.Count -eq 0) {
        Write-Host "No modules matched the removal patterns."
    }
}

function Remove-VbaReferences {
    param (
        [System.Object]$vbProject,
        [string[]]$ReferenceNames
    )

    if (-not $ReferenceNames -or $ReferenceNames.Count -eq 0) {
        return
    }

    foreach ($refName in $ReferenceNames) {
        $found = $false
        foreach ($ref in $vbProject.References) {
            if ($ref.Name -eq $refName) {
                try {
                    $vbProject.References.Remove($ref)
                    Write-Host "Removed reference '$refName'"
                    $found = $true
                }
                catch {
                    Write-Host "Warning: Could not remove reference '$refName': $($_.Exception.Message)"
                }
                break
            }
        }
        if (-not $found) {
            Write-Host "Reference '$refName' not found (may already be absent)."
        }
    }
}

function SafeReleaseComObject($comObject) {
    if ($null -ne $comObject -and $comObject -is [System.__ComObject]) {
        [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($comObject)
    }
}


# read config file
if (-not $ConfigFile) {
    $ConfigFile = Join-Path -Path (Get-Location) -ChildPath "config.json"
}

if (
    -not ([System.IO.Path]::IsPathRooted($ConfigFile))
) {
    $ConfigFile = Join-Path -Path (Get-Location) -ChildPath $ConfigFile.TrimStart('\','/','.')
}

if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Config file not found: $ConfigFile"
    exit 1
}   
$config = Get-Content -Path $ConfigFile | ConvertFrom-Json


[string]$fullPath = $AccessFile
if (-not ([System.IO.Path]::IsPathRooted($fullPath))) {
    $fullPath = Join-Path -Path (Get-Location) -ChildPath $fullPath.TrimStart('\','/','.')
}

[object]$access = $null
[object]$db = $null

Write-Host "Config $fullPath"

try {
    
    $access = New-Object -ComObject Access.Application
    $access.OpenCurrentDatabase($fullPath)

# Pre-compile step: module/reference removal needs an editable .accdb (VBA cannot be
    # changed in a compiled .accde).
    $runPreCompile = ($Stage -eq 'All' -or $Stage -eq 'PreCompile')
    # Post-compile steps: procedures and database properties run on the compiled .accde,
    # matching the behavior of the last release (and leaving the source .accdb unaltered).
    $runPostCompile = ($Stage -eq 'All' -or $Stage -eq 'PostCompile')

# Remove VBA modules matching name patterns (e.g. test modules)
    if ($runPreCompile -and $config.RemoveModules -and $config.RemoveModules.Count -gt 0) {
        Write-Host "Removing VBA modules matching patterns: $($config.RemoveModules -join ', ')"
        $vbProject = $access.VBE.ActiveVBProject
        Remove-VbaModules -vbProject $vbProject -Patterns $config.RemoveModules
    }
    elseif ($runPreCompile) {
        Write-Host "No modules to remove."
    }

# Remove VBA references by name (e.g. Rubberduck)
    if ($runPreCompile -and $config.RemoveReferences -and $config.RemoveReferences.Count -gt 0) {
        Write-Host "Removing VBA references: $($config.RemoveReferences -join ', ')"
        $vbProject = $access.VBE.ActiveVBProject
        Remove-VbaReferences -vbProject $vbProject -ReferenceNames $config.RemoveReferences
    }
    elseif ($runPreCompile) {
        Write-Host "No references to remove."
    }

# Run procedures from config
    if ($runPostCompile -and $config.Procedures -and $config.Procedures.Count -gt 0) {
        
        foreach ($procedure in $config.Procedures) {
            if (-not $procedure.Name) {
                Write-Error "Procedure name is missing in the configuration."
                continue
            }
            $Parameters = if ($procedure.PSObject.Properties.Match('Parameters')) { $procedure.Parameters } else { @() }
            if (-not $Parameters) {
                $Parameters = @()  # Default to empty array if no parameters are defined
            }
            if ($Parameters -and $Parameters.Count -gt 0) {
                Write-Host "Running procedure '$($procedure.Name)' with parameters: $($Parameters -join ', ')"
            } else {
                Write-Host "Running procedure '$($procedure.Name)'"
            }
            Invoke-Procedure -access $access -ProcedureName $procedure.Name -Arguments $Parameters    
        }
    }
    elseif ($runPostCompile) {
        Write-Host "No procedures to run."
    }

# Set database properties from config
    if ($runPostCompile -and $config.DatabaseProperties -and $config.DatabaseProperties.Count -gt 0) {
        
        $db = $access.CurrentDb()

        foreach ($property in $config.DatabaseProperties) {
            $propertyName = $property.Name
            $propertyType = [PropertyType]::Parse([PropertyType], $property.Type)
            $propertyValue = $property.Value

            Write-Host "Setting property '$propertyName' of type '$($propertyType)' to '$propertyValue'"
            Set-DbProperty -db $db -PropertyName $propertyName -PropertyType $propertyType -PropertyValue $propertyValue
        }
    }
    elseif ($runPostCompile) {
        Write-Host "No database properties to set."
    }
}
catch {
    $errorCode = $null
    if ($_.Exception.InnerException -and $_.Exception.InnerException.ErrorCode) {
        $errorCode = $_.Exception.InnerException.ErrorCode
    } elseif ($_.Exception.HResult) {
        $errorCode = $_.Exception.HResult
    }
    $errorMsg = $_.Exception.Message
    Write-Error "An error occurred while setting properties: $($errorCode)  $errorMsg"
    exit 1
}
finally {
    if ($db) {
        SafeReleaseComObject $db
        Remove-Variable -Name db -ErrorAction SilentlyContinue
        
    }
    if ($access) {
        $access.CloseCurrentDatabase()
        $access.Quit()
        SafeReleaseComObject $access
        Remove-Variable -Name access -ErrorAction SilentlyContinue
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}