param(
    [string]$LocationName
)

$accessVersion = "16.0"  # O365

# Registry
$regPath = "HKCU:\Software\Microsoft\Office\$accessVersion\Access\Security\Trusted Locations\$LocationName"

If (-Not (Test-Path $regPath)) {
    exit 0 # nothing to do, location not set
}

# remove trusted location
Remove-Item -Path $regPath -Force -Recurse | Out-Null