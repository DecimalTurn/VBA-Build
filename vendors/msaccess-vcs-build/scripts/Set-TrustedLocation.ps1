param(
    [string]$LocationName,
    [string]$TrustedPath
)

$accessVersion = "16.0"  # O365

# Registry
$regPath = "HKCU:\Software\Microsoft\Office\$accessVersion\Access\Security\Trusted Locations\$LocationName"

If (-Not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

New-ItemProperty -Path $regPath -Name "Path" -Value $TrustedPath -PropertyType String -Force | Out-Null
New-ItemProperty -Path $regPath -Name "AllowSubfolders" -Value 1 -PropertyType DWord -Force | Out-Null
