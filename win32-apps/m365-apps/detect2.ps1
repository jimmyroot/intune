$officeC2RKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$productId = 'O365ProPlusRetail'
$isInstalled = $false

if (Test-Path $officeC2RKey) {
    try {
        $InstalledProducts = Get-ChildItem -Path $OfficeC2RKey | 
            Get-ItemProperty | 
            Where-Object { $_.ProductReleaseIds -like "*$productId*" }
            
        if ($InstalledProducts -ne $null -and $InstalledProducts.Count -gt 0) {
            $isInstalled = $true
        }
    }
    catch {
        # Catch any errors (e.g., registry read issues) and treat it as not detected.
        # Intune will rely on the Exit 1 code below.
    }
}
if ($IsInstalled) {
    Write-Output "SUCCESS: Microsoft 365 Apps ($M365ProductReleaseId) detected."
    Exit 0
}
else {
    # Write-Output "FAILURE: Microsoft 365 Apps not detected."
    Exit 1
}
