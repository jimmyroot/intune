$officeC2RKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$productId = 'O365ProPlusRetail'
$isInstalled = $false
if (Test-Path $OfficeC2RKey) {
    try {
        # 2. Search for the M365 specific ProductReleaseIds value
        $InstalledProducts = Get-ChildItem -Path $OfficeC2RKey | 
            Get-ItemProperty | 
            Where-Object { $_.ProductReleaseIds -like "*$M365ProductReleaseId*" }
            
        # 3. Check if any matching product was found
        if ($InstalledProducts -ne $null -and $InstalledProducts.Count -gt 0) {
            $IsInstalled = $true
        }
    }
    catch {
        # Catch any errors (e.g., registry read issues) and treat it as not detected.
        # Intune will rely on the Exit 1 code below.
    }
}

# --- Intune Detection Logic ---

if ($IsInstalled) {
    # 🌟 Mandatory Intune Success Requirements 🌟
    
    # 1. Write output to STDOUT
    Write-Output "SUCCESS: Microsoft 365 Apps ($M365ProductReleaseId) detected."
    
    # 2. Return Exit Code 0 (Success)
    Exit 0
}
else {
    # 🛑 Mandatory Intune Failure Requirement 🛑
    
    # (Optional, but recommended for logging) Write output to STDOUT
    Write-Output "FAILURE: Microsoft 365 Apps not detected."
    
    # 1. Return a Non-Zero Exit Code (Failure)
    Exit 1
}
