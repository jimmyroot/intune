$officeC2RKey = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
$productId = 'O365ProPlusRetail'
$isInstalled = $false

if (Test-Path $officeC2RKey) {
    try {
        $isInstalled = (Get-ItemPropertyValue -Path $officeC2RKey -Name ProductReleaseIds).contains($productId)
    }
    catch {
        # If we can't read the key just fail
        Exit 1
    }
}
if ($IsInstalled) {
    Write-Output "[SUCCESS]: Microsoft 365 Apps '$productId' detected."
    Exit 0
}
else {
    # If the key isn't what we expect, fail
    # Write-Output "FAILURE: Microsoft 365 Apps not detected."
    Exit 1
}
