function Test-M365Installed {
    $UninstallRegKeys = @(
        'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
        'HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
    )

    $InstalledProducts = [System.Collections.ArrayList]::new()

    foreach ($Key in (Get-ChildItem $UninstallRegKeys)) {
        # Write-Host ($Key.GetValue('DisplayName') -like '*Microsoft 365*')
        If ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
            $Product = $Key.GetValue('DisplayName')
            $InstalledProducts.Add($Product) | Out-Null
        }
    }
   
    If ($InstalledProducts.count -gt 0) {
        foreach ($Product in $InstalledProducts) {
            Write-Host "$Product was installed successfully"
        }
        return $true
    }
    Else {
        #  Write-Host 'No M365 products found, please troubleshoot and try again'
        return  $false
    }
}

$OfficeWasInstalled = Test-M365Installed
