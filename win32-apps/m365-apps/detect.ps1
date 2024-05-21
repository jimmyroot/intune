function Test-M365Installed {
    Write-Verbose 'Checking for M365 Apps...'

    $UninstallRegKeys = @(
        'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
        'HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
    )

    $InstalledProducts = [System.Collections.ArrayList]::new()

    foreach ($Key in (Get-ChildItem $UninstallRegKeys)) {
        If ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
            $Product = $Key.GetValue('DisplayName')
            $InstalledProducts.Add($Product) | Out-Null
        }
    }
   
    If ($InstalledProducts.count -gt 0) {
        Write-Host 'Success'
        Exit 0
    }
    Else {
        Write-Host 'Failure'
        Exit 1
    }
}

Test-M365Installed