# Detection script for Egnyte Drive mappings

$expectedDriveCount = 2
$drivesAreMapped = $false
$regexToMatchSID = "S-1-\d\d?-\d\d?-\d{8,10}-\d{8,10}-\d{8,10}-\d{5,10}$"

function Connect-HKU {
    try {
        New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -ErrorAction Stop -Scope global | Out-Null
    }
    catch {
        $false
    }
    
    Write-Host 'Connected HKU...' -ForegroundColor Green
    $true
}

function Disconnect-HKU {
    try {
        Remove-PSDrive HKU -ErrorAction Stop
        Write-Host 'HKU disconnected...' -ForegroundColor Green
    }
    catch {
        Write-Host 'HKU is already disconnected, nothing to do...' -ForegroundColor DarkYellow
    }
}

function Evaluate-Drives {
    Param (
        [int]$DriveCount
    )

    If ($driveCount -eq $expectedDriveCount) {
        $true
        return
    }
    $false
}

$PSDriveStatus = Connect-HKU

# Check Drive connected successfully we can check the drives
If ($PSDriveStatus) {
    
    $activeUserHives = Get-ChildItem -Path HKU: -ErrorAction SilentlyContinue | Where-Object {$_.Name -match $regexToMatchSID}
    
    ForEach ($hive in $activeUserHives) {
         $pathToUserHive = $hive.PSPath
         $pathToEgnyteDrives = 'Software\Egnyte\Egnyte Drive\'
         $path = Join-Path $pathToUserHive $pathToEgnyteDrives

         If (Test-Path $path) {
             $drivesAreMapped = Evaluate-Drives -DriveCount (Get-ChildItem $path).count
             break
        }
    }
    
    Disconnect-HKU

    If ($drivesAreMapped) {
        Write-Host 'Success! Egnyte drives are mapped'
        Exit 0
    }
    Else {
        Write-Host 'Incorrect number of drive mappings detected. Please contact support@ikpartners.com for assistance.' -ForegroundColor Red
        Exit 1
    }
}
Else {
    Write-Host 'The connection to the registry could not be established. Please contact support@ikpartners.com for assistance.' -ForegroundColor Red
}
