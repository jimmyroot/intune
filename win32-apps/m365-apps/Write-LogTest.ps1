# Write-LogTest.ps1

$OfficeInstallDownloadPath = 'C:\IKP\M365Install'
$LogFolder = Join-Path $OfficeInstallDownloadPath 'Logs'
$LogFileName = (Get-Date -UFormat "%d-%m-%Y") + ".log"
$LogPath = Join-Path $LogFolder $LogFileName

Function Write-Log {

    param(
        [Paramater(Mandatory=$True)]
        [array]$LogOutput,
        [Paramater(Mandatory=$True)]
        [string]$Path
    )

    $currentDate = (Get-Date -UFormat "%d-%m-%T")
    $currentTime = (Get-Date -UFormat "%T")
    $logOutput = $logOutput -join (" ")
    "[$currentDate $currentTime] $logOutput" | Out-File $Path -Append

}