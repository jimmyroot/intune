$VerbosePreference = 'Continue'
$ErrorActionPreference = 'Stop'
$Working = Split-Path $MyInvocation.MyCommand.Path -Parent
$OfficeInstallDownloadPath = 'C:\IKP\M365Install\'

function New-UninstallXMLFile {
    Write-Host 'Creating Uninstall.xml file...'
    $XML = [XML]@"
        <Configuration>
            <Remove All="TRUE"/>
            <Display Level="None" AcceptEULA="TRUE"/>
            <Property Name="AUTOACTIVATE" Value="0" />
            <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
            <Property Name="SharedComputerLicensing" Value="0"/>
            <Property Name="PinIconsToTaskbar" Value="FALSE"/>
        </Configuration>
"@

    $Path = Join-Path $OfficeInstallDownloadPath "Uninstall.xml"
    $XML.Save($Path)
    $Path
}

# We'll use the same working dir as our installer script
If (-not (Test-Path $OfficeInstallDownloadPath)) {
    New-Item -ItemType Directory -Path $OfficeInstallDownloadPath | Out-Null
}

$UninstallXML = New-UninstallXMLFile

try {
    $Setup = Join-Path $OfficeInstallDownloadPath 'setup.exe'
    Write-Host $Setup
    $Arguments = "/configure $UninstallXML"
    Start-Process $Setup -ArgumentList $Arguments -Wait -PassThru
}
catch {
    Write-Warning 'Error running Office365 Uninstaller. The error is below:'
    Write-Warning $_
}