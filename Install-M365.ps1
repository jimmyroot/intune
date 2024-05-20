[CmdletBinding()]
param(
    # Specify this to use a pre-existing XML file
    [String]
    $ConfigurationXMLFile,
    [String]
    $OfficeInstallDownloadPath = 'C:\IKP\M365Install\',
    [Switch]
    $RemoveExisting,
    [Switch]
    $RemoveOnly,
    [Switch]
    $CleanupInstallFiles
)

begin {
    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'

    # Creates an XML file for a new installation
    function New-InstallXMLFile {
        Write-Verbose 'Creating new XML file'
        $InstallXML = [XML]@"
            <Configuration ID="41d6b959-8508-4e35-ba3d-362dd9fc1de7">
                <Info Description="IKP Standard M365 Apps Install" />
                <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" MigrateArch="TRUE">
                    <Product ID="O365BusinessRetail">
                        <Language ID="en-gb" />
                        <ExcludeApp ID="Groove" />
                        <ExcludeApp ID="Lync" />
                        <ExcludeApp ID="OneDrive" />
                        <ExcludeApp ID="Publisher" />
                    </Product>
                    <Product ID="ProofingTools">
                        <Language ID="da-dk" />
                        <Language ID="nl-nl" />
                        <Language ID="fr-fr" />
                        <Language ID="de-de" />
                        <Language ID="lb-lu" />
                        <Language ID="sv-se" />
                    </Product>
                </Add>
                <Updates Enabled="TRUE" />
                <RemoveMSI />
                <AppSettings>
                    <User Key="software\microsoft\office\16.0\common" Name="qmenable" Value="0" Type="REG_DWORD" App="office16" Id="L_EnableCustomerExperienceImprovementProgram" />
                    <User Key="software\microsoft\office\16.0\common" Name="updatereliabilitydata" Value="0" Type="REG_DWORD" App="office16" Id="L_UpdateReliabilityPolicy" />
                </AppSettings>
                <Display Level="None" AcceptEULA="TRUE" />
            </Configuration>
"@
        # Save the above as .xml file
        $Path = Join-Path $OfficeInstallDownloadPath 'Install.xml'
        $InstallXML.Save($Path)
    }    

    # Creates an XML file to remove any pre-existing installation
    function New-UninstallXMLFile {
        Write-Verbose 'Creating Uninstall.xml file...'
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

    # Get the URL of the Office Deployment Tool
    function Get-ODTURL {
        [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'
        $MSWebPage | ForEach-Object {
            If ($_ -match 'url=(https://.*officedeploymenttool.*.exe)') {
                $matches[1]
            }
        }
    }

    # Test if we are executing in elevated context
    function Test-IsElevated {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($id)
        if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Output $true
        } else {
            Write-Output $false
        }
    }

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
            Write-Verbose '...apps found, listing them now...'
            foreach ($Product in $InstalledProducts) {
                Write-Host "Discovered $Product"
            }
            return $true
        }
        Else {
            Write-Warning '...none found, continuing...'
            return  $false
        }
    }

    function Remove-M365Apps {
        If (-not(Test-M365Installed)) {
            Write-Warning 'M365 App removal was requested, however none are present. Continuing...'
            return
        }

        $UninstallXML = New-UninstallXMLFile

        try {
            Write-Verbose 'Attempting to remove M365 Apps, please be patient...'
            $Setup = Join-Path $OfficeInstallDownloadPath 'setup.exe'
            Write-Host $Setup
            $Arguments = "/configure $UninstallXML"
            $Silent = Start-Process $Setup -ArgumentList $Arguments -Wait -PassThru
        }
        catch {
            Write-Warning 'Error running Office365 Uninstaller. The error is below:'
            Write-Warning $_
            exit 1
        }

        If (-not(Test-M365Installed)) {
            Write-Verbose 'Office removal was successful, continuing...'
        }
        Else {
            Write-Warning 'Unable to install M365 Apps. Continuing anyway, but will probably fail...'
        }
    }
}

process {

    # Test for admin
    If (-not (Test-IsElevated)) {
        Write-Error -Message 'Non-elevated context detected. Please re-run this script with elevated permissions.'
    }

    # Create our new working directory
    If (-not (Test-Path $OfficeInstallDownloadPath)) {
        New-Item -ItemType Directory -Path $OfficeInstallDownloadPath | Out-Null
    }

    # Check whether an XML file was specified, if not, create one. If one was specified,
    # check if it exists
    If (-not($ConfigurationXMLFile)) {
        New-InstallXMLFile
    } else {
        If (-not(Test-Path $ConfigurationXMLFile)) {
            Write-Warning 'The configuration XML file specified is not a valid file.'
            Write-Warning 'Please verify that the path is correct and try again.'
            Exit 1
        }
        Else {
            Write-Verbose 'The specified XML file appears to be valid, continuing...'
        }
    }

    # This will work, but only if the supplied XML file is called 'Install.xml' - needs testing
    $ConfigurationXMLFile = Join-Path $OfficeInstallDownloadPath 'Install.xml'
    $ODTInstallerURL = Get-ODTURL

    # Download the ODT Installer
    Write-Verbose 'Downloading the Office Deployment Tool...'

    try {
        $params = @{
            Uri = $ODTInstallerURL
            OutFile = Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe'
        }
        Invoke-WebRequest @params
    }
    catch {
        Write-Warning 'An error occured whilst trying to download the Office Deployment Tool.'
        Write-Warning 'Please verify that the below link is valid:'
        Write-Warning $ODTInstallerURL
        exit 1
    }

    # Run the ODT
    try {
        Write-Verbose 'Running the Office Deployment Tool...'
        $ODTSetup = Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe'
        $Arguments = "/quiet /extract:$OfficeInstallDownloadPath"
        Start-Process $ODTSetup -ArgumentList $Arguments -Wait
    }
    catch {
        Write-Warning 'An error occurred whilst running the Office Deployment tool. The error is below:'
        Write-Warning $_
        exit 1
    }

    # Here's where we remove any existing installations, if a flag was set
    If ($RemoveExisting -or $RemoveOnly) {
        Remove-M365Apps
    }

    # Run the M365 Apps installer
    If (-not($RemoveOnly)) {
        try {
            Write-Verbose 'Downloading and installing Microsoft 365 Apps for business...'
            $Setup = Join-Path $OfficeInstallDownloadPath 'setup.exe'
            $Arguments = "/configure $ConfigurationXMLFile"
            $Silent = Start-Process $Setup -ArgumentList $Arguments -Wait -PassThru
        }
        catch {
            Write-Warning 'Error running the Office install. The error is below:'
            Write-Warning $_
        }

        # Test if M365 Apps were installed successfully
        $OfficeWasInstalled = Test-M365Installed
    }
    Else {
        Write-Verbose 'No more work to do because -RemoveOnly was specified'
    }
}

end {
    If ($CleanupInstallFiles) {
        Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
    }

    If (-not($RemoveOnly)) {
        If ($OfficeWasInstalled) {
            Write-Host "Success! Selected products were successfully installed"
            Exit 0
        }
        Else {
            Write-Warning "Failed. Office was not installed. Please troubleshoot and try again"
            Exit 1
        }
    }
    Else {
        Exit 0
    }
}
