[CmdletBinding()]
param(
    # Specify this to use a pre-existing XML file
    [String]
    $ConfigurationXMLFile,
    [String]
    $OfficeInstallDownloadPath = 'C:\IKP\M365Install',
    [Switch]
    $CleanupInstallFiles = $false
)

begin {

    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'

    # Creates an XML file for a new installation
    function New-InstallXMLFile {
        Write-Host 'Creating new XML file'
        $installXML = [XML]@"
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
        $installXML.Save("$OfficeInstallDownloadPathInstall.xml")
    }    

    # Creates an XML file to remove any pre-existing installation
    function New-UninstallXMLFile {
        
    }

    # Get the URL of the Office Deployment Tool
    function Get-ODTURL {
    
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

    If (!($ConfigurationXMLFile)) {
        New-InstallXMLFile
    } else {
        If (!(Test-Path $ConfigurationXMLFile)) {
            Write-Warning 'The configuration XML file specified is not a valid file.'
            Write-Warning 'Please verify that the path is correct and try again.'
            Exit 1
        }
    }

}

end {

}
