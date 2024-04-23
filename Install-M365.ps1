[CmdletBinding()]
param(
    # Specify this to use a pre-existing XML file
    [String]
    $ConfigurationXMLFile,
    [String]
    $OfficeInstallDownloadPath = 'C:IKP\M365Install',
    [Switch]
    $CleanupInstallFiles = $false
)

begin {

    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'

    # Creates an XML file for a new installation
    function New-InstallXMLFile {

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
        $principal = [New-Object System.Security.Principal.WindowsPrincipal]($id)
    }

}

process {

    # Test for admin
    If (-not (Test-IsElevated)) {
        Write-Error -Message "Please re-run this script with Administrator privileges."
    }

}

end {

}