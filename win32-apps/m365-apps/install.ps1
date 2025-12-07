# Declare this to be an advanced function (access extra built in pwsh stuff) 
[CmdletBinding()]

# Set the parameters we can call the script with e.g. ( -RemoveExisting )
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

# Begin block, contains intialization code for the script, functions are also
# declared here
begin {
    # Set script preferences
    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'
    $officeInstalled = $false
    $logsModule = Join-Path $LogFolder $PSScriptRoot 'modules\JN-Logger.psm1'
    $logsFolder = Join-Path $env:SystemDrive 'IKP\logs'
    $logsFile = Join-Path $logsFolder 'Install-M365Apps.log'

    Import-Module $logsModule -Force -ErrorAction Stop
    Initialize-Logging -LogPath $logPath -LogLevel INFO -logToFile:$true -logToConsole:$true

    # I'd like to add the device name here, at a later date
    Start-LogSection -Title "Initializing M365Apps Installation"

    #region Functions

    # Creates a new XML file from scratch to use with M365 setup.exe
    function New-InstallXMLFile {
        
        Write-LogInfo -Message "Creating install configuration XML file"

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
                <AppSettings>
                    <User Key="software\microsoft\office\16.0\common" Name="qmenable" Value="0" Type="REG_DWORD" App="office16" Id="L_EnableCustomerExperienceImprovementProgram" />
                    <User Key="software\microsoft\office\16.0\common" Name="updatereliabilitydata" Value="0" Type="REG_DWORD" App="office16" Id="L_UpdateReliabilityPolicy" />
                </AppSettings>
                <Display Level="None" AcceptEULA="TRUE" />
            </Configuration>
"@
        # Save the above as an .xml file
        $Path = Join-Path $OfficeInstallDownloadPath 'Install.xml'
        $InstallXML.Save($Path)
    }    

    # Function: Create an XML file to remove office, to be used with M365 setup.exe
    function New-UninstallXMLFile {

        Write-LogInfo -Message "Creating uninstall configuration XML file"
        
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

    # Function: Return the current URL of the Office Deployment Tool
    function Get-ODTURL {
        [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'
        $MSWebPage | ForEach-Object {
            if ($_ -match 'url=(https://.*officedeploymenttool.*.exe)') {
                Write-LogInfo -Message "Found ODT .exe URL: $($matches[1])"
                $matches[1]
            }
        }
    }

    # Function: Test if this script is running in an elevated context
    function Test-IsElevated {

        Write-LogInfo -Message "Are we elevated?"
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($id)
        if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-LogInfo -Message "Affirmative, admin context detected"
            $true
        } else {
            Write-LogWarning -Message "Admin context not detected"
            $false
        }
    }

    # Function: Test M365 installed or not. Looks in the registry to see if there are
    # keys indicating the presence of one or more M365 products
    function Test-M365Installed {
        Write-LogInfo -Message "Testing for the presence of M365 Apps"

        $UninstallRegKeys = @(
            'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
            'HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        )

        # Create an arraylist, use this to store any relevant reg keys we find
        $InstalledProducts = [System.Collections.ArrayList]::new()
        
        foreach ($Key in (Get-ChildItem $UninstallRegKeys)) {
            if ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
                $Product = $Key.GetValue('DisplayName')
                $InstalledProducts.Add($Product) | Out-Null
            }
        }
       
        if ($InstalledProducts.count -gt 0) {
            Write-LogInfo -Message "M365 Apps found, as follows..."
            
            foreach ($Product in $InstalledProducts) {
                Write-LogInfo -Message "Found: $Product"

            }
            $true
        }
        else {
            Write-LogInfo -Message "Didn't find any M365 Apps"
            $false
        }
    }

    # Function: Remove all M365 apps
    function Remove-M365Apps {
        if (-not(Test-M365Installed)) {
            Write-LogWarning -Message "M365 App removal was requested, however none are present. Continuing..."
            return
        }

        $UninstallXML = New-UninstallXMLFile

        try {
            Write-LogInfo -Message "Attempting to remove M365 Apps, please be patient..."
            $setupFilePath = Join-Path $OfficeInstallDownloadPath 'setup.exe'
            # Write-Host $setupFilePath
            $Arguments = "/configure $UninstallXML"
            $process = Start-Process $setupFilePath -ArgumentList $Arguments -Wait -PassThru
        }
        catch {
            Write-LogError -Message "Error uninstalling M365 Apps. The error was: $_"
            Exit
        }

        if (-not(Test-M365Installed)) {
            Write-LogInfo "Successfully removed M365 Apps"
        }
        else {
            Write-Warning "Couldn't remove M365 Apps; continuing anyway, but don't expect too much..."
        }
    }

    #endregion
}

# Process block, main script functionality goes here
process {

    # Test for admin context, exit if false
    if (-not (Test-IsElevated)) {
        Write-LogError -Message 'Exiting...please try again with elevated permissions'
        Exit 1
    }

    # Create a new working directory
    if (-not (Test-Path $OfficeInstallDownloadPath)) {
        Write-LogInfo -Message "Creating working directory: $OfficeInstallDownloadPath"
        New-Item -ItemType Directory -Path $OfficeInstallDownloadPath | Out-Null
    }

    # Check whether an XML file was specified, if not, create one. if one was specified,
    # check if it exists before proceeding, exit if it's not found
    if (-not($RemoveOnly)) {
        if (-not($ConfigurationXMLFile)) {
            Write-LogInfo -Message "Creating new XML installation configuration"

            New-InstallXMLFile
            $ConfigurationXMLFile = Join-Path $OfficeInstallDownloadPath 'Install.xml'
        } else {
            if (-not(Test-Path $ConfigurationXMLFile)) {
                Write-LogError 'The configuration XML file specified is not a valid file. Please verify that the path is correct and try again.'
                Exit 1
            }
            else {
                Write-LogInfo 'The specified XML file appears to be valid, continuing...'
            }
        }
    }

    # Get the ODT URL
    $ODTInstallerURL = Get-ODTURL

    try {
        Write-LogInfo 'Attempting to download the Office Deployment Tool...'

        $params = @{
            Uri = $ODTInstallerURL
            OutFile = Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe'
        }
        Invoke-WebRequest @params
        Write-LogInfo 'Success!'
    }
    catch {
        Write-LogError -Message "An error occured whilst trying to download the Office Deployment Tool.`nPlease make sure the following link is valid: $ODTInstallerURL"
        Exit 1
    }

    # Extract the Office Deployment Tool
    try {
        Write-LogInfo -Message "Attemping to extract the M365 download tool"
        $ODTSetup = Join-Path $OfficeInstallDownloadPath "ODTSetup.exe"
        $Arguments = "/quiet /extract:$OfficeInstallDownloadPath"
        Start-Process $ODTSetup -ArgumentList $Arguments -Wait
    }
    catch {
        Write-LogError -Message "An error occurred whilst extracting the tool. The error is: $_"
        Exit 1
    }

    # Try to remove M365 apps if the script was called with -RemoveExisting or -RemoveOnly
    if ($RemoveExisting -or $RemoveOnly) {
        Remove-M365Apps
    }

    # Download and install office (as long as the script wasn't called with -RemoveOnly)
    if (-not($RemoveOnly)) {
        try {
            Write-LogInfo -Message "Downloading and installing M365 Apps"
            $setupFilePath = Join-Path $OfficeInstallDownloadPath "setup.exe"
            $arguments = "/configure $ConfigurationXMLFile"
            $process = Start-Process $setupFilePath -ArgumentList $arguments -Wait -PassThru
        }
        catch {
            Write-LogWarning -Message "Error installing M365 Apps. The error is: $_"
        }

        # Test if M365 Apps were installed successfully, store result in $officeInstalled ($true or $false)
        $officeInstalled = Test-M365Installed
    }
    else {
        Write-LogInfo -Message 'No more work to do because -RemoveOnly was specified'
    }
}

end {
    # if cleanup flag was set, remove the installation files
    if ($CleanupInstallFiles) {
        Write-LogInfo -Message "Cleaning up installation directory"
        Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
    }

    # if -RemoveOnly was not set, output the result of the attempted Installation
    if (-not($RemoveOnly)) {
        if ($officeInstalled) {
            Write-LogSuccess -Message "M365 Apps were successfully installed ^_^"
            Exit 0
        }
        else {
            Write-LogWarning -Message "Failed to install M365 Apps. Review the logs and try again"
            Exit 1
        }
    }
}