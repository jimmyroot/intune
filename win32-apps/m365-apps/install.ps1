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

# Begin block, contains intialization code for the script, functions are also declared here
begin {

    $VerbosePreference = 'Continue'
    $ErrorActionPreference = 'Stop'
    $officeInstalled = $false
    $logsModule = Join-Path $PSScriptRoot 'modules\JN-Logger.psm1'
    $logsFolder = Join-Path $env:SystemDrive 'IKP\logs'
    $logsFile = Join-Path $logsFolder 'Install-M365Apps.log'

    Import-Module $logsModule -Force -ErrorAction Stop
    Initialize-Logging -LogPath $logsFile -LogLevel INFO -logToFile:$true -logToConsole:$true

    # I'd like to add the device name here, at a later date
    Start-LogSection -Title "Initializing M365Apps Installation"

    #region Functions

    function New-InstallXMLFile {
        
        Write-LogInfo -Message "Creating new ODT XML config file for install"

        $InstallXML = [XML]@"
            <Configuration ID="41d6b959-8508-4e35-ba3d-362dd9fc1de7">
                <Info Description="IKP Standard M365 Apps Install" />
                <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" MigrateArch="TRUE">
                    <Product ID="O365ProPlusRetail">
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
        $Path = Join-Path $OfficeInstallDownloadPath 'Install.xml'
        $InstallXML.Save($Path)
    }    

    function New-UninstallXMLFile {

        Write-LogInfo -Message "Creating ODT XML config file for uninstall operation"
        
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

    function Get-ODTURL {
        [String]$MSWebPage = Invoke-RestMethod "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
        $MSWebPage | ForEach-Object {
            if ($_ -match '"url":"(https://download\.microsoft\.com/[^"]*/officedeploymenttool[^"\s]*\.exe)"') {
                Write-LogInfo -Message "Found ODT .exe URL: $($matches[1])"
                $matches[1]
            }
            else {
                Write-Host "ODT tool not found"
            }
        }
    }

    # Function: Test if this script is running in an elevated context
    function Test-IsElevated {

        Write-LogInfo -Message "Testing for elevation..."
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($id)
        if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-LogInfo -Message "...success, we are admin, let's go!"
            $true
        } else {
            Write-LogWarning -Message "...uh-oh, we don't seem to be running as admin"
            $false
        }
    }

    # Function: Test M365 installed or not. Looks in the registry to see if there are
    # keys indicating the presence of one or more M365 products. Retry the specified number
    # of times, if not found.
    function Test-M365Installed {
        [CmdletBinding()]
        param() 

        $maxAttempts = 10
        $secondsToSleep = 6
        $attempt = 0

        Write-LogInfo -Message "Starting M365 Apps presence test with up to $maxAttempts retries"

        $uninstallRegKeys = @(
            'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
            'HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        )

        $installedProducts = @()
        
        do {
            $Attempt++
            Write-LogInfo -Message "Attempt $attempt of $maxRetries; Searching 'Uninstall' for evidence of M365 Apps"

            $installedProducts = @(
                foreach ($key in (Get-ChildItem -Path $uninstallRegKeys -ErrorAction SilentlyContinue)) {
                    $displayName = $key.GetValue('DisplayName')
                    if ($displayName -like '*Microsoft 365*') {
                        $displayName
                    }
                }
            )

            if ($installedProducts.count -gt 0) {
                Write-LogInfo -Message "M365 apps found on attempt $attempt"
                break
            }

            if ($attempt -lt $maxAttempts) {
                Write-LogInfo -Message "M365 apps not found, yet. We'll try again in $secondsToSleep seconds."
                Start-Sleep -Seconds $secondsToSleep
            }

        } until ($attempt -ge $maxAttempts)

        if ($installedProducts.count -gt 0) {
            foreach ($product in $installedProducts) {
                Write-LogInfo -Message "M365 package found: $product"
            }
            $true
        } else {
            Write-LogWarning -Message "Failed to find any M365 Apps after $maxAttempts attempts."
            $false
        }
    }

    # Function: Remove all M365 apps
    function Remove-M365Apps {

        Write-LogInfo -Message "M365 App removal has been requested"

        if (-not(Test-M365Installed)) {
            Write-LogWarning -Message "It doesn't look like there are any M365 Apps installed, so there's no need to attempt removal"
            return
        }

        $UninstallXML = New-UninstallXMLFile

        try {
            Write-LogInfo -Message "Attempting to remove M365 Apps, please be patient"
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
            Write-LogSuccess "M365 App removal was successful"
        }
        else {
            Write-Warning "Couldn't remove M365 Apps; we'll continuing anyway, but don't get your hopes up"
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
    if (-not $RemoveOnly) {
        if (-not $ConfigurationXMLFile) {
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
        Write-LogInfo '...success!'
    }
    catch {
        Write-LogError -Message "An error occured whilst trying to download the Office Deployment Tool.`nPlease make sure the following link is valid: $ODTInstallerURL"
        Exit 1
    }

    # Extract the Office Deployment Tool
    try {
        Write-LogInfo -Message "Attemping to extract the M365 download tool..."

        $ODTSetup = Join-Path $OfficeInstallDownloadPath "ODTSetup.exe"
        $Arguments = "/quiet /extract:$OfficeInstallDownloadPath"
        Start-Process $ODTSetup -ArgumentList $Arguments -Wait

        Write-LogInfo -Message "...success!"
    }
    catch {
        Write-LogError -Message "...uh-oh, an error occurred whilst extracting the tool. The error is: $($_.Exception.Message)"
        Exit 1
    }

    # Try to remove M365 apps if the script was called with -RemoveExisting or -RemoveOnly
    if ($RemoveExisting -or $RemoveOnly) {
        Remove-M365Apps
    }

    # Download and install office (as long as the script wasn't called with -RemoveOnly)
    if (-not $RemoveOnly) {
        try {
            Write-LogInfo -Message "Downloading and installing M365 Apps..."

            $setupFilePath = Join-Path $OfficeInstallDownloadPath "setup.exe"
            $arguments = "/configure $ConfigurationXMLFile"
            $process = Start-Process $setupFilePath -ArgumentList $arguments -Wait -PassThru

            Write-LogInfo -Message "...looking good..."
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
    if (-not $RemoveOnly) {
        if ($officeInstalled) {
            Write-LogSuccess -Message "M365 Apps were successfully installed!"
            Exit 0
        }
        else {
            Write-LogWarning -Message "Failed to install M365 Apps. Review the logs and try again"
            Exit 1
        }
    }
}
