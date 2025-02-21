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
    $ErrorActionPreference = 'SilentlyContinue'
    $LogFolder = Join-Path $OfficeInstallDownloadPath 'Logs'
    $LogFileName = (Get-Date -UFormat "%d-%m-%Y") + ".log"
    $LogPath = Join-Path $LogFolder $LogFileName

    # Fucntion: Creates a new XML file from scratch to use with M365 setup.exe
    function New-InstallXMLFile {
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
        [String]$MSWebPage = Invoke-RestMethod 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
        $MSWebPage | ForEach-Object {
            if ($_ -match '"url":"(https://download\.microsoft\.com/[^"]*/officedeploymenttool[^"\s]*\.exe)"') {
                $matches[1]
            }
        }
    }

    # Function: Test if this script is running in an elevated context
    function Test-IsElevated {
        $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($id)
        if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
            Write-Output $true
        } else {
            Write-Output $false
        }
    }

    # Function: Test M365 installed or not. Looks in the registry to see if there are
    # keys indicating the presence of one or more M365 products
    function Test-M365Installed {
        Write-Log -LogOutput "Checking for Microsoft 365 Apps..." -Path $LogPath

        $UninstallRegKeys = @(
            'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\',
            'HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\'
        )

        # Create an arraylist, use this to store any relevant reg keys we find
        $InstalledProducts = [System.Collections.ArrayList]::new()
        
        foreach ($Key in (Get-ChildItem $UninstallRegKeys)) {
            If ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
                $Product = $Key.GetValue('DisplayName')
                $InstalledProducts.Add($Product) | Out-Null
            }
        }
       
        If ($InstalledProducts.count -gt 0) {
            Write-Log -LogOutput "Found Microsoft 365 Apps, listing them now..." -Path $LogPath
            foreach ($Product in $InstalledProducts) {
                Write-Log -LogOutput "Found $Product" -Path $LogPath
            }
            return $true
        }
        Else {
            Write-Log -LogOutput "Didn't find any Microsoft 365 Apps" -Path $LogPath
            return  $false
        }
    }

    # Function: Remove all M365 apps
    function Remove-M365Apps {
        If (-not(Test-M365Installed)) {
            Write-Log -LogOutput "M365 App removal was requested, however none are present. Continuing..." -Path $LogPath
            return
        }

        $UninstallXML = New-UninstallXMLFile

        try {
            Write-Log -LogOutput "Attempting to remove Microsoft 365 Apps, please be patient..." -Path $LogPath
            $Setup = Join-Path $OfficeInstallDownloadPath 'setup.exe'
            $Arguments = "/configure $UninstallXML"
            $Silent = Start-Process $Setup -ArgumentList $Arguments -Wait -PassThru
        }
        catch {
            Write-Log -LogOutput "ERROR: Something went wrong during uninstall. The error was: $_" -Path $LogPath
            Write-Log -LogOutput "Exiting..." -Path $LogPath
            exit 1
        }

        If (-not(Test-M365Installed)) {
            Write-Log -LogOutput 'SUCCESS! Microsoft 365 Apps were removed, continuing...' -Path $LogPath
        }
        Else {
            Write-Log -LogOutput'Unable to install M365 Apps. Continuing anyway, but will probably fail...' -Path $LogPath
        }
    }

    # Function: Custom logging
    Function Write-Log {

        param(
            [Parameter(Mandatory=$True)]
            [array]$LogOutput,
            [Parameter(Mandatory=$True)]
            [string]$Path
        )
        
        # Create the log file if it doesn't exist, usually will only happen on first run
        if (-not(Test-Path $Path)) {
            try {
                (New-Item -Path $Path -Force)
                Write-Log -LogOutput "Created log file" -Path $Path
            }
            catch {
                Write-Warning "Couldn't create the log file path, no log will be written :("
                return
            }
        }
    
        # Get date, time, set log output and write the msg
        $currentDate = (Get-Date -UFormat "%d-%m-%Y")
        $currentTime = (Get-Date -UFormat "%T")
        $logOutput = $logOutput -join (" ")
        "[$currentDate $currentTime] $logOutput" | Out-File $Path -Append

    }
    
}


# Process block, main script functionality goes here
process {

    # Test for admin context, exit if false
    If (-not (Test-IsElevated)) {
        # Write-Error -Message 'Non-elevated context detected. Please re-run this script with elevated permissions.'
        Write-Log -LogOutput "Exiting due to lack of elevated permissions, run again as admin" -Path $LogPath
        Exit 1
    }

    # Create a new working directory
    If (-not (Test-Path $OfficeInstallDownloadPath)) {
        Write-Log -LogOutput "Creating working directory $OfficeDownloadPath)..." -Path $LogPath
        New-Item -ItemType Directory -Path $OfficeInstallDownloadPath | Out-Null
        Write-Log -LogOutput "...done" -Path $LogPath
    }

    # Check whether an XML file was specified, if not, create one. If one was specified,
    # check if it exists before proceeding, exit if it's not found
    If (-not($RemoveOnly)) {
        If (-not($ConfigurationXMLFile)) {
            Write-Log -LogOutput "No XML file was specified, creating one..." -Path $LogPath
            New-InstallXMLFile

            $ConfigurationXMLFile = Join-Path $OfficeInstallDownloadPath 'Install.xml'
 
            if (Test-Path $ConfigurationXMLFile) {
                Write-Log -LogOutput "XML file successfully created, proceeding..." -Path $LogPath
            } else {
                Write-Log -LogOutput "XML creation unsuccessful, please try again" -Path $LogPath
            }

        } else {
            If (-not(Test-Path $ConfigurationXMLFile)) {
                Write-Log -LogOutput "An XML file was specified, but it appears to be invalid. Please try again" -Path $LogPath
                Exit 1
            }
            Else {
                Write-Log -LogOutput "The specified XML file appears to be valid, continuing..." -Path $LogPath
            }
        }
    }

    # Get the ODT URL
    $ODTInstallerURL = Get-ODTURL

    # Download the Office Deployment Tool
    Write-Log -LogOutput "Attempting to download the Office Deployment Tool..." -Path $LogPath

    if (-not ([string]::IsNullOrEmpty($ODTInstallerURL))) {
        Write-Log -LogOutput "SUCCESS: Found ODT download url, value is $ODTInstallerURL" -Path $LogPath
        try {
            $params = @{
                Uri = $ODTInstallerURL
                OutFile = Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe'
            }
            Invoke-WebRequest @params

            $ODTPath = Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe'
            if (Test-Path $ODTPath) {
                Write-Log -LogOutput "ODTSetup.exe was successfully downloaded" -Path $LogPath
            } else {
                Write-Log -LogOutput "ERROR: ODTSetup.exe couldn't be found, pease try again" -Path $LogPath
            }
        }
        catch {
            Write-Log -LogOutput "ERROR: Download error. Please verify that the following link is valid: $ODTInstallerURL"
            Write-Log -LogOutput "Exiting..." -Path $LogPath
            exit 1
        }
    } else {
        Write-Log -LogOutput "ERROR: Couldn't retrieve the ODT download URL. Value is $ODTInstallURL" -Path $LogPath
        Write-Log -LogOutput "Exiting..." -Path $LogPath
        exit 1
    }
    

    # Extract the Office Deployment Tool
    try {
        Write-Log -LogOutput 'Attempting to run the Office Deployment Tool...' -Path $LogPath
        $ODTSetup = Join-Path $OfficeInstallDownloadPath 'ODTSetup.exe'
        $Arguments = "/quiet /extract:$OfficeInstallDownloadPath"
        Start-Process $ODTSetup -ArgumentList $Arguments -Wait
    }
    catch {
        Write-Log -LogOutput "ERROR: Something went wrong with the Office Deployment Tool. The error is: $_" -Path $LogPath
        Write-Log -LogOutput "Exiting..." -Path $LogPath
        exit 1
    }

    # Try to remove M365 apps if the script was called with -RemoveExisting or -RemoveOnly
    If ($RemoveExisting -or $RemoveOnly) {
        Write-Log -LogOutput "Remove flag was specified, beginning removal process" -Path $LogPath
        Remove-M365Apps
    }

    # Download and install office (as long as the script wasn't called with -RemoveOnly)
    If (-not($RemoveOnly)) {
        try {
            Write-Log -LogOutput "Downloading and installing Microsoft 365 Apps for business..." -Path $LogPath
            $Setup = Join-Path $OfficeInstallDownloadPath 'setup.exe'
            $Arguments = "/configure $ConfigurationXMLFile"
            $Silent = Start-Process $Setup -ArgumentList $Arguments -Wait -PassThru
        }
        catch {
            Write-Log -LogOutput "ERROR: An error occurred whilst unning the Office installer. The error is $_" -Path $LogPath
            Write-Log -LogOutput "Exiting..." -Path $LogPath
            exit 1
        }

        # Test if M365 Apps were installed successfully, store result in $OfficeWasInstalled ($true or $false)
        $OfficeWasInstalled = Test-M365Installed
    }
    Else {
        Write-Log -LogOutput 'No more work to do because -RemoveOnly was specified' -Path $LogPath
    }
}

end {
    # If cleanup flag was set, remove the installation files
    If ($CleanupInstallFiles) {
        Write-Log -LogOutput "Cleanup flag was specified, cleaning up..." -Path $LogPath
        Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
    }

    # If -RemoveOnly was not set, output the result of the attempted Installation
    If (-not($RemoveOnly)) {
        If ($OfficeWasInstalled) {
            Write-Log -LogOutput "Office install was successful! Yay! Exiting..." -Path $LogPath
            Exit 0
        }
        Else {
            Write-Log -LogOutput "Unsuccessful...Something went wrong and Microsoft 365 Apps were not installed. Please troubleshoot and try again. Exiting..." -Path $LogPath
            Exit 1
        }
    }
    Else {
        Write-Log -LogOutput "Exiting..." -Path $LogPath
        Exit 0
    }
}
