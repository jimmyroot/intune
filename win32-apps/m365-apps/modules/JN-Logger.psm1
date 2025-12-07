<#
.SYNOPSIS
    A modular, re-usable PS logging framework to use in your Powershell scripts

.DESCRIPTION
    Provides structured logging with console output, file logging, and log rotation
    functionality (to control log sizes)
#>

#region Configuration

$script:LogConfig = @{
    LogLevel = 'INFO'
    LogToFile = $true
    LogToConsole = $true
    LogPath = $null
    MaxLogSizeMB = 2
    MaxLogFiles = 5
    IncludeTimestamp = $true
    IncludeCallerInfo = $false
    DateFormat = 'dd-MM-YYYY HH:mm:ss'
}

$script:LogLevels = @{
    DEBUG = 0
    INFO = 1
    WARNING = 2
    ERROR = 3
}

$script:LogColors = @{
    DEBUG = 'Gray'
    INFO = 'Cyan'
    WARNING = 'Yellow'
    ERROR = 'Red'
    SUCCESS = 'Green'
}

#endregion

#region Public functions


function Initialize-Logging {
    <#
    .SYNOPSIS
        Initializes the logging function with custom configuration
    .EXAMPLE
        Initialize-Logging -LogPath C:\Logs\MyScript.log -LogLevel DEBUG
    #>
    [CmdletBinding()]
    param (
        [string]$LogPath,

        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR')]
        [string]$LogLevel = 'INFO',

        [switch]$LogToFile = $false,
        [switch]$LogToConsole = $false,
        [switch]$IncludeCallerInfo,

        [int]$MaxLogSizeMB = 1,
        [int]$MaxLogFiles = 5
    )

    # If a log path isn't supplied, default it to the scripts working dir
    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $callerScript = (Get-PSCallStack)[-1].ScriptName
        if ($callerScript) {
            $scriptName = [System.IO.Path]::GetFileNameWithoutExtension($callerScript)
            $scriptDir = Split-Path $callerScript -Parent
        }
        else {
            $scriptName = "Unknown"
            $scriptDir = $env:TEMP
        }

    $now = (Get-Date -Format 'ddMMyyyy')
        $LogPath = Join-Path $scriptDir "$scriptName`_$now.log"
    Write-Host $LogPath
    }

    $script:LogConfig.LogPath = $LogPath
    $script:LogConfig.LogLevel = $LogLevel
    $script:LogConfig.LogToFile = $LogToFile
    $script:LogConfig.LogToConsole = $LogToConsole
    $script:LogConfig.MaxLogSizeMB = $MaxLogSizeMB
    $script:LogConfig.MaxLogFiles = $MaxLogFiles
    $script:LogConfig.IncludeCallerInfo = $IncludeCallerInfo

    # In case our log path doesn't exist...
    if ($LogToFile) {
        $logDir = Split-Path $LogPath -Parent
        if (-not (Test-Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }

        Invoke-LogRotation
    }

    Write-LogMessage "Logging initialized: $LogPath" -Level INFO -Force
}


function Write-LogMessage {
    <#
    .SYNOPSIS
        Writes a log message using the specified logging level
    .EXAMPLE
        Write-LogMessage "Processing started..." -Level INFO
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [Parameter(Position = 1)]
        [ValidateSet('DEBUG', 'INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO',

        # If we want to bypass log level filtering...
        [switch]$Force
    )

    # Are we logging based on log level?
    if (-not $Force) {
        $currentLevel = $script:LogLevels[$script:LogConfig.LogLevel]
        $messageLevel = $script:LogLevels[$Level]

        if ($Level -eq 'SUCCESS') {
            $messageLevel = $script:LogLevels['INFO']
        }

        if ($messageLevel -lt $currentLevel) {
			Write-Host "Message level is lower than current level"
            return
        }
    }

    $timestamp = if ($script:LogConfig.IncludeTimestamp) { 
        Get-Date -Format $script:LogConfig.DateFormat
    }
    else {
        '$null'
    }

    $caller = if ($script:LogConfig.IncludeCallerInfo) {
        $callStack = Get-PSCallStack
        if ($callStack.Count -gt 1) {
            $callerFrame = $callStack[1]
            $functionName = if ($callerFrame.FunctionName -ne '<ScriptBlock>') {
                $callerFrame.FunctionName
            }
            else {
                'Main'
            }
            " [$functionName]"
        }
        else {
            ''
        }
    }
    else {
        ''
    }

    $logEntry = if ($timestamp) {
        "[$timestamp] [$Level]$caller $Message"
    }
    else {
        "[$Level]$caller $Message"
    }

    if ($script:LogConfig.LogToConsole) {
        $color = $script:LogColors[$Level]
        Write-Host $logEntry -ForegroundColor $color
    }

    if ($script:LogConfig.LogToFile -and $script:LogConfig.LogPath) {
        try {
            $mutex = New-Object System.Threading.Mutex($false, "Global\PSLoggingMutex")
            $mutex.WaitOne(0, $false) | Out-Null
            Add-Content -Path $script:LogConfig.LogPath -Value $logEntry -ErrorAction Stop
            $mutex.ReleaseMutex()
        }
        catch {
            Write-Warning "Failed to write to log file: $_"
        }
        finally {
            if ($mutex) {
                $mutex.Dispose()
            }
        }

        Invoke-LogRotation
    }
}


function Write-LogDebug {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)
    Write-LogMessage -Message $Message -Level DEBUG
}


function Write-LogInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)
    Write-LogMessage -Message $Message -Level INFO
}


function Write-LogWarning {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)
    Write-LogMessage -Message $Message -Level WARNING
}


function Write-LogError {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)
    Write-LogMessage -Message $Message -Level ERROR
}


function Write-LogSuccess {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)
    Write-LogMessage -Message $Message -Level SUCCESS
}


function Write-LogException {
    <#
    .SYNOPSIS
        Logs an exception, this includes full error details
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$Exception,

        [string]$AdditionalMessage
    )
    $errorDetails = @"
        $AdditionalMessage
        Exception: $($Exception.Exception.Message)
        Type: $($Exception.Exception.GetType().FullName)
        Target: $($Exception.TargetObject)
        Script: $($Exception.InvocationInfo.ScriptName):$($Exception.InvocationInfo.ScriptLineNumber)
        Command: $($Exception.InvocationInfo.Line.Trim())
"@

    Write-LogMessage -Message $errorDetails -Level ERROR
}


function Start-LogSection {
    <#
    .SYNOPSIS
        Create a new section in the log file, acts as a 
        visual separator
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        [char]$SeparatorCharacter = '='
    )

    $separator = [string]$SeparatorCharacter * 60
    Write-LogInfo $separator
    Write-LogInfo " $Title"
    Write-LogInfo $separator
}


function Stop-LogSection {
    <#
    .SYNOPSIS
        Ends a log section in the log file, acts as a 
        visual separator
    #>
    [CmdletBinding()]
    param(
        [char]$SeparatorCharacter = '-'
    )

    $separator = [string]$SeparatorCharacter * 60
    Write-LogInfo -Message "$separator"
}


function Get-LogConfiguration {
    <#
    .SYNOPSIS
        Returns the current logging configuration
    #>
    [CmdletBinding()]
    param()

    [PSCustomObject]$script:LogConfig
}

function Get-UserPrincipalName {
    [CmdletBinding()]
    [OutputType([string])]
    param()

    try {
        $upn = whoami /upn 2>$null
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw "Failed to retrieve UPN"
        }

        $atIndex = $upn.IndexOf('@')
        if ($atIndex -le 0) {
            throw "Invalid UPN format: $upn"
        }

        return $upn.Substring(0, $atIndex)
    }
    catch {
        Write-Error "Failed to get user principal name: $_"
    }
}


#end region

#region Private functions


function Invoke-LogRotation {
    <#
    .SYNOPSIS
        Checks if log size limit is exceeded and rotates log files if 
        necessary
    #>
    [CmdletBinding()]
    param()

    # We don't want to run this if we haven't written to a file
    if (-not $script:LogConfig.LogToFile -or -not $script:LogConfig.LogPath) {
        return
    }

    $logFile = $script:LogConfig.LogPath

    if (-not (Test-Path $logFile)) {
        return
    }

    $fileInfo = Get-Item $logFile
    $fileSizeMB = $fileInfo.Length / 1MB

    if ($fileSizeMB -ge $script:LogConfig.MaxLogSizeMB) {
        $timestamp = Get-Date -Format 'ddMMyyyy_HHmmss'
        $directory = Split-Path $logFile -Parent
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($logFile)
        $extension = [System.IO.Path]::GetExtension($logFile)

        $archiveName = "$baseName`_$timestamp$extension"
        $archivePath = Join-Path $directory $archiveName

        try {
            Move-Item -Path $logFile -Destination $archivePath -Force

            $pattern = "$baseName`_*$extension"
            $oldLogs = Get-ChildItem -Path $directory -Filter $pattern |
                       Sort-Object LastWriteTime -Descending |
                       Select-Object -Skip $script:LogConfig.MaxLogFiles

            $oldLogs | Remove-Item -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Warning "Failed to rotate log file: $_"
        }
    }
}


#endregion

#region Module init
# Only uncomment if there's a chance the module could be 
# imported / initialized without a log path
#if (-not $script:LogConfig.LogPath) {
#    Initialize-Logging
#}
#endregion

# Export public functions
Export-ModuleMember -Function @(
    'Initialize-Logging',
    'Write-LogMessage',
    'Write-LogDebug',
    'Write-LogInfo',
    'Write-LogWarning',
    'Write-LogError',
    'Write-LogSuccess',
    'Write-LogException',
    'Start-LogSection',
    'Stop-LogSection',
    'Get-LogConfiguration',
	'Get-UserPrincipalName'
)
