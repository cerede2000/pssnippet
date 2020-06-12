<#
.SYNOPSIS
  Simple log function
.DESCRIPTION
  This function is used to write a log file, it also displays console output
.PARAMETER Message
    Log message, array string
.PARAMETER Logfile
    Out file log, if not passed no log file are generated
.PARAMETER ConsoleOutput
    Switch for view log out in console
.PARAMETER Replace
    Switch for replace existing log file
.PARAMETER Level
    Log level, default Info
.INPUTS
  None
.OUTPUTS
  If log file parameter passed, log are output in log file
  If ConsoleOutput switch enable, log view in console
.NOTES
  Version:        1.0
  Author:         cerede2000
  Creation Date:  2020-06-12
  Purpose/Change: Initial script development
  
.EXAMPLE
  Write-Log 'This is sample log message' -ConsoleOutput
.EXAMPLE
  Write-Log 'This is sample log message' -Logfile C:\Windows\Logs\PSOutput\temp.log
.EXAMPLE
  Write-Log 'This is sample log message' -Level ERROR -Logfile C:\Windows\Logs\PSOutput\temp.log
#>

function Write-Log {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Message
    ,[string]$Logfile
    ,[switch]$ConsoleOutput
    ,[switch]$Replace
    ,[ValidateSet('SUCCESS', 'INFO', 'WARN', 'ERROR', 'DEBUG')]
     [string]$Level = 'INFO'
  )
  
  $LogColors = @{ SUCCESS = 'Green'; INFO = 'White'; WARN = 'Yellow'; ERROR = 'Red'; DEBUG = 'Gray' }

  $TimeStamp = [System.DateTime]::Now.ToString('yyyy-MM-dd HH:mm:ss.fff')
  
  if (-Not [System.String]::IsNullOrEmpty($Logfile)) {
    if(-Not (Test-Path "$(Split-Path $Logfile -Parent)")) { New-Item -ItemType Directory -Path "$(Split-Path $Logfile -Parent)" -Force | Out-Null }
   if($Replace) {
    Out-File -FilePath $Logfile -InputObject "[$TimeStamp] [$Level] <!$Message>"
   } else {
    Out-File -Append -FilePath $Logfile -InputObject "[$TimeStamp] [$Level] <!$Message>"
   }
  }

  if ($ConsoleOutput) {
   Write-Host "[$TimeStamp] [$Level] <!$Message>" -ForegroundColor $LogColors[$Level]
  }
}
