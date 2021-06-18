
# Functions
function Write-CMLog {
  [CmdletBinding()]
  param(
   [Parameter(Mandatory=$True, ValueFromPipeline=$True, ValueFromPipelinebyPropertyName=$True)]
    $Message,
    [Parameter(Mandatory=$True, ValueFromPipeline=$False, ValueFromPipelinebyPropertyName=$True)]
	[string]$File,
    [Parameter(Mandatory=$False, ValueFromPipeline=$False, ValueFromPipelinebyPropertyName=$True)]
    [ValidateSet('Warning','Error','Verbose','Debug', 'Information')] 
    [string]$Type = 'Information',
    [Parameter(Mandatory=$False)]
    [switch]$Clobber
  )

  begin {
    $Now = Get-Date
    if($Clobber -And [System.IO.File]::Exists("$File")) { [System.IO.File]::Delete("$File") }

    switch($Type){
      'Warning' { [int]$Severity = 2}
      'Error' { [int]$Severity = 3}
      'Verbose' { [int]$Severity = 4}
      'Debug' { [int]$Severity = 5}
      'Information' { [int]$Severity = 6}
    }

    $CallingInfo = (Get-PSCallStack)[1]
    $LogDate = $Now.ToString('MM-dd-yyyy')
    $LogTime = $Now.ToString('HH:mm:ss.fffzzz')
    $Context = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Thread = [Threading.Thread]::CurrentThread.ManagedThreadId
    [string]$MessageFormat = ('<![LOG[$Message]LOG]!><time="{0}" date="{1}" component="{2}" context="{3}" type="{4}" thread="{5} : {6}" file="{7} : {8}">' -f
		[string]$LogTime,
		[string]$LogDate,
		[string]$Context,
		[string](Get-Process -Id $PID).ProcessName,
		[int]$Severity,
		[int]$PID,
		[int]$Thread,
		[string]$CallingInfo.Location,
		[string]$CallingInfo.FunctionName
	)

	$FileStreamWriter = [System.IO.File]::AppendText("$File")
  }

  process {
	if($Message -is [array]) {
	  $Messages = $Message
	  foreach($Message in $Messages) {
		$FileStreamWriter.WriteLine($ExecutionContext.InvokeCommand.ExpandString($MessageFormat))
	  }
	} else {
		$FileStreamWriter.WriteLine($ExecutionContext.InvokeCommand.ExpandString($MessageFormat))
	}
  }

  end {
	$FileStreamWriter.Close()
	$FileStreamWriter.Dispose()
  }
}


# Parameters
$ErrorActionPreference = 'Stop'
$Script:LogFile = Join-Path (split-path -parent $MyInvocation.MyCommand.Definition) 'Office2019PatchDownloader.log'
$BaseDir = 'H:\data\Office365'
$BaseURL = 'https://config.office.com/api/filelist'
$UriParameters = @{
    Channel = 'Perpetual2019'
    lid = 'fr-fr', 'en-us', 'pt-br'
}

Add-Type -AssemblyName System.Web

function Step-Main {
    $RunspacePool = [runspacefactory]::CreateRunspacePool(1, $env:NUMBER_OF_PROCESSORS)
    $RunspacePool.Open()
    $Jobs = @()

    $DownloadScriptblock = {
        param($Uri, $File)
    
        $MaxRetries = 3
        $ErrorActionPreference = 'Stop'
        $ResultVar = [PSCustomObject]@{
            File = $File
            DownloadTime = $null
            IsRetried = $false
            Success = $Null
            ErrorMessage = $Null
        }

        $RetriesLeft = $MaxRetries
        $Success = $false
        
        do {
            $ResultVar.Success = $null
            $ResultVar.ErrorMessage = $null

            try {
                #Invoke-WebRequest $Uri -Method Get -OutFile $File -TimeoutSec 10
                
                $DownloadTime = [system.diagnostics.stopwatch]::StartNew()

                $HWebRequest = [System.Net.HttpWebRequest]([System.Net.WebRequest]::Create($Uri))
                $HttpResponse = $HWebRequest.GetResponse()
                $StreamFile = [System.IO.File]::OpenWrite($File)
                $HttpResponse.GetResponseStream().CopyTo($StreamFile)
                $StreamFile.Close()

                $ResultVar.DownloadTime = $DownloadTime.Elapsed.TotalSeconds

                $Success = $HttpResponse.StatusCode.value__ -eq 200
            } catch {
                $ResultVar.IsRetried = $true
                $Success = $false

                $ErrorMessage = $_.Exception.Message
                $LineNumber = $_.InvocationInfo.ScriptLineNumber

                $ResultVar.Success = $False
                $ResultVar.ErrorMessage = "Error : Line [$LineNumber] :: [$ErrorMessage]"

                <#if($_ -is [System.Net.WebException] -and $_.Exception.Status -eq 'Timeout') {
                    
                } else {
                    $ErrorMessage = $_.Exception.Message
                    $LineNumber = $_.InvocationInfo.ScriptLineNumber

                    $ResultVar.Success = $False
                    $ResultVar.ErrorMessage = "Error : Line [$LineNumber] :: [$ErrorMessage]"

                    return $ResultVar
                }#>

                if($RetriesLeft -ge 0) { Start-Sleep -Seconds 10 }
            }

            $RetriesLeft--

        } while(($Success -eq $false) -and ($RetriesLeft -ge 0))

        if(-Not $Success) { return $ResultVar }

        if((Test-Path "$File")) {
            $ResultVar.Success = $True
        } else {
            $ResultVar.Success = $False
            $ResultVar.ErrorMessage = "Error : File [$File] not found !"
        }

        return $ResultVar
    }

    # WSUS Decline superseded updates and get last approved updates
    
    Write-CMLog -Message 'Start clean WSUS Office 2019 updates' -File $Script:LogFile

    [void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")
    $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer()
    $AllUpdates = $WSUS.GetUpdates() | ?{ $_.IsDeclined -eq $False }
    $AllUpdates | ?{ $_.ProductTitles -contains 'Office 365 Client' } | ?{ $_.Title -notmatch '^Office 2019' } | %{
        Write-CMLog -Message "Decline [$($_.Title)] update" -File $Script:LogFile
        $_.Decline()
    }
    $AllUpdates | ?{ $_.ProductTitles -contains 'Office 365 Client' } | ?{ $_.IsSuperseded -eq $True} | %{
        Write-CMLog -Message "Decline [$($_.Title)] update" -File $Script:LogFile
        $_.Decline()
    }

    $AllUpdates = $WSUS.GetUpdates() | ?{ $_.IsDeclined -eq $False }
    $O2019ApprovedUpdates = $AllUpdates | ?{ $_.ProductTitles -contains 'Office 365 Client' } | ?{ $_.IsApproved -eq $True}

    $UpdatesID = @()
    $O2019ApprovedUpdates | %{
        Write-CMLog -Message "Found [$($_.Title)] update" -File $Script:LogFile
        $UpdatesID += $_.Id.UpdateId.Guid
    }

    Write-CMLog -Message 'Define proxy' -File $Script:LogFile
    $Proc = Start-Process -FilePath 'cmd.exe' -ArgumentList '/c netsh winhttp set proxy 172.23.128.21:9090 2>&1>nul' -NoNewWindow -Wait -PassThru

    if(-Not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
}
    $AllProtocols = [System.Net.SecurityProtocolType]'Ssl3,Tls,Tls11,Tls12'
    [System.Net.ServicePointManager]::SecurityProtocol = $AllProtocols
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy

    # Download files

    $DownloadTime = [system.diagnostics.stopwatch]::StartNew()

    $OfficeFiles = @()
    $O2019ApprovedUpdates | %{
        $Guid = $_.Id.UpdateId.Guid

        $url = [Uri]((($_.AdditionalInformationUrls | Select -First 1) -Split 'ServicePath=')[1])
        $HttpQuery = [System.Web.HttpUtility]::ParseQueryString($url)
        $O2019Version = $HttpQuery['Version']
        $Architecture = $HttpQuery['Arch']

        $UriParameters['arch'] = $Architecture
        $UriParameters['Version'] = $O2019Version

        Write-CMLog -Message "Prepare URI [$Guid] - [$O2019Version] - [$Architecture]" -File $Script:LogFile

        $UrlParameterString = '?'
        $UriParameters.GetEnumerator() | %{
            $Name = $_.Key
            $Value = $_.Value
            if(($Value -is [array])) {
                $UrlParameterString += '{0}={1}&' -f $Name, ($Value -join "&$Name=")
            } else {
                $UrlParameterString += '{0}={1}&' -f $Name, $Value
            }
        }

        $RequestUri = $BaseURL + $UrlParameterString

        Write-CMLog -Message "Get URI result [$RequestUri]" -File $Script:LogFile

        $JSon = Invoke-RestMethod -Uri $RequestUri

        Write-CMLog -Message "Create folder strucutre for update id [$Guid]" -File $Script:LogFile

        $basePath = Join-Path $BaseDir $Guid
        if(!(Test-Path "$basePath")) { New-Item -ItemType Directory -Path "$basePath" -Force | Out-Null }

        Write-CMLog -Message "Save ofl in [$basePath]" -File $Script:LogFile
        $JSon | ConvertTo-Json -Compress | Out-File -FilePath "$basePath\ofl.json"

        $JSon.files | %{
	        $filePath = Join-Path $basePath $_.relativePath
	        if(!(Test-Path "$filePath")) { New-Item -ItemType Directory -Path "$filePath" -Force | Out-Null }
	        $filePath = Join-Path $filePath $_.name

            $OfficeFiles += $filePath
        
            if(-Not (Test-Path -Path $filePath)) {
                Write-CMLog -Message "Create download thread for URI [$($_.url)] to file [$filePath]" -File $Script:LogFile

                $PowerShell = [powershell]::Create()
	            $PowerShell.RunspacePool = $RunspacePool
	            $PowerShell.AddScript($DownloadScriptblock).AddParameter('Uri', $_.url).AddParameter('File', $filePath) | Out-Null
	            $Jobs += [PSCustomObject]@{ Powershell = $PowerShell; Handle = $PowerShell.BeginInvoke() }
            }
        }
    }

    $FinishedJobs = @()
    while ($Jobs.Handle.IsCompleted -contains $false) {
	    Start-Sleep 1
        $Jobs | ?{ $_.Handle.IsCompleted } | %{
            $Res = $_.Powershell.EndInvoke($_.Handle)
            
            Write-CMLog -Message "Download file [$($Res.File)] is finish with status success [$($Res.Success)] in [$($Res.DownloadTime)] seconds :: Error Message [$($Res.ErrorMessage)] :: Is retried [$($Res.IsRetried)] :: [$($Jobs.Count)] download jobs left" -File $Script:LogFile
        }

        $FinishedJobs += $Jobs | ?{ $_.Handle.IsCompleted }
        $Jobs = $Jobs | ?{ -Not $_.Handle.IsCompleted }
    }

    Write-CMLog -Message "All files are downloaded in [$($DownloadTime.Elapsed.TotalSeconds)] total seconds" -File $Script:LogFile

    # Check all files is downloaded
    $OfficeFiles | %{
        if(-Not (Test-Path $_)) {
            Write-CMLog -Message "Missing file [$_]" -File $Script:LogFile -Type Error
        }
    }

    # Check files hashes

    Write-CMLog -Message 'Start file hash check' -File $Script:LogFile

    $O2019ApprovedUpdates | %{
        $Guid = $_.Id.UpdateId.Guid
    
        $basePath = Join-Path $BaseDir $Guid
        $JsonData = Get-Content -Path (Join-Path $basePath 'ofl.json') -Raw | ConvertFrom-Json

        $JsonData.files | ?{ -Not [String]::IsNullOrEmpty($_.hashLocation) } | %{
	        $File = Get-ChildItem -Path $basePath -Recurse -Filter $_.Name
	        $HashFile = Get-ChildItem -Path $basePath -Recurse -Filter ($_.hashLocation -Split '/')[0]
	        
	        &expand $HashFile.FullName -F:(($_.hashLocation -Split '/')[1]) C:\Temp\ | Out-Null
	        $c =Get-Content -Path "C:\Temp\$(($_.hashLocation -Split '/')[1])" -Raw
	        $Hash = $c -replace [Convert]::ToChar(0x0).ToString(), ''
	        $FileHash = (Get-FileHash -Path $File.FullName).Hash

            if($Hash -ne $FileHash) {
                Write-CMLog -Message "Hash for file [$($File.FullName)] not match to hash [$Hash] ! Delete it." -File $Script:LogFile -Type Error
                Remove-Item -Path $File.FullName -Force
            } else {
                Write-CMLog -Message "Hash for file [$($File.FullName)] match to hash [$Hash]." -File $Script:LogFile
            }

            Remove-Item -Path "C:\Temp\$(($_.hashLocation -Split '/')[1])" -Force
        }
    }

    Write-CMLog -Message 'Clean proxy' -File $Script:LogFile
    $Proc = Start-Process -FilePath 'cmd.exe' -ArgumentList '/c netsh winhttp reset proxy 2>&1>nul' -NoNewWindow -Wait -PassThru
}

# run main
try {
  Write-CMLog -Message "********** Start $($Script:ScriptName) **********" -File $Script:LogFile
  Step-Main
} catch {
  $ErrorMessage = $_.Exception.Message
  $LineNumber = $_.InvocationInfo.ScriptLineNumber
  
  Write-CMLog -Message "Exception are generated ! Message [$ErrorMessage] at line [$LineNumber] Script exit !" -Type 'Error' -File $Script:LogFile
  $Script:ExitCode = 10
} finally {
  Write-CMLog -Message "********** End script RC [$($Script:ExitCode)] **********" -File $Script:LogFile
  exit $Script:ExitCode
}
