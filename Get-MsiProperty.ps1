function Get-MsiProperty {
  param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [Parameter(Mandatory=$true, ParameterSetName='Property')]
    [string]$Property,
    [Parameter(Mandatory=$true, ParameterSetName='ListProperties')]
    [switch]$ListProperties
  )

  function Get-AbsolutePath ($Path) {
    $Path = [System.IO.Path]::Combine( ((pwd).Path), ($Path) )
    return [System.IO.Path]::GetFullPath($Path)
  }

  function Get-Property($Object, $PropertyName, [object[]]$ArgumentList) {
    return $Object.GetType().InvokeMember($PropertyName, 'Public, Instance, GetProperty', $null, $Object, $ArgumentList)
  }

  function Invoke-Method($Object, $MethodName, $ArgumentList) {
    return $Object.GetType().InvokeMember($MethodName, 'Public, Instance, InvokeMethod', $null, $Object, $ArgumentList)
  }

  $ErrorActionPreference = 'Stop'
  Set-StrictMode -Version Latest

  try {
    $msiOpenDatabaseModeReadOnly = 0
    $Installer = New-Object -ComObject WindowsInstaller.Installer
    $Database = Invoke-Method $Installer OpenDatabase @((Get-AbsolutePath "$Path"), $msiOpenDatabaseModeReadOnly)
    if($ListProperties) {
      $View = Invoke-Method $Database OpenView  @("SELECT Property, Value FROM Property")
    } else {
      $View = Invoke-Method $Database OpenView  @("SELECT Value FROM Property WHERE Property='$Property'")
    }

    Invoke-Method $View Execute

    $Record = Invoke-Method $View Fetch
    $Return = $null
    if ($Record) {
      if($ListProperties) {
        $Properties = @()
        do {
          $Properties += '' | Select @{N='Name'; E={ Get-Property $Record StringData 1 }}, @{N='Value'; E={ Get-Property $Record StringData 2 }}
          $Record = Invoke-Method $View Fetch
        } until($Record -eq $null)
        $Return = $Properties
      } else {
        $Return = [string](Get-Property $Record StringData 1)
      }
    }
  } catch {
    $_
  } finally {
    Invoke-Method $View Close
    Remove-Variable -Name Record, View, Database, Installer
  }

  return $Return
}
