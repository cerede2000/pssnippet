[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
  [Parameter(ParameterSetName = 'Install')]
  [switch]$Install,
  [Parameter(ParameterSetName = 'Uninstall')]
  [switch]$Uninstall
)

if((Get-Command "$PSCommandPath").ParameterSets | ?{ $_.IsDefault -and $_.Name -eq $PSCmdlet.ParameterSetName }) { $Install = $True }

Write-Host $Install
Write-Host $Uninstall

if($Install) {
  'Install'
}

if($Uninstall) {
  'Uninstall'
}
