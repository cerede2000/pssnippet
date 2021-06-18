$sdkserver=""
$siteCode=""
$property=""
$targetDp=""


$DP = Get-WmiObject -Namespace root\sms\site_i01 -Class SMS_SCI_SysResUse -Filter 'RoleName = "SMS Distribution Point" and NetworkOSPath like "%SPX50020%"'
$DP.Get()
$Property = $DP.Props | ?{ $_.PropertyName -eq 'Priority' }
$Property.PropertyName
$Property.Value

$Property.Value = 210
$DP.Props = $Property
$DP.Put()

$priorityValue = 20
$dp = gwmi -computer $sdkserver -namespace "rootsmssite_$sitecode" -query "select * from SMS_SCI_SysResUse where RoleName = 'SMS Distribution Point' and NetworkOSPath = '$targetDp'"
$props = $dp.Props
$prop = $props | where {$_.PropertyName -eq $property}

Write-Output "Current DistributionPoint Priority = " $prop.Value


$prop.Value = $priorityValue


Write-Output "Updating the DistributionPoint Priority to = " $priorityValue


$dp.Props = $props
$dp.Put()

$AllDp = Get-WmiObject -Namespace root\sms\site_i01 -Class SMS_SCI_SysResUse -Filter 'RoleName = "SMS Distribution Point"'
$AllDp | %{
	$_.Get()
	$_.NetworkOSPath
	$Property = $_.Props | ?{ $_.PropertyName -eq 'IsPxe' }
	$Property.PropertyName
	$Property.Value
	
	$Property = $_.Props | ?{ $_.PropertyName -eq 'Priority' }
	$Property.PropertyName
	$Property.Value
}


$AllDp | %{
	$_.Get()
	if(($_.Props | ?{ $_.PropertyName -eq 'IsPxe' -and $_.Value -eq 0 }) -ne $null) {
		$Property = $_.Props | ?{ $_.PropertyName -eq 'Priority' }
		$Property.PropertyName
		$Property.Value
		
		$Property.Value = 210
		$_.Props = $Property
		$_.Put()
	}
}
