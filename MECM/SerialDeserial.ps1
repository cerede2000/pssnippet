[System.Reflection.Assembly]::LoadFrom((Join-Path (Get-Item $env:SMS_ADMIN_UI_PATH).Parent.FullName "Microsoft.ConfigurationManagement.ApplicationManagement.dll"))
$app = gwmi -computer <siteserver> -namespace root\sms\site_<sitecode> -class sms_application -filter "LocalizedDisplayName='<appname>' and isLatest = 'true'"
$app.Get()
$sdmPackageXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString($app.SDMPackageXML, $true)
$sdmPackageXML.AutoInstall = $true
$app.SDMPackageXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::SerializeToString($sdmPackageXML, $true)
$app.Put()