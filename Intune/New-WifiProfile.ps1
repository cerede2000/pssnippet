$jsonWifi = @'
{
  "@odata.type": "#microsoft.graph.windowsWifiConfiguration",
  "displayName": "{{CFGNAME}}",
  "description": null,
  "roleScopeTagIds": [
    "0"
  ],
  "wifiSecurityType": "open",
  "meteredConnectionLimit": "unrestricted",
  "ssid": "{{SSID}}",
  "networkName": "{{SSID}}",
  "connectAutomatically": true,
  "connectWhenNetworkNameIsHidden": true,
  "proxySetting": "none",
  "forceFIPSCompliance": false
}
'@

$jsonWifiProfile = $jsonWifi.Replace('{{CFGNAME}}', 'My Wifi Profile').Replace('{{SSID}}', 'MyWifiSSID')

$restParam = @{
  Method = 'POST'
  Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
  Headers = $AcessToken
  ContentType = 'Application/Json'
  Body = $jsonWifiProfile
}
Invoke-RestMethod @restParam
