param(
  [string]$TenantID,
  [string]$ClientId,
  [string]$ClientSecret,
  [switch]$HeaderToken
)

$Splatting = @{
  URI = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
  Method = 'Post'
  Body = @{
    client_id     = $ClientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
  }
  ContentType = 'application/x-www-form-urlencoded'
  UseBasicParsing = $True
}

try {
  $Token = Invoke-RestMethod @Splatting -ErrorAction Stop
  if($HeaderToken) { @{ Authorization = "$($Token.token_type) $($Token.access_token)" } }
  else { $Token }
} catch {
  write-host $_.Exception.Message -f red
  $null
}

#$AcessToken = Get-MSToken -TenantID $tenantID -ClientId $AppId -ClientSecret $AppSecret -HeaderToken
