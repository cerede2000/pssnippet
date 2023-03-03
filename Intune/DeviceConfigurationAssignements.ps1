function Get-DeviceConfigurationPolicyAssignment {
  param (
    $ConfigurationPolicyId
    $AcessToken
  )
  
  $restParam = @{
    Method = 'Get'
    Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$ConfigurationPolicyId/assignments"
    Headers = $AcessToken
    ContentType = 'Application/Json'
  }
  Invoke-RestMethod @restParam
  
  #.value.target | Select '@odata.type', groupId
}

Function Add-DeviceConfigurationPolicyAssignment(){
  [cmdletbinding()]
  param (
    $ConfigurationPolicyId,
    [string[]]$TargetGroupId,
    $AcessToken
  )
  
  $TargetGroups = [object[]]$TargetGroupId | %{
    @{
      '@odata.type' = '#microsoft.graph.deviceConfigurationGroupAssignment' 
      targetGroupId = $_
    }
  }

  $Data = @{
    deviceConfigurationGroupAssignments = [object[]]$TargetGroups
  }
  
  $JSON = $Data | ConvertTo-Json
  
  $restParam = @{
      Method = 'Post'
      Uri = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$ConfigurationPolicyId/assign"
      Headers = $AcessToken
      ContentType = 'Application/Json'
      Body = $JSON
  }
  Invoke-RestMethod @restParam
}
