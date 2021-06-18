
Get-ADComputer -Filter * -SearchBase 'OU=,DC=,DC=' -Properties Name, operatingSystem, memberOf | %{$i=1}{
  Write-Host $i
  
  $Name = $_.Name
  $OS = $_.operatingSystem
  $Groups = ($_.MemberOf | ?{ $_ -match 'SCCM_Device' } | %{ (Get-AdObject $_ -Properties Name).Name }) -Join ';'
  
  if($_.DistinguishedName -match 'OU=Seven') {
    if(![String]::IsNullOrEmpty($Groups)) {
      "$Name;$OS;$Groups"
    }
  }
  
  $i++
} | Out-File -FilePath C:\temp\ExportADAll_2.csv -Append

############################################################################################################################################################

$Searcher = [ADSISearcher]@{ searchRoot = [ADSI]'LDAP://OU=,DC=,DC='; filter = '(objectCategory=computer)'; PageSize = 1000; CacheResults = $False }
$Searcher.PropertiesToLoad.AddRange(@('distinguishedname', 'name', 'operatingSystem', 'memberOf'))
$GroupSearch = [ADSISearcher]@{ searchRoot = [ADSI]'LDAP://DC=,DC='; PageSize = 100; CacheResults = $False }
$GroupSearch.PropertiesToLoad.AddRange(@('name'))
$Searcher.FindAll() | %{$i=1}{
  Write-Host $i

  if($_.Properties['distinguishedname'] -match 'OU=Seven') {
    $Name = $_.Properties['name']
    $OS = $_.Properties['operatingsystem']
    $Groups = ($_.Properties['memberof'] | ?{ $_ -match 'SCCM_Device' } | %{ $GroupSearch.filter="(distinguishedname=$_)"; $GroupSearch.FindOne().Properties['name'] }) -join ';'
    
    if(![String]::IsNullOrEmpty($Groups)) {
      "$Name;$OS;$Groups"
    }
  }
 
  $i++
} | Out-File -FilePath C:\temp\ExportADAll.csv -Append

############################################################################################################################################################

'GroupName;Computer Name;Distinguished Name' | Out-File -FilePath C:\temp\ExportADAll.csv
$Searcher = [ADSISearcher]@{ searchRoot = [ADSI]'LDAP://OU=,DC=,DC='; filter = '(objectCategory=group)'; SearchScope = 'OneLevel'; PageSize = 1000; CacheResults = $False }
$Searcher.PropertiesToLoad.AddRange(@('name', 'member'))
$Searcher.FindAll() | %{$i=1}{
  Write-Host $i

  $Name = $_.Properties['name']
  $_.Properties['member'] | %{
    if($_ -match 'OU=Seven') {
      $ComputerName = ([regex]::Match($_, '^CN=([^,]*),')).Groups[1].Value
      $DN = $_
      "$Name;$ComputerName;$DN"
    }
  }
 
  $i++
} | Out-File -FilePath C:\temp\ExportADAll.csv -Append

############################################################################################################################################################


'GroupName;Computer Name;Distinguished Name' | Out-File -FilePath C:\temp\ExportADAll.csv
$Searcher = [ADSISearcher]@{ searchRoot = [ADSI]'LDAP://OU=,DC=,DC='; filter = '(&(objectClass=group))'; SearchScope = 'OneLevel'; PageSize = 1000; CacheResults = $False }
$Searcher.PropertiesToLoad.AddRange(@('name', 'parent'))
$Searcher.FindAll() | %{$i=1}{
  Write-Host $i
  
  $Name = $_.Properties['name']
  
  if($_.properties.member.count -eq 0) {
    $GroupSearcher = [ADSISearcher]@{ searchRoot = [ADSI]'LDAP://OU=,DC=,DC='; filter = "(&(objectClass=group)(name=$Name))"; SearchScope = 'OneLevel'; PageSize = 1000; CacheResults = $False }
    $retrievedAllMembers = $false
    $rangeBottom = 0
    $rangeTop = 0
    
    while(-Not $retrievedAllMembers) {
      $rangeTop = $rangeBottom + 1499
      $memberRange = "member;range=$rangeBottom-$rangeTop"
      $rangeBottom += 1500
      
      $GroupSearcher.PropertiesToLoad.Clear() | Out-Null
      $GroupSearcher.PropertiesToLoad.Add("$memberRange") | Out-Null
      
      try {
        $result = $GroupSearcher.FindOne()
        $rangedProperty = $result.Properties.PropertyNames.Where({$_ -match '^member;range'})
        $results = $result.Properties.item($rangedProperty)
        
        if ($results.count -eq 0) {
          $retrievedAllMembers = $true
        } else {
          $results | %{
            if($_ -match 'OU=Seven') {
              $ComputerName = ([regex]::Match($_, '^CN=([^,]*),')).Groups[1].Value
              $DN = $_
              "$Name;$ComputerName;$DN"
            }
          }
        }
      } catch {
        $retrievedAllMembers=$true
      }
      
      $results = $null
      $result = $null
    }
    
    $GroupSearcher.Dispose()
  } else {
    $_.properties.member | %{
      if($_ -match 'OU=Seven') {
        $ComputerName = ([regex]::Match($_, '^CN=([^,]*),')).Groups[1].Value
        $DN = $_
        "$Name;$ComputerName;$DN"
      }
    }
  }
  
  $i++
} | Out-File -FilePath C:\temp\ExportADAll.csv -Append
