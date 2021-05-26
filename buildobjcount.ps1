[datetime]$starttime = Get-Date

$global:totalCountOS = @{}
$global:totalCountOS.Add("NULL",0)
$global:countTotal = 0

$ErrorActionPreference = "Stop"

$arrayDomain = @()
$arrayDomain += "dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=africa,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=europe,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=fareast,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=middleeast,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=northamerica,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=redmond,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=southamerica,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=southpacific,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=mslpa,dc=corp,dc=microsoft,dc=com"
$arrayDomain += "dc=windeploy,dc=ntdev,dc=microsoft,dc=com"
$arrayDomain += "dc=wingroup,dc=windeploy,dc=ntdev,dc=microsoft,dc=com"
$arrayDomain += "dc=xcorp,dc=microsoft,dc=com"
$arrayDomain += "dc=xred,dc=xcorp,dc=microsoft,dc=com"
#>

$ErrorActionPreference = "Continue"

foreach ($strDomain in $arrayDomain) {
  $domHASH = @{}
  $arrCSV = @()

  $domain = $(($strDomain).toupper().substring(3).replace(",DC=", "."))
  $domainname = $domain.split(".")[0]

  $root = [ADSI]"://$strDomain"
  $OUSearch = [System.DirectoryServices.DirectorySearcher]$root
  $OUSearch.Filter = "(objectClass=organizationalUnit)"
  $OUSearch.PageSize = 1000
  $colProplist = "distinguishedname"
  foreach ($i in $colPropList){ [void]$OUSearch.PropertiesToLoad.Add($i) }  
  $result = $OUSearch.Findall()
  foreach ($OU in $result) {
    
    $objRoot = [ADSI]"://$($OU.properties.distinguishedname)"
    $objSearch = [System.DirectoryServices.DirectorySearcher]$objRoot
    $objSearch.PageSize = 1000
    $objSearch.SearchScope = "OneLevel"
    $objSearch.Filter = "(objectClass=computer)"
    $colProplist = "operatingSystemVersion" 
    foreach ($i in $colPropList){ [void]$objSearch.PropertiesToLoad.Add($i) }  
    $objResult = $objSearch.Findall()
    if ($objResult.count -ge 1) {
      $OUHash = @{}
      $OUHash.Add("NULL",0)
      foreach ($obj in $objresult) {
        if ($obj.properties.operatingsystemversion -ne $null) {
          if ($OUHash.ContainsKey($([string]$obj.properties.operatingsystemversion))) {
            $OUHash.Set_Item($([string]$obj.properties.operatingsystemversion), $([int]$OUHash.Get_Item($([string]$obj.properties.operatingsystemversion)) + 1))
          } else {
            $OUHash.Add($([string]$obj.properties.operatingsystemversion),1)
          }
        } else {
          $OUHash.Set_Item("NULL",$([int]$OUHash.Get_Item("NULL") + 1))
        }
      }       
      
      foreach ($thing in $OUHash.keys | sort) {
        if ($domHash.ContainsKey($thing)) {
          $domHash.Set_Item($thing,$([int]$domHash.$thing + [int]$OUHash.$thing))
        } else {
          $domHash.Add($thing,$($OUHash.$thing))
        }
      }
    }
  }
  write-host $domain
  foreach ($thing in $domHash.keys | sort) {

    #write-host ($thing + "," + $($domHash.$thing))

    if ($global:totalCountOS.ContainsKey($thing)) {
      $global:totalCountOS.Set_Item($thing, $([int]$global:totalCountOS.Get_Item($thing) + [int]$domHash.Get_Item($thing)))
    } else {
      $global:totalCountOS.Add($thing,$domHash.Get_Item($thing))
    }  
  }
  #write-host ""
}

write-host "TOTALS:"
foreach ($thing in $global:totalCountOS.keys | sort) {
  write-host ($thing + "," + $($global:totalCountOS.$thing))
}

write-host $(time_to_run $starttime) -foregroundcolor red