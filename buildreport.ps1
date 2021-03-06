$domains = $args[0]

if ($domains -eq $null) {
  write-host "no domain given, exiting"
  return
}

$runninglocation = "g:\dumpster8080"

$ErrorActionPreference = "SilentlyContinue"

$start = Get-Date

#$htmlfile = ([string](get-location) + "\" + $domains + ".inc")
#$buglookup = ([string](get-location) + "\buglookup.txt")  

$htmlfile = ($runninglocation + "\" + $domains + ".inc")
$buglookup = ($runninglocation + "\buglookup.txt")
$weblocation = "G:\Dumpster8080\website"

if (Test-Path -path $buglookup) { 
  $WinDDTBug = Get-Content $buglookup
}

if (Test-Path -path $htmlfile) { 
  remove-item $htmlfile
  Start-Sleep 5
}

function Check_Port($port,$timeout) {   
  $ErrorActionPreference = "SilentlyContinue"
  trap {  return "Failed"  }    
  $tcpclient = new-Object system.Net.Sockets.TcpClient
  $ipconnect = $tcpclient.BeginConnect($strComputer,$port,$null,$null)
  $wait = $ipconnect.AsyncWaitHandle.WaitOne($timeout,$false)
  if(!$wait) {
    $tcpclient.Close()
    $failed = $true     
  } else {
    $error.Clear()
    $tcpclient.EndConnect($ipconnect) | out-Null
    if($error[0]) {
        write-host $error[0]
        $failed = $true
    }
    $tcpclient.Close()
  }
  if($failed) {
    return $NULL
  } else {
    $objDomain = New-Object System.DirectoryServices.DirectoryEntry("://" + $strComputer + ":" + $port)
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.filter ="(&(CN=Domain Admins))" 
    $colResults = $objSearcher.FindOne()
    if ($colResults -ne $null) { return "Success" } else { return "Failed" }  
    $colResults = $null     
  }
}

function isRODC($strComputer,$sitename,$domainDN) {
  $objDomain = New-Object System.DirectoryServices.DirectoryEntry("://" + $domainDN)
  $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
  $objSearcher.SearchRoot = $objDomain
  $objSearcher.filter ="(objectCategory=*)" 
  $colResults = $objSearcher.FindOne()
  if ($colresults.properties.objectcategory -ne $null) { 
    if ($colresults.properties.objectcategory -match "NTDS-DSA-RO") {
      Return "X"
    }
  }
  Return $null
}

function addRow($rowcolor) {
  if ($rowcolor -eq "end") { $end = "/" } else {
    if ($rowcolor -ne $null) { $color = (" bgcolor=`"#" + $rowcolor + "`" ") }
  }
  Add-Content $htmlfile ("<" + $end + "tr" + $color + ">")
  #write-host ("<" + $end + "tr" + $color + ">")
}

function addCell($value,$cellwidth) {
  if ($cellwidth -ne $null) { $width = (" width=`"" + $cellwidth + "`" ") }
  Add-Content $htmlfile ("<td" + $width + "> " + $value + "</td>")
  #write-host ("<td" + $width + "> " + $value + "</td>")
}

function addHeader($domain) {
  Add-Content $htmlfile "<table width=1024 border=`"0`" cellspacing=`"0`">"
  add-Content $htmlfile ("<td><bold><div align=left>Domain: " + $domain.ToUpper() + "</bold></div></td><td><div align=right>Last Checked:  " + $date + " PST</div></td>")
  Add-Content $htmlfile "</table>" 
}

function addRollover($rollovername) {
  Add-Content $htmlfile ("<script type=`"text/javascript`">")
  Add-Content $htmlfile ("addTableRolloverEffect('" + $rollovername + "','tableRollOverEffect1','tableRowClickEffect1');")
  Add-Content $htmlfile ("</script>")
}

function getDCs($domain,$domainDN2) {
  $date = Get-Date
  $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext(“Domain”,$domain)

  try {
    $dclist = [System.DirectoryServices.ActiveDirectory.DomainController]::findall($context) | Sort-Object Name
  } catch {
    [array]$colDClist = nltest /dclist:$strDomain

    $DCList = @()

    $DomainRoles = Get-ADDomain $strDomain
    $ForestRoles = Get-ADForest $strDomain 

    $strInfra = $($DomainRoles.InfraStructureMaster)
    $strRID = $($DomainRoles.RIDMaster)
    $strPDC = $($DomainRoles.PDCEmulator)
    $strSchema = $($ForestRoles.SchemaMaster)
    $strDomainNaming = $($ForestRoles.DomainNamingMaster)

    foreach ($strReturn in $colDCList) {
      if ($strReturn -match "Site:") {
        $colReturn = $($strReturn.trim().replace("`[PDC`]","")) -split ("\s+")
        $objReturn = New-Object System.Object
        $objReturn | Add-Member -type NoteProperty -name Name -Value $colReturn[0]
        $objReturn | Add-Member -type NoteProperty -name SiteName -Value $colReturn[3]
        if ($colReturn[0] -match $strPDC) { $role += "PdcRole," }
        if ($colReturn[0] -match $strRID) { $role += "RidRole," }
        if ($colReturn[0] -match $strInfra) { $role += "InfrastructureRole," }
        if ($colReturn[0] -match $strSchema) { $role += "SchemaRole," }
        if ($colReturn[0] -match $strDomainNaming) { $role += "NamingRole," }
        $objReturn | Add-Member -type NoteProperty -name Roles -Value $role
        $DCList += $objReturn
        $objReturn = $null
        $role = $null
      }
    }
  }


  addHeader($domain)
  Add-Content $htmlfile ("<table width=1024 border=`"0`" cellspacing=`"0`" id=`"" + ($domain.split(".")[0]) + "`"><thead>")
  addRow "7DC86A"
  addCell "Server" "120"
  addCell "<center>IP Address</center>" "110"
  addCell "<center>Version</center>" "60"  
  addCell "<center>SP</center>" "50"    
  addCell "<center>SiteName</center>" "120"
  addCell "<center>PDC</center>" "35"  
  addCell "<center>RID</center>" "35"  
  addCell "<center>INF</center>" "35"  
  addCell "<center>SCH</center>" "35"  
  addCell "<center>DOM</center>" "35"
  addCell "<center>RODC</center>" "40"  
  addCell "<center>DC Query</center>" "70"
  addCell "<center>GC Query</center>" "70" 
  addCell "<center>Sysvol</center>" "55"     
  addCell "<center>DHCP</center>" "60" 
  addCell "<center>Debug</center>" "44"
  addCell "<center>Assigned</center>" "50"      
   
  addCell $null $null  
  addRow "end"  
  Add-Content $htmlfile ("</thead>")

  $count = "0"
    
  foreach ($DC in $DCList) {  
    $strComputer = $dc.name
    if ($dc.roles -match "PdcRole") { $rolePDC = "X" } else { $rolePDC = $null }
    if ($dc.roles -match "RidRole") { $roleRID = "X" } else { $roleRID = $null }
    if ($dc.roles -match "InfrastructureRole") { $roleInfra = "X" } else { $roleInfra = $null }
    if ($dc.roles -match "SchemaRole") { $roleSchema = "X" } else { $roleSchema = $null }
    if ($dc.roles -match "NamingRole") { $roleDomainNaming = "X" } else { $roleDomainNaming = $null }
  
    $strComputer
    
    $domainDN = ("CN=NTDS Settings,CN=" + ($dc.name.split(".")[0]) + ",CN=Servers,CN=" + $dc.sitename + ",CN=Sites,CN=Configuration," + $domainDN2)
    
    $ipaddresslist = [System.Net.Dns]::GetHostAddresses($strComputer) | where {$_.AddressFamily -eq "InterNetwork"}
    foreach ($ip in $ipaddresslist) { $ipaddress = $ip.ipaddresstostring }

    $DHCPQuery = [WMISearcher]‘Select * from Win32_ServerFeature where Name = "DHCP Server"’
    $DHCPQuery.options.timeout = '0:0:8'
    $DHCPQuery.scope.path="\\$strComputer\root\cimv2"
    try {
      if ($($DHCPQuery.Get()).name -eq "DHCP Server") { $dhcp = "X" } else { $dhcp = $null }
    } catch {
      $dhcp = "<font color=`"red`">ERR</font>"
      $errCheck = $true
    }

    $operatingSystem = [WMISearcher]‘Select * from Win32_OperatingSystem’
    $operatingSystem.options.timeout = '0:0:8'
    $operatingSystem.scope.path="\\$strComputer\root\cimv2"
    try {
      if ($($operatingSystem.Get() | select servicepackmajorversion).ServicePackMajorVersion -ne "0") {
        $servicepack = $($($operatingsystem.Get() | select CSDVersion).CSDVersion).trim("Service Pack ")
      } else {
        $servicepack = $null
      }
    } catch {
      $servicepack = "<font color=`"red`">ERR</font>"
      $errCheck = $true
    }

    foreach ($server in $WinDDTBug) {
      if ($server -match ($dc.name.split(".")[0])) {
        ($bugservername,$bugnumber,$bugassignedto) = $server.split(",")
      }
    }

    #$pingreply = ($ping.send($strComputer)).status
    #if ($pingreply -eq "TimedOut") { $pingreply = ($ping.send($strComputer)).status }
  
    #if ($pingreply -eq "Success") {
    
      if ($bugnumber -ne $null) { 
        if ($count -eq "0") { $count = "1" } else { $count = "0" }
        addRow "FFFF00"
      } else {
        if ($count -eq "0") { 
          addRow $null;$count = "1" 
        } else { 
          addRow "D1FBC6"; $count = "0" 
        }     
      } 
      
      if (Test-Path -path \\$strComputer\sysvol) { $sysvol = "X" }
      $rodc = isRODC ($dc.name.split(".")[0]) ($dc.sitename) $domainDN
      
      addCell (($dc.name.split(".")[0]).ToUpper()) $null
      addCell ("<center>" + $ipaddress + "</center>") $null  
      addCell ("<center>" + $($operatingSystem.Get() | select Version).Version + "</center>") $null
      addCell ("<center>" + $servicepack + "</center>") $null        
      addCell ("<center>" + $dc.sitename + "</center>") $null
      addCell ("<center>" + $rolePDC + "</center>") $null
      addCell ("<center>" + $roleRID + "</center>") $null
      addCell ("<center>" + $roleInfra + "</center>") $null
      addCell ("<center>" + $roleSchema + "</center>") $null
      addCell ("<center>" + $roleDomainNaming + "</center>") $null   
      addCell ("<center>" + $rodc + "</center>") $null          
      addCell ("<center>" + (check_port 389 3000) + "</center>") $null 
      addCell ("<center>" + (check_port 3268 3000) + "</center>") $null 
      addCell ("<center>" + ( $sysvol ) + "</center>") $null       
      addCell ("<center>" + ( $dhcp ) + "</center>") $null
      addCell ("<center>" + ( $bugnumber ) + "</center>") $null       
      addCell ("<center>" + ( $bugassignedto ) + "</center>") $null      
      
      addRow "end"

      $ipaddress = $null
      $operatingsystem = $null
      $servicepack = $null
      $dc = $null
      $rolepdc = $null
      $roleinfra = $null
      $roleschema = $null
      $roledomainnaming = $null
      $rodc = $null
      $sysvol = $null
      $dhcp = $null
      $dhcpquery = $null
      $bugnumber = $null
      $bugassignedto = $null        
      
    #}    
  }
  Add-Content $htmlfile "</table></br>"
  addRollover(($domain.split(".")[0]))

  #$date2 = Get-Date
  #write-host ("Running time:  " + (($date2) - ($date)).TotalSeconds + "seconds")
  #write-host ""  
}

$domainXCORP = "DC=xcorp,DC=microsoft,DC=com"
$domainCORP = "DC=corp,DC=microsoft,DC=com"
$domainEXT = "DC=extranet,DC=microsoft,DC=com"
$domainET = "DC=extranettest,DC=microsoft,DC=com"

switch ($domains) {

  corp         { getDCs "corp.microsoft.com" $domainCORP }        
  africa       { getDCs "africa.corp.microsoft.com" $domainCORP }        
  europe       { getDCs "europe.corp.microsoft.com" $domainCORP }
  fareast      { getDCs "fareast.corp.microsoft.com" $domainCORP }
  middleeast   { getDCs "middleeast.corp.microsoft.com" $domainCORP }
  northamerica { getDCs "northamerica.corp.microsoft.com" $domainCORP }
  redmond      { getDCs "redmond.corp.microsoft.com" $domainCORP }
  southamerica { getDCs "southamerica.corp.microsoft.com" $domainCORP }
  southpacific { getDCs "southpacific.corp.microsoft.com" $domainCORP }
  exchange     { getDCs "exchange.corp.microsoft.com" $domainCORP }
  mslpa        { getDCs "mslpa.corp.microsoft.com" $domainCORP }
  wgia         { getDCs "wgia.corp.microsoft.com" $domainCORP }
  windeploy    { getDCs "windeploy.ntdev.microsoft.com" $domainCORP }
  wingroup     { getDCs "wingroup.windeploy.ntdev.microsoft.com" $domainCORP }
  winse        { getDCs "winse.corp.microsoft.com" $domainCORP }
  segroup      { getDCs "segroup.winse.corp.microsoft.com" $domainCORP }          
  xcorp        { getDCs "xcorp.microsoft.com" $domainXCORP }
  xred         { getDCs "xred.xcorp.microsoft.com" $domainXCORP }
  
  extranet     { getDCs "extranet.microsoft.com" $domainEXT }
  partners     { getDCs "partners.extranet.microsoft.com" $domainEXT }
  extranettest { getDCs "extranettest.microsoft.com" $domainET }
  parttest     { getDCs "parttest.extranettest.microsoft.com" $domainET }
  
  ntdev         { getDCs "ntdev.corp.microsoft.com" }
  interactive   { getDCs "interactive.msnbc.com" }
}

#$endtime = Get-Date
#write-host ""
#write-host ""
#write-host ("Start Processing Time: " + $start)
#write-host ("End Processing time:   " + $endtime)
#write-host ("Running time:  " + (($endtime) - ($start)).TotalSeconds + "seconds")

Start-Sleep 5
copy-item $htmlfile $weblocation