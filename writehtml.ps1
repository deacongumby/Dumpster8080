$htmlfile = ([string](get-location) + "\corp.html")

if (Test-Path -path $htmlfile) { 
  remove-item $htmlfile
}

$start = Get-Date

function Check_Port($portname,$port,$timeout)
{   
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
    return “Failed”
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

function check_query($port,$querycheck) {
  if ($querycheck -eq "Failed") {
    write-host $querycheck
  } else {
    $objDomain = New-Object System.DirectoryServices.DirectoryEntry("://" + $FQDN + ":" + $port)
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.filter ="(&(CN=Domain Admins))" 
    $colResults = $objSearcher.FindOne()
    if ($colResults -ne $null) { write-host "Successful" } else { write-host "Failed" }  
    $colResults = $null
  }
}

function addRow($rowcolor) {
  if ($rowcolor -eq "end") { $end = "/" } else {
    if ($rowcolor -ne $null) { $color = (" bgcolor=`"#" + $rowcolor + "`" ") }
  }
  Add-Content $htmlfile ("<" + $end + "tr" + $color + ">")
}

function addCell($value,$cellwidth) {
  if ($cellwidth -ne $null) { $width = (" width=`"" + $cellwidth + "`" ") }
  Add-Content $htmlfile ("<td" + $width + "> " + $value + "</td>")
}

Add-Content $htmlfile "<HTML>"
Add-Content $htmlfile ("<TITLE> Status as of " + (get-date) + " </TITLE>")

Add-Content $htmlfile "<style type=`"text/css`">"
Add-Content $htmlfile "td { font-family: times new roman, times, serif; font-size: 13px; }"
Add-Content $htmlfile "</style>"

Add-Content $htmlfile "<img src=`"directoryservices.jpg`"></BR>"

function getDCs($domain) {
  $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext(“Domain”,$domain)
  $dclist = [System.DirectoryServices.ActiveDirectory.DomainController]::findall($context)
  Add-Content $htmlfile "<table width=1024 border=`"0`" cellspacing=`"0`">"
  add-Content $htmlfile ("<td colspan=8><bold>Domain: " + $domain + "</bold></td>")
  addRow "7DC86A"
  addCell "Server" "120"
  addCell "<center>IP Address</center>" "120"
  addCell "<center>SiteName</center>" "120"
  addCell "<center>PDC</center>" "40"  
  addCell "<center>RID</center>" "40"  
  addCell "<center>INF</center>" "40"  
  addCell "<center>SCH</center>" "40"  
  addCell "<center>DOM</center>" "40"
  
  addCell $null $null  

  $count = "0"
    
  foreach ($DC in $DCList) {
    $strComputer = $dc.name
    if ($dc.roles -match "PdcRole") { $rolePDC = "X" } else { $rolePDC = $null }
    if ($dc.roles -match "RidRole") { $roleRID = "X" } else { $roleRID = $null }
    if ($dc.roles -match "InfrastructureRole") { $roleInfra = "X" } else { $roleInfra = $null }
    if ($dc.roles -match "SchemaRole") { $roleSchema = "X" } else { $roleSchema = $null }
    if ($dc.roles -match "NamingRole") { $roleDomainNaming = "X" } else { $roleDomainNaming = $null }

    #$ipaddress = ([net.dns]::GetHostEntry($strComputer).addresslist | where { $_.addressfamily -ne "interNetworkv6" }).ipaddresstostring
    
    $strComputer
    
    $ipaddresslist = [System.Net.Dns]::GetHostAddresses($strComputer) | where {$_.AddressFamily -eq "InterNetwork"}
    foreach ($ip in $ipaddresslist) { $ipaddress = $ip.ipaddresstostring }
        
    $pingreply = ($ping.send($strComputer)).status
    if ($pingreply -eq "TimedOut") { $pingreply = ($ping.send($strComputer)).status }
  
    if ($pingreply -eq "Success") {  
      if ($count -eq "0") { addRow $null;$count = "1" } else { addRow "D1FBC6"; $count = "0" }
      addCell (($dc.name.split(".")[0]).ToUpper()) $null
      addCell ("<center>" + $ipaddress + "</center>") $null    
      addCell ("<center>" + $dc.sitename + "</center>") $null        
      addCell ("<center>" + $rolePDC + "</center>") $null
      addCell ("<center>" + $roleRID + "</center>") $null
      addCell ("<center>" + $roleInfra + "</center>") $null
      addCell ("<center>" + $roleSchema + "</center>") $null
      addCell ("<center>" + $roleDomainNaming + "</center>") $null
      
      addRow "end"
    }
  }
  Add-Content $htmlfile "</table></br>"
}

getDCs "bacon.local"

Add-Content $htmlfile "<!--#include file=`"testing.html`"-->"

Add-Content $htmlfile "</HTML>"
