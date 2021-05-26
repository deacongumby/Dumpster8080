$domains = $args[0]

if ($domains -eq $null) {
  return
}

$runninglocation = "c:\scripts"

$ErrorActionPreference = "SilentlyContinue"

$start = Get-Date

$htmlfile = ($runninglocation + "\" + $domains + ".inc")
$weblocation = "W:\ADTools\Scripts\Sched_Tasks\DCChkWeb\new"

if (Test-Path -path $buglookup) { 
  $WinDDTBug = Get-Content $buglookup
}

if (Test-Path -path $htmlfile) { 
  remove-item $htmlfile
  Start-Sleep 5
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

function addHeader($domain) {
  Add-Content $htmlfile "<table width=1024 border=`"0`" cellspacing=`"0`">"
  add-Content $htmlfile ("<td><bold><div align=left>Domain: " + $domain.ToUpper() + "</bold></div></td><td><div align=right>Last Checked:  " + $date + " PST</div></td>")
  Add-Content $htmlfile "</table>" 
}

function get-ModelInfo($strDomain) {
  $date = Get-Date
  $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext(“Domain”,$domain)
  $dclist = [System.DirectoryServices.ActiveDirectory.DomainController]::findall($context) | Sort-Object Name

  addHeader($domain)
  Add-Content $htmlfile ("<table width=1024 border=`"0`" cellspacing=`"0`" id=`"" + ($domain.split(".")[0]) + "`"><thead>")
  addRow "7DC86A"
  addCell "Domain" "120"
  addCell "<center>IP Address</center>" "110"   
  addCell $null $null  
  addRow "end"  
  Add-Content $htmlfile ("</thead>")




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






  Add-Content $htmlfile "</table></br>"








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
  iMSNBC        { getDCs "interactive.msnbc.com" }
}

Start-Sleep 5
copy-item $htmlfile $weblocation


