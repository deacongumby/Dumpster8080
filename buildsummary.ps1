[datetime]$starttime = Get-Date

$global:totalcountRODC = 0
$global:totalcountFULL = 0
$global:totalModelHash = @{}
$global:totalModelHash.Add("No Model Listed",0)
$global:totalBuildHash = @{}
$global:totalBuildHash.Add("No Build Listed",0)
$global:countTotal = 0
$global:domainbody = @()

#$runninglocation = "g:\dumpster8080"
#$runninglocation = "C:\Scripts\Build Report"
$runninglocation = "C:\inetpub\wwwroot"

$ErrorActionPreference = "SilentlyContinue"

$htmlfile = ($runninglocation + "\summary.htm")
$weblocation = "G:\Dumpster8080\website"

if (Test-Path -path $htmlfile) { 
  remove-item $htmlfile
  Start-Sleep 5
}

function time_to_run($strStart) {
  [datetime]$strEnd = Get-Date
  $strTimeDiff = new-TimeSpan $strStart $strEnd
  $strDays = $strTimeDiff.days
  $strHours = $strTimeDiff.Hours
  $strMin = $strTimeDiff.Minutes
  $strSec = $strTimeDiff.Seconds
  return "Execution Time: $(if ($strDays -ne 0) {([string]$strDays + " days")})$(if ($strHours -ne 0) {([string]$strHours + " hours")})$(if ($strMin -ne 0) {([string]$strMin + " min")}) $strSec sec"
}

function isRODC($strComputer,$sitename,$forestDN) {
  $domainDN = ("CN=NTDS Settings,CN=" + $strComputer.split(".")[0] + ",CN=Servers,CN=" + $sitename + ",CN=Sites,CN=Configuration," + $forestDN)
  $objDomain = New-Object System.DirectoryServices.DirectoryEntry("://" + $domainDN)
  $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
  $objSearcher.SearchRoot = $objDomain
  $objSearcher.filter ="(&(objectCategory=*))" 
  $colResults = $objSearcher.FindOne()
  if ($colresults.properties.objectcategory -ne $null) { 
    if ($colresults.properties.objectcategory -match "NTDS-DSA-RO") {
      $global:countRODC++
      $global:totalcountRODC++
      return
    }
  }
  $global:countFULL++
  $global:totalcountFULL++
}

function getDCs($domain) {
  write-host $domain
  $date = Get-Date
  $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext(“Domain”,$domain)
  $dclist = [System.DirectoryServices.ActiveDirectory.DomainController]::findall($context) | Sort-Object Name

  $sumDIT = 0
  $sumWS = 0
  $global:countRODC = 0
  $global:countFULL = 0
  $ModelHash = @{}
  $ModelHash.Add("No Model Listed",0)
  $BuildHash = @{}
  $BuildHash.Add("No Build Listed",0)
  
  $count = $dclist.count
  $global:countTotal = $global:countTotal + $count

  foreach ($DC in $DCList) { 

    $strComputer = $dc.name

    $operatingSystem = Get-WmiObject Win32_OperatingSystem –comp $strComputer
    $computerSystem = Get-WmiObject Win32_ComputerSystem -property model –comp $strComputer

    # Model
    if ($computerSystem.model -eq "") {
      $ModelHash.Set_Item("No Model Listed",$([int]$ModelHash.Get_Item("No Model Listed") + 1))
    } elseif ($ModelHash.ContainsKey($([string]$computerSystem.model))) {
      $ModelHash.Set_Item($([string]$computerSystem.model), $([int]$ModelHash.Get_Item($([string]$computerSystem.model)) + 1))
    } else {
      $ModelHash.Add($([string]$computerSystem.model),1)
    }

    # Build
    if ($operatingSystem.Version -eq "") {
      $BuildHash.Set_Item("No Build Listed",$([int]$BuildHash.Get_Item("No Build Listed") + 1))
    } elseif ($BuildHash.ContainsKey($([string]$operatingSystem.Version))) {
      $BuildHash.Set_Item($([string]$operatingSystem.Version), $([int]$BuildHash.Get_Item($([string]$operatingSystem.Version)) + 1))
    } else {
      $BuildHash.Add($([string]$operatingSystem.Version),1)
    }
    
    isRODC $strComputer $dc.sitename $("dc=" + $($dc.forest.name).replace(".",",dc="))

    $query = "Select Message from Win32_NTLogEvent Where LogFile='Directory Service' and Eventcode='1646'and TimeWritten >'"+ (get-date).AddDays(-1) +"'"
    $returnlog = get-wmiobject -query $query -computername $strComputer | select -first 1
    $arrEvent = $ReturnLog.message.split("`n")
    $sumWS = $sumWS + $([Math]::Round($arrEvent[3],0))
    $sumDIT = $sumDIT + $([Math]::Round($arrEvent[5],0))
  }

  $global:domainbody += ("<table width=1024 border=`"0`" cellspacing=`"0`"><tr bgcolor=`"#006600`"><td width=`"50%`"><div align=left><font color=`"#FFFFFF`">Domain: <strong>" + $domain.toUpper() + "</strong></font></div></td>")
  $global:domainbody += "<td><div align=right><font color=`"#FFFFFF`">Last Checked: $(get-date)</font></div></td></tr></table>"

  $global:domainbody += ("<table width=1024 border=`"0`" cellspacing=`"0`"><tr><td width=`"50%`" valign=`"top`">")
  $global:domainbody += ("<table width=490 border=`"0`" cellspacing=`"0`"><thead><tr bgcolor=`"#7DC86A`" ><td width=`"50%`" ><strong>DC Type: </strong></td><td width=`"50%`" >&nbsp;</td></tr></thead>")

  $global:domainbody += ("<tr><td>RODC </td><td><center>$($countRODC)</center></td></tr>")
  $global:domainbody += ("<tr><td>FULL </td><td><center>$($countFULL)</center></td></tr>")
  $global:domainbody += ("<tr bgcolor=`"#D1FBC6`"><td>Total:</td><td><center>$($count)</center></td></table></td>")

  $global:domainbody += ("<td width=`"50%`" align=`"right`" valign=`"top`"><table width=490 border=`"0`" cellspacing=`"0`"><thead><tr bgcolor=`"#7DC86A`"><td width=`"50%`" ><strong>DIT Information:</strong></td><td width=`"50%`" >&nbsp; </td></tr></thead>")
  $global:domainbody += ("<tr><td>Average DIT Size </td><td><center>$([math]::round($($sumDIT / $count), 0))</center></td></tr><tr >")
  $global:domainbody += ("<td>Average WhiteSpace </td><td><center>$([math]::round($($sumWS / $count), 0))</center></td></tr></table></td></tr>")
  $global:domainbody += ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")

  if ($ModelHash.Get_Item("No Model Listed") -eq "0") { $ModelHash.Remove("No Model Listed") }
  if ($BuildHash.Get_Item("No Build Listed") -eq "0") { $BuildHash.Remove("No Build Listed") }

  $global:domainbody += ("<tr><td valign=`"top`"><table width=490 border=`"0`" cellspacing=`"0`"><thead><tr bgcolor=`"#7DC86A`" ><td width=`"50%`" ><strong>Build:</strong></td><td width=`"50%`" >&nbsp; </td></tr></thead>")

  foreach ($build in $BuildHash.keys) {
    $global:domainbody += ("<tr><td>$($build)</td><td><center>$($BuildHash.Get_Item($build))</center></td></tr>")   

    if ($global:totalBuildHash.ContainsKey($build)) {
      $global:totalBuildHash.Set_Item($build, $([int]$global:totalBuildHash.Get_Item($build) + [int]$BuildHash.Get_Item($build)))
    } else {
      $global:totalBuildHash.Add($build,$BuildHash.Get_Item($build))
    }  

  }
  $global:domainbody += ("</table></td>")

  $global:domainbody += ("<td align=`"right`" valign=`"top`"><table width=490 border=`"0`" cellspacing=`"0`"><thead><tr bgcolor=`"#7DC86A`"><td width=`"50%`" ><strong>Model:</strong></td><td width=`"50%`" >&nbsp;</td></tr></thead>")
  foreach ($model in $ModelHash.keys) {
    $global:domainbody += ("<tr><td>$($model)</td><td><center>$($Modelhash.Get_Item($model))</center></td></tr>")

    if ($global:totalModelHash.ContainsKey($model)) {
      $global:totalModelHash.Set_Item($model, $([int]$global:totalModelHash.Get_Item($model) + [int]$Modelhash.Get_Item($model)))
    } else {
      $global:totalModelHash.Add($model,[int]$Modelhash.Get_Item($model))
    } 

  }
  $global:domainbody += ("</table></td></tr></table><BR>")
  $global:countRODC = $null
  $global:countFULL = $null
}


getDCs "bacon.local"      


Add-Content $htmlfile ("<HTML><TITLE> General Domain Summary</TITLE>")
Add-Content $htmlfile ("<style type=`"text/css`">")
Add-Content $htmlfile ("td { font-family: times new roman, times, serif; font-size: 13px; }</style>")
Add-Content $htmlfile ("<body>")

Add-Content $htmlfile ("<table width=1024 border=`"0`" cellspacing=`"0`" background=`"directoryservices.jpg`">")
Add-Content $htmlfile ("<td height=`"93`" ><div align=`"right`"><strong><font size=`"3`" color=`"#FFFFFF`">")
Add-Content $htmlfile ("MSIT Managed Computer Summary&nbsp;&nbsp;</font></strong></div></td>")
Add-Content $htmlfile ("</table>")

Add-Content $htmlfile ("<table width=1024 border=`"0`" cellspacing=`"0`">")
Add-Content $htmlfile ("<td width='85%'><div align=`"right`"><a href=`"http://dumpster.redmond.corp.microsoft.com:8080/`">DC Site and Machine Information</a></div></td>")
Add-Content $htmlfile ("<td width='15%'><div align=`"right`"><a href=`"http://dumpster.redmond.corp.microsoft.com:8080/sch-tasks`">Scheduled Task Status</a></div></td>")
Add-Content $htmlfile ("</table>")

Add-Content $htmlfile ("<table width=1024 border=`"0`" cellspacing=`"0`"><tr><td width=`"50%`"><div align=left><strong><font size=`"5`">MSIT Complete Managed Summary</font></strong><strong></strong></div></td>")
Add-Content $htmlfile ("<td><div align=right>&nbsp;</div></td></tr></table>")

Add-Content $htmlfile ("<table width=1024 border=`"0`" cellspacing=`"0`">")
Add-Content $htmlfile ("<tr><td width=`"50%`" valign=`"top`">")

Add-Content $htmlfile ("<table width=490 border=`"0`" cellspacing=`"0`"><thead><tr bgcolor=`"#7DC86A`" ><td width=`"50%`" ><strong>DC Type: </strong></td><td width=`"50%`" >&nbsp;</td></tr></thead>")
Add-Content $htmlfile ("<tr><td>Total RODC </td><td><center>$($global:totalcountRODC)</center></td></tr>")
Add-Content $htmlfile ("<tr><td>Total FULL </td><td><center>$($global:totalcountFULL)</center></td></tr>")
Add-Content $htmlfile ("<tr bgcolor=`"#D1FBC6`"><td>Total:</td><td><center>$($global:countTotal)</center></td></table><br>")

if ($global:totalBuildHash.Get_Item("No Build Listed") -eq "0") { $global:totalBuildHash.Remove("No Build Listed") }

Add-Content $htmlfile ("<table width=490 border=`"0`" cellspacing=`"0`"><thead><tr bgcolor=`"#7DC86A`" ><td width=`"50%`" ><strong>Build:</strong></td><td width=`"50%`" >&nbsp; </td></tr></thead>")

foreach ($build in $global:totalBuildHash.keys | sort) {
  Add-Content $htmlfile ("<tr><td>$($build)</td><td><center>$($global:totalBuildHash.Get_Item($build))</center></td></tr>")    
}
Add-Content $htmlfile ("</table></td>")

if ($global:totalModelHash.Get_Item("No Model Listed") -eq "0") { $global:totalModelHash.Remove("No Model Listed") }

Add-Content $htmlfile ("<td width=`"50%`" align=`"right`" valign=`"top`">")
Add-Content $htmlfile ("<table width=490 border=`"0`" cellspacing=`"0`"><thead><tr bgcolor=`"#7DC86A`"><td width=`"50%`" ><strong>Model:</strong></td><td width=`"50%`" >&nbsp;</td></tr></thead>")

foreach ($model in $global:totalModelHash.keys | sort) {
  Add-Content $htmlfile ("<tr><td>$($model)</td><td><center>$($global:totalModelHash.Get_Item($model))</center></td></tr>")    
}
Add-Content $htmlfile ("</table></td></tr>")

Add-Content $htmlfile ("<tr><td>&nbsp;</td><td>&nbsp;</td></tr>")

Add-Content $htmlfile ("</table></td>")
Add-Content $htmlfile ("<td>&nbsp;</td>")

Add-Content $htmlfile ("</tr></table><BR>")
Add-Content $htmlfile ("<table width=1024 border=`"0`" cellspacing=`"0`"><tr><td width=`"50%`"><div align=left><strong><font size=`"5`">Domain Specific Summary</font></strong></div></td>")
Add-Content $htmlfile ("<td><div align=right>&nbsp;</div></td></tr></table>")

foreach ($line in $global:domainbody) {
  Add-Content $htmlfile $line
}

Add-Content $htmlfile ("</body></HTML>")

time_to_run $starttime
