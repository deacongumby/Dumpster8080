#$strComputerName = $args[0]
$strComputer = "tk5-red-dc-30"

$global:htmlbody = @()

function addRow($rowcolor) {
  if ($rowcolor -eq "end") { $end = "/" } else {
    if ($rowcolor -ne $null) { $color = (" bgcolor=`"#" + $rowcolor + "`" ") }
  }
  $global:htmlbody += ("<" + $end + "tr" + $color + ">")
}

function addCell($value,$cellwidth) {
  if ($cellwidth -ne $null) { $width = (" width=`"" + $cellwidth + "`" ") }
  $global:htmlbody += ("<td" + $width + "> " + $value + "</td>")
}

function addHeader($rowcolor,$header,$colspan) {
  addRow $rowcolor
  $global:htmlbody += ("<td colspan=$($colspan)>$($header)</td>")
  addRow "end"
}

function addLine2col($rowcolor,$column1,$column2,$indent) {
  addRow $rowcolor
  if ($indent) { $column1 = ($("&nbsp;" * $indent) + [string]$column1) }
  addCell $column1 "128"
  addCell $column2 $null
  addRow "end"
}

function addLine3col($rowcolor,$column1,$column2,$column3,$indent) {
  addRow $rowcolor
  if ($indent) { $column1 = ($("&nbsp;" * $indent) + [string]$column1) }
  addCell ([string]$column1) "128"
  addCell $column2 "128"
  addCell $column3 $null  
  addRow "end"
}

function addTable($header,$colspan) {
  $global:htmlbody += "<table width=1024 border=`"0`" cellspacing=`"0`">"
  addRow "2c6a1f"
  $global:htmlbody += ("<td colspan=$($colspan)><font color=`"FFFFFF`"><strong>$($header)</strong></font></td>")
  addRow "end"
}

function endTable() {
  $global:htmlbody += "</table></br>"
}

$computerSystem = Get-WmiObject Win32_ComputerSystem –comp $strComputer
$operatingSystem = Get-WmiObject Win32_OperatingSystem –comp $strComputer
$IPconfigset = Get-WmiObject Win32_NetworkAdapterConfiguration -comp $strComputer
$cpuInfo = (((Get-WmiObject –class Win32_Processor –comp $strComputer)[0].name).split("@")[-1]).trim() -replace '}',''
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strComputer)         
$regNetlogon = "SYSTEM\\currentcontrolset\\services\\Netlogon\\parameters"          
$context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$($computersystem.domain)) 
$sitename = [string]([System.DirectoryServices.ActiveDirectory.Domain]::getdomain($context).domaincontrollers | select Name, SiteName | where { $_.name -match $($operatingSystem.CSName) }).sitename

if ($sitename -eq $null) { $sitename = $reg.OpenSubKey($regNetlogon).GetValue("DynamicSiteName") }
    
$Today = Get-Date

$regIPAK = $reg.OpenSubKey("software\microsoft\msnipak")           
$returnIPAK = New-Object PsObject;foreach($keyProperty in $regIPAK.GetValueNames()) { $returnIPAK | Add-Member NoteProperty $keyProperty $regIPAK.GetValue($keyProperty) } 
($base, $IPAKvalue) = $returnIPAK.windows -split ":"

addTable "Server Specific Information" 2
addLine2col $null "Server Name:" $operatingSystem.CSName 2
addLine2col $null "Domain:" $operatingSystem.domain 2
addLine2col $null "Site:" $sitename 2
addLine2col $null "Time Zone:" $("GMT" + ($operatingSystem.CurrentTimeZone/60)) 2
addLine2col $null "OS Name:" $operatingSystem.Caption 2
addLine2col $null "OS Build:" $operatingSystem.Version 2 
addLine2col $null "Install Date:" $operatingSystem.ConvertToDateTime($operatingSystem.InstallDate) 2
addLine2col $null "IPAK:" $IPAKValue 2
addLine2col $null "Boot Time:" $operatingSystem.ConvertToDateTime($operatingSystem.LastBootUpTime) 2
addLine2col $null "Server Model:" ([string]$computerSystem.Manufacturer + " " + [string]$computerSystem.Model) 2
addLine2col $null "CPU:" ([string]$computerSystem.NumberOfProcessors + " physical / " + [string]$computerSystem.NumberOfLogicalProcessors + " logical at " + $cpuInfo) 2
addLine2col $null "RAM:" ([string]([math]::round($computerSystem.TotalPhysicalMemory / 1gb)) + " GB") 2
endTable

addTable "Drive Information" 3
addLine3col "D1FBC6" "Drive" "Free" "Total" 2
$getDriveInfo = Get-WMIObject Win32_LogicalDisk -filter "DriveType=3" -computer $strComputer | Select SystemName,DeviceID,VolumeName,@{Name=”size”;Expression={“{0:N1}” -f($_.size/1gb)}},@{Name=”freespace”;Expression={“{0:N1}” -f($_.freespace/1gb)}}
foreach ($drive in $getDriveInfo) {
  addLine3col $null $drive.deviceID ("&nbsp;&nbsp;" + $drive.freespace + " GB") ("&nbsp;&nbsp;" + $drive.size + " GB") 4
} 
endTable

addTable "Network Information" 3
foreach ($IPConfig in $IPConfigSet) {
  if ($Ipconfig.Description -like "Virtual VMBus") { $VirtualMachine = $true }
  if ($Ipconfig.IPaddress) {

    addLine2col $null "NIC:" $IPconfig.Description 2
    addLine2col $null "MAC Address:" $IPconfig.MACAddress 2
           
    $value = $IPconfig.IPAddress -split " "
    foreach ($returned in $value) { addLine2col $null "IP Address:" $returned 2 } 

    addLine2col $null "Subnet Mask:" $IPconfig.IPSubnet 2
        
    $value = $IPconfig.DefaultIPGateway -split " "
    foreach ($returned in $value) { addLine2col $null "Default GW:" $returned 2 } 

    addLine2col $null "Primary WINS:" $IPconfig.WINSPrimaryServer 2
    addLine2col $null "Secondary WINS:" $IPconfig.WINSSecondaryServer 2
      
    foreach ($DNSserver in $IPconfig.dnsserversearchorder) { addLine2col $null "DNS Server:" $DNSserver 2 }
    if ($IPconfig.TcpipNetbiosOptions -eq "1") {
      $NetbiosoverTCP = "Enabled"
    } elseif ($IPconfig.TcpipNetbiosOptions -eq "2") {
      $NetbiosoverTCP = "Disabled"
    } elseif ($IPconfig.TcpipNetbiosOptions -eq "0") {
      $NetbiosoverTCP = "Set in DHCP"
    }
    addLine2col $null "NetBios over TCP/IP:" $NetbiosoverTCP 2       
  }
}
endTable

function get_registry($key,$propertyvalue) {
  $ErrorActionPreference = "SilentlyContinue"
  $regkey = $reg.OpenSubKey($key).GetValue($propertyvalue)
  addLine2col $null "$($propertyvalue):" $regkey 2
  $regkey = $null
}  

addTable "Registry Information" 3 
get_registry $regNetlogon "SrvPriority"
get_registry $regNetlogon "SrvWeight"
get_registry $regNetlogon "MaxConcurrentApi"
endTable

function get-DHCPScopes($strComputer) {
  $result = netsh dhcp server \\$strComputer show scope
  $activecount = 0
  $disabledcount = 0

  foreach ($line in $result) {
    if ($line -match '^\s*\d') {
      if ($line -match 'Active') { $activecount++ }
      if ($line -match 'Disabled') { $disabledcount++ }
    }
  }

  if ($activecount -eq 0) {
    addLine2col $null "Active Scopes:" 0 2
  } else {
    addLine2col $null "Active Scopes:" $activecount 2
  }

  if ($disabledcount -eq 0) {
    addLine2col $null "Disabled Scopes:" 0 2
  } else {
    addLine2col $null "Disabled Scopes:" $disabledcount 2
  }
}

addTable "DHCP Scope Information" 3 
get-DHCPScopes $strComputer
endTable
