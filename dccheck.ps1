$strComputer = "tk5-ext-dc-03.extranet.microsoft.com"

$computerSystem = Get-WmiObject Win32_ComputerSystem –comp $strComputer
$operatingSystem = Get-WmiObject Win32_OperatingSystem –comp $strComputer
$IPconfigset = Get-WmiObject Win32_NetworkAdapterConfiguration -comp $strComputer
$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $strComputer)         
$regNetlogon = "SYSTEM\\currentcontrolset\\services\\Netlogon\\parameters"          

$regIPAK = $reg.OpenSubKey("software\microsoft\msnipak")           
$returnIPAK = New-Object PsObject;foreach($keyProperty in $regIPAK.GetValueNames()) { $returnIPAK | Add-Member NoteProperty $keyProperty $regIPAK.GetValue($keyProperty) } 
($base, $IPAKvalue) = $returnIPAK.windows -split ":Build "
$rawuptime = ($operatingSystem.ConvertToDateTime($operatingSystem.LocalDateTime) - $operatingSystem.ConvertToDateTime($operatingSystem.LastBootUpTime))
$uptime = (([string]($rawuptime.days)+"D")+" "+([string]($rawuptime.hours)+"H")+" "+([string]($rawuptime.minutes)+"M"))

$FQDN = ($operatingSystem.CSName + "." + $computersystem.domain)

$ipaddress = [System.Net.Dns]::GetHostAddresses($FQDN) | where {$_.AddressFamily -eq "InterNetwork"}

$Today = Get-Date

$ping = new-object System.Net.NetworkInformation.Ping
$Reply = $ping.send($FQDN)

if (!(Test-Path -path \\$strComputer\sysvol)) { $sysvol = $false } else { $sysvol = $true }


write-host "========================================================"
write-host $computersystem.domain
write-host "========================================================"

write-host $operatingSystem.CSName  #name
write-host $operatingSystem.Version  #build
if ($operatingsystem.ServicePackMajorVersion -ne "0") { write-host $operatingsystem.ServicePackMajorVersion } #servicepack 
write-host $ipaddress[0]   #ipaddress
write-host $reg.OpenSubKey($regNetlogon).GetValue("DynamicSiteName")  #site
write-host $reply.status #ping
write-host $sysvol #sysvol exists
write-host $IPAKValue  #ipak
write-host $uptime  #uptime


function Check_Port($portname,$port,$timeout)
{   
  $ErrorActionPreference = "SilentlyContinue"
  trap {  Write-host "Failed"  }    
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
    write-host “Failed”
  } else {
    write-host “Successful”    
  }
}

Check_Port "DC Port" 389 3000
Check_Port "GC Port" 3268 3000


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

check_query 389 $dcchecking
check_query 3268 $gcchecking
