
<#

.FUNCTIONALITY
-VMware relevant inventory
-Uses PowerCLI
-use install-module vmware.powercli -allowclobber as required

.SYNOPSIS
-This script was created to help others identify simple issues present in most VMware ESXi environments
-The only input required is the vCenter server name
-A time-stamped HTML report will be created and opened with the default web browser present on the system

.NOTES
Change log

Nov 11, 2020
-Initial version

Nov 12, 2020
-HTML code edited
-Code hygeine

Nov 13, 2020
-Script will exit it VMware PowerCLI failed to install
-Script will exit if PowerShell is not version 5 or above

.DESCRIPTION
Author oreynolds@gmail.com

.EXAMPLE
./Get-VMware-QuickInventory.ps1

.NOTES

.Link
N/A

#>

#$Cred = Get-Credential

### Variables & functions

### Install Nuget and VMware PowerCLI as required

IF (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    write-warning "Please open Powershell as administrator, the script will now exit"
    EXIT

}

if (-not(($PSversionTable.PSVersion).Major -ge 5)) {

    write-warning "Powershell version 5 or above is required to run this script"
    write-warning "Please download/install from here https://www.microsoft.com/en-us/download/details.aspx?id=54616"
    write-warning "The script will now exit"
    EXIT

}

IF (-not(Get-PackageProvider -ListAvailable -name NUget)) {

    Install-PackageProvider -Name NuGet -force -Confirm:$False
}

IF (-not(Get-Module -ListAvailable -name VMware.PowerCLI)) {

    Install-Module -Name VMware.PowerCLI -AllowClobber -force
}

IF (-not(Get-Module -ListAvailable -name VMware.PowerCLI)) {

    write-warning "PowerCLI failed to install. The script will exit"
    EXIT
}

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

IF ($global:DefaultVIServer.Length -eq 0) {

    Connect-VIServer

}

$ShortDate = (Get-Date).ToString('MM-dd-yyyy')
$LogTimeStamp = (Get-Date).ToString('MM-dd-yyyy-hhmm-tt')

$Head = @"
<style>

    h1 {

        font-family: Arial, Helvetica, sans-serif;
        color: #e68a00;
        font-size: 28px;

    }

    
    h2 {

        font-family: Arial, Helvetica, sans-serif;
        color: #000099;
        font-size: 16px;

    }

    
    
   table {
		font-size: 12px;
		border: 0px; 
		font-family: Arial, Helvetica, sans-serif;
	} 
	
    td {
		padding: 4px;
		margin: 0px;
		border: 0;
	}
	
    th {
        background: #395870;
        background: linear-gradient(#49708f, #293f50);
        color: #fff;
        font-size: 11px;
        text-transform: uppercase;
        padding: 10px 15px;
        vertical-align: middle;
	}

    tbody tr:nth-child(even) {
        background: #f0f0f2;
    }

        #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #ff3300;
        font-size: 12px;

    }
    



</style>
"@



### header
$ReportTitle = "<h1>VMware quick environmental report for $($env:USERDNSDOMAIN)"
$ReportTitle += "<h2>Data is from $($LogTimeStamp)</h2>"
$ReportTitle += "<hr></hr>"
#$ReportTitle += "<br><br>"

### 1 
write-host "Collecting VM network card type" -ForegroundColor Cyan
$IntelNICs = get-vm  | get-networkadapter | Where-object {$_.Type -ne 'VMxnet3'} | Select-Object @{E={$_.Parent};Name="VM"}, @{E={$_.Type};Name="Network card type"}

IF ($IntelNICs.Length -eq 0) {

    write-host "All network card types are set correctly to VMware VXnet 3 type" -ForegroundColor Green
    $Pre1 = "<H2>PASS: All Network card type attached to shells are set correctly</H2>"
    #$Pre1 += "<br><br>"
    $Section1HTML = $Pre1
    
}

Else {

    Write-Warning "$($IntelNICs | Measure-Object | Select-object -ExpandProperty Count) VM(s) were found with legacy intel network card types"
    $Pre1 = "<H2>WARNING: Intel E1000 legacy network card type VMs detected</H2>"
    #$Pre1 += "<br><br>"
    $Section1HTML = $IntelNICs | ConvertTo-HTML -Head $Head -PreContent $Pre1 -As Table | Out-String

}

### 2 
write-host "Collecting ESXi power profile" -ForegroundColor Cyan
### https://www.cloudishes.com/2015/09/automating-esxi-host-power-management.html
$HostPowerPolicy = Get-VMHost | Sort-Object | Select-Object -property Name,
@{ N="CurrentPolicy"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.ShortName}},
@{ N="CurrentPolicyKey"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.Key}},
@{ N="AvailablePolicies"; E={$_.ExtensionData.config.PowerSystemCapability.AvailablePolicy.ShortName}} `
| Where-Object {$_.CurrentPolicy -ne "Static"} | Select-Object -Property Name, CurrentPolicy

IF ($HostPowerPolicy.Length -eq 0) {
    
    write-host "All ESXi hosts are set correctly to HIGH PERFORMANCE" -ForegroundColor Green
    $Pre2 = "<H2>PASS: All ESXi hosts are set correctly to HIGH PERFORMANCE</H2>"
    #$Pre2 += "<br><br>"
    $Section2HTML = $Pre2
}

else {

    write-warning "The below ESXi hosts that are NOT set to HIGH PERFORMANCE"
    $Pre2 = "<H2>WARNING: The below ESXi hosts that are NOT set to HIGH PERFORMANCE</H2>"
    #$Pre2 += "<br><br>"
    $Section2HTML = $HostPowerPolicy | ConvertTo-HTML -Head $Head -PreContent $Pre2 -As Table | Out-String

}

### 3 - SCSI controller types
$SCSIControllerTypes = Get-VM | Select-Object Name,@{N="Cluster";E={Get-Cluster -VM $_}},@{N="Controller Type";E={Get-ScsiController -VM $_ | Select-Object -ExpandProperty Type}} `
| Where-Object {$_."Controller Type" -eq "VirtualLsiLogicSAS "} | Select-Object @{E={$_.Name};Name="VM"}, "Controller Type"


IF ($SCSIControllerTypes.Length -eq 0) {
    
    write-host "All SCSI controllers are set correctly to VMWware Paravirtual" -ForegroundColor Green
    $Pre3 = "<H2>PASS: All SCSI controllers are set correctly to VMWware Paravirtual</H2>"
    #$Pre3 += "<br><br>"
    $Section3HTML = $Pre3
}

else {

    write-warning "The below VMs are using legacy LSI SCSI controller types"
    $Pre3 = "<H2>WARNING: The below VMs are using legacy LSI SCSI controller types</H2>"
    #$Pre3 += "<br><br>"
    $Section3HTML = $SCSIControllerTypes | ConvertTo-HTML -Head $Head -PreContent $Pre3 -As Table | Out-String

}

### 4 - VMWare tools version
$VMWareTools = get-vm | Select-Object Name,@{Name="ToolsVersion";Expression={$_.ExtensionData.Guest.ToolsVersion}},@{Name="ToolsStatus";Expression={$_.ExtensionData.Guest.ToolsVersionStatus}} | Sort-Object -Property ToolsStatus

$VMWareTools = get-vm | Select-Object Name,@{Name="ToolsVersion";Expression={$_.ExtensionData.Guest.ToolsVersion}},@{Name="ToolsStatus";Expression={$_.ExtensionData.Guest.ToolsVersionStatus}} `
| Where-object {$_.ToolsStatus -ne "guestToolsCurrent"} | Where-object {$_.ToolsStatus -ne "guestToolsUnmanaged"}

If ($VMWareTools.length -eq 0) {

    write-host "VMware tools is up to date on all VMs" -ForegroundColor Green
    $Pre4 = "<H2>PASS: VMware tools is up to date on all VMs</H2>"
    #$Pre4 += "<br><br>"
    $Section4HTML = $Pre4
}

else {

    write-warning "The below VMs are running older versions of VMware tools and should be updated"
    $Pre4 = "<H2>WARNING: The below VMs are running older versions of VMware tools and should be updated</H2>"
    #$Pre4 += "<br><br>"
    $Section4HTML = $VMWareTools | ConvertTo-HTML -Head $Head -PreContent $Pre4 -As Table | Out-String

}

### 5 - Get ESXi host BIOS version

write-host "Collecting hardware BIOS info"

$ESXiBios = Get-View -ViewType HostSystem | Select-Object Name,@{N="BIOS version";E={$_.Hardware.BiosInfo.BiosVersion}}, @{N="BIOS date";E={$_.Hardware.BiosInfo.releaseDate}}

If ($ESXiBios.length -eq 0) {

    $Pre5 = "<H2>WARNING: ESXi host bios info is not available at this time</H2>"
    #$Pre5 += "<br><br>"
    $Section5HTML = $Pre5
}

Else {

    $Pre5 = "<H2>INFO: ESXi host bios summary</H2>"
    #$Pre5 += "<br><br>"
    $Section5HTML = $ESXiBios | ConvertTo-HTML -Head $Head -PreContent $Pre5 -As Table | Out-String

}

### 6 - ESXi host NTP settings

write-host "Collecting NTP service config"

$NTP = Get-VMHost | Sort-Object Name | Select-Object Name, @{N=“NTPServiceRunning“;E={($_ | Get-VmHostService | Where-Object {$_.key-eq “ntpd“}).Running}},`
@{N=“StartupPolicy“;E={($_ | Get-VmHostService | Where-Object {$_.key-eq “ntpd“}).Policy}}, @{N=“NTPServers“;E={$_ | Get-VMHostNtpServer}}, @{N="Date&Time";E={(get-view $_.ExtensionData.configManager.DateTimeSystem).QueryDateTime()}}

IF ($NTP.length -eq 0) {

    $Pre6 = "<H2>WARNING: ESXi NTP information is not available at this time/H2>"
    #$Pre6 += "<br><br>"
    $Section6HTML = $Pre6

}

Else {

    $Pre6 = "<H2>INFO: ESXi NTP settings</H2>"
    #$Pre6 += "<br><br>"
    $Section6HTML = $NTP | ConvertTo-HTML -Head $Head -PreContent $Pre6 -As Table | Out-String

}


### 7 - VM Hardware version

write-host "Collecting VM shell hardware version"

$VMHardwareVersion = Get-VM | Select-Object name, HardwareVErsion | Sort-Object HardwareVersion

If ($VMHardwareVersion.Length -eq 0) {

    $Pre7 = "<H2>WARNING: VM Shell hardware version info is not available at this time</H2>"
    #$Pre7 += "<br><br>"
    $Section7HTML = $Pre7

}

Else {

    $Pre7 = "<H2>INFO: VM Shell hardware version summary</H2>"
    #$Pre7 += "<br><br>"
    $Section7HTML = $VMHardwareVersion | ConvertTo-HTML -Head $Head -PreContent $Pre7 -As Table | Out-String

}

### 8 - vCPU to pCPU ratio
### Should not be above 5 for Citrix environments

write-host "Collecting vCPU / pCPU ratio"

$ESXihosts = Get-VMhost | Where-Object {$_.model -ne "VMware Virtual Platform"} | Select-Object Name, NumCpu

$RatioSummary = @()

ForEach ($i in $ESXihosts) {

    write-host "Checking $($i.name)"

    $Ratio = (Get-VMHost $i.name | Get-VM | Where-object Name -notlike "vcls*" | Select-Object -expandProperty NumCPU | Measure-Object -sum | Select-Object -ExpandProperty Sum) / $i.NumCpu
    
    $RatioSummary += New-Object -TypeName PSObject -Property @{

    ESXihost = $i.Name
    Ratio = $Ratio

    }

}

if ($RatioSummary.Length -eq 0) {

    $Pre8 = "<H2>WARNING: ESXi vCPU to pCPU info is not available at this time</H2>"
    #$Pre8 += "<br><br>"
    $Section8HTML = $Pre8

}

Else {

    $Pre8 = "<H2>INFO: ESXi vCPU to pCPU ratio summary</H2>"
    #$Pre8 += "<br><br>"
    $Section8HTML = $RatioSummary | ConvertTo-HTML -Head $Head -PreContent $Pre8 -As Table | Out-String
}

$HTMLReport = ""
$HTMLReport = ConvertTo-HTML -Body "$ReportTitle $Section1HTML $Section2HTML $Section3HTML $Section4HTML $Section5HTML $Section6HTML $Section7HTML $Section8HTML" -Title "VMware Quick Environmental report"

$HTMLReport | out-file .\"VMWare-QuickInventory-$LogTimeStamp.html"
Invoke-Item "VMWare-QuickInventory-$LogTimeStamp.html"

write-host "Script is done!" -ForegroundColor Cyan