<#

.FUNCTIONALITY
-VMware relevant inventory
-Uses PowerCLI, install via install-module vmware.powercli -scope AllUsers -force -SkipPublisherCheck -AllowClobber
-use install-module vmware.powercli -allowclobber as required
-install 

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

Nov 17, 2020
-Connect-ViServer code changed
-Hardware collection info changed
-Color coding added for NTP section
-Added CPU Ready Time
-Added DNS servers set on ESXihost

Nov 18, 2020
-Added back missing Connect-VIserver

Nov 19, 2020
-$CurrentDir added
-Estimated seconds processing time for CPU ready
-Reports directory created as required
-Added script processing time
-ImportExcel module will be installed for pulling in related XLS which contains VMware tools/ESXi versions

Nov 20, 2020
-Datastore checks added

June 29, 2022
-Added cluster

Dec 16, 2022
-Updated cluster scan method

Oct 12, 2023
-Added Physical CPU socket count
-Amended NumCPU to NumCPUCore

April 30, 2024
-Text prompt changes
-Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $False

Sept 25, 2025
-Updated to clean up $ESXiSummary code
-Added cluster name, DRS, connected NIC info
-Added hours, minutes, seconds to CPU ready estimate

Sept 26, 2025
-SectionHTMLs changed to actual content description: NTP, ESXi, DataStore, CPU Ready type, etc
-Relevant sections export to XLS, Excel is opened where it's installed to view the report
-Re-titled VMware tools XLS tab in workwork to 'VMWare Tools out of date'
-Added auto filter to ESXi summary and CPU ready time
-tabs for network cards, power profiles, NIC cards only created as required
-Amend warnning to menion only for VDI
-Fix formatting for vCPU to CPU ratio in HTML
-Added code to stop apending total run time to end of HTML report

Sept 29, 2025
-Code hygiene updates

Oct 14, 2025
-VM OS capture added
-Consolidated Get-VM calls for performance optimization ; -CLS VM(s) are filtered out
-Added 'All VMs' XLS tab with comprehensive VM details
-Removed CPU Ready time from HTML output (retained in XLS)
-Added back requirement to launch elevated
-Removed VMToolsXLS matrix code

.DESCRIPTION
Author Owen Reynolds
[https://getvpro.com](https://getvpro.com)

.EXAMPLE
./Get-VMware-QuickInventory.ps1

.NOTES

.Link
N/A

#>

#$Cred = Get-Credential

### Region VARIABLES

$ShortDate = (Get-Date).ToString('MM-dd-yyyy')
$LogTimeStamp = (Get-Date).ToString('MM-dd-yyyy-hhmm-tt')
$ScriptStart = Get-Date

If ($psISE) {

    $CurrentDir = Split-path $psISE.CurrentFile.FullPath
}

Else {

    $CurrentDir = split-path -parent $MyInvocation.MyCommand.Definition

}

If (-not(test-path "$CurrentDir\Reports")) {

    New-item -Path "$CurrentDir\Reports" -ItemType Directory

}

### HTML CSS formatting from [https://adamtheautomator.com/powershell-convertto-html](https://adamtheautomator.com/powershell-convertto-html)
### Colors from [https://www.canva.com/colors/color-wheel/](https://www.canva.com/colors/color-wheel/)

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
        
    .RunningStatus {
    color: #008000;
    }


    .REDStatus {
    color: #ff0000;
    }

    #CreationDate {

        font-family: Arial, Helvetica, sans-serif;
        color: #53ac6b;
        font-size: 12px;
    }

</style>
"@

### header
$ReportTitle = "<h1>VMware quick environmental report for $($env:USERDNSDOMAIN)"
$ReportTitle += "<h2>Data is from $($LogTimeStamp)</h2>"
$ReportTitle += "<hr></hr>"

### REGION FUNCTIONS

function Get-ESXiReady {  
   <#  
   [http://kunaludapi.blogspot.com/2015/01/powercli-cpu-ready-and-usage-from.html](http://kunaludapi.blogspot.com/2015/01/powercli-cpu-ready-and-usage-from.html)
   #>  
   [CmdletBinding()]  
   param(  
   [Parameter(Mandatory=$true,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$true)]  
   [String]$Name) #param   
     $Stattypes = "cpu.usage.average", "cpu.usagemhz.average", "cpu.ready.summation"  
     foreach ($esxi in $(Get-VMHost $Name)) {  
       $vmlist = $AllVMs | Where-Object {$_.VMHost.Name -eq $esxi.Name -and $_.PowerState -eq "PoweredOn"}
       $esxiCPUSockets = $esxi.ExtensionData.Summary.Hardware.NumCpuPkgs   
       $esxiCPUcores = $esxi.ExtensionData.Summary.Hardware.NumCpuCores/$esxiCPUSockets  
       $TotalesxiCPUs = $esxiCPUSockets * $esxiCPUcores  
       foreach ($vm in $vmlist) {  
         $VMCPUNumCpu = $vm.NumCpu  
         $VMCPUCores = $vm.ExtensionData.config.hardware.NumCoresPerSocket  
         $VMCPUSockets = $VMCPUNumCpu / $VMCPUCores  
         $GroupedRealTimestats = Get-Stat -Entity $vm -Stat $Stattypes -Realtime -Instance "" -ErrorAction SilentlyContinue | Group-Object MetricId  
         $RealTimeCPUAverageStat = "{0:N2}" -f $($GroupedRealTimestats | Where-object {$_.Name -eq "cpu.usage.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $RealTimeCPUUsageMhzStat = "{0:N2}" -f $($GroupedRealTimestats | Where-object {$_.Name -eq "cpu.usagemhz.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $RealTimeReadystat = $GroupedRealTimestats | Where-object {$_.Name -eq "cpu.ready.summation"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average  
         $RealTimereadyvalue = [math]::Round($(($RealTimeReadystat / (20 * 1000)) * 100), 2)  
         $Groupeddaystats = Get-Stat -Entity $vm -Stat $Stattypes -Start (get-date).AddDays(-1) -Finish (get-date) -IntervalMins 5 -Instance "" -ErrorAction SilentlyContinue | Group-Object MetricId  
         $dayCPUAverageStat = "{0:N2}" -f $($Groupeddaystats | Where-object {$_.Name -eq "cpu.usage.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $dayCPUUsageMhzStat = "{0:N2}" -f $($Groupeddaystats | Where-object {$_.Name -eq "cpu.usagemhz.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $dayReadystat = $Groupeddaystats | Where-object {$_.Name -eq "cpu.ready.summation"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average  
         $dayreadyvalue = [math]::Round($(($dayReadystat / (300 * 1000)) * 100), 2)  
         $Groupedweekstats = Get-Stat -Entity $vm -Stat $Stattypes -Start (get-date).AddDays(-7) -Finish (get-date) -IntervalMins 30 -Instance "" -ErrorAction SilentlyContinue | Group-Object MetricId  
         $weekCPUAverageStat = "{0:N2}" -f $($Groupedweekstats | Where-object {$_.Name -eq "cpu.usage.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $weekCPUUsageMhzStat = "{0:N2}" -f $($Groupedweekstats | Where-object {$_.Name -eq "cpu.usagemhz.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $weekReadystat = $Groupedweekstats | Where-object {$_.Name -eq "cpu.ready.summation"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average  
         $weekreadyvalue = [math]::Round($(($weekReadystat / (1800 * 1000)) * 100), 2)  
         $Groupedmonthstats = Get-Stat -Entity $vm -Stat $Stattypes -Start (get-date).AddDays(-30) -Finish (get-date) -IntervalMins 120 -Instance "" -ErrorAction SilentlyContinue | Group-Object MetricId  
         $monthCPUAverageStat = "{0:N2}" -f $($Groupedmonthstats | Where-object {$_.Name -eq "cpu.usage.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $monthCPUUsageMhzStat = "{0:N2}" -f $($Groupedmonthstats | Where-object {$_.Name -eq "cpu.usagemhz.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $monthReadystat = $Groupedmonthstats | Where-object {$_.Name -eq "cpu.ready.summation"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average  
         $monthreadyvalue = [math]::Round($(($monthReadystat / (7200 * 1000)) * 100), 2)        
         $Groupedyearstats = Get-Stat -Entity $vm -Stat $Stattypes -Start (get-date).AddDays(-365) -Finish (get-date) -IntervalMins 1440 -Instance "" -ErrorAction SilentlyContinue | Group-Object MetricId  
         $yearCPUAverageStat = "{0:N2}" -f $($Groupedyearstats | Where-object {$_.Name -eq "cpu.usage.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $yearCPUUsageMhzStat = "{0:N2}" -f $($Groupedyearstats | Where-object {$_.Name -eq "cpu.usagemhz.average"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average)  
         $yearReadystat = $Groupedyearstats | Where-object {$_.Name -eq "cpu.ready.summation"} | Select-Object -ExpandProperty Group | Measure-Object -Average Value | Select-Object -ExpandProperty Average  
         $yearreadyvalue = [math]::Round($(($yearReadystat / (86400 * 1000)) * 100), 2)    
         $data = New-Object psobject  
         $data | Add-Member -MemberType NoteProperty -Name VM -Value $vm.name  
         $data | Add-Member -MemberType NoteProperty -Name VMTotalCPUs -Value $VMCPUNumCpu   
         $data | Add-Member -MemberType NoteProperty -Name VMTotalCPUSockets -Value $VMCPUSockets  
         $data | Add-Member -MemberType NoteProperty -Name VMTotalCPUCores -Value $VMCPUCores  
         $data | Add-Member -MemberType NoteProperty -Name "RealTime Usage Average%" -Value $RealTimeCPUAverageStat  
         $data | Add-Member -MemberType NoteProperty -Name "RealTime Usage Mhz" -Value $RealTimeCPUUsageMhzStat  
         $data | Add-Member -MemberType NoteProperty -Name "RealTime Ready%" -Value $RealTimereadyvalue  
         $data | Add-Member -MemberType NoteProperty -Name "Day Usage Average%" -Value $dayCPUAverageStat  
         $data | Add-Member -MemberType NoteProperty -Name "Day Usage Mhz" -Value $dayCPUUsageMhzStat  
         $data | Add-Member -MemberType NoteProperty -Name "Day Ready%" -Value $dayreadyvalue  
         $data | Add-Member -MemberType NoteProperty -Name "week Usage Average%" -Value $weekCPUAverageStat  
         $data | Add-Member -MemberType NoteProperty -Name "week Usage Mhz" -Value $weekCPUUsageMhzStat  
         $data | Add-Member -MemberType NoteProperty -Name "week Ready%" -Value $weekreadyvalue  
         $data | Add-Member -MemberType NoteProperty -Name "month Usage Average%" -Value $monthCPUAverageStat  
         $data | Add-Member -MemberType NoteProperty -Name "month Usage Mhz" -Value $monthCPUUsageMhzStat  
         $data | Add-Member -MemberType NoteProperty -Name "month Ready%" -Value $monthreadyvalue  
         $data | Add-Member -MemberType NoteProperty -Name "Year Usage Average%" -Value $yearCPUAverageStat  
         $data | Add-Member -MemberType NoteProperty -Name "Year Usage Mhz" -Value $yearCPUUsageMhzStat  
         $data | Add-Member -MemberType NoteProperty -Name "Year Ready%" -Value $yearreadyvalue  
         $data | Add-Member -MemberType NoteProperty -Name VMHost -Value $esxi.name  
         $data | Add-Member -MemberType NoteProperty -Name VMHostCPUSockets -Value $esxiCPUSockets  
         $data | Add-Member -MemberType NoteProperty -Name VMHostCPUCores -Value $esxiCPUCores  
         $data | Add-Member -MemberType NoteProperty -Name TotalVMhostCPUs -Value $TotalesxiCPUs  
         $data  
       } #foreach ($vm in $vmlist)  
     }#foreach ($esxi in $(Get-VMHost $Name))  
 } #Function Get-Ready 
 
function Select-CPUReady {
    param (
        [string]$Title = 'CPU Ready detailed analysis'
    )
    #Clear-Host
    Write-Host "================ $Title ================"    
    Write-Host "`r"
    Write-Host "1: Press 'Y' YES"
    Write-Host "`r"
    Write-Host "2: Press 'N' NO"    
    Write-Host "`r"
    Write-Host "Q: Press 'Q' to quit"
}

function Find-And-LaunchExcel {
    [CmdletBinding()]
    param (
        [string]$XLSReport
    )

    $possiblePaths = @(
        "$env:ProgramFiles\Microsoft Office",
        "$env:ProgramFiles(x86)\Microsoft Office",
        "$env:LocalAppData\Microsoft\WindowsApps",
        "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
        "$env:AppData\Microsoft\Windows\Start Menu\Programs",
        "$env:LocalAppData\Programs",
        "$env:ProgramFiles\Common Files\Microsoft Shared\ClickToRun",
        "$env:ProgramFiles\WindowsApps",
        "$env:ProgramData\App-V"
    )

    $excelLocations = @()

    foreach ($basePath in $possiblePaths) {
        if (Test-Path $basePath) {
            try {
                Get-ChildItem -Path $basePath -Recurse -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -eq "EXCEL.EXE" } |
                ForEach-Object { $excelLocations += $_.FullName }
            } catch {
                Write-Verbose "Skipping inaccessible path: $basePath"
            }
        }
    }

    if ($excelLocations.Count -gt 0) {
        Write-Host "Excel found at the following location(s):" -ForegroundColor Green
        $excelLocations | ForEach-Object { Write-Host $_ }

        $excelExe = $excelLocations[0]

        if ($XLSReport -and (Test-Path $XLSReport)) {
            Write-Host "Launching Excel with file $XLSReport" -ForegroundColor Cyan
            Start-Process -FilePath $excelExe -ArgumentList "`"$XLSReport`""
        }
        else {
            Write-Host "Launching Excel..." -ForegroundColor Cyan
            Start-Process -FilePath $excelExe
        }
    } else {
        Write-Warning "Excel is not installed or could not be found in standard locations."
    }
}

function Convert-GuestOSId {
    param (
        [string]$GuestId
    )
    
    $osMap = @{
        'windows9Server64Guest' = 'Windows Server 2016 (64-bit)'
        'windows2019srv_64Guest' = 'Windows Server 2019 (64-bit)'
        'windows2019srvNext_64Guest' = 'Windows Server 2022 (64-bit)'
        'windows11_64Guest' = 'Windows 11 (64-bit)'
        'windows10_64Guest' = 'Windows 10 (64-bit)'
        'windows9_64Guest' = 'Windows 10 (64-bit)'
        'windows8Server64Guest' = 'Windows Server 2012 R2 (64-bit)'
        'windows8_64Guest' = 'Windows 8 (64-bit)'
        'windows7Server64Guest' = 'Windows Server 2008 R2 (64-bit)'
        'windows7_64Guest' = 'Windows 7 (64-bit)'
        'centos8_64Guest' = 'CentOS 8 (64-bit)'
        'centos7_64Guest' = 'CentOS 7 (64-bit)'
        'rhel8_64Guest' = 'Red Hat Enterprise Linux 8 (64-bit)'
        'rhel7_64Guest' = 'Red Hat Enterprise Linux 7 (64-bit)'
        'ubuntu64Guest' = 'Ubuntu Linux (64-bit)'
        'debian10_64Guest' = 'Debian 10 (64-bit)'
        'debian11_64Guest' = 'Debian 11 (64-bit)'
        'otherLinux64Guest' = 'Other Linux (64-bit)'
        'otherGuest64' = 'Other OS (64-bit)'
    }
    
    if ($osMap.ContainsKey($GuestId)) {
        return $osMap[$GuestId]
    } else {
        return $GuestId
    }
}

### END REGION FUNCTIONS

### Install Nuget and VMware PowerCLI as required

IF (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    write-warning "Please open Powershell as administrator, the script will now exit"
    EXIT

}

### REGION MODULES

if (-not(($PSversionTable.PSVersion).Major -ge 5)) {

    write-warning "Powershell version 5 or above is required to run this script"
    write-warning "Please download/install from here https://www.microsoft.com/en-us/download/details.aspx?id=54616"
    write-warning "The script will now exit"
    EXIT

}

IF (-not(Get-PackageProvider -ListAvailable -name NUget)) {
    
    Write-host "Installing Nuget package provider" -foregroundcolor cyan

    Install-PackageProvider -Name NuGet -force -Confirm:$False
}

IF (-not(Get-Module -ListAvailable -name VMware.PowerCLI)) {

    Write-host "Installing VMware PowerCLI" -foregroundcolor cyan

    Install-Module -Name VMware.PowerCLI -AllowClobber -force
}

IF (-not(Get-Module -ListAvailable -name VMware.PowerCLI)) {

    write-warning "PowerCLI failed to install. The script will exit"
    EXIT
}

IF (-not(Get-Module -ListAvailable -name ImportExcel)) {

    Write-host "Installing ImportExcel for use with .XLS" -foregroundcolor cyan

    Install-Module -Name ImportExcel -AllowClobber -force
}

IF (-not(Get-Module -ListAvailable -name ImportExcel)) {

    write-warning "The ImportExcel module failed to install, the script will exit"
    EXIT
}

write-host "Start of script processing" -ForegroundColor Green

import-module ImportExcel

### REGION START OF SCRIPT PROCESSING

Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $False -Confirm:$false
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

write-host "r`n"

write-host "Connecting to vCenter" -foregroundColor cyan

IF ($global:DefaultVIServer.Length -eq 0) {

    $VC = read-host -Prompt "Enter the vCenter name without the leading https://"
    $VcenterCred = get-credential
    Connect-VIServer -Server $VC -Credential $VcenterCred

}

IF ($global:DefaultVIServer.Length -eq 0) {

    write-warning "$VC vCenter is not connected. The script will exit, please re-run"
    EXIT

}

### CONSOLIDATED VM DATA COLLECTION - Single Get-VM call for performance optimization
write-host "Collecting all VM information (single pass - optimized)" -ForegroundColor Cyan
$AllVMs = Get-VM | Where-Object {$_.Name -notlike "vCLS-*"}

### NEW - All VMs detailed summary for XLS export
write-host "Collecting detailed VM information for All VMs tab" -ForegroundColor Cyan

$AllVMsDetails = $AllVMs | ForEach-Object {
    $vm = $_
    $nicType = ($vm | Get-NetworkAdapter | Select-Object -First 1).Type
    $toolsVersion = $vm.ExtensionData.Guest.ToolsVersion
    $toolsStatus = $vm.ExtensionData.Guest.ToolsVersionStatus
    $toolsOutOfDate = if ($toolsStatus -ne "guestToolsCurrent" -and $toolsStatus -ne "guestToolsUnmanaged") { "Out of Date" } else { "Current" }
    $guestOS = Convert-GuestOSId -GuestId $vm.Guest.GuestId
    
    [PSCustomObject]@{
        'VM Name' = $vm.Name
        'Datastore' = ($vm | Get-Datastore | Select-Object -First 1).Name
        'PowerState' = $vm.PowerState
        'Used Space (GB)' = [Math]::Round($vm.UsedSpaceGB, 2)
        'NumCpu' = $vm.NumCpu
        'MemoryGB' = $vm.MemoryGB
        'NIC Type' = $nicType
        'VMware Tools Version' = $toolsVersion
        'Tools Status' = $toolsOutOfDate
        'Guest OS' = $guestOS
    }
}

### 1 
write-host "Collecting VM network card type" -ForegroundColor Cyan
$IntelNICs = $AllVMs | Get-NetworkAdapter | Where-Object {$_.Type -ne 'VMxnet3'} | Select-Object @{E={$_.Parent};Name="VM"}, @{E={$_.Type};Name="Network card type"}

IF ($IntelNICs.Length -eq 0) {

    write-host "All network card types are set correctly to VMware VXnet 3 type" -ForegroundColor Green
    $Pre1 = "<H2>PASS: All Network card type attached to shells are set correctly</H2>"
    #$Pre1 += "<br><br>"
    $IntelNICHTML = $Pre1
    
}

Else {

    Write-Warning "$($IntelNICs | Measure-Object | Select-Object -ExpandProperty Count) VM(s) were found with legacy intel network card types"
    $Pre1 = "<H2>WARNING: Intel E1000 legacy network card type VMs detected</H2>"
    #$Pre1 += "<br><br>"
    $IntelNICHTML = $IntelNICs | ConvertTo-HTML -Head $Head -PreContent $Pre1 -As Table | Out-String

}

### 2 
write-host "Collecting ESXi power profile" -ForegroundColor Cyan
### [https://www.cloudishes.com/2015/09/automating-esxi-host-power-management.html](https://www.cloudishes.com/2015/09/automating-esxi-host-power-management.html)
$HostPowerPolicy = Get-VMHost | Sort-Object | Select-Object -property Name,
@{ N="CurrentPolicy"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.ShortName}},
@{ N="CurrentPolicyKey"; E={$_.ExtensionData.config.PowerSystemInfo.CurrentPolicy.Key}},
@{ N="AvailablePolicies"; E={$_.ExtensionData.config.PowerSystemCapability.AvailablePolicy.ShortName}} `
| Where-Object {$_.CurrentPolicy -ne "Static"} | Select-Object -Property Name, CurrentPolicy

IF ($HostPowerPolicy.Length -eq 0) {
    
    write-host "All ESXi hosts are set correctly to HIGH PERFORMANCE" -ForegroundColor Green
    $Pre2 = "<H2>PASS: All ESXi hosts are set correctly to HIGH PERFORMANCE</H2>"
    #$Pre2 += "<br><br>"
    $ESXIPowerProfile = $Pre2
}

else {

    write-warning "There are ESXi hosts NOT set to HIGH PERFORMANCE"
    $Pre2 = "<H2>WARNING: The below ESXi hosts that are NOT set to HIGH PERFORMANCE</H2>"
    #$Pre2 += "<br><br>"
    $ESXIPowerProfile = $HostPowerPolicy | ConvertTo-HTML -Head $Head -PreContent $Pre2 -As Table | Out-String

}

### 3 - SCSI controller types: [https://www.sqlpassion.at/archive/2019/06/10/benchmarking-the-vmware-lsi-logic-sas-against-the-pvscsi-controller/](https://www.sqlpassion.at/archive/2019/06/10/benchmarking-the-vmware-lsi-logic-sas-against-the-pvscsi-controller/)
$SCSIControllerTypes = $AllVMs | Select-Object Name,@{N="Cluster";E={Get-Cluster -VM $_}},@{N="Controller Type";E={Get-ScsiController -VM $_ | Select-Object -ExpandProperty Type}} `
| Where-Object {$_."Controller Type" -eq "VirtualLsiLogicSAS "} | Select-Object @{E={$_.Name};Name="VM"}, "Controller Type"


IF ($SCSIControllerTypes.Length -eq 0) {
    
    write-host "All SCSI controllers are set correctly to VMWware Paravirtual" -ForegroundColor Green
    $Pre3 = "<H2>PASS: All SCSI controllers are set correctly to VMWware Paravirtual</H2>"
    #$Pre3 += "<br><br>"
    $SCSIControllerTypesHTML = $Pre3
}

else {

    write-warning "There are VMs are using legacy LSI SCSI controller types"
    $Pre3 = "<H2>WARNING: The below VMs are using legacy LSI SCSI controller types, use VMware Paravirtual where possible</H2>"
    #$Pre3 += "<br><br>"
    $SCSIControllerTypesHTML = $SCSIControllerTypes | ConvertTo-HTML -Head $Head -PreContent $Pre3 -As Table | Out-String

}

### 4 - VMware tools version
$VMWareTools = $AllVMs | Select-Object Name,@{Name="ToolsVersion";Expression={$_.ExtensionData.Guest.ToolsVersion}},@{Name="ToolsStatus";Expression={$_.ExtensionData.Guest.ToolsVersionStatus}} `
| Where-Object {$_.ToolsStatus -ne "guestToolsCurrent"} | Where-Object {$_.ToolsStatus -ne "guestToolsUnmanaged"}

If ($VMWareTools.length -eq 0) {

    write-host "VMware tools is up to date on all VMs" -ForegroundColor Green
    $Pre4 = "<H2>PASS: VMware tools is up to date on all VMs</H2>"
    #$Pre4 += "<br><br>"
    $VMToolsHTML = $Pre4
}

else {

    write-warning "There are VMs running older versions of VMware tools and should be updated"
    $Pre4 = "<H2>WARNING: The below VMs are running older versions of VMware tools and should be updated</H2>"
    #$Pre4 += "<br><br>"
    $VMToolsHTML = $VMWareTools | ConvertTo-HTML -Head $Head -PreContent $Pre4 -As Table | Out-String

}

### 5 - Get ESXi host BIOS version

write-host "Collecting ESXi info"

$ESXihosts = Get-VMhost | Where-Object {$_.model -ne "VMware Virtual Platform"} | Select-Object Name, NumCpu

$ESXiSummary = Get-VMHost | Where-Object { $_.Model -ne "VMware Virtual Platform" } | ForEach-Object {
    $vmhost     = $_
    $cluster    = $vmhost.Parent
    $view       = $vmhost | Get-View
    $numSockets = $vmhost.ExtensionData.Summary.Hardware.NumCpuPkgs
    $dns        = ($vmhost | Get-VMHostNetwork).DNSAddress -join ", "
    $drs        = Get-Cluster -Name $cluster.Name | Select-Object -First 1 -Property Name, DrsEnabled, DrsAutomationLevel

    # Gather connected physical NICs with speed, duplex, and model
    $nicDetails = @()
    foreach ($pnic in $view.Config.Network.Pnic) {
        if ($pnic.LinkSpeed -and $pnic.LinkSpeed.SpeedMb -gt 0) {
            $duplex = if ($pnic.LinkSpeed.Duplex) { "Full" } else { "Half" }
            $pciInfo = $view.Hardware.PciDevice | Where-Object { $_.Id -eq $pnic.Pci }
            $nicDetails += "$($pnic.Device): $($pnic.LinkSpeed.SpeedMb)Mb, $duplex Duplex, Model: $($pciInfo.DeviceName)"
        }
    }

    [PSCustomObject]@{
        Cluster         = $cluster.Name
        DRSEnabled      = $drs.DrsEnabled
        DRSLevel        = $drs.DrsAutomationLevel
        Name            = $vmhost.Name
        ConnectionState = $vmhost.ConnectionState
        PowerState      = $vmhost.PowerState
        Model           = $vmhost.Model
        NumCPUCore      = $vmhost.NumCPU
        NumCPUSocket    = $numSockets
        CPUType         = $vmhost.ProcessorType
        Version         = $vmhost.Version
        Build           = $vmhost.Build
        MemGB           = [math]::Round($vmhost.MemoryTotalGB, 2)
        MemGBUsed       = [math]::Round($vmhost.MemoryUsageGB, 2)
        MaxEVCMode      = $vmhost.MaxEVCMode
        BIOSVersion     = $view.Hardware.BiosInfo.BiosVersion
        BIOSDate        = $view.Hardware.BiosInfo.ReleaseDate
        DNSServers      = $dns
        PhysicalNICs    = $nicDetails -join "; "
    }
}

If ($ESXiSummary.length -eq 0) {

    $Pre5 = "<H2>WARNING: ESXi host hardware info is not available at this time</H2>"
    #$Pre5 += "<br><br>"
    $ESXISummaryHTML = $Pre5
}

Else {

    $Pre5 = "<H2>INFO: ESXi host summary</H2>"

    $ESXiSummary = $ESXiSummary | Select-Object Cluster, DRSEnabled, DRSLevel, Name, ConnectionState, PowerState, Model, NumCPUCore, NumCPUSocket, CPUType, Version, Build, MemGB, MemGBUsed, MaxEVCMode, BIOSVersion, BIOSDate, DNSServers, PhysicalNICs
        
    $ESXISummaryHTML = $ESXiSummary | ConvertTo-HTML -Head $Head -PreContent $Pre5 -As Table | Out-String

}

### 6 - ESXi host NTP settings

write-host "Collecting NTP service config"

$NTP = Get-VMHost | Sort-Object Name | Select-Object Name, @{N="NTPServiceRunning";E={($_ | Get-VmHostService | Where-Object {$_.key -eq "ntpd"}).Running}},`
@{N="StartupPolicy";E={($_ | Get-VmHostService | Where-Object {$_.key -eq "ntpd"}).Policy}}, @{N="NTPServers";E={$_ | Get-VMHostNtpServer}}, @{N="Date&Time";E={(Get-View $_.ExtensionData.ConfigManager.DateTimeSystem).QueryDateTime().ToLocalTime()}}

$NTP | Where-Object {$_.NTPServers -notlike "*.ntp.org"} | ForEach-Object {$_ | Add-Member -MemberType NoteProperty -name "NTPServers" -value "Not set to pool.ntp.org" -Force}

IF ($NTP.length -eq 0) {

    $Pre6 = "<H2>WARNING: ESXi NTP information is not available at this time/H2>"
    #$Pre6 += "<br><br>"
    $NTPHTML = $Pre6

}

Else {

    $Pre6 = "<H2>INFO: ESXi NTP settings</H2>"
    $NTPHTML = $NTP | ConvertTo-HTML -Head $Head -PreContent $Pre6 -As Table | Out-String
    $NTPHTML = $NTPHTML -replace '<td>False</td>', '<td class="REDStatus">Stopped</td>'
    $NTPHTML = $NTPHTML -replace '<td>Off</td>', '<td class="REDStatus">off</td>'
    $NTPHTML = $NTPHTML -replace '<td>Not set to pool.ntp.org</td>', '<td class="REDStatus">Not set to ntp.org, please correct</td>'
}


### 7 - VM Hardware version

write-host "Collecting VM shell hardware version"

$VMHardwareVersion = $AllVMs | Select-Object Name, HardwareVersion | Sort-Object HardwareVersion

If ($VMHardwareVersion.Length -eq 0) {

    $Pre7 = "<H2>WARNING: VM Shell hardware version info is not available at this time</H2>"
    #$Pre7 += "<br><br>"
    $VMHardwareVersionHTML = $Pre7

}

Else {

    $Pre7 = "<H2>INFO: VM Shell hardware version summary</H2>"
    #$Pre7 += "<br><br>"
    $VMHardwareVersionHTML = $VMHardwareVersion | ConvertTo-HTML -Head $Head -PreContent $Pre7 -As Table | Out-String

}

### 8 - vCPU to pCPU ratio
### Should not be above 5 for Citrix environments

write-host "Collecting vCPU / pCPU ratio"

$RatioSummary = @()

ForEach ($i in $ESXihosts) {

    write-host "Collecting vCPU to Physical CPU ratio info from $($i.name)"

    $Ratio = ($AllVMs | Where-Object {$_.VMHost.Name -eq $i.Name -and $_.Name -notlike "vcls*"} | Select-Object -ExpandProperty NumCPU | Measure-Object -Sum | Select-Object -ExpandProperty Sum) / $i.NumCpu

    if ($Ratio -ge 5) {

        $Status = "WARNING"

    }

    Else {

        $Status = "vCPU to pCPU ratio is within acceptable limits"
    }


    
    $RatioSummary += New-Object -TypeName PSObject -Property @{

    ESXihost = $i.Name
    Ratio = $Ratio
    Status = $Status

    }

}

If ($RatioSummary.Length -eq 0) {

    $Pre8 = "<H2>WARNING: ESXi vCPU to pCPU info is not available at this time</H2>"    
    $RatioSummaryHTML = $Pre8

}

Else {

    $RatioSummary = $RatioSummary | Select-Object ESXiHost, Status, Ratio
    $Pre8 = "<H2>INFO: ESXi vCPU to pCPU ratio summary</H2>"    
    $RatioSummaryHTML = $RatioSummary | ConvertTo-HTML -Head $Head -PreContent $Pre8 -As Table | Out-String
    $RatioSummaryHTML = $RatioSummaryHTML -replace '<td>WARNING</td>', '<td class="REDStatus">vCPU to pCPU ratio values above 5 can be problematic for VDI workloads</td>'   

}

### 9 CPU Ready time

$VMCount = ($AllVMs | Measure-Object).Count
$EstimatedTimeSeconds = [Math]::Round($VMCount * 1.25, 2)

# Convert seconds to TimeSpan
$timeSpan = [TimeSpan]::FromSeconds($EstimatedTimeSeconds)

Write-Host "Total VMs: $VMCount"

$EstimatedTime = "{0} hours, {1} minutes, {2} seconds" -f $timeSpan.Hours, $timeSpan.Minutes, $timeSpan.Seconds

do {
    Select-CPUReady
    
    Write-Host "`r"
    
    $input = Read-Host "Do you want to collect detailed CPU Ready stats from all VMs in the environment (Y/N) ? Based on a VM count of $VMCount, it should take $EstimatedTime" 
    
    switch ($input) {
        'Y' {
            
            $CPUReadyChoice = "Yes"
        }

        'N' {
            
            $CPUReadyChoice = "No"

        }       

        'q' {
            Write-Warning "Script will now exit"
            Exit-Script
        }
    }

    "You chose $CPUReadyChoice"
    Write-Host "`r"
    #Pause
}

until ($input -ne $null)


IF ($CPUReadyChoice -eq "Yes") {

    IF ($CPUReadySummary) {

        Remove-Variable CPUReadySummary

    }

    Foreach ($ESXiHost in $ESXiHosts) {
    
        write-host "Collecting CPU ready time from ESXi host $($ESXiHost.Name)" -ForegroundColor Cyan    
        $CPUReadySummary += Get-ESXiReady -Name $ESXiHost.Name

    }
 
    $CPUReadySummary = $CPUReadySummary | Sort-Object -Property "RealTime Ready%" -Descending

}

Else {

    write-host "CPU ready stats will not be collected" -ForegroundColor Cyan

}

### 10 - Datastores

$DataStores = Get-DataStore | Select-Object Name, State, @{E={[Math]::Round($_.CapacityGB,2)};Label="Capacity (GB"}, @{E={[Math]::Round($_.FreeSpaceGB,2)};Label="Free Space (GB)"},`
 @{E={$_.Type};Name='File System Type'}, FileSystemVersion | Sort-Object Name

IF ($DataStores.Length -eq 0) {

    $Pre10 = "<H2>WARNING: Datastore info is not available at this time</H2>"    
    $DataStoresHTML = $Pre10

}

Else {    

    $Pre10 = "<H2>INFO: Datastore summary</H2>"    
    $DataStoresHTML = $DataStores | ConvertTo-HTML -Head $Head -PreContent $Pre10 -As Table | Out-String
}

### 

$ScriptEnd = Get-Date

$TotalScriptTime = $ScriptEnd - $ScriptStart | Select-Object Hours, Minutes, Seconds

$Hours = $TotalScriptTime | Select-Object -expand Hours
$Mins = $TotalScriptTime | Select-Object -expand Minutes
$Seconds = $TotalScriptTime | Select-Object -expand Seconds

# Reset the variable to avoid appending to old content
$PostContent = ""

$PostContent += "<hr></hr>"
$PostContent += "<b><p id='CreationDate'>Creation Date: $(Get-Date)"
$PostContent += "<br>"
$PostContent += "Generated by $($Env:UserName)"
$PostContent += "<br>"
$PostContent += "Total processing time: $Hours hours, $Mins minutes, $Seconds seconds</p></b>"
 
### REMOVED CPU Ready from HTML output - kept in XLS only
$HTMLReport = ""
$HTMLReport = ConvertTo-HTML -Body "$ReportTitle $ESXISummaryHTML $ESXIPowerProfile $RatioSummaryHTML $DataStoresHTML $NTPHTML $IntelNICHTML $SCSIControllerTypesHTML $VMToolsHTML $VMHardwareVersionHTML " -Title "VMware Quick Environmental report" -PostContent $PostContent

$HTMLReport | out-file "$CurrentDir\Reports\vSphere_Inventory-$LogTimeStamp.html"

Invoke-Item "$CurrentDir\Reports\vSphere_Inventory-$LogTimeStamp.html"

write-host "XLS report creation and opening, where EXCEL is installed" -ForegroundColor Cyan

$XLSReport = "$CurrentDir\Reports\vSphere_Inventory_$LogTimeStamp.xlsx"

IF (test-path $XLSReport) {

    Remove-item $XLSReport

}

### Export All VMs tab first

if ($ESXiSummary.Length -ne 0)  {
    $ESXiSummary | Export-Excel -Path $XLSReport -WorksheetName "ESXi Summary" -AutoSize -AutoFilter
}
if ($RatioSummary.Length -ne 0) {
    $RatioSummary | Export-Excel -Path $XLSReport -WorksheetName "vCPU to pCPU Ratio" -AutoSize -AutoFilter
}
if ($DataStores.Length -ne 0) {
    $DataStores | Export-Excel -Path $XLSReport -WorksheetName "Datastores" -AutoSize -AutoFilter
}
if ($HostPowerPolicy.Length -ne 0) {
    $HostPowerPolicy | Export-Excel -Path $XLSReport -WorksheetName "Power Profiles" -AutoSize -AutoFilter
}

if ($AllVMsDetails.Length -ne 0) {
    $AllVMsDetails | Export-Excel -Path $XLSReport -WorksheetName "All VMs" -AutoSize -AutoFilter
}

if ($IntelNICs.Length -ne 0) {
    $IntelNICs | Export-Excel -Path $XLSReport -WorksheetName "Intel E1000 legacy NIC on VMS" -AutoSize -AutoFilter
}
if ($SCSIControllerTypes.Length -ne 0) {
    $SCSIControllerTypes | Export-Excel -Path $XLSReport -WorksheetName "LSI SCSI HBA on VMS" -AutoSize -AutoFilter
}
if ($VMWareTools.Length -ne 0) {
    $VMWareTools | Export-Excel -Path $XLSReport -WorksheetName "VMware Tools out-of-date" -AutoSize -AutoFilter
}
if ($NTP.Length -ne 0) {
    $NTP | Export-Excel -Path $XLSReport -WorksheetName "ESXI NTP Settings" -AutoSize -AutoFilter
}
if ($VMHardwareVersion.Length -ne 0) {
    $VMHardwareVersion | Export-Excel -Path $XLSReport -WorksheetName "VM Hardware versions" -AutoSize -AutoFilter
}
if ($CPUReadySummary.Length -ne 0) {
    $CPUReadySummary | Export-Excel -Path $XLSReport -WorksheetName "CPU Ready Time on all VMs" -AutoSize -AutoFilter
}

Find-And-LaunchExcel -XLSReport $XLSReport

write-host "Disconnecting from $($global:DefaultVIServer.Name)" -ForegroundColor Cyan

Disconnect-VIServer -Force -Confirm:$False
$ScriptEnd = Get-Date

write-host "Script is done!" -ForegroundColor Cyan
