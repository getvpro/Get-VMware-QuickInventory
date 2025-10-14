## PS module pre-reqs
1. The script should be run from a recently updated version of windows
1. The script only requires access to vCenter, it does not install components on any vSphere infra: vCenter / ESXi
1. When run, the script will attempt to install the required binaries: Nuget, Import-Excel and VMware PowerCLI
1. If it's preferred, you can install them manually via elevated powershell / PowersShell_ISE window as follows:

```
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

```

## Usage

_NOTE: HTML/XLS Reports genreated are local, nothing is uploaded off client devices_

1. Download the .PS1 to a folder and run the .PS1 as admin

1. You will be prompted to enter in a vCenter address (enter it without the leading https://)

1. You will then receive a second [y]/[n] prompt for collection of additional historical performance data for VMs on each ESXi host

![image](https://github.com/getvpro/Get-VMware-QuickInventory/assets/50507806/1174173b-243c-40f5-be37-f85c43451934)


## Description

This simple script is used to review/report on common issues found on VMware vSphere environments that can cause performance/security issues. 

The script is used only for reporting purposes, a set of recommended actions are shown below

Outputs to a time-stamped HTML report that will open towards the end of the script via the default web browser on the system it's run from

Package provider NUGET will be installed to facilitate installing the VMware PowerCLI module from PS Windows gallery 

The following elements are reported on:

1. Inventory of vSphere clusters capturing ESXI name, version, hardware model, CPU, memory and physical network cards

1. Check that the NTP service is set to ca.ntp.org or pool.ntp.org and to set to "start/stop with host"

1. Check that the power management settings for the ESXi host is set to HIGH PERFORMANCE , else the max performance the ESXi server CPU will run at, will be 80%

1. Scans all VMs to ID which ones are still using Intel E1000 NIC types, which are less performant than the VMware VMNet 3 type NICs. The suggestion would be to change the NIC from E1000 to VMnet 3. Note: This will require re-adding static IP info, if a MAC address was used for firewall rules or DHCP reservation, that will need to be reset as well

1. Scan all VMs to ID which ones are using LSI logic SCSI adapters, which are less performant than the VMware Paravirtual SCSI controller types. Where possible change the type to VMware paravirtual, you will first need to add the VMware paravirtual controller on the related shell, have the OS detect it, then power off the VM shell and amend the SCSI settings to point to the new VMware paravirtual controller: reference page

1. Scan all VMs to ID which are are running older versions of VMware tools. The suggestion would be to update the ESXi host VMware tools repo with the latest version, then set VMs to upgrade to the latest version on reboot. Ensure this is tested first and done in a maintenance window

1. Scan all VMs to ID which are on older VM hardware types on their shell. The suggestion would be to schedule upgrades/reboots of the shells in a maintenance window

1. Scan all VMs to ID which ESXi hosts are running older BIOS (EFI) levels. The suggestion would be to update where required in a maintenance window

1. Checks vCPU to Physical CPU ratio. It's suggested not to exceed a ratio of 5 for Citrix environments , but other workloads can run fine at higher ratios
   
 ## Output
1. an HTML report will be created
2. A single XLS workbook will be created, with multiple tabs will be created for the relevant elements listed in the previous section
 
