<#
    .NOTES
	===========================================================================
	Created by:		Russell Hamker
	Date:			July 30, 2021
	Version:		1.0
	Twitter:		@butch7903
	GitHub:			https://github.com/butch7903
	===========================================================================

	.SYNOPSIS
		This script will generate a UCS Inventory in xlsx xml format.

	.DESCRIPTION
		Use this script to create a xlsx UCS Inventory.
		
	.NOTES
		This script requires a Cisco.UCSManager version 3.0.1.2 or greater and
		ImportExcel version 7.2.1 or greater. 
		
	.TROUBLESHOOTING
		
#>

##Check if Modules are installed, if so load them, else install them
if (Get-InstalledModule -Name Cisco.UCSManager -MinimumVersion 3.0.1.2) {
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	Write-Host "PowerShell Module Cisco.UCSManager required minimum version was found previously installed"
	Write-Host "Importing PowerShell Module Cisco.UCSManager"
	Import-Module -Name Cisco.UCSManager
	Write-Host "Importing PowerShell Module Cisco.UCSManager Completed"
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	#CLEAR
} else {
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	Write-Host "PowerShell Module Cisco.UCSManager does not exist"
	Write-Host "Setting TLS Security Protocol to TLS1.2, this is needed for Proxy Access"
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
	Write-Host "Setting Micrsoft PowerShell Gallery as a Trusted Repository"
	Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
	Write-Host "Verifying that NuGet is at minimum version 2.8.5.201 to proceed with update"
	Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
	Write-Host "Uninstalling any older versions of the PowerShellGet Module"
	Get-Module PowerShellGet | Uninstall-Module -Force
	Write-Host "Installing PowerShellGet Module"
	Install-Module -Name PowerShellGet -Scope AllUsers -Force
	Write-Host "Setting Execution Policy to RemoteSigned"
	Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force
	Write-Host "Uninstalling any older versions of the Cisco.UCSManager Module"
	Get-Module Cisco.UCSManager | Uninstall-Module -Force
	Write-Host "Installing Newest version of Cisco.UCSManager PowerShell Module"
	Install-Module -Name Cisco.UCSManager -MinimumVersion 3.0.1.2 -Scope AllUsers -Force
	Write-Host "Importing PowerShell Module Cisco.UCSManager"
	Import-Module -Name Cisco.UCSManager
	Write-Host "PowerShell Module Cisco.UCSManager Loaded"
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	#Clear
}

##Check if Modules are installed, if so load them, else install them
if (Get-InstalledModule -Name ImportExcel -MinimumVersion 7.2.1) {
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	Write-Host "PowerShell Module ImportExcel required minimum version was found previously installed"
	Write-Host "Importing PowerShell Module ImportExcel"
	Import-Module -Name ImportExcel
	Write-Host "Importing PowerShell Module ImportExcel Completed"
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	#CLEAR
} else {
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	Write-Host "PowerShell Module ImportExcel does not exist"
	Write-Host "Setting TLS Security Protocol to TLS1.2, this is needed for Proxy Access"
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
	Write-Host "Setting Micrsoft PowerShell Gallery as a Trusted Repository"
	Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
	Write-Host "Verifying that NuGet is at minimum version 2.8.5.201 to proceed with update"
	Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
	Write-Host "Uninstalling any older versions of the ImportExcel Module"
	Get-Module ImportExcel | Uninstall-Module -Force
	Write-Host "Installing Newest version of ImportExcel PowerShell Module"
	Install-Module -Name ImportExcel -MinimumVersion 7.2.1 -Scope AllUsers -Force
	Write-Host "Importing PowerShell Module ImportExcel"
	Import-Module -Name ImportExcel
	Write-Host "PowerShell Module ImportExcel Loaded"
	Write-Host "-----------------------------------------------------------------------------------------------------------------------"
	#Clear
}

##Document Start Time
$STARTTIME = Get-Date -format "MMM-dd-yyyy HH-mm-ss"
$STARTTIMESW = [Diagnostics.Stopwatch]::StartNew()

##Set Variables
##Get Current Path
$pwd = pwd
$CURRENTDATE = Get-Date

#Set Months to keep data
$MONTHSTOKEEP  = -3 #-3 equals 3 months

#if ucs-domains.csv exists, import it 
$AnswerFile = $pwd.path+"\"+"ucs-domains.csv"
If (Test-Path $AnswerFile){
Write-Host "Answer file found, importing answer file"$AnswerFile
$UCSLIST = Get-Content -Path $AnswerFile | Sort
}Else{
$Answers_List = @()
$Answers="" | Select UCSList
$UCSLIST = Read-Host "Please input the FQDN or IP of your UCSM
Note: If you wish to do multiple UCSMs, edit the CSV to include many
Example 1: hamucs01.hamker.local
Example 2: 
hamucs01.hamker.local
hamucs02.hamker.local
hamucs03.hamker.local
"
$Answers.UCSList = $UCSLIST
$Answers_List += $Answers
$Answers_List | Format-Table -AutoSize
Write-Host "Exporting Information to File"$AnswerFile
$Answers_List | Export-CSV -NoTypeInformation $AnswerFile
}

##Get Date Info for Logging
$LOGDATE = Get-Date -format "MMM-dd-yyyy_HH-mm"
##Specify Log File Info
$LOGFILENAME = "Log_UCS-Inventory_" + $LOGDATE + ".txt"
#Create Log Folder
$LogFolder = $pwd.path+"\Log"
If (Test-Path $LogFolder){
	Write-Host "Log Directory Created. Continuing..."
}Else{
	New-Item $LogFolder -type directory
}
#Specify Log File
$LOGFILE = $pwd.path+"\Log\"+$LOGFILENAME

##Starting Logging
Start-Transcript -path $LOGFILE -Append
Write-Host "-----------------------------------------------------------------------------------------------------------------------"
Write-Host (Get-Date -format "MMM-dd-yyyy_HH-mm-ss")
Write-Host "Script Logging Started"
Write-Host (Get-Date -format "MMM-dd-yyyy_HH-mm-ss")
Write-Host "-----------------------------------------------------------------------------------------------------------------------"

##Create Secure AES Keys for User and Password Management
$KeyFile = $pwd.path+"\"+"UCSAES.key"
If (Test-Path $KeyFile){
	Write-Host "AES File Exists"
	$Key = Get-Content $KeyFile
	Write-Host "Continuing..."
}
Else {
	$Key = New-Object Byte[] 16   # You can use 16, 24, or 32 for AES
	[Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($Key)
	$Key | out-file $KeyFile
}

##Create Secure XML Credential File for UCS
$UCSCreds = $pwd.path+"\"+"UCSCreds.xml"
If (Test-Path $UCSCreds){
	Write-Host "UCSCreds.xml file found"
	Write-Host "Continuing..."
	$ImportObject = Import-Clixml $UCSCreds
	$SecureString = ConvertTo-SecureString -String $ImportObject.Password -Key $Key
	$MyCredential = New-Object System.Management.Automation.PSCredential($ImportObject.UserName, $SecureString)
}
Else {
	$newPScreds = Get-Credential -message "Enter UCS admin creds here:"
	#$rng = [System.Security.Cryptography.RNGCryptoServiceProvider]::Create()
	#$rng.GetBytes($Key)
	$exportObject = New-Object psobject -Property @{
		UserName = $newPScreds.UserName
		Password = ConvertFrom-SecureString -SecureString $newPScreds.Password -Key $Key
	}
	$exportObject | Export-Clixml UCSCreds.xml
	$MyCredential = $newPScreds
}

$UCSCredentials = $MyCredential

function GenerateReport()
{
	Param([Parameter(Mandatory=$true)][string]$UCSM,
				[Parameter(Mandatory=$true)][string]$OutFile,
				[Parameter(Mandatory=$true)][System.Management.Automation.PSCredential]$UCSCredentials)

	# Connect to the UCS
	[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 #Added for proxy requirement. may need to comment this out to disable if not using TLS1.2
	$ucsc = Connect-Ucs -Name $UCSM -Credential $UCSCredentials

	# Test connection
	$connected = Get-UcsPSSession
	if ($connected -eq $null) {
		Write-Host "Error connecting to UCS Manager!"
		Return
	}

	Write-Host "Connected to: $UCSM"
	Write-Host "Starting inventory collection and outputting to: "
	Write-Host "$OutFile"

	#Remove spreadsheet if needed
	Remove-Item -Path $Outfile -ErrorAction Ignore

	# Get Fabric Interconnects
	$WORKSHEETNAME =  "Fabric Interconnects"
	$FI = Get-UcsNetworkElement | Select-Object Ucs,Rn,OobIfIp,OobIfMask,OobIfGw,Operability,Model,Serial | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Fabric Interconnect inventory
	$WORKSHEETNAME =  "FI Inventory"
	Get-UcsFiModule | Sort-Object -Property Dn | Select-Object Dn,Model,Descr,OperState,State,Serial | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Cluster State
	$WORKSHEETNAME =  "Cluster Status"
	Get-UcsStatus | Select-Object Name,VirtualIpv4Address,HaConfiguration,HaReadiness,HaReady,EthernetState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent
	Get-UcsStatus | Select-Object FiALeadership,FiAOobIpv4Address,FiAManagementServicesState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsStatus | Select-Object FiBLeadership,FiBOobIpv4Address,FiBManagementServicesState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append

	# Get Chassis info
	$WORKSHEETNAME =  "Chassis Inventory"
	Get-UcsChassis | Sort-Object -Property Rn | Select-Object Rn,AdminState,Model,OperState,LicState,Power,Thermal,Serial | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get chassis IOM info
	$WORKSHEETNAME =  "IOM Inventory"
	Get-UcsIom | Sort-Object -Property Dn | Select-Object ChassisId,Rn,Model,Discovery,ConfigState,OperState,Side,Thermal,Serial | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Fabric Interconnect to Chassis port mapping
	$WORKSHEETNAME =  "FI to IOM Connections"
	Get-UcsEtherSwitchIntFIo | Select-Object ChassisId,Discovery,Model,OperState,SwitchId,PeerSlotId,PeerPortId,SlotId,PortId,XcvrType | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Global chassis discovery policy
	$chassisDiscoveryPolicy = Get-UcsChassisDiscoveryPolicy | Select-Object Rn,LinkAggregationPref,Action
	$WORKSHEETNAME =  "Chassis Discovery Policy"
	$chassisDiscoveryPolicy | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Global chassis power redundancy policy
	$chassisPowerRedPolicy = Get-UcsPowerControlPolicy
	$WORKSHEETNAME =  "Chassis Power Redundancy Policy"
	$chassisPowerRedPolicy | Select-Object Rn,Redundancy | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get all UCS servers and server info

	# Does the system have blade servers? return those
	if (Get-UcsBlade) {
		$WORKSHEETNAME =  "Server Inventory - Blades"
		Get-UcsBlade | Select-Object ServerId,Model,AvailableMemory,@{N='CPUs';E={$_.NumOfCpus}},@{N='Cores';E={$_.NumOfCores}},@{N='Adaptors';E={$_.NumOfAdaptors}},@{N='eNICs';E={$_.NumOfEthHostIfs}},@{N='fNICs';E={$_.NumOfFcHostIfs}},AssignedToDn,OperPower,Serial | Sort-Object -Property ChassisID,SlotID | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent
	}
	# Does the system have rack servers? return those
	if (Get-UcsRackUnit) {
		$WORKSHEETNAME =  "Server Inventory - Rack-mounts"
		Get-UcsRackUnit | Select-Object ServerId,Model,AvailableMemory,@{N='CPUs';E={$_.NumOfCpus}},@{N='Cores';E={$_.NumOfCores}},@{N='Adaptors';E={$_.NumOfAdaptors}},@{N='eNICs';E={$_.NumOfEthHostIfs}},@{N='fNICs';E={$_.NumOfFcHostIfs}},AssignedToDn,OperPower,Serial | Sort-Object { [int]$_.ServerId } | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent
	}

	# Get server adaptor (mezzanine card) info
	$WORKSHEETNAME =  "Server Adaptor Inventory"
	Get-UcsAdaptorUnit | Sort-Object -Property Dn | Select-Object ChassisId,BladeId,Rn,Model | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get server adaptor port expander info
	$WORKSHEETNAME =  "Servers with Adpt Port Expandrs"
	Get-UcsAdaptorUnitExtn | Sort-Object -Property Dn | Select-Object Dn,Model,Presence | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get server processor info
	$WORKSHEETNAME =  "Server CPU Inventory"
	Get-UcsProcessorUnit | Sort-Object -Property Dn | Select-Object Dn,SocketDesignation,Cores,CoresEnabled,Threads,Speed,OperState,Thermal,Model | Where-Object {$_.OperState -ne "removed"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get server memory info
	$WORKSHEETNAME =  "Server Memory Inventory"
	Get-UcsMemoryUnit | Sort-Object -Property Dn,Location | where {$_.Capacity -ne "unspecified"} | Select-Object -Property Dn,Location,Capacity,Clock,OperState,Model | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get server storage controller info
	$WORKSHEETNAME =  "Server Storage Controller Inv"
	Get-UcsStorageController | Sort-Object -Property Dn | Select-Object Vendor,Model | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get server local disk info
	$WORKSHEETNAME =  "Server Local Disk Inventory"
	Get-UcsStorageLocalDisk | Sort-Object -Property Dn | Select-Object Dn,Model,Size,Serial | where {$_.Size -ne "unknown"}  | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get UCSM firmware version
	$WORKSHEETNAME =  "UCS Manager"
	Get-UcsFirmwareRunning | Select-Object Dn,Type,Version | Sort-Object -Property Dn | Where-Object {$_.Type -eq "mgmt-ext"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Fabric Interconnect firmware
	$WORKSHEETNAME =  "Fabric Interconnect"
	Get-UcsFirmwareRunning | Select-Object Dn,Type,Version | Sort-Object -Property Dn | Where-Object {$_.Type -eq "switch-kernel" -OR $_.Type -eq "switch-software"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get IOM firmware
	$WORKSHEETNAME =  "IOM"
	Get-UcsFirmwareRunning | Select-Object Deployment,Dn,Type,Version | Sort-Object -Property Dn | Where-Object {$_.Type -eq "iocard"} | Where-Object -FilterScript {$_.Deployment -notlike "boot-loader"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Server Adapter firmware
	$WORKSHEETNAME =  "Server Adapters"
	Get-UcsFirmwareRunning | Select-Object Deployment,Dn,Type,Version | Sort-Object -Property Dn | Where-Object {$_.Type -eq "adaptor"} | Where-Object -FilterScript {$_.Deployment -notlike "boot-loader"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Server CIMC firmware
	$WORKSHEETNAME =  "Server CIMC"
	Get-UcsFirmwareRunning | Select-Object Deployment,Dn,Type,Version | Sort-Object -Property Dn | Where-Object {$_.Type -eq "blade-controller"} | Where-Object -FilterScript {$_.Deployment -notlike "boot-loader"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Server BIOS versions
	$WORKSHEETNAME =  "Server BIOS"
	Get-UcsFirmwareRunning | Select-Object Dn,Type,Version | Sort-Object -Property Dn | Where-Object {$_.Type -eq "blade-bios" -Or $_.Type -eq "rack-bios"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Host Firmware Packages
	$WORKSHEETNAME =  "Host Firmware Packages"
	Get-UcsFirmwareComputeHostPack | Select-Object Dn,Name,BladeBundleVersion,RackBundleVersion | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	##########################################################################################################################################################
	##########################################################################################################################################################
	###################################################  SERVICE CONFIGURATION  ##############################################################################
	##########################################################################################################################################################
	##########################################################################################################################################################

	# Get Service Profile Templates
	$WORKSHEETNAME =  "Service Profile Templates"
	Get-UcsServiceProfile | Where-object {$_.Type -ne "instance"}  | Sort-object -Property Name | Select-Object Dn,Name,BiosProfileName,BootPolicyName,HostFwPolicyName,LocalDiskPolicyName,MaintPolicyName,VconProfileName | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Service Profiles
	$WORKSHEETNAME =  "Service Profiles"
	Get-UcsServiceProfile | Where-object {$_.Type -eq "instance"}  | Sort-object -Property Name | Select-Object Dn,Name,OperSrcTemplName,AssocState,PnDn,BiosProfileName,IdentPoolName,Uuid,BootPolicyName,HostFwPolicyName,LocalDiskPolicyName,MaintPolicyName,VconProfileName,OperState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Maintenance Policies
	$WORKSHEETNAME =  "Maintenance Policies"
	Get-UcsMaintenancePolicy | Select-Object Name,Dn,UptimeDisr,Descr | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Boot Policies
	$WORKSHEETNAME =  "Boot Policies"
	Get-UcsBootPolicy | sort-object -Property Dn | Select-Object Dn,Name,Purpose,RebootOnUpdate | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get SAN Boot Policies
	$WORKSHEETNAME =  "SAN Boot Policies"
	Get-UcsLsbootSanImagePath | sort-object -Property Dn | Select-Object Dn,Type,Vnicname,Lun,Wwn | Where-Object -FilterScript {$_.Dn -notlike "sys/chassis*"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Local Disk Policies
	$WORKSHEETNAME =  "Local Disk Policies"
	Get-UcsLocalDiskConfigPolicy | Select-Object Dn,Name,Mode,Descr | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Scrub Policies
	$WORKSHEETNAME =  "Scrub Policies"
	Get-UcsScrubPolicy | Select-Object Dn,Name,BiosSettingsScrub,DiskScrub | Where-Object {$_.Name -ne "policy"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get BIOS Policies
	$WORKSHEETNAME =  "BIOS Policies"
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"} | Select-Object Dn,Name | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get BIOS Policy Settings
	$WORKSHEETNAME =  "BIOS Policy Settings"
	"UcsBiosVfQuietBoot" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfQuietBoot | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosVfPOSTErrorPause" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfPOSTErrorPause | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosVfResumeOnACPowerLoss" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfResumeOnACPowerLoss | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosVfFrontPanelLockout" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfFrontPanelLockout | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosTurboBoost	" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosTurboBoost | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosEnhancedIntelSpeedStep" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosEnhancedIntelSpeedStep | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosHyperThreading" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosHyperThreading | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosVfCoreMultiProcessing" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfCoreMultiProcessing | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosExecuteDisabledBit" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosExecuteDisabledBit | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosVfIntelVirtualizationTechnology" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfIntelVirtualizationTechnology | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosVfDirectCacheAccess" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfDirectCacheAccess | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfProcessorCState" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfProcessorCState | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosVfProcessorC1E" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfProcessorC1E | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfProcessorC3Report" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfProcessorC3Report | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfProcessorC6Report" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfProcessorC6Report | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfProcessorC7Report" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfProcessorC7Report | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfCPUPerformance" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfCPUPerformance | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfMaxVariableMTRRSetting" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfMaxVariableMTRRSetting | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosIntelDirectedIO" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosIntelDirectedIO | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfSelectMemoryRASConfiguration" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfSelectMemoryRASConfiguration | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"UcsBiosNUMA" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosNUMA | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append 
	"UcsBiosLvDdrMode" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosLvDdrMode | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append 
	"UcsBiosVfUSBBootConfig" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfUSBBootConfig | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append 
	"UcsBiosVfUSBFrontPanelAccessLock" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfUSBFrontPanelAccessLock | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append 
	"UcsBiosVfUSBSystemIdlePowerOptimizingSetting" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfUSBSystemIdlePowerOptimizingSetting | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append 
	"AppendUcsBiosVfMaximumMemoryBelow4GB" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfMaximumMemoryBelow4GB | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfMemoryMappedIOAbove4GB" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfMemoryMappedIOAbove4GB | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfBootOptionRetry" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfBootOptionRetry | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfIntelEntrySASRAIDModule" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfIntelEntrySASRAIDModule | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	"AppendUcsBiosVfOSBootWatchdogTimer" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	Get-UcsBiosPolicy | Where-Object {$_.Name -ne "SRIOV"}  | Get-UcsBiosVfOSBootWatchdogTimer | Sort-Object Dn | Select-Object Dn,Vp* | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append


	# Get Service Profiles vNIC/vHBA Assignments
	$WORKSHEETNAME =  "Service Profile vNIC Placements"
	Get-UcsLsVConAssign -Transport ethernet | Select-Object Dn,Vnicname,Adminvcon,Order | Sort-Object Dn | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Ethernet VLAN to vNIC Mappings #
	$WORKSHEETNAME =  "Ethernet VLAN to vNIC Mappings"
	Get-UcsAdaptorVlan | sort-object Dn |Select-Object Dn,Name,Id,SwitchId | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get UUID Suffix Pools
	$WORKSHEETNAME =  "UUID Pools"
	Get-UcsUuidSuffixPool | Select-Object Dn,Name,AssignmentOrder,Prefix,Size,Assigned | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get UUID Suffix Pool Blocks
	$WORKSHEETNAME =  "UUID Pool Blocks"
	Get-UcsUuidSuffixBlock | Select-Object Dn,From,To | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get UUID UUID Pool Assignments
	$WORKSHEETNAME =  "UUID Pool Assignments"
	Get-UcsUuidpoolAddr | Where-Object {$_.Assigned -ne "no"} | select-object AssignedToDn,Id | sort-object -property AssignedToDn | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Server Pools
	$WORKSHEETNAME =  "Server Pools"
	Get-UcsServerPool | Select-Object Dn,Name,Assigned | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Server Pool Assignments
	$WORKSHEETNAME =  "Server Pool Assignments"
	"UcsComputePooledSlot" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	$SLOT = Get-UcsComputePooledSlot | Select-Object Dn,Rn 
	IF($SLOT){$SLOT | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append}
	"UcsComputePooledRackUnit" | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append
	$RACK = Get-UcsComputePooledRackUnit | Select-Object Dn,PoolableDn 
	IF($RACK){$RACK | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent -Append}


	##########################################################################################################################################################
	##########################################################################################################################################################
	###################################################    LAN CONFIGURATION    ##############################################################################
	##########################################################################################################################################################
	##########################################################################################################################################################

	# Get LAN Switching Mode
	$WORKSHEETNAME =  "FI Eth Switching Mode"
	Get-UcsLanCloud | Select-Object Rn,Mode | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Fabric Interconnect Ethernet port usage and role info
	$WORKSHEETNAME =  "FI Eth Port Configuration"
	Get-UcsFabricPort | Select-Object Dn,IfRole,LicState,Mode,OperState,OperSpeed,XcvrType | Where-Object {$_.OperState -eq "up"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Ethernet LAN Uplink Port Channel info
	$WORKSHEETNAME =  "FI Eth Uplink Port Channels"
	Get-UcsUplinkPortChannel | Sort-Object -Property Name | Select-Object Dn,Name,OperSpeed,OperState,Transport | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Ethernet LAN Uplink Port Channel port membership info
	$WORKSHEETNAME =  "FI Eth Uplink Port Channel Mbrs"
	Get-UcsUplinkPortChannelMember | Sort-Object -Property Dn |Select-Object Dn,Membership | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get QoS Class Configuration
	$WORKSHEETNAME =  "QoS System Class Configuration"
	Get-UcsQosClass | Select-Object Priority,AdminState,Cos,Weight,Drop,Mtu | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get QoS Policies
	$WORKSHEETNAME =  "QoS Policies"
	Get-UcsQosPolicy | Select-Object Dn,Name | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get QoS vNIC Policy Map
	$WORKSHEETNAME =  "QoS vNIC Policy Map"
	Get-UcsVnicEgressPolicy | Sort-Object -Property Prio | Select-Object Dn,Prio | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Ethernet VLANs
	$WORKSHEETNAME =  "Ethernet VLANs"
	Get-UcsVlan | where {$_.IfRole -eq "network"} | Sort-Object -Property Id | Select-Object Id,Name,SwitchId | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Network Control Policies
	$WORKSHEETNAME =  "Network Control Policies"
	Get-UcsNetworkControlPolicy | Select-Object Dn,Name,Cdp,UplinkFailAction | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get vNIC Templates
	$vnicTemplates = Get-UcsVnicTemplate | Select-Object Dn,Name,Descr,SwitchId,TemplType,IdentPoolName,Mtu,NwCtrlPolicyName,QosPolicyName
	$WORKSHEETNAME =  "vNIC Templates"
	$vnicTemplates | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Ethernet VLAN to vNIC Mappings #
	$WORKSHEETNAME =  "Ethernet VLAN to vNIC Mappings"
	Get-UcsAdaptorVlan | sort-object Dn |Select-Object Dn,Name,Id,SwitchId | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get IP Pools
	$WORKSHEETNAME =  "IP Pools"
	Get-UcsIpPool | Select-Object Dn,Name,AssignmentOrder,Size | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get IP Pool Blocks
	$WORKSHEETNAME =  "IP Pool Blocks"
	Get-UcsIpPoolBlock | Select-Object Dn,From,To,Subnet,DefGw | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get IP CIMC MGMT Pool Assignments
	$WORKSHEETNAME =  "CIMC IP Pool Assignments"
	Get-UcsIpPoolAddr | Sort-Object -Property AssignedToDn | where {$_.Assigned -eq "yes"} | Select-Object AssignedToDn,Id | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get MAC Address Pools
	$WORKSHEETNAME =  "MAC Address Pools"
	Get-UcsMacPool | Select-Object Dn,Name,AssignmentOrder,Size,Assigned | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get MAC Address Pool Blocks
	$WORKSHEETNAME =  "MAC Address Pool Blocks"
	Get-UcsMacMemberBlock | Select-Object Dn,From,To | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get MAC Pool Assignments
	$WORKSHEETNAME =  "MAC Address Pool Assignments"
	Get-UcsVnic | Sort-Object -Property Dn | Select-Object Dn,IdentPoolName,Addr | where {$_.Addr -ne "derived"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	##########################################################################################################################################################
	##########################################################################################################################################################
	###################################################    SAN CONFIGURATION    ##############################################################################
	##########################################################################################################################################################
	##########################################################################################################################################################

	# Get SAN Switching Mode
	$WORKSHEETNAME =  "FI Fibre Channel Switching Mode"
	Get-UcsSanCloud | Select-Object Rn,Mode | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Fabric Interconnect FC Uplink Ports
	$WORKSHEETNAME =  "FI FC Uplink Ports"
	Get-UcsFiFcPort | Select-Object EpDn,SwitchId,SlotId,PortId,LicState,Mode,OperSpeed,OperState,wwn | sort-object -descending  | where-object {$_.OperState -ne "sfp-not-present"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get SAN Fiber Channel Uplink Port Channel info
	$WORKSHEETNAME =  "FI FC Uplink Port Channels"
	Get-UcsFcUplinkPortChannel | Select-Object Dn,Name,OperSpeed,OperState,Transport | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Fabric Interconnect FCoE Uplink Ports
	$WORKSHEETNAME =  "FI FCoE Uplink Ports"
	Get-UcsFabricPort | Where-Object {$_.IfRole -eq "fcoe-uplink"} | Select-Object IfRole,EpDn,LicState,OperState,OperSpeed | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get SAN FCoE Uplink Port Channel info
	$WORKSHEETNAME =  "FI FCoE Uplink Port Channels"
	Get-UcsFabricFcoeSanPc | Select-Object Dn,Name,FcoeState,OperState,Transport,Type | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get SAN FCoE Uplink Port Channel Members
	$WORKSHEETNAME =  "FI FCoE Uplink Port Channel Mbr"
	Get-UcsFabricFcoeSanPcEp | Select-Object Dn,IfRole,LicState,Membership,OperState,SwitchId,PortId,Type | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get FC VSAN info
	$WORKSHEETNAME =  "FC VSANs"
	Get-UcsVsan | Select-Object Dn,Id,FcoeVlan,DefaultZoning | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get FC Port Channel VSAN Mapping
	$WORKSHEETNAME =  "FC VSAN to FC Port Mappings"
	Get-UcsVsanMemberFcPortChannel | Select-Object EpDn,IfType | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get vHBA Templates
	$vhbaTemplates = Get-UcsVhbaTemplate | Select-Object Dn,Name,Descr,SwitchId,TemplType,QosPolicyName
	$WORKSHEETNAME =  "vHBA Templates"
	$vhbaTemplates | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Service Profiles vNIC/vHBA Assignments
	$WORKSHEETNAME =  "Service Profile vHBA Placements"
	Get-UcsLsVConAssign -Transport fc | Select-Object Dn,Vnicname,Adminvcon,Order | Sort-Object dn | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get vHBA to VSAN Mappings
	$WORKSHEETNAME =  "vHBA to VSAN Mappings"
	Get-UcsVhbaInterface | Select-Object Dn,OperVnetName,Initiator | Where-Object {$_.Initiator -ne "00:00:00:00:00:00:00:00"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get WWNN Pools
	$WORKSHEETNAME =  "WWN Pools"
	Get-UcsWwnPool | Select-Object Dn,Name,AssignmentOrder,Purpose,Size,Assigned | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get WWNN/WWPN Pool Assignments
	$WORKSHEETNAME =  "WWN Pool Assignments"
	Get-UcsVhba | Sort-Object -Property Addr | Select-Object Dn,IdentPoolName,NodeAddr,Addr | where {$_.NodeAddr -ne "vnic-derived"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get WWNN/WWPN vHBA and adaptor Assignments
	$WORKSHEETNAME =  "vHBA Details"
	Get-UcsAdaptorHostFcIf | sort-object -Property VnicDn -Descending | Select-Object VnicDn,Vendor,Model,LinkState,SwitchId,NodeWwn,Wwn | Where-Object {$_.NodeWwn -ne "00:00:00:00:00:00:00:00"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	##########################################################################################################################################################
	##########################################################################################################################################################
	###################################################   ADMIN CONFIGURATION   ##############################################################################
	##########################################################################################################################################################
	##########################################################################################################################################################

	# Get Organizations
	$WORKSHEETNAME =  "Organizations"
	Get-UcsOrg | Select-Object Name,Dn | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Fault Policy
	$WORKSHEETNAME =  "Fault Policy"
	Get-UcsFaultPolicy | Select-Object Rn,AckAction,ClearAction,ClearInterval,FlapInterval,RetentionInterval | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Syslog Remote Destinations
	$WORKSHEETNAME =  "Remote Syslog"
	Get-UcsSyslogClient | Where-Object {$_.AdminState -ne "disabled"} | Select-Object Rn,Severity,Hostname,ForwardingFacility | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Syslog Sources
	$WORKSHEETNAME =  "Syslog Sources"
	Get-UcsSyslogSource | Select-Object Rn,Audits,Events,Faults | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Syslog Local File
	$WORKSHEETNAME =  "Syslog Local File"
	Get-UcsSyslogFile | Select-Object Rn,Name,AdminState,Severity,Size | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Full State Backup Policy
	$WORKSHEETNAME =  "Full State Backup Policy"
	Get-UcsMgmtBackupPolicy | Select-Object Descr,Host,LastBackup,Proto,Schedule,AdminState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get All Config Backup Policy
	$WORKSHEETNAME =  "All Configuration Backup Policy"
	Get-UcsMgmtCfgExportPolicy | Select-Object Descr,Host,LastBackup,Proto,Schedule,AdminState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Native Authentication Source
	$WORKSHEETNAME =  "Native Authentication"
	Get-UcsNativeAuth | Select-Object Rn,DefLogin,ConLogin,DefRolePolicy | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get local users
	$WORKSHEETNAME =  "Local users"
	Get-UcsLocalUser | Sort-Object Name | Select-Object Name,Email,AccountStatus,Expiration,Expires,PwdLifeTime | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get LDAP server info
	$WORKSHEETNAME =  "LDAP Providers"
	Get-UcsLdapProvider | Select-Object Name,Rootdn,Basedn,Attribute | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get LDAP group mappings
	$WORKSHEETNAME =  "LDAP Group Mappings"
	Get-UcsLdapGroupMap | Select-Object Name | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get user and LDAP group roles
	$WORKSHEETNAME =  "LDAP User Roles"
	Get-UcsUserRole | Select-Object Name,Dn | Where-Object {$_.Dn -like "sys/ldap-ext*"} | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get tacacs providers
	$WORKSHEETNAME =  "TACACS+ Providers"
	Get-UcsTacacsProvider | Sort-Object -Property Order,Name | Select-Object Order,Name,Port,KeySet,Retries,Timeout | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Call Home config
	$callHome = Get-UcsCallhome
	$WORKSHEETNAME =  "Call Home Configuration"
	$callHome | Sort-Object -Property Ucs | Select-Object AdminState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Call Home SMTP Server
	$WORKSHEETNAME =  "Call Home SMTP Server"
	Get-UcsCallhomeSmtp | Sort-Object -Property Ucs | Select-Object Host | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Call Home Recipients
	$WORKSHEETNAME =  "Call Home Recipients"
	Get-UcsCallhomeRecipient | Sort-Object -Property Ucs | Select-Object Dn,Email | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get SNMP Configuration
	$WORKSHEETNAME =  "SNMP Configuration"
	Get-UcsSnmp | Sort-Object -Property Ucs | Select-Object AdminState,Community,SysContact,SysLocation | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get DNS Servers
	$dnsServers = Get-UcsDnsServer | Select-Object Name
	$WORKSHEETNAME =  "DNS Servers"
	$dnsServers | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Timezone
	$WORKSHEETNAME =  "Timezone"
	Get-UcsTimezone | Select-Object Timezone | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get NTP Servers
	$ntpServers = Get-UcsNtpServer | Select-Object Name
	$WORKSHEETNAME =  "NTP Servers"
	$ntpServers | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Cluster Configuration and State
	$WORKSHEETNAME =  "Cluster Configuration"
	Get-UcsStatus | Select-Object Name,VirtualIpv4Address,HaConfiguration,HaReadiness,HaReady,FiALeadership,FiAOobIpv4Address,FiAOobIpv4DefaultGateway,FiAManagementServicesState,FiBLeadership,FiBOobIpv4Address,FiBOobIpv4DefaultGateway,FiBManagementServicesState | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Management Interface Monitoring Policy
	$WORKSHEETNAME =  "Management Int Monitoring Pol"
	Get-UcsMgmtInterfaceMonitorPolicy | Select-Object AdminState,EnableHAFailover,MonitorMechanism | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get host-id information
	$WORKSHEETNAME =  "FI HostIDs"
	Get-UcsLicenseServerHostId | Sort-Object -Property Scope | Select-Object Scope,HostId | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get installed license information
	$ucsLicenses = Get-UcsLicense
	$WORKSHEETNAME =  "Installed Licenses"
	$ucsLicenses | Sort-Object -Property Scope,Feature | Select-Object Scope,Feature,Sku,AbsQuant,UsedQuant,GracePeriodUsed,OperState,PeerStatus | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	##########################################################################################################################################################
	##########################################################################################################################################################
	###################################################        STATISTICS       ##############################################################################
	##########################################################################################################################################################
	##########################################################################################################################################################

	# Get all UCS Faults sorted by severity
	$WORKSHEETNAME =  "Faults"
	Get-UcsFault | Sort-Object -Property @{Expression = {$_.Severity}; Ascending = $true}, Created -Descending | Select-Object Severity,Created,Descr,dn | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get chassis power usage stats
	$WORKSHEETNAME =  "Chassis Power"
	Get-UcsChassisStats | Select-Object Dn,InputPower,InputPowerAvg,InputPowerMax,InputPowerMin,OutputPower,OutputPowerAvg,OutputPowerMax,OutputPowerMin,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get chassis and FI power status
	$WORKSHEETNAME =  "Chas and FI Power Supply Status"
	Get-UcsPsu | Sort-Object -Property Dn | Select-Object Dn,OperState,Perf,Power,Thermal,Voltage | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get chassis PSU stats
	$WORKSHEETNAME =  "Chassis Power Supplies"
	Get-UcsPsuStats | Sort-Object -Property Dn | Select-Object Dn,AmbientTemp,AmbientTempAvg,Input210v,Input210vAvg,Output12v,Output12vAvg,OutputCurrentAvg,OutputPowerAvg,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get chassis and FI fan stats
	$WORKSHEETNAME =  "Chassis and FI Fan"
	Get-UcsFan | Sort-Object -Property Dn | Select-Object Dn,Module,Id,Perf,Power,OperState,Thermal | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get chassis IOM temp stats
	$WORKSHEETNAME =  "Chassis IOM Temperatures"
	Get-UcsEquipmentIOCardStats | Sort-Object -Property Dn | Select-Object Dn,AmbientTemp,AmbientTempAvg,Temp,TempAvg,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get server power usage
	$WORKSHEETNAME =  "Server Power"
	Get-UcsComputeMbPowerStats | Sort-Object -Property Dn | Select-Object Dn,ConsumedPower,ConsumedPowerAvg,ConsumedPowerMax,InputCurrent,InputCurrentAvg,InputVoltage,InputVoltageAvg,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get server temperatures
	$WORKSHEETNAME =  "Server Temperatures"
	Get-UcsComputeMbTempStats | Sort-Object -Property Dn | Select-Object Dn,FmTempSenIo,FmTempSenIoAvg,FmTempSenIoMax,FmTempSenRear,FmTempSenRearAvg,FmTempSenRearMax,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get Memory temperatures
	$WORKSHEETNAME =  "Memory Temperatures"
	Get-UcsMemoryUnitEnvStats | Sort-Object -Property Dn | Select-Object Dn,Temperature,TemperatureAvg,TemperatureMax,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get CPU power and temperatures
	$WORKSHEETNAME =  "CPU Power and Temperatures"
	Get-UcsProcessorEnvStats | Sort-Object -Property Dn | Select-Object Dn,InputCurrent,InputCurrentAvg,InputCurrentMax,Temperature,TemperatureAvg,TemperatureMax,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get LAN Uplink Port Channel Loss Stats
	$WORKSHEETNAME =  "LAN Uplink Port Channel Loss"
	Get-UcsUplinkPortChannel | Get-UcsEtherLossStats | Sort-Object -Property Dn | Select-Object Dn,ExcessCollision,ExcessCollisionDeltaAvg,LateCollision,LateCollisionDeltaAvg,MultiCollision,MultiCollisionDeltaAvg,SingleCollision,SingleCollisionDeltaAvg | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get LAN Uplink Port Channel Receive Stats
	$WORKSHEETNAME =  "LAN Uplink Port Channel Receive"
	Get-UcsUplinkPortChannel | Get-UcsEtherRxStats | Sort-Object -Property Dn | Select-Object Dn,BroadcastPackets,BroadcastPacketsDeltaAvg,JumboPackets,JumboPacketsDeltaAvg,MulticastPackets,MulticastPacketsDeltaAvg,TotalBytes,TotalBytesDeltaAvg,TotalPackets,TotalPacketsDeltaAvg,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get LAN Uplink Port Channel Transmit Stats
	$WORKSHEETNAME =  "LAN Uplink Port Channel Transm"
	Get-UcsUplinkPortChannel | Get-UcsEtherTxStats | Sort-Object -Property Dn | Select-Object Dn,BroadcastPackets,BroadcastPacketsDeltaAvg,JumboPackets,JumboPacketsDeltaAvg,MulticastPackets,MulticastPacketsDeltaAvg,TotalBytes,TotalBytesDeltaAvg,TotalPackets,TotalPacketsDeltaAvg,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get vNIC Stats
	$WORKSHEETNAME =  "vNICs"
	Get-UcsAdaptorVnicStats | Sort-Object -Property Dn | Select-Object Dn,BytesRx,BytesRxDeltaAvg,BytesTx,BytesTxDeltaAvg,PacketsRx,PacketsRxDeltaAvg,PacketsTx,PacketsTxDeltaAvg,DroppedRx,DroppedRxDeltaAvg,DroppedTx,DroppedTxDeltaAvg,ErrorsTx,ErrorsTxDeltaAvg,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get FC Uplink Port Channel Loss Stats
	$WORKSHEETNAME =  "FC Uplink Ports"
	Get-UcsFcErrStats | Sort-Object -Property Dn | Select-Object Dn,CrcRx,CrcRxDeltaAvg,DiscardRx,DiscardRxDeltaAvg,DiscardTx,DiscardTxDeltaAvg,LinkFailures,SignalLosses,Suspect | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Get FCoE Uplink Port Channel Stats
	$WORKSHEETNAME =  "FCoE Uplink Port Channels"
	Get-UcsEtherFcoeInterfaceStats | Select-Object DN,BytesRx,BytesTx,DroppedRx,DroppedTx,ErrorsRx,ErrorsTx | Export-Excel -Path $OUTFILE -Worksheetname $WORKSHEETNAME -ShowPercent

	# Disconnect
	Write-Host "Disconnecting from UCS $UCSM"
	Disconnect-Ucs

	Write-Host "Done generating report for $UCSM"
}

#Run for each UCSM Environment
Write-Host "Starting Cisco UCS Inventory Script (UIS).."
ForEach($UCS in $UCSLIST)
{

##Specify Export File Info
$EXPORTFILENAME = "UCS_Inventory_"+"$UCS"+"_"+$LOGDATE+".xlsx"
#Create Info Folder
$INFOFOLDER = $pwd.path+"\info"
If (Test-Path $INFOFOLDER){
	Write-Host "Info Directory Created. Continuing..."
}Else{
	New-Item $INFOFOLDER -type directory
}
#Create UCS Folder
$AFOLDER = $INFOFOLDER+"\UCS"
If (Test-Path $AFOLDER){
	Write-Host "UCS Directory Created. Continuing..."
}Else{
	New-Item $AFOLDER -type directory
}
#Create UCS Folder
$UCSFOLDER = $AFOLDER+"\$UCS"
If (Test-Path $UCSFOLDER){
	Write-Host "UCS Directory Created. Continuing..."
}Else{
	New-Item $UCSFOLDER -type directory
}
#Create Inventory Folder
$EXPORTFOLDER = $UCSFOLDER+"\Inventory"
If (Test-Path $EXPORTFOLDER){
	Write-Host "Inventory Directory Created. Continuing..."
}Else{
	New-Item $EXPORTFOLDER -type directory
}
#Specify Log File
$OutFile = $EXPORTFOLDER+"\"+$EXPORTFILENAME
Write-Host "Completed creating Export Folders and Variables"
Write-Host (Get-Date -format "MMM-dd-yyyy_HH-mm-ss")
Write-Host "-----------------------------------------------------------------------------------------------------------------------"

##Clean up Old XLSX Files Prior to Starting
Write-Host "Deleting Old XLSX Exports"
$DatetoDelete = $CurrentDate.AddMonths($MONTHSTOKEEP)
Get-ChildItem $EXPORTFOLDER -Filter *.xlsx | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Confirm:$false
Write-Host "Completed deleting Old XLSX Exports"

##Disconnect from any open UCS Sessions
#This can cause problems if there are any
Write-Host "-----------------------------------------------------------------------------------------------------------------------"
Write-Host (Get-Date -format "MMM-dd-yyyy_HH-mm-ss")
Write-Host "Disconnecting from any Open UCS Sessions"
TRY
{Disconnect-Ucs}
CATCH
{Write-Host "No Open UCS Sessions found"}
Write-Host (Get-Date -format "MMM-dd-yyyy_HH-mm-ss")
Write-Host "-----------------------------------------------------------------------------------------------------------------------"

#
Write-Host "Starting Cisco UCS Inventory Script (UIS) for $UCS.."
GenerateReport -UCSM $UCS -Outfile $OutFile -UCSCredentials $UCSCredentials
Write-Host (Get-Date -format "MMM-dd-yyyy_HH-mm-ss")
Write-Host "-----------------------------------------------------------------------------------------------------------------------"
}
