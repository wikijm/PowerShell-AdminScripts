#requires -version 2
<#
.SYNOPSIS
  Delete Hyper-V VM with an input .CSV file

.DESCRIPTION
  Delete Hyper-V Virtual Machines with an input .CSV file (contains Name;DiskCapacityInGB;Generation;CPUNb;StartupRAMinMB;MinimumRAMinMB;SwitchName)

.INPUTS
 .CSV file selected by user during the script

.OUTPUTS
  Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log

.NOTES
  Version:        1.0
  Author:         ALBERT Jean-Marc
  Creation Date:  26/08/2015
  Purpose/Change: 2015.08.24 - ALBERT Jean-Marc - Initial script development
				  2015.08.26 - ALBERT Jean-Marc - Replace .CSV fixed name per file dialog selection
                                                 
  
.EXAMPLE
  <None>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "2.0"

#Write script directory path on "ScriptDir" variable
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Log file creation, similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD].log
$ActualDate = Get-Date -uformat % Y_% m_% d
$ScriptLogFile = "$ScriptDir\$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))" + "_" + $ActualDate + ".log"

# Declare Environment and Hyper-V parameters
$VMLoc = "E:\Hyper-V\"
$NetworkSwitch1 = "Externe"
$NetworkSwitch1Type = "Public"
$NetworkSwitch2 = "LAN 192.168.1.0"
$NetworkSwitch2Type = "Private"

#-----------------------------------------------------------[Functions]------------------------------------------------------------
function Stop-TranscriptOnLog
{
	Stop-Transcript
	# Add EOL required for Notepad.exe application usage
	[string]::Join("`r`n", (Get-Content $ScriptLogFile)) | Out-File $ScriptLogFile
}

function Select-FileDialog
{
	param ([string]$Title, [string]$Filter = "All files *.*|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
	$fileDialogBox = New-Object Windows.Forms.OpenFileDialog
	$fileDialogBox.ShowHelp = $false
	$fileDialogBox.initialDirectory = $ScriptDir
	$fileDialogBox.filter = $Filter
	$fileDialogBox.Title = $Title
	$Show = $fileDialogBox.ShowDialog()
	
	If ($Show -eq "OK")
	{
		Return $fileDialogBox.FileName
	}
	Else
	{
		Write-Error "Canceled operation"
		[System.Windows.Forms.MessageBox]::Show("Script is not able to continue. Operation stopped.", "Operation canceled", 0, [Windows.Forms.MessageBoxIcon]::Error)
		Stop-TranscriptOnLog
		Exit
	}
	
}

#------------------------------------------------------------[Actions]-------------------------------------------------------------
# Start of log completion
Start-Transcript $ScriptLogFile | Out-Null

# Import CSV file
[System.Windows.Forms.MessageBox]::Show(
"
Select on this window the CSV file who contains VM parameters.
Its content must be similar to:
  
Name;DiskCapacityInGB;Generation;CPUNb;StartupRAMinMB;MinimumRAMinMB;SwitchName			(Required line)
VM01;200;2;2;1024;512;LAN 192.168.1.0
VM02;200;2;2;1024;512;LAN 192.168.1.0
VM03;200;2;2;1024;512;LAN 192.168.1.0
", "VM parameters list", 0, [Windows.Forms.MessageBoxIcon]::Question)

$CSVInputFile = Select-FileDialog -Title "Select CSV file" -Filter "CSV File (*.csv) |*.csv"

# Import VM parameters list
$csvValues = Import-Csv $CSVInputFile -Delimiter ';'


# Loop to handle Virtual Machine generation & launch
foreach ($line in $csvValues)
{
	$VMName = $VMList.Name
	$DiskCapacityinGB = $VMList.DiskCapacityInGB
	[int64]$DiskCapacity = 1GB*$DiskCapacityinGB
	$Generation = $VMList.Generation
	$CPU = $VMList.CPUNb
	$MemoryStartupBytes = $VMList.StartupRAMinMB
	$MemoryMinimumBytes = $VMList.MinimumRAMinMB
	[int64]$startupmem = 1MB*$MemoryStartupBytes
	[int64]$minimummem = 1MB*$MemoryMinimumBytes
	$SwitchName = $VMList.SwitchName
	
	# Create Virtual Machines & VHDX
	New-VHD -Path "$VMLoc\$VMName\Virtual Hard Disks\$VMName.vhdx" -SizeBytes $DiskCapacity -Dynamic
	New-VM -Path $VMLoc -Name $VMName -Generation $Generation -MemoryStartupBytes $startupmem -VHDPath "$VMLoc\$VMName\Virtual Hard Disks\$VMName.vhdx" -SwitchName $SwitchName
	Set-VMProcessor –VMName $VMName –count $CPU
	Set-VMMemory -VMName $VMName -DynamicMemoryEnabled $true -StartupBytes $startupmem -MinimumBytes $minimummem
}

$OUTPUT = [System.Windows.Forms.MessageBox]::Show("VM were create well. Would you like to launch them?", "Ended VM creation", 4, [Windows.Forms.MessageBoxIcon]::Question)
if ($OUTPUT -eq "YES")
{
	# Import CSV file
	$VMList = Import-Csv $CSVVMConfigFile -delimiter ';'
	# Start all created VM
	foreach ($VMList in $VMList)
	{
		$VMName = $VMList.Name
		Start-VM $VMName
	}
}
else
{
	[System.Windows.Forms.MessageBox]::Show("Created VM were not started.", "Ended VM creation", 0, [Windows.Forms.MessageBoxIcon]::Information)
}

# Get VM
Get-VM | Sort-Object Name | Select Name, State, CPUUsage, MemoryAssigned | Export-CliXML $ScriptDir\Get-VM.xml
Import-CliXML $ScriptDir\Get-VM.xml | Out-GridView -Title Get-VM -PassThru

# Stop the log transcript
Stop-TranscriptOnLog