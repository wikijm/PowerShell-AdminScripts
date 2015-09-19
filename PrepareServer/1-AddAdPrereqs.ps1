#requires -version 2
<#
.SYNOPSIS
  <None>

.DESCRIPTION
  <None>

.INPUTS
  <None>

.OUTPUTS
  Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
  Create a list of AD Objects, similar to $ScriptDir\[SCRIPTNAME]_FoundADObjects_[YYYY_MM_DD]_[HHhMMmSSs].csv
   
   
.NOTES
  Version:        0.1
  Author:         ALBERT Jean-Marc
  Creation Date:  DD/MM/YYYY (DD/MM/YYYY
  Purpose/Change: 1.0 - YYYY.MM.DD - ALBERT Jean-Marc - Initial script development
                    
                                                  
.SOURCES
  <None>
  
  
.EXAMPLE
  <None>

#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Set-StrictMode -version Latest

#Set Error Action to Silently Continue
#$ErrorActionPreference = "SilentlyContinue"

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$scriptFile = $MyInvocation.MyCommand.Definition
$launchDate = get-date -f "yyyy.MM.dd-HHhmmmss"
$logDirectoryPath = $scriptPath + "\"
$buffer = "$scriptPath\bufferCommand.txt"
$fullScriptPath = (Resolve-Path -Path $buffer).Path

$loggingFunctions = "$scriptPath\logging\Logging.ps1"
$utilsFunctions = "$scriptPath\utilities\Utils.ps1"
$addsFunctions = "$scriptPath\utilities\ADDS.ps1"


#----------------------------------------------------------[Declarations]----------------------------------------------------------

$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "0.1"

if(!(Test-Path $logDirectoryPath)) {
    New-Item $logDirectoryPath -type directory | Out-Null
}

$logFileName = "Log_" + $launchDate + ".log"
$logPathName = "$logDirectoryPath\$logFileName"

$global:streamWriter = New-Object System.IO.StreamWriter $logPathName

#-----------------------------------------------------------[Functions]------------------------------------------------------------

. $loggingFunctions
. $utilsFunctions
. $addsFunctions

#----------------------------------------------------------[Execution]----------------------------------------------------------

Start-Log -scriptName $scriptName -scriptVersion $scriptVersion -streamWriter $global:streamWriter
cls
Write-Host "================================================================================================"

# Prerequisites
Test-LocalAdminRights
if($adminFlag -eq $false){
    Write-Host "You have to launch this script with " -nonewline; Write-Host "local Administrator rights!" -f Red    
    $scriptPath = Split-Path $MyInvocation.InvocationName    
    $RWMC = $scriptPath + "\$scriptName.ps1"
    $ArgumentList = 'Start-Process -FilePath powershell.exe -ArgumentList \"-ExecutionPolicy Bypass -File "{0}"\" -Verb Runas' -f $RWMC;
    Start-Process -FilePath powershell.exe -ArgumentList $ArgumentList -Wait -NoNewWindow;    
    Stop-Script
                         }
else {
    Write-Host "Administrator rights detected." -nonewline; Write-Host "The script will continue." -f Green
     }

Write-Host "================================================================================================"

# Modify the Execution Policy
Write-Progress -Activity "Modify the Execution Policy" -status "Running..." -id 1
Set-ExecutionPolicy Bypass -Scope Process

# Set static IP address
Write-Progress -Activity "Set static IP address" -status "Running..." -id 1
$ipaddress = "192.168.10.11"
$ipprefix = "24"
$ipgw = "192.168.10.1"
$ipdns = "127.0.0.1"
$ipif = (Get-NetAdapter).ifIndex
New-NetIPAddress -IPAddress $ipaddress -PrefixLength $ipprefix `
				 -InterfaceIndex $ipif -DefaultGateway $ipgw
# Rename the NetAdapter
Get-NetAdapter -Name * | ? status -eq up | Rename-NetAdapter -NewName "LAN $ipgw"

# Rename the computer
$newname = "AD1FR1010"
Rename-Computer -NewName $newname -force

#install features
$addsTools = "RSAT-AD-Tools"
Add-WindowsFeature $addsTools

# Add informations to logfile
Get-Content env:computername >> $logPathName
Get-NetIPAddress | Where { $_.PrefixLength -eq 24 } >> $logPathName
Get-WindowsFeature | Where installed >> $logPathName




# Writing informations in the log file
Write-Progress -Activity "Write informations in the log file" -status "Running..." -id 1
End-Log -streamWriter $global:streamWriter
notepad $logPathName

# Restart the computer
Write-Progress -Activity "Server will restart in 15 seconds" -status "Running..." -id 1
Start-Sleep -s 15
Restart-Computer