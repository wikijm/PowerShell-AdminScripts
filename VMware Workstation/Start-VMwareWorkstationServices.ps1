#requires -version 2
<#
.SYNOPSIS
  Show a menu to easily check and start VMware Workstation services

.DESCRIPTION
  Show a menu to easily check and start VMware Workstation services. Type a number from 0 to 4 to launch and action.

.INPUTS
  <None>

.OUTPUTS
  <None>
   
   
.NOTES
  Version:        0.1
  Author:         ALBERT Jean-Marc
  Creation Date:  23/04/2016 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2016.04.23 - ALBERT Jean-Marc - Initial script development
                    
                                                  
.SOURCES
  <None>
  
  
.EXAMPLE
  <None>

#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Set-StrictMode -version Latest

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$ScriptName = [System.IO.Path]::GetFileName($ScriptFile)
$ScriptPath = Split-Path $MyInvocation.InvocationName    
$ScriptFilePath = $ScriptPath + "\$ScriptName.ps1"
$ScriptVersion = "0.1"

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Show-Menu
{
    param (
        [string]$selection
    )

    Clear-Host
    Write-Host "================ $Title ================"
    
    Write-Host "1: List VMmware services (with status)"
    Write-Host "2: Start System services (Authorization and USB Arbitration)"
    Write-Host "3: Start Network services (DHCP and NAT)"
    Write-Host "4: Start Remote Administration service (Workstation Server)"
    Write-Host "0: Exit"
}

function Get-VMwareServicesStatus
{ 
 Get-Service | Where-Object {$_.DisplayName -Like "VMware*"} | Sort-Object Status,Name,DisplayName | %{
    if ( $_.Status -eq "Stopped")
        {
         Write-Host $_.Status "|" $_.Name "|" $_.DisplayName  -ForegroundColor red
        }
    elseif ( $_.Status -eq "Running")
        {
        Write-Host $_.Status "|" $_.Name "|" $_.DisplayName  -ForegroundColor green}
        }
}

#----------------------------------------------------------[Execution]----------------------------------------------------------

#Show menu
do
 {
     Show-Menu
     $selection = Read-Host "Please make a selection"
     switch ($selection)
     {
         '1' {
             Clear-Host
             Get-VMwareServicesStatus
             }
         '2' {
             Start-Service "VMAuthdService"
             Start-Service "VMUSBArbService"
             Get-VMwareServicesStatus
             }
         '3' {
             Start-Service "VMnetDHCP"
             Start-Service "VMware NAT Service"
             Get-VMwareServicesStatus
             }
         '4' {
             Start-Service "VMwareHostd"
             Get-VMwareServicesStatus
             }
         '0' {
             Exit
             }
     }
     pause
 }
 until ($selection -ceq '0')

Exit