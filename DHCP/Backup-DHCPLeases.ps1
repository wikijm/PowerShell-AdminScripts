#requires -version 2 -RunAsAdministrator
<#
    .SYNOPSIS
                    List all DHCP lease Objects coming from a targeted DHCP server, and export results into a CSV

    .DESCRIPTION
                    List all DHCP lease Objects coming from a targeted DHCP server, and export results into a CSV

    .INPUTS
                    DHCP server name as an argument

    .OUTPUTS
                    Create transcript log file similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD]_[HHhMMmSSs].log
                    Create a list of DHCP lease Objects, similar to $ScriptDir\[SCRIPTNAME]_DHCPServer_Reservations_[YYYY_MM_DD]_[HHhMMmSSs].csv
   
   
    .NOTES
    Version:        1.0
    Author:         ALBERT Jean-Marc
    Creation Date:  13/02/2019
    Purpose/Change: 1.0 - 2019.02.13 - ALBERT Jean-Marc - Initial script development
                                    
                                                  
    .SOURCES
                    https://blogs.technet.microsoft.com/leesteve/2017/06/09/ps-without-bs-extracting-dhcp-reservations-to-a-csv/
  
    .EXAMPLE
                    <None>
#>
 
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
 
#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

# Powershell DHCPServer module importation
Import-Module DHCPServer

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Write script directory path on "ScriptDir" variable
$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

#Log file creation, similar to $ScriptDir\[SCRIPTNAME]_[YYYY_MM_DD].log
$SystemTime = Get-Date -uformat %Hh%Mm%Ss
$SystemDate = Get-Date -uformat %Y_%m_%d
$ScriptLogFile = "$ScriptDir\$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))" + "_" + $DHCPServer + "_" + $SystemDate + "_" + $SystemTime + ".log"

#Define CSV file export of DHCP leases objects
$CSVExport = "$ScriptDir\$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Definition))" + "_" + $DHCPServer + "_" + "Reservations" + "_" + $SystemDate + "_" + $SystemTime + ".csv"


#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Stop-TranscriptOnLog {   
  Stop-Transcript
    # We put in the transcript the line breaks necessary for Notepad
    [string]::Join("`r`n",(Get-Content $ScriptLogFile)) | Out-File $ScriptLogFile
                              }

#------------------------------------------------------------[Actions]-------------------------------------------------------------

 # Start of log completion
   Start-Transcript $ScriptLogFile | Out-Null

 
   # Catch DHCP Server name
   $DHCPServer = $env:COMPUTERNAME
   <#
   param ([string]$DHCPServer = $args[0])
   
   If ($DHCPServer -eq $null) {
        $DHCPServer = $env:COMPUTERNAME
   }#>

  # Get DHCP scope list on targeted server(s)  
    $DHCPServerScope = Get-DHCPServerV4Scope -ComputerName $DHCPServer

  # For each scope, list and export every leases to a .CSV file
    $DHCPServerScope | ForEach {
                                 Get-DHCPServerv4Lease -ScopeID $_.ScopeID | where {$_.AddressState -like '*Reservation'}
                       } | Select-Object ScopeId,IPAddress,HostName,ClientID,AddressState | Export-Csv $CSVExport -Delimiter ";" -NoTypeInformation
 
        # MessageBox who inform of the end of the process
            Write-Information "Process done"
 
        # Stop the log transcript
            Stop-TranscriptOnLog

        # Open the log file | Debug purpose only, you may leave it as a comment
            #Invoke-Item $ScriptLogFile
