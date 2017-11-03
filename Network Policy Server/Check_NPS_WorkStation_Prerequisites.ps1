#requires -version 2
<#
.SYNOPSIS
  Check NPS prerequisites on Workstation

.DESCRIPTION
  Check NPS prerequisites on Workstation, like AD membership, certificates presence and Wifi config

.INPUTS
  <None>

.OUTPUTS
  <None> 
   
.NOTES
  Version:        0.1
  Author:         ALBERT Jean-Marc
  Creation Date:  03/11/2017 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2017.11.03 - ALBERT Jean-Marc - Initial script development
                                                                  
.SOURCES
  <None>
 
.EXAMPLE
  <None>

#>

#region ---------------------------------------------------------[Initialisations]--------------------------------------------------------
Set-StrictMode -version Latest

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Modify Window width
Clear-Host
$h = Get-Host
$win = $h.ui.rawui.windowsize
$win.Width  = 68
$h.ui.rawui.set_windowsize($win)
#endregion

#region ----------------------------------------------------------[Declarations]----------------------------------------------------------
$ComputerName = $env:COMPUTERNAME
$DomainName = 'contoso.com'
$DomainMembership = Test-Domain $ComputerName | Select-String Domain
$TrustedRootCertificationAuthorityCertName = 'CN=contoso-rootCA, DC=contoso, DC=com'
$TargetedADGroup = "Computer Wifi"
$WifiCorrectAccessPointName = "*AccessPoint Name*"
$WifiActiveConnection = netsh wlan show interfaces
$WifiActiveConnectionSSID = $WifiActiveConnection | Select-String \sSSID
#endregion

#region -----------------------------------------------------------[Functions]------------------------------------------------------------
function Test-Domain {             
    [CmdletBinding()]             
    param (             
    [parameter(Position=0,            
        Mandatory=$true,            
        ValueFromPipeline=$true,             
        ValueFromPipelineByPropertyName=$true)]            
        [string]$computerName=$env:COMPUTERNAME             
    )             
    BEGIN{}
    PROCESS{Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName | Select Domain
    }
    END{}   
}
function Get-ComputerADGroupMembership {
    Get-ADComputer $ComputerName -Properties MemberOf
}
#endregion

#region -----------------------------------------------------------[Execution]------------------------------------------------------------
#Load PowerShell AD module before execution
$ImportModuleADid = (Import-Module ActiveDirectory).id
Wait-Process -Id $ImportModuleADid -Timeout 10


Write-Host "------------------------------------------------------------------`n
- Show actual computer name:                                     -`n
    Computer:                                       $ComputerName"
Write-Host "------------------------------------------------------------------`
- Check domain membership status:                                -"
If ($DomainMembership -like "*$DomainName*" ) {
   Write-Host '    AD Group membership (need AD PowerShell module):      OK' -ForegroundColor Green
}

Else {
   Write-Host '    AD Group membership (need AD PowerShell module):      KO' -ForegroundColor Red
}
Write-Host "------------------------------------------------------------------`
- Check authorized computer membership status:                   -"
If (Get-ComputerADGroupMembership | Where {$_.MemberOf -like "*$TargetedADGroup*"} ) {
   Write-Host '    Domain membership:                                    OK' -ForegroundColor Green
}

Else {
   Write-Host '    Domain membership:                                    KO' -ForegroundColor Red
}
 #Get-ComputerADGroupMembership | %{if ($_.MemberOf -like "*$TargetedADGroup*") {Write-Host "found"} 


Write-Host "------------------------------------------------------------------`
- Check workstation certificate status:                          -"
If (Get-ChildItem cert:\LocalMachine\My | Where Subject -like *$ComputerName* ) {
   Write-Host '    Computer Certificate:                                 Present' -ForegroundColor Green

}

Else {
   Write-Host '    Computer Certificate:                                 Missing' -ForegroundColor Red
}

Write-Host "------------------------------------------------------------------`
- Check trusted root certification authority certificate status: -"

If (Get-ChildItem cert:\LocalMachine\CA | Where Subject -eq $TrustedRootCertificationAuthorityCertName ) {
   Write-Host '    Trusted Root Certification Authority certificate:     Present' -ForegroundColor Green

}

Else {
   Write-Host '    Trusted Root Certification Authority certificate::    Missing'-ForegroundColor Red
}


Write-Host "------------------------------------------------------------------`
- Show actual Wifi connection:                                   -"
$WifiActiveConnection | Select-String SSID,Authentification,Chiffrement,'Mode de connexion',Signal
If (!$WifiActiveConnection) {
   Write-Host 'No Wifi connection is active!' -ForegroundColor Red

}

If ($WifiActiveConnectionSSID -like $WifiCorrectAccessPointName) {
   Write-Host 'You are on the right AccessPoint' -ForegroundColor Green
    If (($WifiActiveConnection | Select-String 'Mode de connexion') -like "*Profil*") {
        Write-Host 'You received the right GPO' -ForegroundColor Green
    }
    Else {
        Write-Host ''
        Write-Host 'You have not received the right GPO!' -ForegroundColor Red
    }
}
Else {
       Write-Host ''
       Write-Host 'You are not on the right AccessPoint!' -ForegroundColor Red
}

Write-Host "------------------------------------------------------------------"
Pause
#endregion
