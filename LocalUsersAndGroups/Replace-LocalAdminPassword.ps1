#requires -version 2
<#
.SYNOPSIS
  Replace password of local admin account

.DESCRIPTION
  Get local account (user object) with $childObjectSID ends with "-500" (Admin account) then replace password

.INPUTS
  <None>

.OUTPUTS
  <None>
   
   
.NOTES
  Version:        0.1
  Author:         ALBERT Jean-Marc
  Creation Date:  29/04/2016 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2016.04.29 - ALBERT Jean-Marc - Initial script development
                    
                                                  
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

$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "0.1"

$ComputerName = $env:COMPUTERNAME
$Computer = [ADSI] "WinNT://$ComputerName,Computer"
$DecodedPassword = "Str0n6P@ssw0rd!!"

#-----------------------------------------------------------[Execution]------------------------------------------------------------
Write-Host "======================================================="
Write-Host {}{}{}{}{}{}{}{}{}{}{}{}"$scriptName"
Write-Host "======================================================="


#Get local admin account name
Write-Progress -Activity "Get local admin account name" -status "Running..." -id 1 
foreach ( $childObject in $Computer.Children ) {
  # Skip objects that are not users.
  if ( $childObject.Class -ne "User" ) {
    continue
  }  
  $type = "System.Security.Principal.SecurityIdentifier"
  # BEGIN CALLOUT A
  $childObjectSID = new-object $type($childObject.objectSid[0],0)
  # END CALLOUT A
  if ( $childObjectSID.Value.EndsWith("-500") ) {
    $LocalAdminAccount = $($childObject.Name[0])
    break
  }   
}

#Show local Admin account
Write-Progress -Activity "Show local Admin account" -status "Running..." -id 1 
Write-Host -ForegroundColor Green "Local user account: $LocalAdminAccount"

#Define new password to local Admin account
Write-Progress -Activity "Define new password to local Admin account" -status "Running..." -id 1 
$Computer
$User = [adsi]"WinNT://$ComputerName/$LocalAdminAccount,user"
$User.SetPassword($DecodedPassword)
$User.SetInfo()
Write-Host -ForegroundColor Green "Password changed successfully"