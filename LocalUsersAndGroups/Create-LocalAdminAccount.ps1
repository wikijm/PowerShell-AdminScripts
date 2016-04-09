#requires -version 2
<#
.SYNOPSIS
  Create a local admin user, with local admin group adaptation

.DESCRIPTION
  Create a local admin user, with local admin group adaptation

.INPUTS
  <None>

.OUTPUTS
  <None>
     
   
.NOTES
  Version:        0.1
  Author:         ALBERT Jean-Marc
  Creation Date:  08/04/2016 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2016.04.08 - ALBERT Jean-Marc - Initial script development
                    
                                                  
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

$scriptVersion = "0.1"

$userName = "test"
$password = "P@ssw0rd!!"
$description = 'TYPE_A_DESCRIPTION'

$computer = $Env:COMPUTERNAME

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function LocalGroupExist ($groupName) { 
 return [ADSI]::Exists("WinNT://$Env:COMPUTERNAME/$groupName,group")
}

function LocalUserExist ($userName) {
  # Local user account creation:
  $colUsers = ($Computer.psbase.children | Where-Object {$_.psBase.schemaClassName -eq "User"} | Select-Object -expand Name)
  $userFound = $colUsers -contains $userName
  return $userFound
}

function CreateLocalUser ($userName,$password) {
 $userExist = LocalUserExist($userName)
 
 if($userExist -eq $false)
 {
  $User = $Computer.Create("User", $userName)
  $User.SetPassword($password)
  $User.SetInfo()
  $User.FullName = $userName
  $User.SetInfo()
  $user.description = $description
  $user.SetInfo()
  $User.UserFlags = 64 + 65536 # ADS_UF_PASSWD_CANT_CHANGE + ADS_UF_DONT_EXPIRE_PASSWD
  $User.SetInfo()
 }
 else {
  "User : $userName already exist."
 }
}

function AddUserToGroup ($groupName, $userName) {
 $group = [ADSI]"WinNT://$Env:COMPUTERNAME/$groupName"
 $user = [ADSI]"WinNT://$Env:COMPUTERNAME/$userName"
 $memberExist = CheckGroupMember $groupName $userName
 if($memberExist -eq $false)
 {
  $group = [ADSI]"WinNT://$Env:COMPUTERNAME/$groupName"  
  $user = [ADSI]"WinNT://$Env:COMPUTERNAME/$userName" 
  $group.Add($user.Path)
 }
}

#----------------------------------------------------------[Execution]----------------------------------------------------------

#Get local admin group
 $LocalAdminGroup = (get-wmiobject win32_group | Where-Object {$_.Name -Like "Administr*"}).Name

 #Create $userName local account
 CreateLocalUser $userName $password
 
 #Check $userName local account creation
  $IsAccountExist = LocalUserExist($userName)
  if($IsAccountExist -eq $true)
  {
   "$userName now exist"
  }
  else
  {
   "/!\ Error: $userName don't exist /!\"
  }
 
 
 #Add $userName local account to local admin group
 AddUserToGroup $LocalAdminGroup