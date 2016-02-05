$GroupName = ""
Get-ADGroupMember -Identity $GroupName | Select name,objectClass,SamAccountName | Export-CSV C:\exportfile.csv -NoType -Force

#requires -version 2
<#
.SYNOPSIS  
    This script get list of users on an AD Group and export it to CSV file
.DESCRIPTION  
    This script get list of users on an AD Group and export it to CSV file with name, objectClass and SamAccountName
.INPUTS
  AD Group name
.OUTPUTS
    CSV file widefined with $CSVExport
.NOTES
  Version:        1.0
  Author:         ALBERT Jean-Marc
  Creation Date:  05/02/2016
  Purpose/Change: 1.0 - 2016.02.05 - ALBERT Jean-Marc - Initial script development
.EXAMPLE  
  N/A
#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#Script Version
$sScriptVersion = "1.0"

$ADGroupName = ""
$CSVExport = ""

#-----------------------------------------------------------[Functions]------------------------------------------------------------
    <#
        Empty
    #>

#------------------------------------------------------------[Actions]-------------------------------------------------------------

Get-ADGroupMember -Identity $ADGroupName | Select name,objectClass,SamAccountName | Export-CSV $CSVExport -NoType -Force

# Show an information message
[System.Windows.Forms.MessageBox]::Show("All .pst from $PSTfolder were imported to Outlook" , "Information" , 0, [Windows.Forms.MessageBoxIcon]::Information)

# End Script