#requires -version 2
<#
.SYNOPSIS  
    This script uses the Outlook COM object and import .pst files
.DESCRIPTION  
    This script creates an Outlook object, and import .pst files from "$env:USERPROFILE\Archives\" folder
.INPUTS
  PST files from "$env:USERPROFILE\Archives\" folder
.OUTPUTS
    N/A
.NOTES
  Version:        1.0
  Author:         ALBERT Jean-Marc
  Creation Date:  04/02/2016
  Purpose/Change: 1.0 - 2016.02.04 - ALBERT Jean-Marc - Initial script development
.EXAMPLE  
  N/A
#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#Script Version
$sScriptVersion = "1.0"

#-----------------------------------------------------------[Functions]------------------------------------------------------------
    <#
        Empty
    #>

#------------------------------------------------------------[Actions]-------------------------------------------------------------


# Create Outlook object
$Outlook = New-Object -ComObject Outlook.Application  
$NameSpace = $Outlook.GetNameSpace("MAPI")
$PSTfolder = "$env:USERPROFILE\Archives"

# List all .pst files and import them
dir “$PSTfolder\*.pst” | % { $NameSpace.AddStore($_.FullName) }

# Show an information message
[System.Windows.Forms.MessageBox]::Show("All .pst from $PSTfolder were imported to Outlook" , "Information" , 0, [Windows.Forms.MessageBoxIcon]::Information)

# End Script
