#requires -version 2
<#
.SYNOPSIS
  Encode selected file to Base64, and export result to "SELECTEDFILE-encoded.txt"

.DESCRIPTION
  Encode selected file (with dialog) to Base64, and export the result to "SELECTEDFILE-encoded.txt"

.INPUTS
  File selected by user with "Select file" dialog

.OUTPUTS
  File selected by user with "Select file" dialog + "-encoded.txt" suffix
     
.NOTES
  Version:        0.1
  Author:         ALBERT Jean-Marc
  Creation Date:  17/03/2016 (DD/MM/YYYY)
  Purpose/Change: 1.0 - 2016.03.17 - ALBERT Jean-Marc - Initial script development
                    
                                                  
.SOURCES
  <None>
  
  
.EXAMPLE
  <None>

#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Set-StrictMode -version Latest

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"
$scriptFile = $MyInvocation.MyCommand.Definition


#----------------------------------------------------------[Declarations]----------------------------------------------------------

$scriptName = [System.IO.Path]::GetFileName($scriptFile)
$scriptVersion = "0.1"


#-----------------------------------------------------------[Functions]------------------------------------------------------------

function Select-FileDialog
 {
    param([string]$Title,[string]$Filter="All files *.*|*.*")
	[System.Reflection.Assembly]::LoadWithPartialName( 'System.Windows.Forms' ) | Out-Null
	$fileDialogBox = New-Object Windows.Forms.OpenFileDialog
	$fileDialogBox.ShowHelp = $false
	$fileDialogBox.initialDirectory = $ScriptDir
	$fileDialogBox.filter = $Filter
    $fileDialogBox.Title = $Title
	$Show = $fileDialogBox.ShowDialog( )

        If ($Show -eq "OK")
            {
                Return $fileDialogBox.FileName
            }
        Else
            {
                Write-Error "Canceled operation"
		          [System.Windows.Forms.MessageBox]::Show("Script is not able to continue. Operation stopped." , "Operation canceled" , 0, [Windows.Forms.MessageBoxIcon]::Error)
                Stop-TranscriptOnLog
		        Exit
            }

 }


#----------------------------------------------------------[Execution]----------------------------------------------------------

cls

# Import file to encode
[System.Windows.Forms.MessageBox]::Show(
"
Select on this window the file that you want to encode to Base64
", "Select file", 0, [Windows.Forms.MessageBoxIcon]::Question)

$FileToEncode = Select-FileDialog -Title "Select file"
$Content = Get-Content -Path $FileToEncode -Encoding Byte
$Base64 = [System.Convert]::ToBase64String($Content)
$EncodedOutput = $FileToEncode + "-encoded.txt"
$Base64 | Out-File $EncodedOutput

# Inform user to the end of the process, and let the possibility to open the output
$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Encoded version was create well. Would you like to open it?", "Ended Base64 encoding", 4, [Windows.Forms.MessageBoxIcon]::Question)
if ($OUTPUT -eq "YES")
    {
    notepad $EncodedOutput
	}

else
{
	[System.Windows.Forms.MessageBox]::Show("End of the script.", "Ended Base64 encoding process", 0, [Windows.Forms.MessageBoxIcon]::Information)
}