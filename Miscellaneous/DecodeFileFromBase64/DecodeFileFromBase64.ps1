#requires -version 2
<#
.SYNOPSIS
  Decode selected file from Base64, and export result to original (decoded) file

.DESCRIPTION
  Decode selected file (with dialog) from Base64, and export result to original (decoded) file (without "-encoded.txt" suffix)

.INPUTS
  File selected by user with "Select file" dialog

.OUTPUTS
  File selected by user with "Select file" dialog without "-encoded.txt" suffix
     
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

# Import file to decode
[System.Windows.Forms.MessageBox]::Show(
"
Select on this window the file that you want to decode from Base64
", "Select file", 0, [Windows.Forms.MessageBoxIcon]::Question)

$FileToDecode = Select-FileDialog -Title "Select file"
$Content = [System.Convert]::FromBase64String($Base64)
$DecodedFile = $FileToDecode.Replace("-encoded.txt","")
Set-Content -Path $DecodedFile -Value $Content -Encoding Byte

# Inform user to the end of the process, and let the possibility to open the output
$OUTPUT = [System.Windows.Forms.MessageBox]::Show("Decoded version was create well. Would you like to open it?", "Ended Base64 decoding", 4, [Windows.Forms.MessageBoxIcon]::Question)
if ($OUTPUT -eq "YES")
    {
    notepad $DecodedFile
	}

else
{
	[System.Windows.Forms.MessageBox]::Show("End of the script.", "Ended Base64 decoding process", 0, [Windows.Forms.MessageBoxIcon]::Information)
}