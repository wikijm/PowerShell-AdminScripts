Function CreateShortcut
{
    [CmdletBinding()]
    param (	
	    [parameter(Mandatory=$true)]
	    [ValidateScript( {[IO.File]::Exists($_)} )]
	    [System.IO.FileInfo] $Target,
	
	    [ValidateScript( {[IO.Directory]::Exists($_)} )]
	    [System.IO.DirectoryInfo] $OutputDirectory,
	
	    [string] $Name,
	    [string] $Description,
	
	    [string] $Arguments,
	    [System.IO.DirectoryInfo] $WorkingDirectory,
	
	    [string] $HotKey,
	    [int] $WindowStyle = 1,
	    [string] $IconLocation,
	    [switch] $Elevated
    )

    try {
	    #region Create Shortcut
	    if ($Name) {
		    [System.IO.FileInfo] $LinkFileName = [System.IO.Path]::ChangeExtension($Name, "lnk")
	    } else {
		    [System.IO.FileInfo] $LinkFileName = [System.IO.Path]::ChangeExtension($Target.Name, "lnk")
	    }
	
	    if ($OutputDirectory) {
		    [System.IO.FileInfo] $LinkFile = [IO.Path]::Combine($OutputDirectory, $LinkFileName)
	    } else {
		    [System.IO.FileInfo] $LinkFile = [IO.Path]::Combine($Target.Directory, $LinkFileName)
	    }

       
	    $wshshell = New-Object -ComObject WScript.Shell
	    $shortCut = $wshShell.CreateShortCut($LinkFile) 
	    $shortCut.TargetPath = $Target
	    $shortCut.WindowStyle = $WindowStyle
	    $shortCut.Description = $Description
	    $shortCut.WorkingDirectory = $WorkingDirectory
	    $shortCut.HotKey = $HotKey
	    $shortCut.Arguments = $Arguments
	    if ($IconLocation) {
		    $shortCut.IconLocation = $IconLocation
	    }
	    $shortCut.Save()
	    #endregion

	    #region Elevation Flag
	    if ($Elevated) {
		    $tempFileName = [IO.Path]::GetRandomFileName()
		    $tempFile = [IO.FileInfo][IO.Path]::Combine($LinkFile.Directory, $tempFileName)
		
		    $writer = new-object System.IO.FileStream $tempFile, ([System.IO.FileMode]::Create)
		    $reader = $LinkFile.OpenRead()
		
		    while ($reader.Position -lt $reader.Length)
		    {		
			    $byte = $reader.ReadByte()
			    if ($reader.Position -eq 22) {
				    $byte = 34
			    }
			    $writer.WriteByte($byte)
		    }
		
		    $reader.Close()
		    $writer.Close()
		
		    $LinkFile.Delete()
		
		    Rename-Item -Path $tempFile -NewName $LinkFile.Name
	    }
	    #endregion
    } catch {
	    Write-Error "Failed to create shortcut. The error was '$_'."
	    return $null
    }
    return $LinkFile
}

CreateShortcut -name "Notepad++ Admin" -Target "${env:ProgramFiles(x86)}\Notepad++\notepad++.exe" -OutputDirectory "C:\Users\Public\Desktop" -Elevated True
