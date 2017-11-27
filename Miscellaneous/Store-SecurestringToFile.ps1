param(
    [Parameter(Mandatory=$true)]
    [string]$Path,
    [Parameter(Mandatory=$true)]
    [string]$password
    )

    # Convert the password to a secure string
    $SecurePassword = $password | ConvertTo-SecureString -AsPlainText -Force
    
    # Store the credential in the path
    $SecurePassword | ConvertFrom-SecureString | Out-File $Path
    
    # Write What we did
    Write-Host "Wrote password to $path"    
    
<# 
.SYNOPSIS
Stores a password in a file on the local computer for retrevial by scripts.

.DESCRIPTION
Used for securely storing a password on a machine for use with automated scripts.

Takes a password and encrypts it using the local account, then stores that password in a file you specify.
Only the account that creates the output file can decrypt the password stored in the file.

.PARAMETER Path
Path and file name for the password file that will be created.

.PARAMETER Password
Plain text version of password.

.OUTPUTS
File Specified in Path variable will contain an encrypted version of the password.

.EXAMPLE
.\Store-SecurestringToFile.ps1 -Path c:\scripts\scriptname.key -Password "Password123"

Puts the encrypted version of Password123 into the c:\scripts\scriptname.key file
#>
