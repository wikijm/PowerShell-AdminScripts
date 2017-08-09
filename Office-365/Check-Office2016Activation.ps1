$launchdate = Get-Date -f "yyyy.MM.dd-HHhmmmss"
$Office2016FolderPath = 'C:\Program Files\Microsoft Office\Office16'
$CookieFile = 'C:\Users\' + $env:username + '\AppData\Local\Microsoft\Office\16.0\office365ForBusiness.txt'
$Office365LicenceStatusResult = 'C:\' + $env:HOMEPATH + '\office365forbusinessactstat.txt'


function Check-Office365ForBusinessActivation () {
    C:\Windows\System32\cscript.exe 'C:\Program Files\Microsoft Office\Office16\OSPP.VBS' /dstatus | Out-File -Force $Office365LicenceStatusResult
    $ActivationStatus = $(get-content $Office365LicenceStatusResult -ReadCount 24 | foreach { $_ -match "LICENSE STATUS" })
                        
    $Var = "Office Activated "
        if ($ActivationStatus -match "LICENSE STATUS:  ---LICENSED--- ") {
            $Var = $Var + "OK "
        }
 
    else {
        $Var = $Var + "Bad "
         }
 
    if ($Var -like "*Bad*") {
        echo "Office Not Activated"
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show("La mise à jour de la suite Office vient d'être effectuée.
        Afin de l'activer, merci de suivre la procédure PDF qui s'ouvrira lorsque vous cliquerez sur le bouton 'OK'.",
        "Office 2016 - Activation",
        0,
        [Windows.Forms.MessageBoxIcon]::Warning)

        Start-Process 'X:\SomewhereOverTheRainbow\1stLaunch.pdf' -WindowStyle Maximized -Verb Open
    }
    else
    {
        echo "Office Activated"
        New-Item $CookieFile -ItemType File
        Set-Content -Value $launchdate -Path $CookieFile 
    }
}

if(!(Test-Path $CookieFile)) {
    Check-Office365ForBusinessActivation
} 
