$launchDate = get-date -f "dd/MM/yyyy"
$launchHour = get-date -f "HH:mm:ss"
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$logFileName = "pmad.log"
$logPathName = "P:\Exploitation\Suivi\PMAD\$logFileName"

Add-Type -AssemblyName System.Windows.Forms

$Form = New-Object system.Windows.Forms.Form
$Form.Text = "Contrôle à distance"
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.MinimizeBox = $false
$Form.TopMost = $true
$Form.Width = 460
$Form.Height = 305

$StartRC = New-Object system.windows.Forms.Button
$StartRC.Text = "Lancer contrôle à distance"
$StartRC.Width = 142
$StartRC.Height = 52
$StartRC.Add_MouseClick({
    $Result = "$env:COMPUTERNAME,$env:USERNAME,$($textBoxutilisateur.Text),$($textBoxComputerName.Text),$($textBoxTicket.Text)"
	$launchDate + ',' + $launchHour + ',' +$env:COMPUTERNAME + ',' + $env:USERNAME + ',' + $textBoxutilisateur.Text + ',' + $textBoxComputerName.Text + ',' + $textBoxTicket.Text | Out-File -filepath $logPathName -Append -Encoding UTF8
    msra.exe /offerRA $textBoxComputerName.Text
})
$StartRC.location = new-object system.drawing.point(223,157)
$StartRC.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($StartRC)

$textBoxutilisateur = New-Object system.windows.Forms.TextBox
$textBoxutilisateur.Text = "xxxN"
$textBoxutilisateur.Width = 100
$textBoxutilisateur.Height = 20
$textBoxutilisateur.location = new-object system.drawing.point(101,151)
$textBoxutilisateur.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxutilisateur)

$textBoxComputerName = New-Object system.windows.Forms.TextBox
$textBoxComputerName.Text = "Uxnnnnnnnnn"
$textBoxComputerName.Width = 100
$textBoxComputerName.Height = 20
$textBoxComputerName.location = new-object system.drawing.point(101,189)
$textBoxComputerName.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxComputerName)

$Utilisateur = New-Object system.windows.Forms.Label
$Utilisateur.Text = "Utilisateur"
$Utilisateur.AutoSize = $true
$Utilisateur.Width = 25
$Utilisateur.Height = 10
$Utilisateur.location = new-object system.drawing.point(24,151)
$Utilisateur.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($Utilisateur)

$LabelComputerName = New-Object system.windows.Forms.Label
$LabelComputerName.Text = "Poste"
$LabelComputerName.AutoSize = $true
$LabelComputerName.Width = 25
$LabelComputerName.Height = 10
$LabelComputerName.location = new-object system.drawing.point(23,188)
$LabelComputerName.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($LabelComputerName)

$textBoxTicket = New-Object system.windows.Forms.TextBox
$textBoxTicket.Width = 100
$textBoxTicket.Height = 20
$textBoxTicket.location = new-object system.drawing.point(98,230)
$textBoxTicket.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxTicket)

$label15 = New-Object system.windows.Forms.Label
$label15.Text = "Ticket"
$label15.AutoSize = $true
$label15.Width = 25
$label15.Height = 10
$label15.location = new-object system.drawing.point(22,230)
$label15.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($label15)

$Description = New-Object system.windows.Forms.Label
$Description.Text = "Merci de remplir les champs ci-dessous afin de prendre la main
sur un poste MIEL.

Tous sont obligatoires."
$Description.AutoSize = $true
$Description.Width = 25
$Description.Height = 10
$Description.location = new-object system.drawing.point(10,20)
$Description.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($Description)

$buttonLog = New-Object system.windows.Forms.Button
$buttonLog.Text = "Log"
$buttonLog.Width = 60
$buttonLog.Height = 30
$buttonLog.location = new-object system.drawing.point(264,214)
$buttonLog.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($buttonLog)
$buttonLog.Add_MouseClick({
     Invoke-Item $logPathName
})


[void]$Form.ShowDialog()
$Form.Dispose()
