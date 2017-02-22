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
$Form.TopMost = $false
$Form.Width = 460
$Form.Height = 305

$buttonStartRC = New-Object system.windows.Forms.Button
$buttonStartRC.Text = "Lancer contrôle à distance"
$buttonStartRC.Width = 142
$buttonStartRC.Height = 52
$buttonStartRC.Add_MouseClick({
    $Result = "$env:COMPUTERNAME,$env:USERNAME,$($textBoxUtilisateur.Text),$($textBoxComputerName.Text),$($textBoxTicket.Text)"
	$launchDate + ',' + $launchHour + ',' +$env:COMPUTERNAME + ',' + $env:USERNAME + ',' + $textBoxUtilisateur.Text + ',' + $textBoxComputerName.Text + ',' + $textBoxTicket.Text | Out-File -filepath $logPathName -Append -Encoding UTF8
    msra.exe /offerRA $textBoxComputerName.Text
})
$buttonStartRC.location = new-object system.drawing.point(223,157)
$buttonStartRC.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($buttonStartRC)

$textBoxUtilisateur = New-Object system.windows.Forms.TextBox
$textBoxUtilisateur.Text = "xxxN"
$textBoxUtilisateur.Width = 100
$textBoxUtilisateur.Height = 20
$textBoxUtilisateur.location = new-object system.drawing.point(101,151)
$textBoxUtilisateur.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxUtilisateur)

$textBoxComputerName = New-Object system.windows.Forms.TextBox
$textBoxComputerName.Text = "Uxnnnnnnnnn"
$textBoxComputerName.Width = 100
$textBoxComputerName.Height = 20
$textBoxComputerName.location = new-object system.drawing.point(101,189)
$textBoxComputerName.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($textBoxComputerName)

$LabelUtilisateur = New-Object system.windows.Forms.Label
$LabelUtilisateur.Text = "Utilisateur"
$LabelUtilisateur.AutoSize = $true
$LabelUtilisateur.Width = 25
$LabelUtilisateur.Height = 10
$LabelUtilisateur.location = new-object system.drawing.point(24,151)
$LabelUtilisateur.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($LabelUtilisateur)

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

$labelTicket = New-Object system.windows.Forms.Label
$labelTicket.Text = "Ticket"
$labelTicket.AutoSize = $true
$labelTicket.Width = 25
$labelTicket.Height = 10
$labelTicket.location = new-object system.drawing.point(22,230)
$labelTicket.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelTicket)

$labelDescription = New-Object system.windows.Forms.Label
$labelDescription.Text = "Merci de remplir les champs ci-dessous afin de prendre la main
sur un poste COMPANY.

Tous sont obligatoires."
$labelDescription.AutoSize = $true
$labelDescription.Width = 25
$labelDescription.Height = 10
$labelDescription.location = new-object system.drawing.point(10,20)
$labelDescription.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelDescription)

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
