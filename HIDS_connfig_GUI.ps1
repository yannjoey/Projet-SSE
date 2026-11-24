# C:\HIDS\HIDS_Config_GUI.ps1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

# --- I/O ---
$workDir = 'C:\HIDS'
if (-not (Test-Path $workDir)) { New-Item $workDir -ItemType Directory | Out-Null }
$configPath = Join-Path $workDir 'config.json'

function Load-Config {
    if (Test-Path $configPath) {
        try { return Get-Content $configPath -Raw | ConvertFrom-Json } catch { }
    }
    [pscustomobject]@{
        FilesToMonitor = @(
            'C:\Users\Administrateur\Documents\HIDS',
            'C:\Windows\System32\drivers\etc\hosts'
        )
        AlertEmail = 'yann-joey.assongou-komnek@efrei.net'
        UseOutlook = $true
    }
}

function Save-Config($cfg) {
    ($cfg | ConvertTo-Json -Depth 5) | Set-Content -Encoding UTF8 $configPath
}

function Send-OutlookMail {
    param([string]$to,[string]$subject,[string]$body)
    try {
        $out = New-Object -ComObject Outlook.Application
        $mail = $out.CreateItem(0)
        $mail.To = $to
        $mail.Subject = $subject
        $mail.Body = $body
        $mail.Send()
        return $true
    } catch {
        [Windows.Forms.MessageBox]::Show("Échec envoi Outlook :`r`n$($_.Exception.Message)","Outlook",0,16) | Out-Null
        return $false
    }
}

# --- UI ---
$cfg = Load-Config

$form = New-Object Windows.Forms.Form
$form.Text = 'Configuration HIDS'
$form.Size = New-Object Drawing.Size(720,520)
$form.StartPosition = 'CenterScreen'
$form.MaximizeBox = $false

$lblPaths = New-Object Windows.Forms.Label
$lblPaths.Text = "Fichiers / dossiers surveillés :"
$lblPaths.Location = '12,12'
$lblPaths.AutoSize = $true

$list = New-Object Windows.Forms.ListBox
$list.Location = '15,35'
$list.Size = New-Object Drawing.Size(560,360)
$list.HorizontalScrollbar = $true
$list.SelectionMode = 'MultiExtended'
$list.Items.AddRange($cfg.FilesToMonitor)

$btnAddFile = New-Object Windows.Forms.Button
$btnAddFile.Text = 'Ajouter Fichier'
$btnAddFile.Location = '585,35';  $btnAddFile.Width = 110

$btnAddFolder = New-Object Windows.Forms.Button
$btnAddFolder.Text = 'Ajouter Dossier'
$btnAddFolder.Location = '585,70'; $btnAddFolder.Width = 110

$btnRemove = New-Object Windows.Forms.Button
$btnRemove.Text = 'Supprimer'
$btnRemove.Location = '585,105';  $btnRemove.Width = 110

$lblMail = New-Object Windows.Forms.Label
$lblMail.Text = 'Adresse de destination (alerte) :'
$lblMail.Location = '12,410'; $lblMail.AutoSize = $true

$txtMail = New-Object Windows.Forms.TextBox
$txtMail.Location = '220,407'
$txtMail.Size = New-Object Drawing.Size(355,24)
$txtMail.Text = $cfg.AlertEmail

$chkOutlook = New-Object Windows.Forms.CheckBox
$chkOutlook.Text = 'Envoyer via Outlook'
$chkOutlook.Location = '585,140'
$chkOutlook.Checked = $cfg.UseOutlook

$btnTest = New-Object Windows.Forms.Button
$btnTest.Text = 'Tester envoi'
$btnTest.Location = '585,400'; $btnTest.Width = 110

$btnSave = New-Object Windows.Forms.Button
$btnSave.Text = 'Sauvegarder'
$btnSave.Location = '585,435'; $btnSave.Width = 110

$btnClose = New-Object Windows.Forms.Button
$btnClose.Text = 'Fermer'
$btnClose.Location = '470,440'; $btnClose.Width = 105

# --- Events ---
$btnAddFile.Add_Click({
    $dlg = New-Object Windows.Forms.OpenFileDialog
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq 'OK') {
        foreach($p in $dlg.FileNames){ if(-not $list.Items.Contains($p)){ [void]$list.Items.Add($p) } }
    }
})

$btnAddFolder.Add_Click({
    $dlg = New-Object Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq 'OK') {
        $p = $dlg.SelectedPath
        if(-not $list.Items.Contains($p)){ [void]$list.Items.Add($p) }
    }
})

$btnRemove.Add_Click({
    $sel = @($list.SelectedItems)
    foreach($i in $sel){ $list.Items.Remove($i) }
})

$btnTest.Add_Click({
    if (-not $chkOutlook.Checked) {
        [Windows.Forms.MessageBox]::Show("Coche 'Envoyer via Outlook' d'abord.","Test",0,48) | Out-Null; return
    }
    $to = $txtMail.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($to)) {
        [Windows.Forms.MessageBox]::Show("Renseigne l'adresse destinataire.","Test",0,48) | Out-Null; return
    }
    $ok = Send-OutlookMail -to $to -subject "[HIDS] Test configuration" -body "Test à $(Get-Date)"
    if ($ok) { [Windows.Forms.MessageBox]::Show("Mail envoyé à $to via Outlook.","Test",0,64) | Out-Null }
})

$btnSave.Add_Click({
    $cfg = [pscustomobject]@{
        FilesToMonitor = @($list.Items)
        AlertEmail     = $txtMail.Text.Trim()
        UseOutlook     = $chkOutlook.Checked
    }
    Save-Config $cfg
    [Windows.Forms.MessageBox]::Show("Configuration sauvegardée dans `n$configPath","HIDS",0,64) | Out-Null
})

$btnClose.Add_Click({ $form.Close() })

$form.Controls.AddRange(@($lblPaths,$list,$btnAddFile,$btnAddFolder,$btnRemove,$lblMail,$txtMail,$chkOutlook,$btnTest,$btnSave,$btnClose))
[void]$form.ShowDialog()
