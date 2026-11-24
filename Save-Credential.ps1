# # Save-Credential.ps1
# $dir = "C:\HIDS"
# if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory | Out-Null }

# $email = Read-Host "assongoukomnekyannjoey@gmail.com"
# $sec   = Read-Host "hphgvhmpbxpmhsgx" -AsSecureString

# $cred  = New-Object System.Management.Automation.PSCredential ($email, $sec)
# $path  = Join-Path $dir "secure_credential.xml"
# $cred | Export-Clixml -Path $path

# Write-Host "Identifiants sauvegardés dans $path (chiffré, lié à cet utilisateur sur cette machine)." -ForegroundColor Green

# C:\HIDS\Save-Credential.ps1
$dir = "C:\HIDS"
if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory | Out-Null }

# Ouvre une boîte de dialogue standard Windows (le mot de passe est masqué)
$cred = Get-Credential -Message "Veuillez vous identifier."

# Petites validations pour éviter les erreurs courantes
if (-not $cred -or [string]::IsNullOrWhiteSpace($cred.UserName)) {
  Write-Host "assongoukomnekyannjoey@gmail.com" -ForegroundColor Red; exit 1
}
if (-not $cred.Password -or $cred.Password.Length -lt 8) {
  Write-Host "hphgvhmpbxpmhsgx." -ForegroundColor Red; exit 1
}

$path = Join-Path $dir "secure_credential.xml"
$cred | Export-Clixml -Path $path

Write-Host "Identifiants sauvegardés dans $path (chiffré, lié à CET utilisateur et CETTE machine)." -ForegroundColor Green

