[CmdletBinding()]
param(
  [switch]$Init,              # crée/écrase la baseline puis s'arrête
  [switch]$UpdateBaseline,    # accepte l’état courant comme nouvelle baseline
  [string[]]$Targets = @(
    "C:\Users\hpspectre",     # dossier
    "C:\Users\Public\Documents\texte.txt",        # fichier
    "C:\Windows\System32\drivers\etc\hosts" ,
    "C:\Users\Public\projetsse.txt"     # fichier
  ),
  [string]$MailTo = "yann-joey.assongou-komnek@efrei.net" # destinataire Outlook
)

# --- Vérif admin ---
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Host "Lancer PowerShell en administrateur." -ForegroundColor Red; exit 1
}

# --- Dossier de travail ---
$WorkDir = "C:\HIDS"
if (-not (Test-Path $WorkDir)) { New-Item $WorkDir -ItemType Directory | Out-Null }
$BaselineFile = Join-Path $WorkDir "baseline_local.xml"
$LogFile      = Join-Path $WorkDir "hids.log"

# === Utilitaires ===
function Expand-Targets {
  param([string[]]$Paths)
  $files = @()
  foreach ($p in $Paths) {
    if (Test-Path $p -PathType Container) {
      $files += Get-ChildItem -Path $p -Recurse -File -ErrorAction SilentlyContinue |
                Select-Object -Expand FullName
    } elseif (Test-Path $p -PathType Leaf) {
      $files += (Get-Item $p -ErrorAction SilentlyContinue).FullName
    } else { $files += $p } # on garde pour signaler INTROUVABLE
  }
  $files | Sort-Object -Unique
}

function Get-FileHashes {
  param([string[]]$Files)
  $map=@{}
  foreach ($f in $Files) {
    if (Test-Path $f -PathType Leaf) {
      try { $map[$f] = (Get-FileHash -Path $f -Algorithm SHA256 -ErrorAction Stop).Hash }
      catch { $map[$f] = "ERREUR_HASH" }
    } else { $map[$f] = "INTROUVABLE" }
  }
  $map
}

function Save-Baseline($Files) { (Get-FileHashes -Files $Files) | Export-Clixml $BaselineFile }

# --- Alerte via Outlook (MAPI) ---
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
    Write-Host "Outlook: échec d'envoi - $($_.Exception.Message)" -ForegroundColor Yellow
    return $false
  }
}

# === Flux principal ===
$AllFiles = Expand-Targets -Paths $Targets

if ($Init -or -not (Test-Path $BaselineFile)) {
  Write-Host "Création de la baseline..." -ForegroundColor Yellow
  Save-Baseline -Files $AllFiles
  Write-Host "Baseline enregistrée : $BaselineFile" -ForegroundColor Green
  if ($Init) { return }
}

$Baseline = Import-Clixml $BaselineFile
$Current  = Get-FileHashes -Files $AllFiles

Write-Host "`n--- RESULTAT DE LA SURVEILLANCE ---`n"
$AllPaths = @($Baseline.Keys + $Current.Keys) | Sort-Object -Unique
$changes  = New-Object System.Text.StringBuilder

foreach ($p in $AllPaths) {
  $inB = $Baseline.ContainsKey($p)
  $inC = $Current.ContainsKey($p)

  if (-not $inB -and $inC -and $Current[$p] -ne "INTROUVABLE") {
    Write-Host "[NOUVEAU ] $p" -ForegroundColor Yellow
    [void]$changes.AppendLine("NOUVEAU  : $p")
    Add-Content $LogFile "[$(Get-Date)] NOUVEAU  : $p"
  }
  elseif ($inB -and (-not $inC -or $Current[$p] -eq "INTROUVABLE")) {
    Write-Host "[NOTHING] $p" -ForegroundColor Red
    [void]$changes.AppendLine("NOTHING : $p")
    Add-Content $LogFile "[$(Get-Date)] NOTHING : $p"
  }
  elseif ($inB -and $inC -and $Baseline[$p] -ne $Current[$p]) {
    Write-Host "[MODIFIE ] $p" -ForegroundColor Red
    [void]$changes.AppendLine("MODIFIE  : $p")
    Add-Content $LogFile "[$(Get-Date)] MODIFIE  : $p"
  }
  elseif ($inB -and $inC) {
    Write-Host "[OK      ] $p" -ForegroundColor Green
  }
}

if ($changes.Length -gt 0) {
  $subject = "[HIDS] $env:COMPUTERNAME : changements détectés"
  $body    = $changes.ToString()
  $ok = Send-OutlookMail -to $MailTo -subject $subject -body $body
  if ($ok) { Write-Host "Alerte envoyée via Outlook à $MailTo" -ForegroundColor Green }
} else {
  Write-Host "Aucun changement détecté." -ForegroundColor DarkGray
}

if ($UpdateBaseline) {
  Save-Baseline -Files $AllFiles
  Write-Host "`nBaseline mise à jour." -ForegroundColor Green
}
