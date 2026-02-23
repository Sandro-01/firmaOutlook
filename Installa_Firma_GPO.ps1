# ==============================================================================
# Script: Installa_Firma_GPO.ps1
# Versione: 14.1 - Aspetta Outlook + uscita silenziosa senza VPN
# ==============================================================================

$NetworkPath = "\\DEAZRADS101\Firme"
$LocalPath = "$env:APPDATA\Microsoft\Signatures"
$LogFile = "$env:TEMP\Firma_Install_Log.txt"
$SignatureName = "Firma_Aziendale"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    Add-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message" -EA SilentlyContinue
}

Write-Log "=========================================="
Write-Log "INSTALLAZIONE FIRMA v14.1"
Write-Log "Utente: $env:USERNAME | PC: $env:COMPUTERNAME"

# === VERIFICA RETE - Esci silenziosamente se non connesso ===
if (-not (Test-Path $NetworkPath -EA SilentlyContinue)) {
    Write-Log "Rete aziendale non disponibile, uscita silenziosa" "INFO"
    exit 0
}

# === ASPETTA OUTLOOK ===
$maxWait = 300
$waited = 0
while (-not (Get-Process outlook -EA SilentlyContinue) -and $waited -lt $maxWait) {
    Start-Sleep -Seconds 10
    $waited += 10
}

if (-not (Get-Process outlook -EA SilentlyContinue)) {
    Write-Log "Outlook non avviato, esco" "WARNING"
    exit 0
}

Write-Log "Outlook rilevato, attendo sync..."
Start-Sleep -Seconds 45

try {
    if (-not (Test-Path $LocalPath)) {
        New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
    }

    # === TROVA CARTELLA UTENTE ===
    $UserFolder = $null
    try {
        $adUser = ([ADSISEARCHER]"samaccountname=$env:USERNAME").FindOne()
        if ($adUser -and $adUser.Properties.mail) {
            $email = $adUser.Properties.mail[0]
            if (Test-Path "$NetworkPath\$email") { $UserFolder = "$NetworkPath\$email" }
        }
    } catch {}

    if (-not $UserFolder -and (Test-Path "$NetworkPath\$env:USERNAME")) {
        $UserFolder = "$NetworkPath\$env:USERNAME"
    }
    if (-not $UserFolder -and (Test-Path "$NetworkPath\Default")) {
        $UserFolder = "$NetworkPath\Default"
    }
    if (-not $UserFolder) {
        Write-Log "Nessuna cartella firma" "ERROR"
        exit 1
    }
    Write-Log "Cartella: $UserFolder" "SUCCESS"

    # === COPIA FILE ===
    $copiedFiles = 0
    Get-ChildItem -Path $UserFolder -File -EA SilentlyContinue | ForEach-Object {
        $ext = $_.Extension.ToLower()
        if ($ext -in @(".htm", ".html", ".txt", ".rtf")) {
            $destName = if ($ext -eq ".html") { "$SignatureName.htm" } else { "$SignatureName$ext" }
            Copy-Item -Path $_.FullName -Destination (Join-Path $LocalPath $destName) -Force
            $copiedFiles++
        }
    }
    Write-Log "File copiati: $copiedFiles" "SUCCESS"

    # === IMPOSTA FIRMA REGISTRO ===
    $signatureBinary = [System.Text.Encoding]::Unicode.GetBytes($SignatureName + "`0")
    $profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $signaturesSet = 0

    if (Test-Path $profilesPath) {
        Get-ChildItem $profilesPath | ForEach-Object {
            $accountsPath = Join-Path $_.PSPath "9375CFF0413111d3B88A00104B2A6676"
            if (Test-Path $accountsPath) {
                Get-ChildItem $accountsPath | ForEach-Object {
                    Remove-ItemProperty -Path $_.PSPath -Name "New Signature" -Force -EA SilentlyContinue
                    Remove-ItemProperty -Path $_.PSPath -Name "Reply-Forward Signature" -Force -EA SilentlyContinue
                    New-ItemProperty -Path $_.PSPath -Name "New Signature" -Value $signatureBinary -PropertyType Binary -Force | Out-Null
                    New-ItemProperty -Path $_.PSPath -Name "Reply-Forward Signature" -Value $signatureBinary -PropertyType Binary -Force | Out-Null
                    $signaturesSet++
                }
            }
        }
    }

    Write-Log "Firme OK: $signaturesSet" "SUCCESS"
    Write-Log "=========================================="
    exit 0
}
catch {
    Write-Log "ERRORE: $_" "ERROR"
    exit 1
}
