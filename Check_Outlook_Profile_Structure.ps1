# ============================================================================
# VERIFICA STRUTTURA PROFILO OUTLOOK - v1.0
# Analizza struttura completa registro per identificare tipo Outlook
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "ANALISI STRUTTURA PROFILO OUTLOOK" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# 1. VERIFICA TIPO OUTLOOK
Write-Host "1. IDENTIFICAZIONE TIPO OUTLOOK" -ForegroundColor Yellow
Write-Host "   ----------------------------" -ForegroundColor Yellow

$newOutlookPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup"
if (Test-Path $newOutlookPath) {
    $setupProps = Get-ItemProperty $newOutlookPath -ErrorAction SilentlyContinue
    Write-Host "   Outlook Setup trovato:" -ForegroundColor Green
    Write-Host "   - FirstRun: $($setupProps.FirstRun)" -ForegroundColor Gray
    Write-Host "   - ImportPRF: $($setupProps.ImportPRF)" -ForegroundColor Gray
}

# Verifica se è New Outlook (web-based)
$newOutlookCheck = Get-Process | Where-Object { $_.ProcessName -like "*olk*" -or $_.MainWindowTitle -like "*Outlook (new)*" }
if ($newOutlookCheck) {
    Write-Host "`n   [IMPORTANTE] Rilevato NUOVO OUTLOOK (web-based)!" -ForegroundColor Red
    Write-Host "   Il nuovo Outlook NON usa il registro per le firme!" -ForegroundColor Red
    Write-Host "   Le firme sono gestite online tramite Exchange/365." -ForegroundColor Red
}

# 2. LISTA TUTTI I PROFILI
Write-Host "`n2. PROFILI OUTLOOK DISPONIBILI" -ForegroundColor Yellow
Write-Host "   ---------------------------" -ForegroundColor Yellow

$profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
try {
    $profiles = Get-ChildItem $profilesPath -ErrorAction Stop

    Write-Host "   Trovati $($profiles.Count) profilo/i:`n" -ForegroundColor Green

    foreach ($profile in $profiles) {
        $profileName = $profile.PSChildName
        Write-Host "   === PROFILO: $profileName ===" -ForegroundColor Cyan

        # Verifica profilo predefinito
        $defaultProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook"
        $defaultProfileName = (Get-ItemProperty $defaultProfilePath -ErrorAction SilentlyContinue).'DefaultProfile'
        if ($profileName -eq $defaultProfileName) {
            Write-Host "   [DEFAULT] Questo è il profilo predefinito ✅" -ForegroundColor Green
        }

        # Esplora struttura profilo
        Write-Host "   Struttura chiavi:" -ForegroundColor Gray
        $profileKeys = Get-ChildItem $profile.PSPath -ErrorAction SilentlyContinue
        foreach ($key in $profileKeys) {
            Write-Host "   - $($key.PSChildName)" -ForegroundColor Gray

            # Se è la chiave account standard, analizza
            if ($key.PSChildName -eq "9375CFF0413111d3B88A00104B2A6676") {
                Write-Host "     ✅ Chiave account STANDARD trovata!" -ForegroundColor Green

                $accounts = Get-ChildItem $key.PSPath -ErrorAction SilentlyContinue
                Write-Host "     Account configurati: $($accounts.Count)" -ForegroundColor Green

                foreach ($account in $accounts | Select-Object -First 3) {
                    $accountProps = Get-ItemProperty $account.PSPath -ErrorAction SilentlyContinue
                    $email = $accountProps.'Account Name'
                    Write-Host "       - Account: $($account.PSChildName)" -ForegroundColor Gray
                    if ($email) {
                        Write-Host "         Email: $email" -ForegroundColor Gray
                    }
                }
            }
        }
        Write-Host ""
    }

} catch {
    Write-Host "   [ERRORE] Impossibile leggere profili: $_" -ForegroundColor Red
}

# 3. VERIFICA VERSIONE OUTLOOK
Write-Host "`n3. VERSIONE OUTLOOK INSTALLATA" -ForegroundColor Yellow
Write-Host "   ---------------------------" -ForegroundColor Yellow

# Verifica Office 16.0
$office16Path = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
if (Test-Path $office16Path) {
    $officeProps = Get-ItemProperty $office16Path -ErrorAction SilentlyContinue
    Write-Host "   [OK] Office 365 / 2016+ rilevato" -ForegroundColor Green
    Write-Host "   - Versione: $($officeProps.VersionToReport)" -ForegroundColor Gray
    Write-Host "   - Build: $($officeProps.ClientVersionToReport)" -ForegroundColor Gray
    Write-Host "   - UpdateChannel: $($officeProps.UpdateChannel)" -ForegroundColor Gray
}

# Verifica quale Outlook è in esecuzione
Write-Host "`n   Processo Outlook in esecuzione:" -ForegroundColor Yellow
$outlookProc = Get-Process -Name "outlook" -ErrorAction SilentlyContinue
if ($outlookProc) {
    foreach ($proc in $outlookProc) {
        Write-Host "   - Path: $($proc.Path)" -ForegroundColor Gray
        Write-Host "   - Titolo finestra: $($proc.MainWindowTitle)" -ForegroundColor Gray

        # Verifica se è New Outlook
        if ($proc.MainWindowTitle -like "*new*" -or $proc.Path -like "*HxOutlook*") {
            Write-Host "   [CRITICO] NUOVO OUTLOOK RILEVATO!" -ForegroundColor Red
            Write-Host "   Il nuovo Outlook non supporta firma via registro!" -ForegroundColor Red
        } else {
            Write-Host "   [OK] Outlook CLASSICO rilevato ✅" -ForegroundColor Green
        }
    }
}

# 4. CERCA CHIAVI FIRMA IN TUTTI I PERCORSI POSSIBILI
Write-Host "`n4. RICERCA CHIAVI FIRMA" -ForegroundColor Yellow
Write-Host "   --------------------" -ForegroundColor Yellow

# Cerca in tutte le versioni Office
$officeVersions = @("14.0", "15.0", "16.0")
foreach ($ver in $officeVersions) {
    $basePath = "HKCU:\Software\Microsoft\Office\$ver\Common\MailSettings"
    if (Test-Path $basePath) {
        Write-Host "`n   [TROVATO] Office $ver - MailSettings" -ForegroundColor Green
        $props = Get-ItemProperty $basePath -ErrorAction SilentlyContinue

        $newSig = $props.'NewSignature'
        $replySig = $props.'ReplySignature'

        if ($newSig) {
            Write-Host "   - NewSignature: $newSig" -ForegroundColor Gray
        } else {
            Write-Host "   - NewSignature: NON IMPOSTATA" -ForegroundColor Gray
        }

        if ($replySig) {
            Write-Host "   - ReplySignature: $replySig" -ForegroundColor Gray
        } else {
            Write-Host "   - ReplySignature: NON IMPOSTATA" -ForegroundColor Gray
        }
    }
}

# 5. VERIFICA PERCORSO LOG ALTERNATIVO
Write-Host "`n5. RICERCA LOG INSTALLAZIONE" -ForegroundColor Yellow
Write-Host "   -------------------------" -ForegroundColor Yellow

$possibleLogPaths = @(
    "C:\Temp\Firma_Install_Log.txt",
    "$env:TEMP\Firma_Install_Log.txt",
    "$env:TEMP\Installazione_Firma_*.log"
)

foreach ($logPath in $possibleLogPaths) {
    if ($logPath -like "*`**") {
        # Pattern con wildcard
        $logFiles = Get-ChildItem -Path ($logPath -replace '\*.*$', '') -Filter ($logPath -replace '^.*\\', '') -ErrorAction SilentlyContinue
        if ($logFiles) {
            Write-Host "   [TROVATO] $($logFiles.Count) log file(s):" -ForegroundColor Green
            foreach ($log in $logFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 3) {
                Write-Host "   - $($log.FullName)" -ForegroundColor Gray
                Write-Host "     Ultima modifica: $($log.LastWriteTime)" -ForegroundColor Gray
            }
        }
    } else {
        if (Test-Path $logPath) {
            Write-Host "   [TROVATO] $logPath" -ForegroundColor Green
            $logInfo = Get-Item $logPath
            Write-Host "   Ultima modifica: $($logInfo.LastWriteTime)" -ForegroundColor Gray
        }
    }
}

# RIEPILOGO
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RIEPILOGO" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Prossime azioni consigliate:" -ForegroundColor Yellow
Write-Host "1. Se stai usando NUOVO Outlook:" -ForegroundColor Gray
Write-Host "   - Le firme NON si gestiscono via registro" -ForegroundColor Gray
Write-Host "   - Devi usare Outlook Web / Exchange Online per impostare firme" -ForegroundColor Gray
Write-Host "   - OPPURE torna a Outlook CLASSICO" -ForegroundColor Gray
Write-Host "2. Se stai usando Outlook CLASSICO:" -ForegroundColor Gray
Write-Host "   - Chiudi Outlook completamente" -ForegroundColor Gray
Write-Host "   - Riesegui lo script: \\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1" -ForegroundColor Gray
Write-Host "   - Verifica che il profilo abbia la chiave 9375CFF0413111d3B88A00104B2A6676" -ForegroundColor Gray
Write-Host "`n"
