# ============================================================================
# SCRIPT DIAGNOSTICO FIRMA OUTLOOK - v1.0
# Verifica stato firma e registry dopo installazione v12.1
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "DIAGNOSI FIRMA OUTLOOK - v1.0" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# 1. VERIFICA PROCESSO OUTLOOK
Write-Host "1. VERIFICA PROCESSO OUTLOOK" -ForegroundColor Yellow
Write-Host "   -------------------------" -ForegroundColor Yellow
$outlookProcesses = Get-Process -Name "outlook" -ErrorAction SilentlyContinue
if ($outlookProcesses) {
    Write-Host "   [ATTENZIONE] Outlook è in esecuzione!" -ForegroundColor Red
    foreach ($proc in $outlookProcesses) {
        Write-Host "   - PID: $($proc.Id) | Memoria: $(($proc.WorkingSet64 / 1MB).ToString('F2')) MB" -ForegroundColor Red
    }
    Write-Host "`n   IMPORTANTE: Chiudi Outlook e riesegui questo script!`n" -ForegroundColor Red
} else {
    Write-Host "   [OK] Outlook non in esecuzione" -ForegroundColor Green
}

# 2. VERIFICA FILE FIRMA
Write-Host "`n2. VERIFICA FILE FIRMA" -ForegroundColor Yellow
Write-Host "   -------------------" -ForegroundColor Yellow
$sigPath = "$env:APPDATA\Microsoft\Signatures"
Write-Host "   Percorso: $sigPath" -ForegroundColor Gray

if (Test-Path $sigPath) {
    $files = Get-ChildItem $sigPath -ErrorAction SilentlyContinue
    if ($files) {
        Write-Host "   [OK] Trovati $($files.Count) file:" -ForegroundColor Green
        foreach ($file in $files) {
            $size = ($file.Length / 1KB).ToString('F2')
            Write-Host "   - $($file.Name) ($size KB) - $($file.LastWriteTime)" -ForegroundColor Gray
        }
    } else {
        Write-Host "   [ERRORE] Cartella esistente ma vuota!" -ForegroundColor Red
    }
} else {
    Write-Host "   [ERRORE] Cartella firme non trovata!" -ForegroundColor Red
}

# 3. VERIFICA PROFILO OUTLOOK
Write-Host "`n3. VERIFICA PROFILO OUTLOOK" -ForegroundColor Yellow
Write-Host "   ------------------------" -ForegroundColor Yellow
$profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

try {
    $profiles = Get-ChildItem $profilesPath -ErrorAction Stop

    if ($profiles.Count -eq 0) {
        Write-Host "   [ERRORE] Nessun profilo Outlook trovato!" -ForegroundColor Red
    } else {
        $profile = $profiles | Select-Object -First 1
        $profileName = $profile.PSChildName
        Write-Host "   [OK] Profilo trovato: $profileName" -ForegroundColor Green

        # 4. VERIFICA ACCOUNT E REGISTRY
        Write-Host "`n4. VERIFICA ACCOUNT E REGISTRY" -ForegroundColor Yellow
        Write-Host "   ---------------------------" -ForegroundColor Yellow

        $accountsPath = "$profilesPath\$profileName\9375CFF0413111d3B88A00104B2A6676"

        try {
            $accounts = Get-ChildItem $accountsPath -ErrorAction Stop

            if ($accounts.Count -eq 0) {
                Write-Host "   [ERRORE] Nessun account trovato!" -ForegroundColor Red
            } else {
                Write-Host "   [OK] Trovati $($accounts.Count) account(s)`n" -ForegroundColor Green

                foreach ($account in $accounts) {
                    Write-Host "   === ACCOUNT: $($account.PSChildName) ===" -ForegroundColor Cyan

                    try {
                        $props = Get-ItemProperty $account.PSPath -ErrorAction Stop

                        # Email
                        $email = $props.'Account Name'
                        if ($email) {
                            Write-Host "   Email: $email" -ForegroundColor Gray
                        } else {
                            Write-Host "   Email: NON TROVATA" -ForegroundColor Red
                        }

                        # New Signature
                        $newSig = $props.'New Signature'
                        if ($newSig) {
                            $type = $newSig.GetType().Name
                            if ($type -eq "Byte[]") {
                                try {
                                    $value = [System.Text.Encoding]::Unicode.GetString($newSig).TrimEnd("`0")
                                    Write-Host "   New Signature: [OK] Binary - Valore: '$value'" -ForegroundColor Green
                                } catch {
                                    Write-Host "   New Signature: [ERRORE] Binary ma non decodificabile" -ForegroundColor Red
                                }
                            } else {
                                Write-Host "   New Signature: [ERRORE] Tipo errato: $type (dovrebbe essere Byte[])" -ForegroundColor Red
                                Write-Host "   Valore: $newSig" -ForegroundColor Red
                            }
                        } else {
                            Write-Host "   New Signature: [ERRORE] NON IMPOSTATA" -ForegroundColor Red
                        }

                        # Reply-Forward Signature
                        $replySig = $props.'Reply-Forward Signature'
                        if ($replySig) {
                            $type = $replySig.GetType().Name
                            if ($type -eq "Byte[]") {
                                try {
                                    $value = [System.Text.Encoding]::Unicode.GetString($replySig).TrimEnd("`0")
                                    Write-Host "   Reply-Forward Signature: [OK] Binary - Valore: '$value'" -ForegroundColor Green
                                } catch {
                                    Write-Host "   Reply-Forward Signature: [ERRORE] Binary ma non decodificabile" -ForegroundColor Red
                                }
                            } else {
                                Write-Host "   Reply-Forward Signature: [ERRORE] Tipo errato: $type (dovrebbe essere Byte[])" -ForegroundColor Red
                                Write-Host "   Valore: $replySig" -ForegroundColor Red
                            }
                        } else {
                            Write-Host "   Reply-Forward Signature: [ERRORE] NON IMPOSTATA" -ForegroundColor Red
                        }

                        Write-Host ""

                    } catch {
                        Write-Host "   [ERRORE] Impossibile leggere proprietà account: $_" -ForegroundColor Red
                    }
                }
            }

        } catch {
            Write-Host "   [ERRORE] Impossibile accedere agli account: $_" -ForegroundColor Red
        }
    }

} catch {
    Write-Host "   [ERRORE] Impossibile accedere ai profili Outlook: $_" -ForegroundColor Red
}

# 5. VERIFICA LOG INSTALLAZIONE
Write-Host "`n5. VERIFICA LOG INSTALLAZIONE" -ForegroundColor Yellow
Write-Host "   --------------------------" -ForegroundColor Yellow
$logPath = "C:\Temp\Firma_Install_Log.txt"

if (Test-Path $logPath) {
    $logInfo = Get-Item $logPath
    Write-Host "   [OK] Log trovato: $logPath" -ForegroundColor Green
    Write-Host "   Ultima modifica: $($logInfo.LastWriteTime)" -ForegroundColor Gray
    Write-Host "`n   === ULTIME 20 RIGHE DEL LOG ===" -ForegroundColor Cyan
    Get-Content $logPath -Tail 20 | ForEach-Object {
        if ($_ -match "\[SUCCESS\]") {
            Write-Host "   $_" -ForegroundColor Green
        } elseif ($_ -match "\[ERROR\]") {
            Write-Host "   $_" -ForegroundColor Red
        } elseif ($_ -match "\[WARNING\]") {
            Write-Host "   $_" -ForegroundColor Yellow
        } else {
            Write-Host "   $_" -ForegroundColor Gray
        }
    }
} else {
    Write-Host "   [ERRORE] Log non trovato: $logPath" -ForegroundColor Red
}

# RIEPILOGO FINALE
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RIEPILOGO DIAGNOSTICA" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Prossimi passi suggeriti:" -ForegroundColor Yellow
Write-Host "1. Se Outlook era in esecuzione, chiudilo e riesegui lo script di installazione" -ForegroundColor Gray
Write-Host "2. Se i registry sono Binary ma Outlook non mostra le firme, prova:" -ForegroundColor Gray
Write-Host "   - Chiudi completamente Outlook (verifica Task Manager)" -ForegroundColor Gray
Write-Host "   - Riesegui: \\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1" -ForegroundColor Gray
Write-Host "   - Riapri Outlook e verifica dropdown: File -> Opzioni -> Posta -> Firme" -ForegroundColor Gray
Write-Host "`n"
