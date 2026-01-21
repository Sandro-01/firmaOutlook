# ============================================================================
# VERIFICA FIRME NEL PROFILO PREDEFINITO - v1.0
# Controlla se le firme sono impostate nel profilo "Outlook" predefinito
# ============================================================================

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "VERIFICA FIRME PROFILO PREDEFINITO" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

$profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
$defaultProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook"
$defaultProfileName = (Get-ItemProperty $defaultProfilePath -ErrorAction SilentlyContinue).'DefaultProfile'

Write-Host "Profilo predefinito: $defaultProfileName`n" -ForegroundColor Yellow

$accountsPath = "$profilesPath\$defaultProfileName\9375CFF0413111d3B88A00104B2A6676"

try {
    $accounts = Get-ChildItem $accountsPath -ErrorAction Stop

    Write-Host "Trovati $($accounts.Count) account(s):`n" -ForegroundColor Green

    foreach ($account in $accounts) {
        Write-Host "=== ACCOUNT: $($account.PSChildName) ===" -ForegroundColor Cyan

        try {
            $props = Get-ItemProperty $account.PSPath -ErrorAction Stop

            # Email
            $email = $props.'Account Name'
            if ($email) {
                Write-Host "Email: $email" -ForegroundColor White
            } else {
                Write-Host "Email: NON TROVATA" -ForegroundColor Red
            }

            # New Signature
            $newSig = $props.'New Signature'
            if ($newSig) {
                $type = $newSig.GetType().Name
                if ($type -eq "Byte[]") {
                    try {
                        $value = [System.Text.Encoding]::Unicode.GetString($newSig).TrimEnd("`0")
                        Write-Host "New Signature: ✅ Binary - '$value'" -ForegroundColor Green
                    } catch {
                        Write-Host "New Signature: ❌ Binary ma non decodificabile" -ForegroundColor Red
                    }
                } else {
                    Write-Host "New Signature: ❌ Tipo: $type (dovrebbe essere Byte[])" -ForegroundColor Red
                }
            } else {
                Write-Host "New Signature: ❌ NON IMPOSTATA" -ForegroundColor Red
            }

            # Reply-Forward Signature
            $replySig = $props.'Reply-Forward Signature'
            if ($replySig) {
                $type = $replySig.GetType().Name
                if ($type -eq "Byte[]") {
                    try {
                        $value = [System.Text.Encoding]::Unicode.GetString($replySig).TrimEnd("`0")
                        Write-Host "Reply-Forward Signature: ✅ Binary - '$value'" -ForegroundColor Green
                    } catch {
                        Write-Host "Reply-Forward Signature: ❌ Binary ma non decodificabile" -ForegroundColor Red
                    }
                } else {
                    Write-Host "Reply-Forward Signature: ❌ Tipo: $type (dovrebbe essere Byte[])" -ForegroundColor Red
                }
            } else {
                Write-Host "Reply-Forward Signature: ❌ NON IMPOSTATA" -ForegroundColor Red
            }

            Write-Host ""

        } catch {
            Write-Host "Errore lettura account: $_`n" -ForegroundColor Red
        }
    }

} catch {
    Write-Host "Errore accesso account: $_" -ForegroundColor Red
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "RIEPILOGO" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

Write-Host "Se vedi ✅ per account $($email):" -ForegroundColor Yellow
Write-Host "  1. Apri Outlook" -ForegroundColor Gray
Write-Host "  2. Vai a: File → Opzioni → Posta → Firme" -ForegroundColor Gray
Write-Host "  3. I dropdown DEVONO essere popolati con 'Firma_Aziendale'" -ForegroundColor Gray
Write-Host "`nSe vedi ❌ NON IMPOSTATA:" -ForegroundColor Yellow
Write-Host "  1. Chiudi completamente Outlook" -ForegroundColor Gray
Write-Host "  2. Esegui: \\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1" -ForegroundColor Gray
Write-Host "  3. Riesegui questo script per verificare" -ForegroundColor Gray
Write-Host "`n"
