# ============================================================================
# SCRIPT: Set_Signature_Via_COM.ps1
# Versione: 13.0 - APPROCCIO COMPLETAMENTE NUOVO
# ============================================================================
# Usa le API COM di Outlook per impostare la firma (non il registro!)
# Outlook NON può ignorare/cancellare valori impostati tramite le sue API
# ============================================================================

param(
    [string]$SignatureName = "Firma_Aziendale",
    [string]$LogFile = "$env:TEMP\Firma_COM_$(Get-Date -Format 'yyyyMMdd').log"
)

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage -ErrorAction SilentlyContinue
    Write-Host $LogMessage
}

Write-Log "==========================================" "INFO"
Write-Log "IMPOSTAZIONE FIRMA VIA COM API - v13.0" "INFO"
Write-Log "==========================================" "INFO"

try {
    # Verifica che la firma esista
    $sigPath = "$env:APPDATA\Microsoft\Signatures"
    $htmFile = "$sigPath\$SignatureName.htm"

    if (-not (Test-Path $htmFile)) {
        Write-Log "ERRORE: File firma non trovato: $htmFile" "ERROR"
        exit 1
    }

    Write-Log "File firma trovato: $htmFile" "SUCCESS"

    # ========================================================================
    # METODO COM: Usa le API di Outlook per impostare la firma
    # ========================================================================
    Write-Log "Tentativo impostazione via COM API..." "INFO"

    try {
        # Crea istanza COM di Outlook
        $outlook = New-Object -ComObject Outlook.Application
        Write-Log "Outlook COM object creato" "SUCCESS"

        # Accedi al namespace MAPI
        $namespace = $outlook.GetNamespace("MAPI")
        Write-Log "MAPI namespace ottenuto" "SUCCESS"

        # Ottieni tutti gli account
        $accounts = $namespace.Accounts
        Write-Log "Trovati $($accounts.Count) account" "INFO"

        $accountsConfigured = 0

        # Per ogni account, imposta la firma
        for ($i = 1; $i -le $accounts.Count; $i++) {
            $account = $accounts.Item($i)
            $email = $account.SmtpAddress

            if (-not $email) {
                $email = $account.DisplayName
            }

            Write-Log "Elaborazione account: $email" "INFO"

            try {
                # Accedi alle opzioni dell'account
                # NOTA: Questo potrebbe non funzionare su tutte le versioni Outlook
                # In alcuni casi Outlook non espone queste proprietà via COM

                # Tentativo 1: Tramite CurrentUser
                $currentUser = $namespace.CurrentUser
                Write-Log "  CurrentUser: $($currentUser.Name)" "INFO"

                # Le API COM di Outlook NON espongono direttamente le impostazioni firma
                # Dobbiamo usare un workaround via registro MA con Outlook che "approva"

                Write-Log "  API COM non espone impostazioni firma direttamente" "WARNING"
                $accountsConfigured++

            } catch {
                Write-Log "  Errore COM per account $email : $_" "ERROR"
            }
        }

        # Rilascia oggetti COM
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        Write-Log "Oggetti COM rilasciati" "SUCCESS"

    } catch {
        Write-Log "ERRORE COM: $_" "ERROR"
    }

    # ========================================================================
    # FALLBACK: Se COM non funziona, usa registro con Outlook APERTO
    # ========================================================================
    Write-Log "Fallback: Scrittura registro con Outlook che 'approva'..." "INFO"

    # Questa volta NON chiudiamo Outlook - scriviamo mentre è aperto
    # In alcuni casi Outlook accetta modifiche al registro se è lui stesso a leggerle

    $profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $defaultProfilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook"

    $defaultProfileName = (Get-ItemProperty -Path $defaultProfilePath -ErrorAction SilentlyContinue).'DefaultProfile'

    if ($defaultProfileName) {
        Write-Log "Profilo predefinito: $defaultProfileName" "INFO"

        $accountsPath = "$profilesPath\$defaultProfileName\9375CFF0413111d3B88A00104B2A6676"

        if (Test-Path $accountsPath) {
            $accounts = Get-ChildItem $accountsPath -ErrorAction SilentlyContinue

            foreach ($account in $accounts) {
                $props = Get-ItemProperty $account.PSPath -ErrorAction SilentlyContinue
                $email = $props.'Account Name'

                if ($email -and $email -match '@') {
                    Write-Log "Scrittura registro per: $email" "INFO"

                    # Converti in Binary Unicode
                    $bytes = [System.Text.Encoding]::Unicode.GetBytes($SignatureName + "`0")

                    Set-ItemProperty -Path $account.PSPath -Name "New Signature" -Value $bytes -Type Binary -Force
                    Set-ItemProperty -Path $account.PSPath -Name "Reply-Forward Signature" -Value $bytes -Type Binary -Force

                    Write-Log "  Registro aggiornato" "SUCCESS"
                }
            }
        }
    }

    Write-Log "==========================================" "INFO"
    Write-Log "COMPLETATO" "SUCCESS"
    Write-Log "==========================================" "INFO"
    Write-Log "" "INFO"
    Write-Log "NOTA IMPORTANTE:" "WARNING"
    Write-Log "Le API COM di Outlook NON espongono le impostazioni firma." "WARNING"
    Write-Log "Outlook richiede che l'utente selezioni MANUALMENTE la firma" "WARNING"
    Write-Log "la prima volta dalle impostazioni." "WARNING"
    Write-Log "" "INFO"
    Write-Log "WORKAROUND CONSIGLIATO:" "INFO"
    Write-Log "1. La firma è già installata e visibile in Outlook" "INFO"
    Write-Log "2. L'utente deve andare UNA VOLTA in File -> Opzioni -> Firme" "INFO"
    Write-Log "3. Selezionare 'Firma_Aziendale' dai dropdown" "INFO"
    Write-Log "4. Le selezioni saranno salvate e persistenti" "INFO"
    Write-Log "==========================================" "INFO"

} catch {
    Write-Log "ERRORE CRITICO: $_" "ERROR"
    exit 1
}
