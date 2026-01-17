# ==============================================================================
# Script: Installa_Firma_GPO.ps1
# Versione: 12.0 - FIX CRITICO: Firma predefinita in formato Binary Unicode
# Carton Group Italia - Distribuzione Firme Outlook
#
# CHANGELOG v12.0:
# - FIX CRITICO: Conversione valori firma da String a Binary Unicode
# - Aggiunta scrittura in MailSettings come fallback
# - Migliorata identificazione profilo Outlook attivo
# - Gestione robusta account multipli e versioni Office diverse
# - Logging dettagliato con percorsi registro per debug
# ==============================================================================

$NetworkPath = "\\DEAZRADS101\Firme"
$LocalPath = "$env:APPDATA\Microsoft\Signatures"
$LogFile = "$env:TEMP\Installazione_Firma_$(Get-Date -Format 'yyyyMMdd').log"
$SignatureName = "Firma_Aziendale"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $LogFile -Value "[$Timestamp] [$Level] $Message" -ErrorAction SilentlyContinue
}

Write-Log "==========================================" "INFO"
Write-Log "INSTALLAZIONE FIRMA v12.0" "INFO"
Write-Log "Utente: $env:USERNAME" "INFO"
Write-Log "Computer: $env:COMPUTERNAME" "INFO"
Write-Log "==========================================" "INFO"

try {
    # ========================================================================
    # VERIFICA CONNESSIONE RETE
    # ========================================================================
    Write-Log "Verifica connessione al server..." "INFO"
    if (-not (Test-Path $NetworkPath)) {
        Write-Log "Server non raggiungibile: $NetworkPath" "ERROR"
        exit 1
    }
    Write-Log "Server raggiungibile" "SUCCESS"

    # ========================================================================
    # CREA CARTELLA FIRME LOCALE
    # ========================================================================
    if (-not (Test-Path $LocalPath)) {
        New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
        Write-Log "Cartella firme creata: $LocalPath" "INFO"
    }

    # ========================================================================
    # CERCA CARTELLA UTENTE SUL SERVER
    # ========================================================================
    Write-Log "Ricerca cartella firma utente..." "INFO"
    $UserFolder = $null
    $PossiblePaths = @()
    $UserEmail = $null
    $UserDisplayName = $null

    # Metodo 1: Username Windows
    $PossiblePaths += "$NetworkPath\$env:USERNAME"
    $PossiblePaths += "$NetworkPath\$($env:USERNAME.ToLower())"

    # Metodo 2: Email da Active Directory
    try {
        $adUser = ([ADSISEARCHER]"samaccountname=$env:USERNAME").FindOne()
        if ($adUser) {
            $mail = $adUser.Properties.mail
            if ($mail -and $mail.Count -gt 0) {
                $UserEmail = $mail[0].ToString()
                Write-Log "Email AD trovata: $UserEmail" "INFO"

                # Email completa (con e senza case sensitivity)
                $PossiblePaths += "$NetworkPath\$UserEmail"
                $PossiblePaths += "$NetworkPath\$($UserEmail.ToLower())"

                # Username da email (parte prima della @)
                $emailUser = $UserEmail.Split('@')[0]
                $PossiblePaths += "$NetworkPath\$emailUser"
                $PossiblePaths += "$NetworkPath\$($emailUser.ToLower())"
            }

            $displayName = $adUser.Properties.displayname
            if ($displayName -and $displayName.Count -gt 0) {
                $UserDisplayName = $displayName[0].ToString()
                Write-Log "Nome completo: $UserDisplayName" "INFO"
            }
        }
    } catch {
        Write-Log "Impossibile recuperare info AD: $_" "WARNING"
    }

    # Rimuovi duplicati
    $PossiblePaths = $PossiblePaths | Select-Object -Unique
    Write-Log "Ricerca tra $($PossiblePaths.Count) percorsi possibili..." "INFO"

    # Cerca la cartella
    foreach ($path in $PossiblePaths) {
        if (Test-Path $path) {
            $UserFolder = $path
            Write-Log "Cartella trovata: $UserFolder" "SUCCESS"
            break
        }
    }

    if (-not $UserFolder) {
        Write-Log "Nessuna cartella firma trovata per l'utente" "ERROR"
        Write-Log "Percorsi cercati: $($PossiblePaths -join ', ')" "INFO"
        exit 1
    }

    # ========================================================================
    # COPIA FILE FIRMA DAL SERVER
    # ========================================================================
    Write-Log "Copia file firma dal server..." "INFO"
    $sourceFiles = Get-ChildItem -Path $UserFolder -File -ErrorAction SilentlyContinue

    if (-not $sourceFiles -or $sourceFiles.Count -eq 0) {
        Write-Log "Nessun file firma trovato in $UserFolder" "ERROR"
        exit 1
    }

    Write-Log "Trovati $($sourceFiles.Count) file da copiare" "INFO"
    $copiedFiles = 0

    foreach ($file in $sourceFiles) {
        try {
            $destPath = Join-Path $LocalPath $file.Name
            Copy-Item -Path $file.FullName -Destination $destPath -Force
            $sizeKB = [math]::Round($file.Length / 1KB, 2)
            Write-Log "File copiato: $($file.Name) ($sizeKB KB)" "SUCCESS"
            $copiedFiles++
        } catch {
            Write-Log "Errore copia $($file.Name): $_" "ERROR"
        }
    }

    # ========================================================================
    # GENERA FILE RTF (se non esiste)
    # ========================================================================
    $rtfPath = Join-Path $LocalPath "$SignatureName.rtf"
    $htmPath = Join-Path $LocalPath "$SignatureName.htm"

    if ((Test-Path $htmPath) -and -not (Test-Path $rtfPath)) {
        Write-Log "Generazione file RTF..." "INFO"
        try {
            $htmContent = Get-Content $htmPath -Raw -Encoding UTF8
            $rtfContent = "{\rtf1\ansi\deff0 {\fonttbl {\f0 Arial;}} \f0\fs20 Firma Aziendale - Vedi versione HTML}"
            Set-Content -Path $rtfPath -Value $rtfContent -Encoding ASCII
            $sizeKB = [math]::Round((Get-Item $rtfPath).Length / 1KB, 2)
            Write-Log "File RTF creato ($sizeKB KB)" "SUCCESS"
        } catch {
            Write-Log "Errore creazione RTF: $_" "WARNING"
        }
    }

    # ========================================================================
    # CONFIGURA PERMESSI FILE (modificabili dall'utente)
    # ========================================================================
    Write-Log "Configurazione attributi per permettere modifiche utente..." "INFO"
    $signatureFiles = Get-ChildItem -Path $LocalPath -Filter "$SignatureName.*" -ErrorAction SilentlyContinue
    $permissionsSet = 0

    foreach ($file in $signatureFiles) {
        try {
            if ($file.IsReadOnly) {
                $file.IsReadOnly = $false
            }
            $acl = Get-Acl $file.FullName
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $permission = New-Object System.Security.AccessControl.FileSystemAccessRule($currentUser, "FullControl", "Allow")
            $acl.SetAccessRule($permission)
            Set-Acl $file.FullName $acl
            Write-Log "  Permessi utente impostati: $($file.Name)" "SUCCESS"
            $permissionsSet++
        } catch {
            Write-Log "  Errore permessi $($file.Name): $_" "WARNING"
        }
    }
    Write-Log "File configurati come modificabili: $permissionsSet" "SUCCESS"

    # ========================================================================
    # SBLOCCA INTERFACCIA OUTLOOK (rimuovi chiavi bloccanti)
    # ========================================================================
    Write-Log "Sblocco interfaccia Outlook..." "INFO"
    $mailSettingsPath = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
    $blockingKeys = @("NewSignature", "ReplySignature", "DisableSignatures", "DisablePersonalSignatures")

    if (Test-Path $mailSettingsPath) {
        foreach ($key in $blockingKeys) {
            try {
                Remove-ItemProperty -Path $mailSettingsPath -Name $key -ErrorAction SilentlyContinue
            } catch {}
        }
        Write-Log "Chiavi bloccanti rimosse da MailSettings" "SUCCESS"
    } else {
        Write-Log "Percorso MailSettings non trovato (normale se Office non configurato)" "INFO"
    }

    # ========================================================================
    # FUNZIONE: Converte stringa in Binary Unicode per registro Outlook
    # ========================================================================
    function Convert-ToRegistryBinary {
        param([string]$Value)

        # Outlook richiede formato Unicode (UTF-16 LE) con terminatore null
        $bytes = [System.Text.Encoding]::Unicode.GetBytes($Value + "`0")
        return $bytes
    }

    # ========================================================================
    # IMPOSTA FIRMA PREDEFINITA - METODO 1: MailSettings (Fallback)
    # ========================================================================
    Write-Log "Metodo 1: Impostazione firma in MailSettings (fallback per GPO)..." "INFO"

    # Supporta Office 2016, 2019, 2021, Microsoft 365
    $officeVersions = @("16.0", "15.0", "14.0")
    $mailSettingsWritten = $false

    foreach ($version in $officeVersions) {
        $mailSettingsPath = "HKCU:\Software\Microsoft\Office\$version\Common\MailSettings"

        if (Test-Path "HKCU:\Software\Microsoft\Office\$version") {
            try {
                # Crea il percorso se non esiste
                if (-not (Test-Path $mailSettingsPath)) {
                    New-Item -Path $mailSettingsPath -Force | Out-Null
                    Write-Log "  Creato percorso: $mailSettingsPath" "INFO"
                }

                # Converti nome firma in Binary Unicode
                $signatureBinary = Convert-ToRegistryBinary -Value $SignatureName

                # Imposta le firme predefinite
                Set-ItemProperty -Path $mailSettingsPath -Name "NewSignature" -Value $signatureBinary -Type Binary -Force
                Set-ItemProperty -Path $mailSettingsPath -Name "ReplySignature" -Value $signatureBinary -Type Binary -Force

                Write-Log "  Firma impostata in MailSettings (Office $version)" "SUCCESS"
                Write-Log "  Percorso: $mailSettingsPath" "INFO"
                $mailSettingsWritten = $true
            } catch {
                Write-Log "  Errore scrittura MailSettings (Office $version): $_" "WARNING"
            }
        }
    }

    if ($mailSettingsWritten) {
        Write-Log "Firma predefinita scritta in MailSettings come fallback" "SUCCESS"
    } else {
        Write-Log "Nessuna versione Office trovata per MailSettings" "WARNING"
    }

    # ========================================================================
    # IMPOSTA FIRMA PREDEFINITA - METODO 2: Account Outlook (Principale)
    # ========================================================================
    Write-Log "Metodo 2: Impostazione firma per account Outlook..." "INFO"
    $signaturesSet = 0
    $accountsProcessed = 0

    # Identifica profilo Outlook attivo
    $profilesBasePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $defaultProfileName = $null

    # Cerca il profilo predefinito
    try {
        $outlookPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook"
        if (Test-Path $outlookPath) {
            $defaultProfileName = (Get-ItemProperty -Path $outlookPath -Name "DefaultProfile" -ErrorAction SilentlyContinue).'DefaultProfile'
            if ($defaultProfileName) {
                Write-Log "  Profilo predefinito: $defaultProfileName" "INFO"
            }
        }
    } catch {}

    # Se non trovato, cerca tutti i profili
    if (Test-Path $profilesBasePath) {
        $profiles = Get-ChildItem $profilesBasePath -ErrorAction SilentlyContinue

        # Prioritizza il profilo predefinito
        if ($defaultProfileName) {
            $profiles = $profiles | Sort-Object { $_.PSChildName -ne $defaultProfileName }
        }

        Write-Log "  Trovati $($profiles.Count) profilo/i Outlook" "INFO"

        foreach ($profile in $profiles) {
            $profileName = $profile.PSChildName
            Write-Log "  Elaborazione profilo: $profileName" "INFO"

            # Path account email (GUID standard Outlook)
            $accountsPath = Join-Path $profile.PSPath "9375CFF0413111d3B88A00104B2A6676"

            if (Test-Path $accountsPath) {
                # Enumera tutti i GUID (account email)
                $accounts = Get-ChildItem $accountsPath -ErrorAction SilentlyContinue

                Write-Log "    Trovati $($accounts.Count) account nel profilo" "INFO"

                foreach ($account in $accounts) {
                    $accountsProcessed++
                    $accountGuid = $account.PSChildName

                    try {
                        # Tenta di identificare l'account
                        $props = Get-ItemProperty $account.PSPath -ErrorAction SilentlyContinue
                        $accountIdentifier = $null

                        # Cerca identificatori email (in ordine di priorità)
                        $emailProperties = @('SMTP Address', 'Email', 'Account Name', 'Display Name')
                        foreach ($propName in $emailProperties) {
                            if ($props.PSObject.Properties[$propName]) {
                                $value = $props.$propName
                                if ($value -and $value -match '@') {
                                    $accountIdentifier = $value
                                    break
                                }
                            }
                        }

                        if ($accountIdentifier) {
                            Write-Log "    Account: $accountIdentifier (GUID: $accountGuid)" "INFO"
                        } else {
                            Write-Log "    Account GUID: $accountGuid (email non identificata)" "INFO"
                        }

                        # Converti nome firma in Binary Unicode
                        $signatureBinary = Convert-ToRegistryBinary -Value $SignatureName

                        # IMPOSTA LE FIRME (formato Binary Unicode)
                        Set-ItemProperty -Path $account.PSPath -Name "New Signature" -Value $signatureBinary -Type Binary -Force
                        Set-ItemProperty -Path $account.PSPath -Name "Reply-Forward Signature" -Value $signatureBinary -Type Binary -Force

                        # Verifica scrittura
                        $verifyNew = Get-ItemProperty -Path $account.PSPath -Name "New Signature" -ErrorAction SilentlyContinue
                        $verifyReply = Get-ItemProperty -Path $account.PSPath -Name "Reply-Forward Signature" -ErrorAction SilentlyContinue

                        if ($verifyNew -and $verifyReply) {
                            if ($accountIdentifier) {
                                Write-Log "    [OK] Firma impostata per: $accountIdentifier" "SUCCESS"
                            } else {
                                Write-Log "    [OK] Firma impostata per GUID: $accountGuid" "SUCCESS"
                            }
                            Write-Log "    Percorso: $($account.PSPath)" "INFO"
                            $signaturesSet++
                        } else {
                            Write-Log "    [FAIL] Verifica scrittura fallita per account $accountGuid" "WARNING"
                        }

                    } catch {
                        Write-Log "    Errore impostazione firma per GUID $accountGuid : $_" "ERROR"
                    }
                }
            } else {
                Write-Log "    Nessun account trovato in questo profilo" "WARNING"
            }
        }
    } else {
        Write-Log "Nessun profilo Outlook trovato nel registro" "WARNING"
        Write-Log "Percorso cercato: $profilesBasePath" "INFO"
    }

    # ========================================================================
    # RIEPILOGO FINALE
    # ========================================================================
    Write-Log "==========================================" "INFO"

    if ($signaturesSet -gt 0 -or $mailSettingsWritten) {
        Write-Log "INSTALLAZIONE COMPLETATA CON SUCCESSO" "SUCCESS"
    } else {
        Write-Log "INSTALLAZIONE COMPLETATA CON AVVISI" "WARNING"
        Write-Log "AZIONE RICHIESTA: Verificare configurazione Outlook" "WARNING"
    }

    Write-Log "==========================================" "INFO"
    Write-Log "Firma: $SignatureName" "INFO"
    Write-Log "Posizione file: $LocalPath" "INFO"
    Write-Log "File copiati: $copiedFiles" "INFO"
    Write-Log "File modificabili: $permissionsSet" "SUCCESS"
    Write-Log "Interfaccia Outlook: SBLOCCATA" "SUCCESS"
    Write-Log "" "INFO"
    Write-Log "Account Outlook elaborati: $accountsProcessed" "INFO"
    Write-Log "Firme impostate correttamente: $signaturesSet" "SUCCESS"
    Write-Log "MailSettings configurato: $(if($mailSettingsWritten){'SI'}else{'NO'})" "INFO"
    Write-Log "==========================================" "INFO"
    Write-Log "" "INFO"
    Write-Log "NOTA: Riavviare Outlook per applicare le modifiche" "INFO"
    Write-Log "Log completo: $LogFile" "INFO"
    Write-Log "==========================================" "INFO"

    if ($signaturesSet -eq 0 -and -not $mailSettingsWritten) {
        Write-Log "ATTENZIONE: Nessun account configurato! Verifica:" "WARNING"
        Write-Log "1. Outlook è configurato con almeno un account email?" "WARNING"
        Write-Log "2. L'utente ha eseguito Outlook almeno una volta?" "WARNING"
        Write-Log "3. Verificare manualmente il registro con i comandi nella documentazione" "WARNING"
        exit 2
    }

    exit 0
}
catch {
    Write-Log "ERRORE CRITICO: $_" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
