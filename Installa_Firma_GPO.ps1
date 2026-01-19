# ==============================================================================
# Script: Installa_Firma_GPO.ps1
# Versione: 12.0 - FIX CRITICO: Firma predefinita in formato Binary Unicode
# Autore: Sandro - IT Specialist Carton Group
# ==============================================================================
# Novità v12.0:
# - FIX CRITICO: Conversione valori firma da String a Binary Unicode
# - Outlook ora legge correttamente la firma predefinita (dropdown popolati)
# - Aggiunto fallback MailSettings per compatibilità GPO centralizzate
# - Supporto versioni Office multiple (2016, 2019, 2021, M365)
# - Logging migliorato con percorsi registro per debug
# - Verifica post-scrittura per conferma successo
# ==============================================================================
# Novità v11.0:
# - Imposta automaticamente la firma come predefinita per nuovi messaggi e risposte
# - Funziona anche con account Exchange/M365 con dropdown disabilitato
# ==============================================================================

$NetworkPath = "\\DEAZRADS101\Firme"
$LocalPath = "$env:APPDATA\Microsoft\Signatures"
$LogFile = "$env:TEMP\Installazione_Firma_$(Get-Date -Format 'yyyyMMdd').log"
$SignatureName = "Firma_Aziendale"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage -ErrorAction SilentlyContinue
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

Write-Log "==========================================" "INFO"
Write-Log "INSTALLAZIONE FIRMA v12.0" "INFO"
Write-Log "Utente: $env:USERNAME" "INFO"
Write-Log "Computer: $env:COMPUTERNAME" "INFO"
Write-Log "Data: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "INFO"
Write-Log "==========================================" "INFO"

try {
    # ========================================================================
    # SBLOCCO INTERFACCIA FIRMA OUTLOOK
    # ========================================================================
    Write-Log "Sblocco interfaccia firma Outlook..." "INFO"

    $policyPaths = @(
        "HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Options\Mail",
        "HKCU:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings",
        "HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Setup"
    )

    foreach ($policyPath in $policyPaths) {
        if (Test-Path $policyPath) {
            $blockingKeys = @("DisableSignatures", "NewSignature", "ReplySignature", "DisablePersonalSignatures")
            foreach ($key in $blockingKeys) {
                if (Get-ItemProperty -Path $policyPath -Name $key -ErrorAction SilentlyContinue) {
                    Remove-ItemProperty -Path $policyPath -Name $key -Force -ErrorAction SilentlyContinue
                    Write-Log "  Rimossa policy: $policyPath\$key" "SUCCESS"
                }
            }
        }
    }

    $mailSettingsPath = "HKCU:\Software\Microsoft\Office\16.0\Common\MailSettings"
    if (Test-Path $mailSettingsPath) {
        $blockingKeys = @("NewSignature", "ReplySignature")
        foreach ($key in $blockingKeys) {
            if (Get-ItemProperty -Path $mailSettingsPath -Name $key -ErrorAction SilentlyContinue) {
                Remove-ItemProperty -Path $mailSettingsPath -Name $key -Force -ErrorAction SilentlyContinue
                Write-Log "  Rimossa chiave: $mailSettingsPath\$key" "SUCCESS"
            }
        }
    }

    Write-Log "Interfaccia firma sbloccata" "SUCCESS"

    # ========================================================================
    # VERIFICA RETE E PREREQUISITI
    # ========================================================================
    Write-Log "Verifica connessione server..." "INFO"

    if (-not (Test-Path $NetworkPath)) {
        Write-Log "ERRORE: Server non raggiungibile: $NetworkPath" "ERROR"
        exit 1
    }
    Write-Log "Server raggiungibile" "SUCCESS"

    if (-not (Test-Path $LocalPath)) {
        New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
        Write-Log "Cartella firme creata: $LocalPath" "SUCCESS"
    }

    # ========================================================================
    # CERCA CARTELLA UTENTE SUL SERVER
    # ========================================================================
    $UserFolder = $null
    $PossiblePaths = @()

    Write-Log "Ricerca cartella firma utente..." "INFO"

    $PossiblePaths += "$NetworkPath\$env:USERNAME"
    $PossiblePaths += "$NetworkPath\$($env:USERNAME.ToLower())"

    try {
        $adUser = ([ADSISEARCHER]"samaccountname=$env:USERNAME").FindOne()
        if ($adUser) {
            $UserEmail = $adUser.Properties.mail
            if ($UserEmail) {
                $PossiblePaths += "$NetworkPath\$UserEmail"
                $emailUser = $UserEmail.Split('@')[0]
                $PossiblePaths += "$NetworkPath\$emailUser"
                Write-Log "Email AD trovata: $UserEmail" "INFO"
            }
            $displayName = $adUser.Properties.displayname
            if ($displayName) {
                Write-Log "Nome completo: $displayName" "INFO"
            }
        }
    } catch {
        Write-Log "Impossibile recuperare info da AD (continuo con altri metodi)" "WARNING"
    }

    if ($env:USERNAME -match '\.') {
        $parts = $env:USERNAME.Split('.')
        if ($parts.Count -ge 2) {
            $PossiblePaths += "$NetworkPath\$($parts[1])"
            $PossiblePaths += "$NetworkPath\$($parts[0]).$($parts[1])"
        }
    }

    Write-Log "Ricerca tra $($PossiblePaths.Count) percorsi possibili..." "INFO"

    foreach ($Path in $PossiblePaths) {
        if (Test-Path $Path) {
            $UserFolder = $Path
            Write-Log "Cartella trovata: $UserFolder" "SUCCESS"
            break
        }
    }

    if (-not $UserFolder) {
        $UserFolder = "$NetworkPath\Default"
        if (Test-Path $UserFolder) {
            Write-Log "Uso cartella Default (firma generica)" "WARNING"
        } else {
            Write-Log "Nessuna cartella firma trovata per questo utente" "ERROR"
            Write-Log "Percorsi testati:" "ERROR"
            $PossiblePaths | ForEach-Object { Write-Log "  - $_" "ERROR" }
            exit 1
        }
    }

    # ========================================================================
    # COPIA FILE FIRMA DAL SERVER
    # ========================================================================
    Write-Log "Copia file firma dal server..." "INFO"

    $Files = Get-ChildItem -Path $UserFolder -File -ErrorAction SilentlyContinue

    if (-not $Files -or $Files.Count -eq 0) {
        Write-Log "Nessun file nella cartella $UserFolder" "WARNING"
        exit 0
    }

    Write-Log "Trovati $($Files.Count) file da copiare" "INFO"

    $HtmlContent = ""
    $TextContent = ""
    $copiedFiles = 0

    foreach ($File in $Files) {
        $NewName = $null

        switch ($File.Extension.ToLower()) {
            ".htm"  {
                $NewName = "$SignatureName.htm"
                $HtmlContent = Get-Content $File.FullName -Raw -Encoding UTF8
            }
            ".html" {
                $NewName = "$SignatureName.htm"
                $HtmlContent = Get-Content $File.FullName -Raw -Encoding UTF8
            }
            ".txt"  {
                $NewName = "$SignatureName.txt"
                $TextContent = Get-Content $File.FullName -Raw -Encoding UTF8
            }
        }

        if ($NewName) {
            $DestFile = Join-Path $LocalPath $NewName

            try {
                Copy-Item -Path $File.FullName -Destination $DestFile -Force -ErrorAction Stop
                Write-Log "File copiato: $NewName ($(($File.Length / 1KB).ToString('F2')) KB)" "SUCCESS"
                $copiedFiles++
            } catch {
                Write-Log "Errore copiando $NewName : $_" "ERROR"
            }
        }
    }

    if ($copiedFiles -eq 0) {
        Write-Log "Nessun file copiato" "WARNING"
        exit 0
    }

    # ========================================================================
    # GENERA FILE RTF PER COMPATIBILITÀ OUTLOOK
    # ========================================================================
    Write-Log "Generazione file RTF..." "INFO"

    $RtfText = if ($TextContent) {
        $TextContent
    } else {
        if ($HtmlContent) {
            $HtmlContent -replace '<[^>]+>', '' -replace '&nbsp;', ' ' -replace '&amp;', '&'
        } else {
            "Firma Aziendale"
        }
    }

    $RtfContent = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat\deflang1040{\fonttbl{\f0\fnil\fcharset0 Calibri;}{\f1\fnil Calibri;}}
{\*\generator Riched20 10.0.19041}\viewkind4\uc1
\pard\sa200\sl276\slmult1\f0\fs22\lang16 $($RtfText -replace "`n", "\par`n" -replace "`r", "")\f1\par
}
"@

    $RtfFile = Join-Path $LocalPath "$SignatureName.rtf"

    try {
        $RtfContent | Out-File $RtfFile -Encoding ASCII -Force
        Write-Log "File RTF creato ($(((Get-Item $RtfFile).Length / 1KB).ToString('F2')) KB)" "SUCCESS"
    } catch {
        Write-Log "Errore creazione RTF - $_" "WARNING"
    }

    # ========================================================================
    # RIMOZIONE ATTRIBUTO READONLY - FIRMA MODIFICABILE
    # ========================================================================
    Write-Log "Configurazione attributi per permettere modifiche utente..." "INFO"

    $firmaFiles = @(
        "$LocalPath\$SignatureName.htm",
        "$LocalPath\$SignatureName.txt",
        "$LocalPath\$SignatureName.rtf"
    )

    $filesFixed = 0

    foreach ($filePath in $firmaFiles) {
        if (Test-Path $filePath) {
            try {
                $fileItem = Get-Item $filePath

                if ($fileItem.IsReadOnly) {
                    $fileItem.IsReadOnly = $false
                    Write-Log "  ReadOnly rimosso: $(Split-Path $filePath -Leaf)" "SUCCESS"
                }

                $fileItem.Attributes = $fileItem.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly)
                $fileItem.Attributes = $fileItem.Attributes -band (-bnot [System.IO.FileAttributes]::System)
                $fileItem.Attributes = $fileItem.Attributes -band (-bnot [System.IO.FileAttributes]::Hidden)

                try {
                    $acl = Get-Acl $filePath
                    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

                    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                        $currentUser,
                        "Modify",
                        "Allow"
                    )

                    $acl.SetAccessRule($rule)
                    Set-Acl -Path $filePath -AclObject $acl

                    Write-Log "  Permessi utente impostati: $(Split-Path $filePath -Leaf)" "SUCCESS"
                } catch {
                    Write-Log "  Impossibile modificare permessi ACL (file comunque utilizzabile)" "WARNING"
                }

                $filesFixed++

            } catch {
                Write-Log "  Errore configurazione $(Split-Path $filePath -Leaf): $_" "WARNING"
            }
        }
    }

    Write-Log "File configurati come modificabili: $filesFixed" "SUCCESS"

    # ========================================================================
    # IMPOSTAZIONE FIRMA PREDEFINITA - METODO 1: MailSettings (Fallback)
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

                # ✅ FIX v12.0: Converti nome firma in Binary Unicode
                $signatureBinary = Convert-ToRegistryBinary -Value $SignatureName

                # Imposta le firme predefinite (Binary, non String!)
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
    # IMPOSTAZIONE FIRMA PREDEFINITA - METODO 2: Account Outlook (Principale)
    # ========================================================================
    Write-Log "Metodo 2: Impostazione firma per account Outlook..." "INFO"

    $profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $signaturesSet = 0
    $accountsProcessed = 0

    # Identifica profilo Outlook predefinito
    $defaultProfileName = $null
    try {
        $outlookPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook"
        if (Test-Path $outlookPath) {
            $defaultProfileName = (Get-ItemProperty -Path $outlookPath -Name "DefaultProfile" -ErrorAction SilentlyContinue).'DefaultProfile'
            if ($defaultProfileName) {
                Write-Log "  Profilo predefinito: $defaultProfileName" "INFO"
            }
        }
    } catch {}

    if (Test-Path $profilesPath) {
        $profiles = Get-ChildItem $profilesPath -ErrorAction SilentlyContinue

        # Prioritizza il profilo predefinito
        if ($defaultProfileName) {
            $profiles = $profiles | Sort-Object { $_.PSChildName -ne $defaultProfileName }
        }

        Write-Log "  Trovati $($profiles.Count) profilo/i Outlook" "INFO"

        foreach ($profile in $profiles) {
            $profileName = $profile.PSChildName
            Write-Log "  Elaborazione profilo: $profileName" "INFO"

            $accountsPath = Join-Path $profile.PSPath "9375CFF0413111d3B88A00104B2A6676"

            if (Test-Path $accountsPath) {
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

                        # ✅ FIX v12.0: Converti nome firma in Binary Unicode
                        $signatureBinary = Convert-ToRegistryBinary -Value $SignatureName

                        # IMPOSTA LE FIRME (Binary Unicode, NON String!)
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
        Write-Log "Nessun profilo Outlook trovato (Outlook mai avviato?)" "WARNING"
        Write-Log "Percorso cercato: $profilesPath" "INFO"
    }

    # ========================================================================
    # RIEPILOGO FINALE
    # ========================================================================
    Write-Log "==========================================" "INFO"

    if ($signaturesSet -gt 0 -or $mailSettingsWritten) {
        Write-Log "INSTALLAZIONE COMPLETATA CON SUCCESSO" "SUCCESS"
    } else {
        Write-Log "INSTALLAZIONE COMPLETATA CON AVVISI" "WARNING"
        Write-Log "NOTA: La firma verrà impostata al prossimo avvio Outlook" "INFO"
    }

    Write-Log "==========================================" "INFO"
    Write-Log "Firma: $SignatureName" "INFO"
    Write-Log "Posizione file: $LocalPath" "INFO"
    Write-Log "File copiati: $copiedFiles" "INFO"
    Write-Log "File modificabili: $filesFixed" "SUCCESS"
    Write-Log "Interfaccia Outlook: SBLOCCATA" "SUCCESS"
    Write-Log "" "INFO"
    Write-Log "Account Outlook elaborati: $accountsProcessed" "INFO"
    Write-Log "Firme impostate correttamente: $signaturesSet" "SUCCESS"
    Write-Log "MailSettings configurato: $(if($mailSettingsWritten){'SI'}else{'NO'})" "INFO"
    Write-Log "==========================================" "INFO"
    Write-Log "" "INFO"
    Write-Log "IMPORTANTE: Riavviare Outlook per applicare le modifiche" "INFO"
    Write-Log "Log completo: $LogFile" "INFO"
    Write-Log "==========================================" "INFO"

    if ($signaturesSet -eq 0 -and -not $mailSettingsWritten) {
        Write-Log "ATTENZIONE: Nessun account configurato! Verificare:" "WARNING"
        Write-Log "1. Outlook è configurato con almeno un account email?" "WARNING"
        Write-Log "2. L'utente ha eseguito Outlook almeno una volta?" "WARNING"
        exit 2
    }

    exit 0
}
catch {
    Write-Log "ERRORE CRITICO: $_" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
