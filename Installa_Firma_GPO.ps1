# ==============================================================================
# Script: Installa_Firma_GPO.ps1
# Versione: 12.2 - FIX: Scrittura Binary garantita + rimozione policy HKLM
# Autore: Sandro - IT Specialist Carton Group
# ==============================================================================
# Novità v12.2:
# - FIX CRITICO: Usa Remove + New-ItemProperty invece di Set-ItemProperty
#   (Set-ItemProperty può scrivere come String invece di Binary)
# - Rimozione policy HKLM che bloccano dropdown firma (richiede Admin)
# - Verifica tipo registro (Binary vs String) dopo scrittura
# - Retry automatico se prima scrittura fallisce
# ==============================================================================
# Novità v12.1:
# - FIX CRITICO: Verifica post-scrittura ora controlla il valore effettivo
#   (il bug in v12.0 contava come successo anche scritture fallite)
# - Aggiunto confronto byte-array per confermare scrittura corretta
# - Logging migliorato con dettagli sul tipo di errore per debug
# ==============================================================================
# Novità v12.0:
# - FIX CRITICO: Conversione valori firma da String a Binary Unicode
# - Outlook ora legge correttamente la firma predefinita (dropdown popolati)
# - Aggiunto fallback MailSettings per compatibilità GPO centralizzate
# - Supporto versioni Office multiple (2016, 2019, 2021, M365)
# - Logging migliorato con percorsi registro per debug
# ==============================================================================
# Novità v11.0:
# - Imposta automaticamente la firma come predefinita per nuovi messaggi e risposte
# - Funziona anche con account Exchange/M365 con dropdown disabilitato
# ==============================================================================

$NetworkPath = "\\DEAZRADS101\Firme"
$LocalPath = "$env:APPDATA\Microsoft\Signatures"
$LogFile = "$env:TEMP\Firma_Install_Log.txt"
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

# ========================================================================
# FUNZIONE: Verifica se due byte array sono identici (FIX v12.1)
# ========================================================================
function Compare-ByteArrays {
    param(
        [byte[]]$Array1,
        [byte[]]$Array2
    )

    if ($null -eq $Array1 -or $null -eq $Array2) {
        return $false
    }

    if ($Array1.Length -ne $Array2.Length) {
        return $false
    }

    for ($i = 0; $i -lt $Array1.Length; $i++) {
        if ($Array1[$i] -ne $Array2[$i]) {
            return $false
        }
    }

    return $true
}

# ========================================================================
# FUNZIONE: Scrittura sicura Binary nel registro (FIX v12.2)
# Usa Remove + New invece di Set per garantire il tipo Binary
# ========================================================================
function Set-RegistryBinaryValue {
    param(
        [string]$Path,
        [string]$Name,
        [byte[]]$Value
    )

    try {
        # Rimuovi valore esistente (potrebbe essere tipo sbagliato)
        Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue

        # Crea nuovo valore come Binary
        New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType Binary -Force | Out-Null

        # Verifica tipo scritto
        $regKey = Get-Item -Path $Path -ErrorAction SilentlyContinue
        if ($regKey) {
            $valueKind = $regKey.GetValueKind($Name)
            if ($valueKind -eq [Microsoft.Win32.RegistryValueKind]::Binary) {
                return $true
            } else {
                Write-Log "      [WARN] Tipo scritto: $valueKind (atteso: Binary)" "WARNING"
                return $false
            }
        }
        return $true
    } catch {
        Write-Log "      [ERROR] Scrittura fallita: $_" "ERROR"
        return $false
    }
}

Write-Log "==========================================" "INFO"
Write-Log "INSTALLAZIONE FIRMA v12.2" "INFO"
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

    # ✅ FIX v12.2: Rimozione policy HKLM che bloccano dropdown firma (richiede Admin)
    Write-Log "Rimozione policy HKLM bloccanti (richiede privilegi Admin)..." "INFO"

    $hklmPolicyPaths = @(
        "HKLM:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings",
        "HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Options\Mail"
    )

    $hklmBlockingKeys = @(
        "DisableRoamingSignatures",
        "DisableRoamingSignaturesTemporaryToggle",
        "DisableSignatures",
        "NewSignature",
        "ReplySignature"
    )

    foreach ($hklmPath in $hklmPolicyPaths) {
        if (Test-Path $hklmPath) {
            foreach ($key in $hklmBlockingKeys) {
                try {
                    $existingValue = Get-ItemProperty -Path $hklmPath -Name $key -ErrorAction SilentlyContinue
                    if ($existingValue) {
                        Remove-ItemProperty -Path $hklmPath -Name $key -Force -ErrorAction Stop
                        Write-Log "  Rimossa policy HKLM: $hklmPath\$key" "SUCCESS"
                    }
                } catch {
                    if ($_.Exception.Message -match "richiesta|denied|access|accesso") {
                        Write-Log "  [ADMIN] Impossibile rimuovere $key - eseguire come Amministratore" "WARNING"
                    }
                }
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
    $signaturesFailed = 0
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

    # Prepara il byte array di riferimento per la verifica (FIX v12.1)
    $expectedSignatureBinary = Convert-ToRegistryBinary -Value $SignatureName

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

                        # ✅ FIX v12.2: Usa funzione sicura con Remove + New-ItemProperty
                        $signatureBinary = Convert-ToRegistryBinary -Value $SignatureName

                        # Converti PSPath in path standard per Get-Item
                        $accountRegPath = $account.PSPath -replace 'Microsoft\.PowerShell\.Core\\Registry::', ''

                        # IMPOSTA LE FIRME con metodo sicuro (Remove + New)
                        $writeNewOK = Set-RegistryBinaryValue -Path $accountRegPath -Name "New Signature" -Value $signatureBinary
                        $writeReplyOK = Set-RegistryBinaryValue -Path $accountRegPath -Name "Reply-Forward Signature" -Value $signatureBinary

                        # ✅ FIX v12.2: Se prima scrittura fallisce, riprova una volta
                        if (-not $writeNewOK -or -not $writeReplyOK) {
                            Write-Log "    [RETRY] Ritento scrittura..." "WARNING"
                            Start-Sleep -Milliseconds 500
                            if (-not $writeNewOK) {
                                $writeNewOK = Set-RegistryBinaryValue -Path $accountRegPath -Name "New Signature" -Value $signatureBinary
                            }
                            if (-not $writeReplyOK) {
                                $writeReplyOK = Set-RegistryBinaryValue -Path $accountRegPath -Name "Reply-Forward Signature" -Value $signatureBinary
                            }
                        }

                        # ✅ FIX v12.1: Verifica scrittura CORRETTA - confronta i byte effettivi
                        $verifyProps = Get-ItemProperty -Path $accountRegPath -ErrorAction SilentlyContinue
                        $actualNewSig = $verifyProps.'New Signature'
                        $actualReplySig = $verifyProps.'Reply-Forward Signature'

                        $newSigOK = Compare-ByteArrays -Array1 $actualNewSig -Array2 $expectedSignatureBinary
                        $replySigOK = Compare-ByteArrays -Array1 $actualReplySig -Array2 $expectedSignatureBinary

                        if ($newSigOK -and $replySigOK) {
                            if ($accountIdentifier) {
                                Write-Log "    [OK] Firma impostata per: $accountIdentifier" "SUCCESS"
                            } else {
                                Write-Log "    [OK] Firma impostata per GUID: $accountGuid" "SUCCESS"
                            }
                            Write-Log "    Percorso: $accountRegPath" "INFO"
                            $signaturesSet++
                        } else {
                            # ✅ FIX v12.1: Logging dettagliato del fallimento
                            $failReason = ""
                            if (-not $newSigOK -and -not $replySigOK) {
                                $failReason = "entrambe le firme non scritte"
                            } elseif (-not $newSigOK) {
                                $failReason = "New Signature non scritta"
                            } else {
                                $failReason = "Reply-Forward Signature non scritta"
                            }

                            if ($accountIdentifier) {
                                Write-Log "    [FAIL] Verifica fallita per: $accountIdentifier ($failReason)" "WARNING"
                            } else {
                                Write-Log "    [FAIL] Verifica fallita per GUID: $accountGuid ($failReason)" "WARNING"
                            }

                            # ✅ FIX v12.2: Debug info - mostra tipo e valore
                            if ($null -eq $actualNewSig) {
                                Write-Log "    [DEBUG] New Signature: NULL" "WARNING"
                            } elseif ($actualNewSig -is [string]) {
                                Write-Log "    [DEBUG] New Signature: TIPO ERRATO (String invece di Binary)" "WARNING"
                            } else {
                                Write-Log "    [DEBUG] New Signature: $($actualNewSig.Length) bytes (attesi: $($expectedSignatureBinary.Length))" "WARNING"
                            }

                            $signaturesFailed++
                        }

                    } catch {
                        Write-Log "    [ERROR] Errore impostazione firma per GUID $accountGuid : $_" "ERROR"
                        $signaturesFailed++
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
        if ($signaturesFailed -eq 0) {
            Write-Log "INSTALLAZIONE COMPLETATA CON SUCCESSO" "SUCCESS"
        } else {
            Write-Log "INSTALLAZIONE COMPLETATA CON AVVISI" "WARNING"
            Write-Log "Account con errori: $signaturesFailed (potrebbero richiedere intervento manuale)" "WARNING"
        }
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
    if ($signaturesFailed -gt 0) {
        Write-Log "Firme con errori: $signaturesFailed" "WARNING"
    }
    Write-Log "MailSettings configurato: $(if($mailSettingsWritten){'SI'}else{'NO'})" "INFO"
    Write-Log "==========================================" "INFO"
    Write-Log "" "INFO"
    Write-Log "IMPORTANTE: Notifica utente..." "INFO"

    # Notifica utente (se possibile)
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $notify = New-Object System.Windows.Forms.NotifyIcon
        $notify.Icon = [System.Drawing.SystemIcons]::Information
        $notify.BalloonTipTitle = "Firma Aziendale"
        $notify.BalloonTipText = "La firma email è stata installata. Riavvia Outlook per applicarla."
        $notify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $notify.Visible = $true
        $notify.ShowBalloonTip(10000)
        Start-Sleep -Seconds 2
        $notify.Dispose()
        Write-Log "Notifica mostrata all'utente" "SUCCESS"
    } catch {
        Write-Log "Impossibile mostrare notifica: $_" "WARNING"
    }

    Write-Log "" "INFO"
    Write-Log "Log completo: $LogFile" "INFO"
    Write-Log "==========================================" "INFO"

    if ($signaturesSet -eq 0 -and -not $mailSettingsWritten) {
        Write-Log "ATTENZIONE: Nessun account configurato! Verificare:" "WARNING"
        Write-Log "1. Outlook è configurato con almeno un account email?" "WARNING"
        Write-Log "2. L'utente ha eseguito Outlook almeno una volta?" "WARNING"
        exit 2
    }

    # ✅ FIX v12.1: Exit code diversi per distinguere successo parziale
    if ($signaturesFailed -gt 0) {
        exit 3  # Successo parziale - alcuni account non configurati
    }

    exit 0
}
catch {
    Write-Log "ERRORE CRITICO: $_" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
    exit 1
}
