# ==============================================================================
# Script: Installa_Firma_GPO.ps1
# Versione: 12.3 - FIX: Scrittura inline (risolve path PSPath)
# ==============================================================================

$NetworkPath = "\\DEAZRADS101\Firme"
$LocalPath = "$env:APPDATA\Microsoft\Signatures"
$LogFile = "$env:TEMP\Firma_Install_Log.txt"
$SignatureName = "Firma_Aziendale"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    Add-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message" -ErrorAction SilentlyContinue
}

function Convert-ToRegistryBinary {
    param([string]$Value)
    return [System.Text.Encoding]::Unicode.GetBytes($Value + "`0")
}

function Compare-ByteArrays {
    param([byte[]]$Array1, [byte[]]$Array2)
    if ($null -eq $Array1 -or $null -eq $Array2) { return $false }
    if ($Array1.Length -ne $Array2.Length) { return $false }
    for ($i = 0; $i -lt $Array1.Length; $i++) {
        if ($Array1[$i] -ne $Array2[$i]) { return $false }
    }
    return $true
}

Write-Log "=========================================="
Write-Log "INSTALLAZIONE FIRMA v12.3"
Write-Log "Utente: $env:USERNAME | PC: $env:COMPUTERNAME"
Write-Log "=========================================="

try {
    # === SBLOCCO INTERFACCIA ===
    Write-Log "Sblocco interfaccia firma Outlook..."

    $policyPaths = @(
        "HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Options\Mail",
        "HKCU:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings",
        "HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\Setup"
    )
    $blockingKeys = @("DisableSignatures", "NewSignature", "ReplySignature", "DisablePersonalSignatures")

    foreach ($path in $policyPaths) {
        if (Test-Path $path) {
            foreach ($key in $blockingKeys) {
                Remove-ItemProperty -Path $path -Name $key -Force -ErrorAction SilentlyContinue
            }
        }
    }

    # HKLM (richiede Admin)
    $hklmPaths = @(
        "HKLM:\Software\Policies\Microsoft\Office\16.0\Common\MailSettings",
        "HKLM:\Software\Policies\Microsoft\Office\16.0\Outlook\Options\Mail"
    )
    $hklmKeys = @("DisableRoamingSignatures", "DisableRoamingSignaturesTemporaryToggle", "DisableSignatures")

    foreach ($path in $hklmPaths) {
        if (Test-Path $path) {
            foreach ($key in $hklmKeys) {
                try { Remove-ItemProperty -Path $path -Name $key -Force -ErrorAction Stop } catch {}
            }
        }
    }
    Write-Log "Interfaccia sbloccata" "SUCCESS"

    # === VERIFICA RETE ===
    if (-not (Test-Path $NetworkPath)) {
        Write-Log "Server non raggiungibile: $NetworkPath" "ERROR"
        exit 1
    }
    Write-Log "Server raggiungibile" "SUCCESS"

    if (-not (Test-Path $LocalPath)) {
        New-Item -ItemType Directory -Path $LocalPath -Force | Out-Null
    }

    # === CERCA CARTELLA UTENTE ===
    $UserFolder = $null
    $PossiblePaths = @("$NetworkPath\$env:USERNAME", "$NetworkPath\$($env:USERNAME.ToLower())")

    try {
        $adUser = ([ADSISEARCHER]"samaccountname=$env:USERNAME").FindOne()
        if ($adUser -and $adUser.Properties.mail) {
            $email = $adUser.Properties.mail[0]
            $PossiblePaths += "$NetworkPath\$email"
            $PossiblePaths += "$NetworkPath\$($email.Split('@')[0])"
            Write-Log "Email AD: $email"
        }
    } catch {}

    foreach ($p in $PossiblePaths) {
        if (Test-Path $p) { $UserFolder = $p; break }
    }
    if (-not $UserFolder) {
        $UserFolder = "$NetworkPath\Default"
        if (-not (Test-Path $UserFolder)) { Write-Log "Nessuna cartella firma" "ERROR"; exit 1 }
    }
    Write-Log "Cartella: $UserFolder" "SUCCESS"

    # === COPIA FILE ===
    $Files = Get-ChildItem -Path $UserFolder -File -ErrorAction SilentlyContinue
    $copiedFiles = 0; $HtmlContent = ""; $TextContent = ""

    foreach ($File in $Files) {
        $NewName = $null
        switch ($File.Extension.ToLower()) {
            ".htm"  { $NewName = "$SignatureName.htm"; $HtmlContent = Get-Content $File.FullName -Raw -Encoding UTF8 }
            ".html" { $NewName = "$SignatureName.htm"; $HtmlContent = Get-Content $File.FullName -Raw -Encoding UTF8 }
            ".txt"  { $NewName = "$SignatureName.txt"; $TextContent = Get-Content $File.FullName -Raw -Encoding UTF8 }
        }
        if ($NewName) {
            Copy-Item -Path $File.FullName -Destination (Join-Path $LocalPath $NewName) -Force
            $copiedFiles++
        }
    }
    Write-Log "File copiati: $copiedFiles" "SUCCESS"

    # === GENERA RTF ===
    $RtfText = if ($TextContent) { $TextContent } else { $HtmlContent -replace '<[^>]+>', '' -replace '&nbsp;', ' ' }
    "{\rtf1\ansi $($RtfText -replace "`n", "\par ")}" | Out-File (Join-Path $LocalPath "$SignatureName.rtf") -Encoding ASCII -Force

    # === RIMUOVI READONLY ===
    @("htm", "txt", "rtf") | ForEach-Object {
        $f = Join-Path $LocalPath "$SignatureName.$_"
        if (Test-Path $f) { (Get-Item $f).IsReadOnly = $false }
    }
    Write-Log "File modificabili" "SUCCESS"

    # === MAILSETTINGS ===
    $signatureBinary = Convert-ToRegistryBinary -Value $SignatureName
    foreach ($ver in @("16.0", "15.0", "14.0")) {
        $msPath = "HKCU:\Software\Microsoft\Office\$ver\Common\MailSettings"
        if (Test-Path "HKCU:\Software\Microsoft\Office\$ver") {
            if (-not (Test-Path $msPath)) { New-Item -Path $msPath -Force | Out-Null }
            Set-ItemProperty -Path $msPath -Name "NewSignature" -Value $signatureBinary -Type Binary -Force
            Set-ItemProperty -Path $msPath -Name "ReplySignature" -Value $signatureBinary -Type Binary -Force
            Write-Log "MailSettings $ver OK" "SUCCESS"
        }
    }

    # === ACCOUNT OUTLOOK ===
    Write-Log "Impostazione firma per account..."
    $profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
    $signaturesSet = 0; $signaturesFailed = 0
    $expectedBinary = Convert-ToRegistryBinary -Value $SignatureName

    if (Test-Path $profilesPath) {
        foreach ($profile in (Get-ChildItem $profilesPath)) {
            Write-Log "  Profilo: $($profile.PSChildName)"
            $accountsPath = Join-Path $profile.PSPath "9375CFF0413111d3B88A00104B2A6676"

            if (Test-Path $accountsPath) {
                foreach ($account in (Get-ChildItem $accountsPath)) {
                    $guid = $account.PSChildName
                    $props = Get-ItemProperty $account.PSPath -ErrorAction SilentlyContinue
                    $email = $null
                    foreach ($p in @('SMTP Address', 'Email', 'Account Name')) {
                        if ($props.$p -and $props.$p -match '@') { $email = $props.$p; break }
                    }
                    Write-Log "    Account: $(if($email){$email}else{$guid})"

                    try {
                        # ✅ SCRITTURA INLINE - usa $account.PSPath direttamente
                        Remove-ItemProperty -Path $account.PSPath -Name "New Signature" -Force -ErrorAction SilentlyContinue
                        Remove-ItemProperty -Path $account.PSPath -Name "Reply-Forward Signature" -Force -ErrorAction SilentlyContinue
                        New-ItemProperty -Path $account.PSPath -Name "New Signature" -Value $signatureBinary -PropertyType Binary -Force | Out-Null
                        New-ItemProperty -Path $account.PSPath -Name "Reply-Forward Signature" -Value $signatureBinary -PropertyType Binary -Force | Out-Null

                        $check = Get-ItemProperty -Path $account.PSPath
                        if ((Compare-ByteArrays $check.'New Signature' $expectedBinary) -and (Compare-ByteArrays $check.'Reply-Forward Signature' $expectedBinary)) {
                            Write-Log "    [OK]" "SUCCESS"
                            $signaturesSet++
                        } else {
                            Write-Log "    [FAIL] Verifica fallita" "WARNING"
                            $signaturesFailed++
                        }
                    } catch {
                        Write-Log "    [ERROR] $_" "ERROR"
                        $signaturesFailed++
                    }
                }
            }
        }
    }

    # === RIEPILOGO ===
    Write-Log "=========================================="
    Write-Log "Firme OK: $signaturesSet" "SUCCESS"
    if ($signaturesFailed -gt 0) { Write-Log "Firme FAIL: $signaturesFailed" "WARNING" }
    Write-Log "=========================================="

    if ($signaturesFailed -gt 0) { exit 3 }
    exit 0
}
catch {
    Write-Log "ERRORE CRITICO: $_" "ERROR"
    exit 1
}
