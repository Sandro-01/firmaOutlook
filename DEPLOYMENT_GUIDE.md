# 🚀 DEPLOYMENT GUIDE - v12.0 FIX FIRMA PREDEFINITA

**Script:** Installa_Firma_GPO.ps1 v12.0
**Bug risolto:** Dropdown Outlook vuoti nonostante firma installata
**Deployment stimato:** 15 minuti (test pilota) + 24-48h (rollout completo)

---

## ⚡ QUICK START (Per IT Admin)

### 1️⃣ BACKUP VERSIONE CORRENTE

```powershell
# Sul Domain Controller o file server
$scriptPath = "\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1"
$backupPath = "\\DEAZRADS101\Firme\Scripts\Backup\Installa_Firma_GPO_v11_$(Get-Date -Format 'yyyyMMdd_HHmmss').ps1"

Copy-Item $scriptPath $backupPath
Write-Host "✅ Backup creato: $backupPath" -ForegroundColor Green
```

### 2️⃣ DEPLOY v12.0

**Opzione A: Direct Replacement (produzione immediata)**

```powershell
# Sostituisci direttamente lo script
Copy-Item ".\Installa_Firma_GPO.ps1" "\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1" -Force

# Verifica deployment
$version = Get-Content "\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1" | Select-String "# Versione:"
Write-Host $version -ForegroundColor Cyan

# Output atteso: # Versione: 12.0 - FIX CRITICO: Firma predefinita in formato Binary Unicode
```

**Opzione B: Test Pilota (raccomandato)**

```powershell
# 1. Copia script con nome test
Copy-Item ".\Installa_Firma_GPO.ps1" "\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO_v12.ps1" -Force

# 2. Crea VBS wrapper per test
$vbsTest = @'
Set objShell = CreateObject("WScript.Shell")
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ""\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO_v12.ps1""", 0, True
'@

Set-Content "\\DEAZRADS101\Firme\Scripts\Installa_Firma_Invisibile_v12_TEST.vbs" -Value $vbsTest

# 3. Esegui su 5-10 utenti test
# Metodo manuale: doppio click su Installa_Firma_Invisibile_v12_TEST.vbs
# Oppure via GPO temporanea solo su OU test
```

### 3️⃣ TEST IMMEDIATO

```powershell
# Esegui su PC utente test
\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1

# Verifica log (ultimi 30 righe)
Get-ChildItem $env:TEMP\Installazione_Firma*.log |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1 |
    Get-Content -Tail 30

# 🔍 CERCA QUESTE RIGHE NEL LOG:
# [SUCCESS] Firma impostata in MailSettings (Office 16.0)
# [OK] Firma impostata per: nome.cognome@cartongrp.com
# [INFO] Percorso: Microsoft.PowerShell.Core\Registry::HKEY_CURRENT_USER\...
# Firme impostate correttamente: 1 (o più se account multipli)
```

### 4️⃣ VERIFICA IN OUTLOOK

```
1. Chiudi completamente Outlook (anche icona system tray)
2. Riapri Outlook
3. Nuovo messaggio → Firma → Verifica dropdown "Firma_Aziendale" selezionato ✅
4. File → Opzioni → Posta → Firme → Verifica dropdown popolati ✅
```

### 5️⃣ ROLLOUT COMPLETO (dopo 24-48h test OK)

```powershell
# Sostituisci script produzione
Copy-Item "\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO_v12.ps1" `
          "\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1" -Force

# Lo script verrà eseguito automaticamente su tutti i PC al prossimo:
# - Logon utente
# - Unlock workstation
# - Daily alle 12:00
# (trigger configurati nella GPO "CARTONGRP - Distribuzione Firme Outlook")

Write-Host "✅ ROLLOUT COMPLETO ATTIVATO" -ForegroundColor Green
```

---

## 🧪 TEST CHECKLIST

### Pre-Deployment
- [ ] Backup v11.0 eseguito
- [ ] v12.0 scaricato e verificato
- [ ] 5-10 utenti test identificati (includere almeno 1 VPN user)

### Post-Deployment Test Pilota (utenti test)
- [ ] Script eseguito senza errori
- [ ] Log contiene `[SUCCESS]` per firma impostata
- [ ] Dropdown Outlook popolati
- [ ] Firma applicata automaticamente a nuove email
- [ ] Firma applicata a risposte/inoltri
- [ ] Test con account multipli (se applicabile)
- [ ] Test da VPN (se applicabile)

### Post-Deployment Produzione (48h dopo rollout)
- [ ] Nessun ticket helpdesk per firme mancanti
- [ ] Spot check su 20-30 utenti random
- [ ] Verifica utenti VPN/remote
- [ ] Log GPO execution (Event Viewer)

---

## 📊 MONITORING POST-DEPLOYMENT

### Verifica esecuzione GPO via Event Viewer

```powershell
# Su PC utente, apri Event Viewer o usa PowerShell:
Get-WinEvent -LogName "Microsoft-Windows-TaskScheduler/Operational" -MaxEvents 20 |
    Where-Object { $_.Message -like "*Aggiornamento_Firma_Email*" } |
    Format-Table TimeCreated, Id, Message -AutoSize
```

### Raccolta log centralizzata (opzionale)

```powershell
# Crea share per log centralizzati
$logShare = "\\DEAZRADS101\Firme\Logs"
if (-not (Test-Path $logShare)) {
    New-Item -ItemType Directory -Path $logShare -Force
}

# Modifica script per copiare log:
# Aggiungi alla fine di Installa_Firma_GPO.ps1 (riga ~390):

$centralLog = "$logShare\$env:COMPUTERNAME`_$env:USERNAME`_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Copy-Item $LogFile $centralLog -ErrorAction SilentlyContinue

# Poi analizza con:
Get-ChildItem $logShare\*.log |
    Where-Object { $_.LastWriteTime -gt (Get-Date).AddDays(-1) } |
    ForEach-Object {
        $success = Select-String -Path $_.FullName -Pattern "\[SUCCESS\]" -AllMatches
        [PSCustomObject]@{
            File = $_.Name
            User = $_.Name.Split('_')[1]
            SuccessCount = $success.Matches.Count
            Size = $_.Length
        }
    } | Format-Table -AutoSize
```

---

## 🔧 TROUBLESHOOTING RAPIDO

### Problema: Script si esegue ma dropdown ancora vuoti

**Diagnosi:**

```powershell
# 1. Verifica tipo registro (DEVE essere Binary)
Get-ItemProperty "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\*" |
    ForEach-Object {
        try {
            $path = $_.PSPath
            $type = (Get-Item $path).GetValueKind('New Signature')
            Write-Host "Tipo: $type (deve essere Binary)" -ForegroundColor $(if($type -eq 'Binary'){'Green'}else{'Red'})
        } catch {}
    }
```

**Soluzione rapida (manuale):**

```powershell
# Forza esecuzione script con log verbose
\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1

# Controlla log completo
notepad $env:TEMP\Installazione_Firma_$(Get-Date -Format 'yyyyMMdd').log

# Se ancora non funziona: reset completo firma
Remove-Item "$env:APPDATA\Microsoft\Signatures\*" -Force -Recurse
\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1
```

### Problema: "Nessun account email trovato"

**Diagnosi:**

```powershell
# Verifica profili Outlook
Get-ChildItem "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

# Se vuoto: utente non ha mai aperto Outlook
```

**Soluzione:**

1. Apri Outlook e configura account email
2. Chiudi Outlook completamente
3. Esegui script: `\\DEAZRADS101\Firme\Scripts\Installa_Firma_GPO.ps1`

### Problema: "Server non raggiungibile"

**Diagnosi:**

```powershell
# Verifica connettività
Test-Path "\\DEAZRADS101\Firme"

# Se False: problema di rete o VPN
```

**Soluzione:**

```powershell
# Verifica DNS
nslookup DEAZRADS101

# Ping server
Test-Connection DEAZRADS101 -Count 2

# Verifica VPN (se utente remoto)
Get-NetIPConfiguration | Where-Object { $_.InterfaceDescription -like "*VPN*" }

# Test SMB
Test-NetConnection DEAZRADS101 -Port 445
```

### Problema: GPO non si applica

**Diagnosi:**

```powershell
# Verifica GPO applicata
gpresult /r /scope:user | Select-String "CARTONGRP"

# Output atteso: CARTONGRP - Distribuzione Firme Outlook

# Verifica Scheduled Task
Get-ScheduledTask -TaskName "Aggiornamento_Firma_Email" -ErrorAction SilentlyContinue
```

**Soluzione:**

```powershell
# Forza aggiornamento GPO
gpupdate /force

# Verifica OU corretta
whoami /fqdn
# Output esempio: CN=User,OU=Treviso,OU=rIT,OU=Client,DC=adds,DC=cartongrp,DC=com

# Esegui manualmente task
schtasks /Run /TN "Aggiornamento_Firma_Email"
```

---

## 📞 SUPPORTO UTENTI

### Script diagnostico per helpdesk

```powershell
# Salva come: Diagnosi_Firma_Outlook.ps1

Write-Host "=== DIAGNOSI FIRMA OUTLOOK ===" -ForegroundColor Cyan
Write-Host ""

# 1. Info utente
Write-Host "Utente: $env:USERNAME" -ForegroundColor Yellow
Write-Host "Computer: $env:COMPUTERNAME" -ForegroundColor Yellow
Write-Host "Dominio: $env:USERDOMAIN" -ForegroundColor Yellow
Write-Host ""

# 2. Connettività server
Write-Host "Connettività server firme..." -ForegroundColor Yellow
$serverOk = Test-Path "\\DEAZRADS101\Firme"
Write-Host "  Server raggiungibile: $(if($serverOk){'SI ✅'}else{'NO ❌'})" -ForegroundColor $(if($serverOk){'Green'}else{'Red'})
Write-Host ""

# 3. File firma locali
Write-Host "File firma locali..." -ForegroundColor Yellow
$signatures = Get-ChildItem "$env:APPDATA\Microsoft\Signatures" -ErrorAction SilentlyContinue
if ($signatures) {
    Write-Host "  Trovati $($signatures.Count) file" -ForegroundColor Green
    $signatures | ForEach-Object { Write-Host "    - $($_.Name)" }
} else {
    Write-Host "  NESSUN FILE FIRMA TROVATO ❌" -ForegroundColor Red
}
Write-Host ""

# 4. Registro firma predefinita
Write-Host "Impostazioni registro firma..." -ForegroundColor Yellow
$accounts = Get-ChildItem "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676" -ErrorAction SilentlyContinue

if ($accounts) {
    foreach ($acc in $accounts) {
        $props = Get-ItemProperty $acc.PSPath -ErrorAction SilentlyContinue
        $newSig = $props.'New Signature'

        if ($newSig) {
            $type = (Get-Item $acc.PSPath).GetValueKind('New Signature')
            $typeOk = $type -eq 'Binary'

            Write-Host "  Account: $($acc.PSChildName)" -ForegroundColor $(if($typeOk){'Green'}else{'Red'})
            Write-Host "    Tipo: $type $(if($typeOk){'✅'}else{'❌ ERRORE: deve essere Binary'})" -ForegroundColor $(if($typeOk){'Green'}else{'Red'})

            if ($typeOk) {
                $sigName = [System.Text.Encoding]::Unicode.GetString($newSig).TrimEnd("`0")
                Write-Host "    Firma: $sigName" -ForegroundColor Green
            }
        } else {
            Write-Host "  Account $($acc.PSChildName): FIRMA NON IMPOSTATA ❌" -ForegroundColor Red
        }
    }
} else {
    Write-Host "  NESSUN ACCOUNT OUTLOOK CONFIGURATO ❌" -ForegroundColor Red
}
Write-Host ""

# 5. Log ultimo script
Write-Host "Log ultimo script..." -ForegroundColor Yellow
$lastLog = Get-ChildItem $env:TEMP\Installazione_Firma*.log -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime -Descending |
           Select-Object -First 1

if ($lastLog) {
    Write-Host "  Log: $($lastLog.FullName)" -ForegroundColor Green
    Write-Host "  Data: $($lastLog.LastWriteTime)" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Ultime 10 righe:" -ForegroundColor Cyan
    Get-Content $lastLog.FullName -Tail 10 | ForEach-Object { Write-Host "    $_" }
} else {
    Write-Host "  NESSUN LOG TROVATO ❌" -ForegroundColor Red
}
Write-Host ""

# 6. Outlook in esecuzione
Write-Host "Stato Outlook..." -ForegroundColor Yellow
$outlookRunning = Get-Process outlook -ErrorAction SilentlyContinue
if ($outlookRunning) {
    Write-Host "  Outlook in esecuzione: SI (chiudere e riaprire per applicare modifiche)" -ForegroundColor Yellow
} else {
    Write-Host "  Outlook in esecuzione: NO ✅" -ForegroundColor Green
}
Write-Host ""

Write-Host "=== FINE DIAGNOSI ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Premi un tasto per chiudere..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
```

### Istruzioni per utenti finali

```
PROBLEMA: La firma non è impostata come predefinita in Outlook

SOLUZIONE RAPIDA:

1. Chiudi completamente Outlook (anche icona vicino orologio)

2. Apri Esplora File e digita nella barra indirizzo:
   \\DEAZRADS101\Firme\Scripts

3. Doppio click su: Installa_Firma_Invisibile.vbs

4. Attendi 10 secondi (non appare nulla)

5. Riapri Outlook

6. Nuovo messaggio → Verifica firma presente in basso

Se il problema persiste:
- Contatta IT Helpdesk
- Specifica: "Firma Outlook dropdown vuoti dopo v12.0"
- Allega: %TEMP%\Installazione_Firma_[DATA].log
```

---

## 📋 CHECKLIST ROLLOUT COMPLETO

### Prima del deployment
- [ ] README.md del progetto aggiornato con info v12.0
- [ ] Team IT notificato del deployment
- [ ] Helpdesk informato del fix e possibili chiamate
- [ ] Backup script v11.0 eseguito
- [ ] Script v12.0 testato in laboratorio

### Durante deployment
- [ ] Utenti test selezionati (5-10 utenti, inclusi VPN)
- [ ] Script v12.0 deployato su server test
- [ ] Test manuale eseguito su utenti pilota
- [ ] Log verificati per conferma successo
- [ ] Outlook verificato manualmente su 3+ utenti test

### 24h dopo test pilota
- [ ] Nessun errore segnalato da utenti test
- [ ] Log analizzati: 100% successi
- [ ] Dropdown Outlook verificati funzionanti
- [ ] Firme applicate automaticamente
- [ ] Approvazione per rollout produzione

### Rollout produzione
- [ ] Script v12.0 copiato su share produzione
- [ ] VBS wrapper aggiornato (se necessario)
- [ ] GPO verificata attiva su tutte le OU
- [ ] Comunicazione interna inviata
- [ ] Monitoring attivato

### 48h dopo rollout produzione
- [ ] Spot check 20-30 utenti random
- [ ] Ticket helpdesk analizzati (< 5% issue rate accettabile)
- [ ] Log centralizzati analizzati
- [ ] Utenti VPN/remote verificati
- [ ] Documentazione aggiornata
- [ ] Chiusura progetto fix

---

## ✅ SUCCESS CRITERIA

**Il deployment è considerato SUCCESSO se:**

- ✅ 95%+ utenti hanno dropdown Outlook popolati
- ✅ 0 errori critici nei log script
- ✅ < 5% ticket helpdesk correlati
- ✅ Firma applicata automaticamente a nuove email
- ✅ Supporto account multipli funzionante
- ✅ Utenti VPN/remote senza problemi

**KPI Target:**
- Risoluzione bug: 100% (da 5% utenti affetti a 0%)
- Tempo deployment: < 72h (incluso test pilota)
- Ticket helpdesk: < 10 (su ~300 utenti)

---

## 📚 DOCUMENTAZIONE CORRELATA

- `BUG_ANALYSIS_AND_FIX.md` - Analisi tecnica dettagliata del bug
- `Installa_Firma_GPO.ps1` - Script v12.0 corretto
- `README.md` - Documentazione progetto completa

---

## 🆘 CONTATTI SUPPORTO

**IT Specialist:** Sandro - Carton Group Italia
**Email:** [inserire email IT]
**Ticket System:** [inserire sistema ticketing]

**Per emergenze:**
- Rollback immediato: Ripristinare backup v11.0
- Script manuale: Eseguire direttamente .ps1 su PC utente
- Workaround temporaneo: Impostare manualmente firma in Outlook (File → Opzioni → Posta → Firme)

---

*Deployment Guide v1.0 - Generato automaticamente per il fix v12.0*
*Last Update: 2026-01-17*
