# ETL KNIME + Excel/VBA Batch Runner

Questo script illustra il livello di organizzazione di una pipeline di automazione completa. 
Sebbene il flusso di lavoro KNIME e i moduli VBA siano specifici e non possa pubblicarli per questionni di know-how, il framework mostra come integrare più tecnologie con una gestione degli errori e una pianificazione affidabili.

##  Problema risolto
Automazione giornaliera di report che prima richiedeva:
- 1 ora di lavoro manuale (con relativo crosscheck di più fonti)
- Intervento umano per aggiornamenti Excel
- Rischio di errori a causa di processo ripetitivo

## Tecnologie
- **Batch Script** (orchestrazione)
- **VBScript** (automazione Excel - non presente in Repo)
- **KNIME** (ETL workflow - non presente in Repo)
- **Excel VBA** (business logic - non presente in Repo)
- **Windows Task Scheduler** (scheduling)

## Features
- Verifica prerequisiti e sicurezza
- Gestione errori completa con messaggi chiari
- Calcolo automatico date (mese corrente + successivo)
- Import dinamico moduli VBA
- Salvataggio con naming convention standardizzato
- Zero intervento umano richiesto

---

## Caratteristiche

- Verifica accesso all’unità di rete `R:\`.
- Controllo del **Trust access VBA** in Excel.
- Avvio workflow KNIME (`.knwf`) da riga di comando con `-reset`.
- Esecuzione silenziosa di Excel/macros via **VBScript**.
- Import dinamico del modulo `DataProcessing.bas` e chiamata `DataComparison.AVVIO`.
- Export XLSX con naming automatico basato su data corrente e **mese successivo**.
- **Email automatica (VBA)**: genera un messaggio Outlook con:
  - oggetto parametrico (es. “Report YYMM”),
  - **allegati**: i due file generati,
  - **corpo HTML** con **tabella riassuntiva KPI** (totali, variazioni, note).
- Messaggi d’errore esplicativi e **exit code** non-zero.
- Pulizia file temporanei.

---

## Prerequisiti

- **Windows 10/11**
- **KNIME Analytics Platform** in  
  `%USERPROFILE%\AppData\Local\Programs\KNIME\knime.exe`
- **Microsoft Excel** con:
  - *File → Opzioni → Centro protezione → Impostazioni Centro protezione → Impostazioni macro*  
  - Abilitare **“Considera attendibile l’accesso al modello a oggetti dei progetti VBA”**
- **Microsoft Outlook** (per bozza/invio email)
- **Unità `R:\`** mappata e raggiungibile

---

## Come funziona (in breve)

1. **Pre-check**: console, accesso `R:\`, trust Excel (via VBS).
2. **KNIME**: esecuzione workflow batch con `-application org.knime.product.KNIME_BATCH_APPLICATION`.
3. **Date**: calcolo `YY/MM/DD` e **mese/anno successivo** via PowerShell.
4. **Excel/VBA**: per ciascun file (`DataComparison.xlsx`, `DataComparisonNext.xlsx`)
 - import di `DataProcessing.bas`,
 - esecuzione macro `DataComparison.AVVIO`,
 - salvataggio:
   - `YYMM - Report Mensile.xlsx`
   - `YY(next)MM(next) - Report Giornaliero.xlsx`
5. **Email (VBA)**: il modulo crea una **bozza** in Outlook (o invia, se configurato) con **allegati** e **tabella riassuntiva**.

---

