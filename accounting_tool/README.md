# Accounting Reconciliation Tool
**Riconciliazione contabile automatica** — estratti conto ↔ partitari NAV ↔ gestione contestazioni

---

## Struttura cartelle

```
accounting_tool/
│
├── reconcilia.py          ← script principale
├── genera_demo.py         ← genera file demo e template
├── README.md
│
├── input/
│   ├── estratti/          ← metti qui gli estratti conto (PDF o Excel)
│   ├── partitari/         ← metti qui i partitari (Excel/CSV da NAV)
│   ├── fatture/           ← opzionale: registri fatture separati
│   │
│   ├── TEMPLATE_estratto_conto.xlsx   ← template da compilare
│   ├── TEMPLATE_partitario.xlsx
│   └── TEMPLATE_contestazioni.xlsx
│
├── output/                ← i report Excel vengono salvati qui
└── logs/                  ← log di ogni esecuzione
```

---

## Installazione (una volta sola)

```bash
pip install pandas openpyxl pdfplumber python-dateutil
```

---

## Utilizzo rapido

### 1. Genera i file demo per vedere come funziona
```bash
python genera_demo.py
```

### 2. Esegui la riconciliazione con i file demo
```bash
python reconcilia.py \
  --estratto  input/estratti/estratto_giugno2025_DEMO.xlsx \
  --fornitori input/partitari/partitario_fornitori_DEMO.xlsx \
  --clienti   input/partitari/partitario_clienti_DEMO.xlsx \
  --contestazioni input/contestazioni_DEMO.xlsx \
  --output    report_giugno2025.xlsx
```

### 3. Con i tuoi file reali
```bash
python reconcilia.py \
  --estratto  input/estratti/estratto_luglio.pdf \
  --fornitori input/partitari/partitario_forn_nav.xlsx \
  --clienti   input/partitari/partitario_cli_nav.xlsx \
  --emesse    input/fatture/fatture_emesse.xlsx \
  --ricevute  input/fatture/fatture_ricevute.xlsx \
  --contestazioni input/contestazioni.xlsx \
  --output    report_luglio2025.xlsx
```

---

## Opzioni avanzate

| Parametro | Default | Descrizione |
|-----------|---------|-------------|
| `--tol-importo` | 0.01 € | Tolleranza su importo (es. arrotondamenti) |
| `--tol-giorni` | 3 | Giorni di scarto accettati tra data estratto e scadenza |
| `--no-desc` | off | Disabilita l'abbinamento tramite descrizione |
| `--usa-data` | off | Attiva abbinamento per data scadenza |

---

## Formato file di input

### Estratto conto (Excel/CSV)
| Colonna | Obbligatoria | Note |
|---------|-------------|------|
| Data Operazione | ✅ | Qualsiasi formato data italiano |
| Descrizione | ✅ | Testo libero — lo script cerca il numero fattura |
| Dare | ✅* | *Almeno Dare o Avere |
| Avere | ✅* | |
| Saldo | ❌ | Opzionale |

> **Per i PDF**: lo script legge automaticamente le tabelle. Se il PDF è scansionato (immagine) contatta il tuo responsabile IT per attivare OCR.

### Partitario (da NAV o Excel manuale)
| Colonna | Obbligatoria | Note |
|---------|-------------|------|
| Numero Documento | ✅ | Deve corrispondere a quanto scritto nelle descrizioni dell'estratto |
| Tipo | ✅ | `Emessa` (cliente) o `Ricevuta` (fornitore) |
| Controparte | ✅ | Ragione sociale |
| Data Documento | ✅ | |
| Data Scadenza | ❌ | Usata se attivi `--usa-data` |
| Importo | ✅ | |
| Stato | ❌ | Es. "Aperta", "Pagata" |
| Contestata | ❌ | `Sì` / `No` |
| Motivo Contestazione | ❌ | Testo libero |

### File contestazioni (opzionale — integra dati del gestionale)
| Colonna | Obbligatoria | Note |
|---------|-------------|------|
| Numero Documento | ✅ | Deve corrispondere al partitario |
| Stato Contestazione | ✅ | Valori: `contestata_da_noi`, `contestata_dal_cliente`, `in_attesa_nota_credito`, `risolta`, `aperta` |
| Data Contestazione | ❌ | |
| Motivo | ❌ | |
| Responsabile | ❌ | |

---

## Report di output (Excel)

Il file Excel generato contiene 6 fogli:

| Foglio | Contenuto |
|--------|-----------|
| 📊 **Riepilogo** | KPI: tasso riconciliazione, totale entrate/uscite, conteggi per stato |
| 🏦 **Estratto Conto** | Tutti i movimenti con esito abbinamento e documento abbinato |
| 📋 **Partitario Fornitori** | Fatture passive con stato riconciliazione e flag contestazione |
| 📋 **Partitario Clienti** | Fatture attive con stato riconciliazione e flag contestazione |
| 🚩 **Contestazioni** | Solo le partite contestate (da noi o dal cliente) |
| ⚠️ **Anomalie** | Movimenti non identificati + fatture non riconciliate + contestate |

### Colori nei fogli
| Colore | Significato |
|--------|-------------|
| 🟢 Verde | Abbinata |
| 🟡 Giallo | Parziale (importo diverso) |
| 🔴 Rosa | Non riconciliata |
| 🟣 Viola | Movimento non identificato |
| 🔴 Rosso scuro | Contestata |

---

## Logica di abbinamento

Lo script usa un sistema a punteggio:

1. **Importo** — confronto diretto con tolleranza configurabile
2. **Numero documento nella descrizione** — cerca il codice fattura nel campo descrizione dell'estratto (es. "FRN-0088" o "FATT-2025-041")
3. **Data** — opzionale, confronto data scadenza ↔ data movimento ±N giorni

Una partita viene dichiarata **Abbinata** se il miglior candidato supera una soglia di score. **Parziale** se l'importo differisce ma il documento è identificabile.

---

## Integrazione con NAV

Per esportare i partitari da NAV/Business Central:
1. Vai su **Movimenti contabili fornitori** o **Movimenti contabili clienti**
2. Applica filtri (es. per periodo)
3. Esporta su Excel
4. Salva in `input/partitari/`

Le colonne riconosciute automaticamente includono le naming convention standard di NAV (Nr. documento, Data registrazione, Importo, Descrizione, ecc.).

---

## Processo consigliato mensile

```
1. Esporta estratto conto da banca (Excel o PDF)
2. Esporta partitari da NAV per il mese
3. Aggiorna file contestazioni se ci sono nuove dispute
4. Lancia: python reconcilia.py [parametri]
5. Apri il report → foglio Anomalie → gestisci le partite non riconciliate
6. Archivia il report in output/ con nome mese/anno
```
