# Guida al deploy — Riconciliazione Contabile Web

## Struttura della webapp

```
webapp/
├── app.py                  ← backend FastAPI (API + sessioni + routing)
├── requirements.txt
├── render.yaml             ← config deploy su Render.com
├── templates/
│   ├── login.html          ← pagina di accesso
│   ├── dashboard.html      ← lista report (operatore + revisore)
│   ├── nuova.html          ← caricamento file (solo operatore)
│   └── report.html         ← dettaglio + download (tutti)
├── static/css/style.css    ← design pulito, responsive
├── uploads/                ← file caricati (temporanei per run)
├── reports/                ← Excel generati (persistenti)
└── data/reports.json       ← storico riconciliazioni
```

## Utenti

| Username | Password default | Ruolo | Cosa può fare |
|----------|-----------------|-------|---------------|
| `operatore` | `operatore123` | Operatore | Upload file, avvia riconciliazione, scarica report |
| `revisore` | `revisore123` | Revisore | Solo visualizza lista e scarica report |

> **Importante:** cambia le password prima del deploy tramite variabili d'ambiente.

---

## Deploy su Render.com (gratuito, 10 minuti)

### 1. Metti il codice su GitHub
```bash
# Crea un repo privato su github.com, poi:
git init
git add .
git commit -m "Prima versione"
git remote add origin https://github.com/TUO-USERNAME/riconciliazione.git
git push -u origin main
```

### 2. Crea il servizio su Render
1. Vai su [render.com](https://render.com) → New → **Web Service**
2. Collega il tuo repo GitHub
3. Render legge automaticamente `render.yaml` — clicca **Deploy**

### 3. Imposta le password (variabili d'ambiente)
Nel pannello Render → Environment:
```
PASS_OPERATORE = la-tua-password-sicura
PASS_REVISORE  = password-del-capo
```

### 4. Deploy completato
Render ti dà un URL tipo: `https://riconciliazione-contabile.onrender.com`

Condividi con il tuo capo solo le credenziali da **revisore**.

---

## Note tecniche

- **Storage**: i report Excel sono salvati sul disco persistente di Render (1 GB incluso nel piano gratuito)
- **Sessioni**: in memoria (si resettano al riavvio — normale per un piano gratuito)
- **Sicurezza**: HTTPS automatico su Render, password hashate SHA-256, cookie httponly
- **Upgrade futuro**: se si aggiungono utenti, basta sostituire `USERS` in `app.py` con un DB PostgreSQL

## Avvio locale (test)

```bash
pip install -r requirements.txt
uvicorn app:app --reload
# apri http://localhost:8000
```
