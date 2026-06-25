"""
Accounting Reconciliation Web App
Utenti: Operatore (tu) + Revisore (capo, solo lettura)
"""
import os
import json
import uuid
import hashlib
import secrets
from pathlib import Path
from datetime import datetime
from typing import Optional

from fastapi import (
    FastAPI, Request, UploadFile, File, Form,
    HTTPException, Depends, Response
)
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

# importa il motore di riconciliazione (stesso script già costruito)
import sys
sys.path.insert(0, str(Path(__file__).parent.parent / "accounting_tool"))
from reconcilia import AccountingProcessor

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
BASE_DIR    = Path(__file__).parent
UPLOAD_DIR  = BASE_DIR / "uploads"
REPORT_DIR  = BASE_DIR / "reports"
DATA_FILE   = BASE_DIR / "data" / "reports.json"
SESSION_DIR = BASE_DIR / "data"

# Credenziali — in produzione usa variabili d'ambiente
USERS = {
    "operatore": {
        "password_hash": hashlib.sha256(os.getenv("PASS_OPERATORE", "operatore123").encode()).hexdigest(),
        "role": "operatore",
    },
    "revisore": {
        "password_hash": hashlib.sha256(os.getenv("PASS_REVISORE", "revisore123").encode()).hexdigest(),
        "role": "revisore",
    },
}

app = FastAPI(title="Riconciliazione Contabile")
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
app.mount("/reports-file", StaticFiles(directory=REPORT_DIR), name="reports-file")
templates = Jinja2Templates(directory=BASE_DIR / "templates")

# ---------------------------------------------------------------------------
# SESSION HELPERS (cookie-based, semplice)
# ---------------------------------------------------------------------------
SESSIONS: dict = {}

def get_session(request: Request) -> Optional[dict]:
    sid = request.cookies.get("session_id")
    return SESSIONS.get(sid)

def require_login(request: Request) -> dict:
    session = get_session(request)
    if not session:
        raise HTTPException(status_code=302, headers={"Location": "/login"})
    return session

def require_operatore(request: Request) -> dict:
    session = require_login(request)
    if session["role"] != "operatore":
        raise HTTPException(status_code=403, detail="Accesso negato")
    return session

# ---------------------------------------------------------------------------
# REPORT STORE (JSON semplice, no DB)
# ---------------------------------------------------------------------------
def load_reports() -> list:
    if not DATA_FILE.exists():
        return []
    return json.loads(DATA_FILE.read_text())

def save_report(meta: dict):
    reports = load_reports()
    reports.insert(0, meta)
    DATA_FILE.write_text(json.dumps(reports, indent=2, ensure_ascii=False))

# ---------------------------------------------------------------------------
# ROUTES — AUTH
# ---------------------------------------------------------------------------
@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, error: str = ""):
    return templates.TemplateResponse("login.html", {"request": request, "error": error})

@app.post("/login")
async def login(
    response: Response,
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
):
    user = USERS.get(username)
    ph   = hashlib.sha256(password.encode()).hexdigest()
    if not user or user["password_hash"] != ph:
        return templates.TemplateResponse(
            "login.html",
            {"request": request, "error": "Credenziali non valide"},
            status_code=401,
        )
    sid = secrets.token_hex(32)
    SESSIONS[sid] = {"username": username, "role": user["role"]}
    resp = RedirectResponse("/", status_code=302)
    resp.set_cookie("session_id", sid, httponly=True, samesite="lax")
    return resp

@app.get("/logout")
async def logout(request: Request):
    sid = request.cookies.get("session_id")
    SESSIONS.pop(sid, None)
    resp = RedirectResponse("/login", status_code=302)
    resp.delete_cookie("session_id")
    return resp

# ---------------------------------------------------------------------------
# ROUTES — DASHBOARD (tutti i ruoli)
# ---------------------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    session = get_session(request)
    if not session:
        return RedirectResponse("/login", status_code=302)
    reports = load_reports()
    return templates.TemplateResponse("dashboard.html", {
        "request": request,
        "session": session,
        "reports": reports,
    })

# ---------------------------------------------------------------------------
# ROUTES — NUOVA RICONCILIAZIONE (solo operatore)
# ---------------------------------------------------------------------------
@app.get("/nuova", response_class=HTMLResponse)
async def nuova_page(request: Request):
    session = get_session(request)
    if not session or session["role"] != "operatore":
        return RedirectResponse("/", status_code=302)
    return templates.TemplateResponse("nuova.html", {"request": request, "session": session})

@app.post("/elabora")
async def elabora(
    request: Request,
    estratto:    UploadFile = File(...),
    fornitori:   Optional[UploadFile] = File(None),
    clienti:     Optional[UploadFile] = File(None),
    emesse:      Optional[UploadFile] = File(None),
    ricevute:    Optional[UploadFile] = File(None),
    contestazioni: Optional[UploadFile] = File(None),
    periodo:     str = Form(""),
    note:        str = Form(""),
    tol_importo: float = Form(0.01),
    tol_giorni:  int   = Form(3),
):
    session = get_session(request)
    if not session or session["role"] != "operatore":
        return RedirectResponse("/login", status_code=302)

    run_id   = str(uuid.uuid4())[:8]
    run_dir  = UPLOAD_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    async def save_upload(f: Optional[UploadFile], name: str) -> Optional[Path]:
        if not f or not f.filename:
            return None
        dest = run_dir / name
        dest.write_bytes(await f.read())
        return dest

    p_estratto      = await save_upload(estratto,      "estratto" + Path(estratto.filename).suffix)
    p_fornitori     = await save_upload(fornitori,     "fornitori.xlsx") if fornitori else None
    p_clienti       = await save_upload(clienti,       "clienti.xlsx")   if clienti  else None
    p_emesse        = await save_upload(emesse,        "emesse.xlsx")    if emesse   else None
    p_ricevute      = await save_upload(ricevute,      "ricevute.xlsx")  if ricevute else None
    p_contestazioni = await save_upload(contestazioni, "contestazioni.xlsx") if contestazioni else None

    report_name = f"report_{run_id}.xlsx"
    report_path = REPORT_DIR / report_name

    try:
        proc = AccountingProcessor(config={
            "tolleranza_importo": tol_importo,
            "tolleranza_giorni":  tol_giorni,
            "usa_desc": True,
        })
        # redirige l'output nella cartella reports della webapp
        import accounting_tool.reconcilia as rec_mod
        orig_output = rec_mod.OUTPUT_DIR
        rec_mod.OUTPUT_DIR = REPORT_DIR

        proc.run(
            estratto_path         = p_estratto,
            partitario_forn_path  = p_fornitori,
            partitario_cli_path   = p_clienti,
            fatture_emesse_path   = p_emesse,
            fatture_ricevute_path = p_ricevute,
            contestazioni_path    = p_contestazioni,
            output_name           = report_name,
        )
        rec_mod.OUTPUT_DIR = orig_output
        stato = "completato"
        errore = ""
    except Exception as e:
        stato  = "errore"
        errore = str(e)

    meta = {
        "id":          run_id,
        "periodo":     periodo or datetime.now().strftime("%B %Y"),
        "note":        note,
        "data":        datetime.now().strftime("%d/%m/%Y %H:%M"),
        "operatore":   session["username"],
        "report_file": report_name,
        "stato":       stato,
        "errore":      errore,
    }
    save_report(meta)
    return RedirectResponse(f"/report/{run_id}", status_code=302)

# ---------------------------------------------------------------------------
# ROUTES — DETTAGLIO REPORT (tutti i ruoli)
# ---------------------------------------------------------------------------
@app.get("/report/{run_id}", response_class=HTMLResponse)
async def report_detail(request: Request, run_id: str):
    session = get_session(request)
    if not session:
        return RedirectResponse("/login", status_code=302)
    reports = load_reports()
    report  = next((r for r in reports if r["id"] == run_id), None)
    if not report:
        raise HTTPException(status_code=404, detail="Report non trovato")
    return templates.TemplateResponse("report.html", {
        "request": request,
        "session": session,
        "report":  report,
    })

@app.get("/download/{run_id}")
async def download_report(request: Request, run_id: str):
    session = get_session(request)
    if not session:
        return RedirectResponse("/login", status_code=302)
    reports = load_reports()
    report  = next((r for r in reports if r["id"] == run_id), None)
    if not report:
        raise HTTPException(status_code=404)
    path = REPORT_DIR / report["report_file"]
    if not path.exists():
        raise HTTPException(status_code=404, detail="File non trovato")
    return FileResponse(
        path,
        filename=f"riconciliazione_{report['periodo'].replace(' ','_')}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
