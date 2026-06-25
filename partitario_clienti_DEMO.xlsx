"""
=============================================================================
ACCOUNTING RECONCILIATION TOOL
Processo completo: lettura estratti conto (PDF/Excel) → normalizzazione →
riconciliazione con partitari NAV → gestione contestazioni → report Excel
=============================================================================
"""

import os
import re
import sys
import json
import logging
from pathlib import Path
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import pdfplumber
from dateutil import parser as dateparser
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# CONFIGURAZIONE
# ---------------------------------------------------------------------------
BASE_DIR     = Path(__file__).parent
INPUT_DIR    = BASE_DIR / "input"
OUTPUT_DIR   = BASE_DIR / "output"
LOG_DIR      = BASE_DIR / "logs"

TOLERANZA_IMPORTO  = 0.01   # euro – tolleranza su amount match
TOLLERANZA_GIORNI  = 3      # giorni – finestra date match
SOGLIA_MATCH_DESC  = 0.75   # similitudine minima su descrizione (0-1)

# Colori Excel (ARGB)
C_HEADER     = "FF1F3A5F"
C_ABBINATA   = "FFD6F5E3"
C_PARZIALE   = "FFFFF3CD"
C_NON_REC    = "FFFDE8E8"
C_NON_ID     = "FFEDE8FD"
C_CONTESTATA = "FFFFD6D6"
C_TEXT_HDR   = "FFFFFFFF"
C_GRAY_ROW   = "FFF8F8F8"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger(__name__)


# ===========================================================================
# 1.  READER – legge PDF e Excel normalizzando in DataFrame standard
# ===========================================================================

class BankStatementReader:
    """Legge estratto conto da PDF o Excel/CSV e restituisce un DataFrame
    con colonne: data, descrizione, dare, avere, saldo, riferimento"""

    COLONNE_STD = ["data", "descrizione", "dare", "avere", "saldo", "riferimento", "fonte"]

    # pattern comuni nelle descrizioni per estrarre numero fattura
    _RIF_PATTERN = re.compile(
        r"(fatt[a-z]*[\s\-#\.]*[\d/\-]+|n[°º\.\s]*[\d/\-]+|inv[\s\-#\.]*[\d/\-]+)",
        re.IGNORECASE,
    )

    def read(self, path: Path) -> pd.DataFrame:
        ext = path.suffix.lower()
        log.info(f"Lettura estratto conto: {path.name}  (formato: {ext})")
        if ext == ".pdf":
            return self._read_pdf(path)
        elif ext in (".xlsx", ".xls"):
            return self._read_excel(path)
        elif ext == ".csv":
            return self._read_csv(path)
        else:
            raise ValueError(f"Formato non supportato: {ext}")

    # --- PDF ---
    def _read_pdf(self, path: Path) -> pd.DataFrame:
        rows = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                if tables:
                    for table in tables:
                        for row in table[1:]:  # skip header
                            if row and any(row):
                                rows.append(self._parse_pdf_row(row))
                else:
                    # fallback testo libero
                    text = page.extract_text() or ""
                    for line in text.splitlines():
                        parsed = self._parse_text_line(line)
                        if parsed:
                            rows.append(parsed)
        return self._to_dataframe(rows, fonte=path.name)

    def _parse_pdf_row(self, row: list) -> dict:
        """Adatta in base al numero di colonne rilevato."""
        row = [str(c).strip() if c else "" for c in row]
        result = {"data": None, "descrizione": "", "dare": 0.0, "avere": 0.0, "saldo": None}
        if len(row) >= 4:
            result["data"]        = self._parse_date(row[0])
            result["descrizione"] = row[1]
            dare_val  = self._parse_amount(row[2])
            avere_val = self._parse_amount(row[3])
            result["dare"]  = dare_val  if dare_val  else 0.0
            result["avere"] = avere_val if avere_val else 0.0
            if len(row) >= 5:
                result["saldo"] = self._parse_amount(row[4])
        result["riferimento"] = self._extract_ref(result["descrizione"])
        return result

    def _parse_text_line(self, line: str) -> Optional[dict]:
        date_match = re.search(r"\b(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{2,4})\b", line)
        amount_matches = re.findall(r"[\-]?\d{1,3}(?:\.\d{3})*(?:,\d{2})?", line)
        if not date_match or not amount_matches:
            return None
        desc = line[date_match.end():].strip()
        amounts = [self._parse_amount(a) for a in amount_matches]
        amounts = [a for a in amounts if a is not None]
        return {
            "data": self._parse_date(date_match.group()),
            "descrizione": desc[:120],
            "dare":  amounts[0] if amounts and amounts[0] < 0 else 0.0,
            "avere": amounts[0] if amounts and amounts[0] > 0 else 0.0,
            "saldo": amounts[1] if len(amounts) > 1 else None,
            "riferimento": self._extract_ref(desc),
        }

    # --- Excel / CSV ---
    def _read_excel(self, path: Path) -> pd.DataFrame:
        xl = pd.read_excel(path, sheet_name=0, header=None, dtype=str)
        return self._detect_and_parse(xl, fonte=path.name)

    def _read_csv(self, path: Path) -> pd.DataFrame:
        for sep in [";", ",", "\t"]:
            try:
                df = pd.read_csv(path, sep=sep, dtype=str, encoding="utf-8-sig")
                if len(df.columns) >= 3:
                    return self._detect_and_parse(df, fonte=path.name)
            except Exception:
                continue
        raise ValueError(f"Impossibile leggere {path.name} come CSV")

    def _detect_and_parse(self, df: pd.DataFrame, fonte: str) -> pd.DataFrame:
        """Rileva automaticamente le colonne tramite euristica."""
        header_row = self._find_header_row(df)
        if header_row is not None:
            df.columns = df.iloc[header_row].str.lower().str.strip()
            df = df.iloc[header_row + 1:].reset_index(drop=True)

        col_map = self._map_columns(df.columns.tolist())
        rows = []
        for _, row in df.iterrows():
            r = {}
            r["data"]        = self._parse_date(str(row.get(col_map.get("data",""), "")))
            r["descrizione"] = str(row.get(col_map.get("descrizione",""), ""))[:120]
            r["dare"]        = self._parse_amount(str(row.get(col_map.get("dare",""), ""))) or 0.0
            r["avere"]       = self._parse_amount(str(row.get(col_map.get("avere",""), ""))) or 0.0
            r["saldo"]       = self._parse_amount(str(row.get(col_map.get("saldo",""), "")))
            # se c'è una colonna importo unica con segno
            if col_map.get("importo"):
                imp = self._parse_amount(str(row.get(col_map["importo"],"")) or "") or 0.0
                r["dare"]  = abs(imp) if imp < 0 else 0.0
                r["avere"] = imp       if imp > 0 else 0.0
            r["riferimento"] = self._extract_ref(r["descrizione"])
            if r["data"]:
                rows.append(r)
        return self._to_dataframe(rows, fonte=fonte)

    # --- Utilities ---
    def _find_header_row(self, df: pd.DataFrame) -> Optional[int]:
        kw = {"data", "date", "descrizione", "description", "importo", "dare", "avere"}
        for i, row in df.iterrows():
            cells = {str(c).lower().strip() for c in row if pd.notna(c)}
            if len(cells & kw) >= 2:
                return i
        return None

    def _map_columns(self, cols: list) -> dict:
        mapping = {}
        patterns = {
            "data":        ["data", "date", "data valuta", "data operazione", "data mov"],
            "descrizione": ["descrizione", "description", "causale", "motivo", "note"],
            "dare":        ["dare", "addebito", "uscita", "debit"],
            "avere":       ["avere", "accredito", "entrata", "credit"],
            "saldo":       ["saldo", "balance", "saldo finale"],
            "importo":     ["importo", "amount", "valore"],
        }
        for key, terms in patterns.items():
            for col in cols:
                if any(t in str(col).lower() for t in terms):
                    mapping[key] = col
                    break
        return mapping

    def _parse_date(self, s: str) -> Optional[datetime]:
        if not s or s in ("nan", "None", ""):
            return None
        try:
            return dateparser.parse(s, dayfirst=True)
        except Exception:
            return None

    def _parse_amount(self, s: str) -> Optional[float]:
        if not s or s in ("nan", "None", ""):
            return None
        s = s.replace(" ", "").replace("€", "").replace("EUR", "")
        # formato italiano: 1.234,56
        if re.search(r"\d\.\d{3},", s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None

    def _extract_ref(self, desc: str) -> str:
        m = self._RIF_PATTERN.search(desc)
        return m.group().strip() if m else ""

    def _to_dataframe(self, rows: list, fonte: str) -> pd.DataFrame:
        if not rows:
            log.warning("Nessun movimento estratto.")
            return pd.DataFrame(columns=self.COLONNE_STD)
        df = pd.DataFrame(rows)
        df["fonte"] = fonte
        for col in self.COLONNE_STD:
            if col not in df.columns:
                df[col] = None
        df["dare"]  = pd.to_numeric(df["dare"],  errors="coerce").fillna(0.0)
        df["avere"] = pd.to_numeric(df["avere"], errors="coerce").fillna(0.0)
        df = df[df["data"].notna()].copy()
        df.sort_values("data", inplace=True)
        df.reset_index(drop=True, inplace=True)
        log.info(f"  → {len(df)} movimenti estratti")
        return df[self.COLONNE_STD]


# ===========================================================================
# 2.  LEDGER READER – legge partitario NAV / Excel
# ===========================================================================

class LedgerReader:
    """Legge partitari (fornitori/clienti) e registri fatture da Excel/CSV."""

    COLONNE_STD = [
        "numero_documento", "tipo",          # Emessa / Ricevuta
        "controparte",      "data_documento",
        "data_scadenza",    "importo",
        "stato",            "note",
        "contestata",       "motivo_contestazione",
    ]

    def read(self, path: Path, tipo_default: str = "Emessa") -> pd.DataFrame:
        ext = path.suffix.lower()
        log.info(f"Lettura partitario: {path.name}")
        if ext in (".xlsx", ".xls"):
            df_raw = pd.read_excel(path, dtype=str)
        elif ext == ".csv":
            df_raw = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        else:
            raise ValueError(f"Formato non supportato: {ext}")
        return self._normalize(df_raw, tipo_default)

    def _normalize(self, df: pd.DataFrame, tipo_default: str) -> pd.DataFrame:
        df.columns = df.columns.str.lower().str.strip()
        col_map = self._map_columns(df.columns.tolist())
        out = pd.DataFrame()
        out["numero_documento"]     = df.get(col_map.get("numero_documento", "__x__"), pd.Series(dtype=str))
        out["tipo"]                 = df.get(col_map.get("tipo", "__x__"), tipo_default)
        out["controparte"]          = df.get(col_map.get("controparte", "__x__"), "")
        out["data_documento"]       = df.get(col_map.get("data_documento", "__x__"), "").apply(self._parse_date)
        out["data_scadenza"]        = df.get(col_map.get("data_scadenza", "__x__"), "").apply(self._parse_date)
        out["importo"]              = pd.to_numeric(
            df.get(col_map.get("importo", "__x__"), "").apply(self._clean_amount),
            errors="coerce"
        ).fillna(0.0)
        out["stato"]                = df.get(col_map.get("stato", "__x__"), "Aperta")
        out["note"]                 = df.get(col_map.get("note", "__x__"), "")
        out["contestata"]           = df.get(col_map.get("contestata", "__x__"), "No")
        out["motivo_contestazione"] = df.get(col_map.get("motivo_contestazione", "__x__"), "")
        # colonne mancanti
        for col in self.COLONNE_STD:
            if col not in out.columns:
                out[col] = ""
        out = out[self.COLONNE_STD].copy()
        out["_matched"]     = False
        out["_match_banca"] = ""
        out["_esito"]       = "Non riconciliata"
        log.info(f"  → {len(out)} documenti caricati")
        return out

    def _map_columns(self, cols: list) -> dict:
        mapping = {}
        patterns = {
            "numero_documento":     ["num", "numero", "n°", "doc", "fattura", "invoice", "rif"],
            "tipo":                 ["tipo", "type", "direzione"],
            "controparte":          ["fornitore", "cliente", "controparte", "partner", "ragione"],
            "data_documento":       ["data doc", "data fatt", "data emis", "invoice date", "data"],
            "data_scadenza":        ["scadenza", "due date", "pagamento"],
            "importo":              ["importo", "amount", "totale", "valore", "imponibile"],
            "stato":                ["stato", "status", "pagata", "saldato"],
            "note":                 ["note", "notes", "memo"],
            "contestata":           ["contest", "disputa", "reclamo"],
            "motivo_contestazione": ["motivo", "reason", "descrizione contest"],
        }
        for key, terms in patterns.items():
            for col in cols:
                if any(t in col.lower() for t in terms):
                    mapping[key] = col
                    break
        return mapping

    def _parse_date(self, s) -> Optional[datetime]:
        if not s or str(s) in ("nan", "None", ""):
            return None
        try:
            return dateparser.parse(str(s), dayfirst=True)
        except Exception:
            return None

    def _clean_amount(self, s) -> str:
        if not s:
            return ""
        s = str(s).replace(" ", "").replace("€", "").replace("EUR", "")
        if re.search(r"\d\.\d{3},", s):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", ".")
        return s


# ===========================================================================
# 3.  RECONCILIATION ENGINE
# ===========================================================================

class ReconciliationEngine:
    """Abbina movimenti bancari con partite contabili."""

    def __init__(
        self,
        tolleranza_importo: float = TOLERANZA_IMPORTO,
        tolleranza_giorni:  int   = TOLLERANZA_GIORNI,
        usa_desc:           bool  = True,
        usa_data:           bool  = False,
    ):
        self.tol_imp  = tolleranza_importo
        self.tol_days = tolleranza_giorni
        self.usa_desc = usa_desc
        self.usa_data = usa_data

    def reconcile(
        self,
        estratto:  pd.DataFrame,
        partitario: pd.DataFrame,
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        """Restituisce (estratto_annotato, partitario_annotato)."""

        estratto  = estratto.copy()
        partitario = partitario.copy()

        estratto["_esito"]      = "Non identificato"
        estratto["_match_doc"]  = ""
        estratto["_match_cont"] = ""

        for idx_p, part in partitario.iterrows():
            if part["_matched"]:
                continue
            importo_target = abs(part["importo"])
            candidates = []

            for idx_e, mov in estratto.iterrows():
                if mov["_esito"] != "Non identificato":
                    continue

                # importo
                amount_banca = mov["avere"] if part["tipo"] == "Emessa" else mov["dare"]
                diff_imp = abs(amount_banca - importo_target)
                if diff_imp > self.tol_imp + importo_target * 0.001:
                    # prova tolleranza
                    if diff_imp > self.tol_imp:
                        continue

                score = 1.0 - (diff_imp / (importo_target + 0.01))

                # numero documento nella descrizione
                num_doc = str(part["numero_documento"])
                if self.usa_desc and num_doc and num_doc != "nan":
                    if num_doc.lower() in str(mov["descrizione"]).lower():
                        score += 1.0
                    elif str(mov["riferimento"]).lower() in num_doc.lower():
                        score += 0.5

                # data
                if self.usa_data and part["data_scadenza"] and mov["data"]:
                    delta = abs((mov["data"] - part["data_scadenza"]).days)
                    if delta <= self.tol_days:
                        score += 0.5 * (1 - delta / self.tol_days)

                candidates.append((idx_e, score, diff_imp))

            if not candidates:
                continue

            # miglior candidato
            candidates.sort(key=lambda x: (-x[1], x[2]))
            best_idx, best_score, best_diff = candidates[0]

            if best_score > 0.5:
                esito = "Abbinata" if best_diff <= self.tol_imp else "Parziale"
                partitario.at[idx_p, "_matched"]     = True
                partitario.at[idx_p, "_match_banca"] = str(estratto.at[best_idx, "data"])[:10] + \
                                                        f" | {estratto.at[best_idx, 'descrizione'][:40]}"
                partitario.at[idx_p, "_esito"]       = esito

                estratto.at[best_idx, "_esito"]      = esito
                estratto.at[best_idx, "_match_doc"]  = num_doc
                estratto.at[best_idx, "_match_cont"] = str(part["controparte"])

        # partite non abbinate
        partitario.loc[~partitario["_matched"], "_esito"] = "Non riconciliata"

        log.info(
            f"Riconciliazione completata: "
            f"{(partitario['_esito']=='Abbinata').sum()} abbinate, "
            f"{(partitario['_esito']=='Non riconciliata').sum()} non riconciliate, "
            f"{(partitario['_esito']=='Parziale').sum()} parziali"
        )
        return estratto, partitario


# ===========================================================================
# 4.  CONTESTAZIONI MANAGER
# ===========================================================================

class ContestationManager:
    """Carica e applica lo stato contestazioni al partitario."""

    STATI_VALIDI = {
        "contestata_da_noi",       # noi contestiamo al fornitore
        "contestata_dal_cliente",  # cliente contesta a noi
        "in_attesa_nota_credito",
        "risolta",
        "aperta",
    }

    def load_contestazioni(self, path: Path) -> pd.DataFrame:
        """Legge file Excel/CSV con colonne: numero_documento, stato_contestazione,
        data_contestazione, motivo, responsabile."""
        log.info(f"Lettura contestazioni: {path.name}")
        if path.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(path, dtype=str)
        else:
            df = pd.read_csv(path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
        df.columns = df.columns.str.lower().str.strip()
        return df

    def apply(self, partitario: pd.DataFrame, contestazioni: pd.DataFrame) -> pd.DataFrame:
        """Unisce le contestazioni al partitario."""
        if contestazioni.empty:
            return partitario
        col_doc = self._find_col(contestazioni.columns, ["numero_documento", "num", "doc", "fattura"])
        col_stato = self._find_col(contestazioni.columns, ["stato_contestazione", "stato", "status"])
        col_motivo = self._find_col(contestazioni.columns, ["motivo", "reason", "note"])
        col_resp = self._find_col(contestazioni.columns, ["responsabile", "resp", "owner"])

        merged = partitario.merge(
            contestazioni[[c for c in [col_doc, col_stato, col_motivo, col_resp] if c]].rename(
                columns={
                    col_doc:   "numero_documento",
                    col_stato: "_stato_cont",
                    col_motivo:"_motivo_cont",
                    col_resp:  "_responsabile",
                }
            ),
            on="numero_documento",
            how="left",
        )
        # aggiorna campi contestazione
        mask = merged["_stato_cont"].notna()
        merged.loc[mask, "contestata"]           = "Sì"
        merged.loc[mask, "motivo_contestazione"] = merged.loc[mask, "_motivo_cont"].fillna("")
        merged.loc[mask & (merged["_stato_cont"].str.lower() != "risolta"), "_esito"] = "Contestata"
        for c in ["_stato_cont", "_motivo_cont", "_responsabile"]:
            if c in merged.columns:
                merged.drop(columns=[c], inplace=True)
        log.info(f"  → {mask.sum()} contestazioni applicate")
        return merged

    def _find_col(self, cols, terms):
        for c in cols:
            if any(t in c.lower() for t in terms):
                return c
        return None


# ===========================================================================
# 5.  EXCEL REPORT GENERATOR
# ===========================================================================

class ReportGenerator:
    """Genera report Excel professionale con fogli separati."""

    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)

    def _header_style(self, cell, bg=C_HEADER, fg=C_TEXT_HDR):
        cell.font      = Font(bold=True, color=fg, size=10)
        cell.fill      = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def _border(self):
        s = Side(style="thin", color="FFD0D0D0")
        return Border(left=s, right=s, top=s, bottom=s)

    def _row_fill(self, esito: str) -> PatternFill:
        c = {
            "Abbinata":        C_ABBINATA,
            "Parziale":        C_PARZIALE,
            "Non riconciliata":C_NON_REC,
            "Non identificato":C_NON_ID,
            "Contestata":      C_CONTESTATA,
        }.get(esito, "FFFFFFFF")
        return PatternFill("solid", fgColor=c)

    def _write_sheet(self, title: str, headers: list, rows: list, col_widths: list = None):
        ws = self.wb.create_sheet(title=title)
        ws.freeze_panes = "A2"
        for j, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=j, value=h)
            self._header_style(cell)
            ws.row_dimensions[1].height = 30

        for i, row in enumerate(rows, 2):
            esito = row[-1] if row else ""
            fill  = self._row_fill(esito)
            bg    = C_GRAY_ROW if i % 2 == 0 else "FFFFFFFF"
            for j, val in enumerate(row, 1):
                cell = ws.cell(row=i, column=j, value=val)
                cell.fill      = fill if esito not in ("", None) else PatternFill("solid", fgColor=bg)
                cell.border    = self._border()
                cell.alignment = Alignment(vertical="center", wrap_text=False)
                cell.font      = Font(size=9)
                if isinstance(val, datetime):
                    cell.number_format = "DD/MM/YYYY"
                elif isinstance(val, float):
                    cell.number_format = '#,##0.00 €'

        if col_widths:
            for j, w in enumerate(col_widths, 1):
                ws.column_dimensions[get_column_letter(j)].width = w
        return ws

    # ---- fogli ----

    def add_riepilogo(self, estratto: pd.DataFrame, partitario: pd.DataFrame):
        tot_entrate    = estratto["avere"].sum()
        tot_uscite     = estratto["dare"].sum()
        abbinate       = (partitario["_esito"] == "Abbinata").sum()
        parziali       = (partitario["_esito"] == "Parziale").sum()
        non_rec        = (partitario["_esito"] == "Non riconciliata").sum()
        contestate     = (partitario["_esito"] == "Contestata").sum()
        non_id         = (estratto["_esito"] == "Non identificato").sum()
        totale_doc     = len(partitario)
        perc           = round(abbinate / totale_doc * 100, 1) if totale_doc else 0

        rows = [
            ["ESTRATTO CONTO", "", ""],
            ["Totale movimenti",      len(estratto),                  ""],
            ["Totale entrate",        tot_entrate,                    "€"],
            ["Totale uscite",         tot_uscite,                     "€"],
            ["Saldo netto",           tot_entrate - tot_uscite,       "€"],
            ["", "", ""],
            ["RICONCILIAZIONE", "", ""],
            ["Totale documenti",      totale_doc,                     ""],
            ["Abbinate",              abbinate,                       ""],
            ["Parziali",              parziali,                       ""],
            ["Non riconciliate",      non_rec,                        ""],
            ["Contestate",            contestate,                     ""],
            ["Movimenti non id.",     non_id,                         ""],
            ["Tasso riconciliazione", f"{perc}%",                     ""],
            ["", "", ""],
            ["Generato il",           datetime.now().strftime("%d/%m/%Y %H:%M"), ""],
        ]
        ws = self.wb.create_sheet(title="📊 Riepilogo", index=0)
        ws.freeze_panes = "A1"
        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 8
        for i, (label, val, unit) in enumerate(rows, 1):
            ws.cell(row=i, column=1, value=label).font = Font(
                bold=("ESTRATTO" in label or "RICONCILIAZIONE" in label), size=10
            )
            cell_val = ws.cell(row=i, column=2, value=val)
            cell_val.font = Font(size=10)
            if isinstance(val, float):
                cell_val.number_format = '#,##0.00'
            ws.cell(row=i, column=3, value=unit).font = Font(size=10, color="FF888888")

    def add_estratto(self, df: pd.DataFrame):
        headers = ["Data", "Descrizione", "Dare (€)", "Avere (€)", "Saldo (€)", "Riferimento", "Fonte", "Esito", "Doc. abbinato", "Controparte"]
        rows = []
        for _, r in df.iterrows():
            rows.append([
                r["data"],
                r["descrizione"],
                r["dare"] if r["dare"] else None,
                r["avere"] if r["avere"] else None,
                r["saldo"],
                r["riferimento"],
                r["fonte"],
                r["_esito"],
                r["_match_doc"],
                r["_match_cont"],
            ])
        self._write_sheet(
            "🏦 Estratto Conto", headers, rows,
            [12, 40, 12, 12, 12, 14, 16, 16, 14, 20],
        )

    def add_partitario(self, df: pd.DataFrame, title: str = "Partitario"):
        headers = [
            "N° Documento", "Tipo", "Controparte", "Data Doc.", "Scadenza",
            "Importo (€)", "Stato", "Contestata", "Motivo Contestazione",
            "Note", "Esito Ric.", "Match Banca",
        ]
        rows = []
        for _, r in df.iterrows():
            rows.append([
                r["numero_documento"],
                r["tipo"],
                r["controparte"],
                r["data_documento"],
                r["data_scadenza"],
                r["importo"],
                r["stato"],
                r["contestata"],
                r["motivo_contestazione"],
                r["note"],
                r["_esito"],
                r["_match_banca"],
            ])
        self._write_sheet(
            title, headers, rows,
            [14, 10, 22, 12, 12, 12, 12, 10, 28, 20, 16, 40],
        )

    def add_anomalie(self, estratto: pd.DataFrame, partitario: pd.DataFrame):
        headers = ["Tipo Anomalia", "Riferimento", "Controparte", "Importo (€)", "Data", "Dettaglio", "Esito"]
        rows = []
        for _, r in estratto[estratto["_esito"] == "Non identificato"].iterrows():
            rows.append([
                "Movimento non identificato",
                r["riferimento"] or "-",
                "-",
                r["dare"] if r["dare"] else r["avere"],
                r["data"],
                r["descrizione"][:60],
                r["_esito"],
            ])
        for _, r in partitario[partitario["_esito"] == "Non riconciliata"].iterrows():
            rows.append([
                "Fattura non riconciliata",
                r["numero_documento"],
                r["controparte"],
                r["importo"],
                r["data_scadenza"],
                r["note"][:60] if r["note"] else "",
                r["_esito"],
            ])
        for _, r in partitario[partitario["_esito"] == "Contestata"].iterrows():
            rows.append([
                "Documento contestato",
                r["numero_documento"],
                r["controparte"],
                r["importo"],
                r["data_documento"],
                r["motivo_contestazione"][:60],
                r["_esito"],
            ])
        self._write_sheet(
            "⚠️ Anomalie", headers, rows,
            [26, 16, 22, 12, 12, 40, 16],
        )

    def add_contestazioni_detail(self, partitario: pd.DataFrame):
        mask = partitario["contestata"].str.lower().isin(["sì", "si", "yes", "true", "1"])
        if not mask.any():
            return
        df_c = partitario[mask]
        headers = ["N° Documento", "Tipo", "Controparte", "Importo (€)", "Data Doc.", "Stato", "Contestata", "Motivo", "Note", "Esito"]
        rows = []
        for _, r in df_c.iterrows():
            rows.append([
                r["numero_documento"], r["tipo"], r["controparte"],
                r["importo"], r["data_documento"], r["stato"],
                r["contestata"], r["motivo_contestazione"], r["note"], r["_esito"],
            ])
        self._write_sheet(
            "🚩 Contestazioni", headers, rows,
            [14, 10, 22, 12, 12, 14, 10, 32, 24, 16],
        )

    def save(self, path: Path):
        self.wb.save(path)
        log.info(f"Report salvato: {path}")


# ===========================================================================
# 6.  ORCHESTRATOR – collega tutto
# ===========================================================================

class AccountingProcessor:
    """Entry point principale del processo."""

    def __init__(self, config: dict = None):
        self.cfg = config or {}
        self.bank_reader    = BankStatementReader()
        self.ledger_reader  = LedgerReader()
        self.engine         = ReconciliationEngine(
            tolleranza_importo = self.cfg.get("tolleranza_importo", TOLERANZA_IMPORTO),
            tolleranza_giorni  = self.cfg.get("tolleranza_giorni",  TOLLERANZA_GIORNI),
            usa_desc           = self.cfg.get("usa_desc", True),
            usa_data           = self.cfg.get("usa_data", False),
        )
        self.cont_manager   = ContestationManager()
        self.reporter       = ReportGenerator()

    def run(
        self,
        estratto_path:         Path,
        partitario_forn_path:  Optional[Path] = None,
        partitario_cli_path:   Optional[Path] = None,
        fatture_emesse_path:   Optional[Path] = None,
        fatture_ricevute_path: Optional[Path] = None,
        contestazioni_path:    Optional[Path] = None,
        output_name:           str = None,
    ) -> Path:
        log.info("=" * 60)
        log.info("AVVIO PROCESSO DI RICONCILIAZIONE")
        log.info("=" * 60)

        # 1. leggi estratto conto
        estratto = self.bank_reader.read(estratto_path)

        # 2. costruisci partitario unificato
        frames = []
        if partitario_forn_path:
            df = self.ledger_reader.read(partitario_forn_path, tipo_default="Ricevuta")
            frames.append(df)
        if partitario_cli_path:
            df = self.ledger_reader.read(partitario_cli_path, tipo_default="Emessa")
            frames.append(df)
        if fatture_emesse_path:
            df = self.ledger_reader.read(fatture_emesse_path, tipo_default="Emessa")
            frames.append(df)
        if fatture_ricevute_path:
            df = self.ledger_reader.read(fatture_ricevute_path, tipo_default="Ricevuta")
            frames.append(df)

        if not frames:
            log.warning("Nessun partitario/registro fatture fornito. Report solo estratto conto.")
            partitario = pd.DataFrame(columns=LedgerReader.COLONNE_STD + ["_matched", "_match_banca", "_esito"])
        else:
            partitario = pd.concat(frames, ignore_index=True)

        # 3. applica contestazioni
        if contestazioni_path:
            cont_df = self.cont_manager.load_contestazioni(contestazioni_path)
            partitario = self.cont_manager.apply(partitario, cont_df)

        # 4. riconcilia
        estratto, partitario = self.engine.reconcile(estratto, partitario)

        # 5. genera report
        self.reporter.add_riepilogo(estratto, partitario)
        self.reporter.add_estratto(estratto)
        if not partitario.empty:
            mask_forn = partitario["tipo"] == "Ricevuta"
            mask_cli  = partitario["tipo"] == "Emessa"
            if mask_forn.any():
                self.reporter.add_partitario(partitario[mask_forn], "📋 Partitario Fornitori")
            if mask_cli.any():
                self.reporter.add_partitario(partitario[mask_cli], "📋 Partitario Clienti")
            self.reporter.add_contestazioni_detail(partitario)
        self.reporter.add_anomalie(estratto, partitario)

        out_name = output_name or f"riconciliazione_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        out_path = OUTPUT_DIR / out_name
        self.reporter.save(out_path)

        log.info("=" * 60)
        log.info("PROCESSO COMPLETATO")
        log.info("=" * 60)
        return out_path


# ===========================================================================
# 7.  CLI
# ===========================================================================

def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Riconciliazione contabile: estratto conto ↔ partitari NAV",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi:
  python reconcilia.py --estratto input/estratti/giugno.pdf
  python reconcilia.py --estratto input/estratti/giugno.xlsx \\
                       --fornitori input/partitari/fornitori.xlsx \\
                       --clienti input/partitari/clienti.xlsx \\
                       --emesse input/fatture/emesse.xlsx \\
                       --ricevute input/fatture/ricevute.xlsx \\
                       --contestazioni input/contestazioni.xlsx \\
                       --output report_giugno.xlsx
        """,
    )
    parser.add_argument("--estratto",       required=True,  help="Estratto conto bancario (PDF/Excel/CSV)")
    parser.add_argument("--fornitori",      default=None,   help="Partitario fornitori (Excel/CSV)")
    parser.add_argument("--clienti",        default=None,   help="Partitario clienti (Excel/CSV)")
    parser.add_argument("--emesse",         default=None,   help="Registro fatture emesse (Excel/CSV)")
    parser.add_argument("--ricevute",       default=None,   help="Registro fatture ricevute (Excel/CSV)")
    parser.add_argument("--contestazioni",  default=None,   help="File contestazioni (Excel/CSV)")
    parser.add_argument("--output",         default=None,   help="Nome file output (default: timestamped)")
    parser.add_argument("--tol-importo",    type=float, default=TOLERANZA_IMPORTO, help="Tolleranza su importo (€)")
    parser.add_argument("--tol-giorni",     type=int,   default=TOLLERANZA_GIORNI, help="Tolleranza date (giorni)")
    parser.add_argument("--no-desc",        action="store_true", help="Non usare descrizione per l'abbinamento")
    parser.add_argument("--usa-data",       action="store_true", help="Usa data scadenza nell'abbinamento")

    args = parser.parse_args()

    cfg = {
        "tolleranza_importo": args.tol_importo,
        "tolleranza_giorni":  args.tol_giorni,
        "usa_desc": not args.no_desc,
        "usa_data": args.usa_data,
    }

    processor = AccountingProcessor(config=cfg)
    out = processor.run(
        estratto_path         = Path(args.estratto),
        partitario_forn_path  = Path(args.fornitori)      if args.fornitori     else None,
        partitario_cli_path   = Path(args.clienti)        if args.clienti       else None,
        fatture_emesse_path   = Path(args.emesse)         if args.emesse        else None,
        fatture_ricevute_path = Path(args.ricevute)       if args.ricevute      else None,
        contestazioni_path    = Path(args.contestazioni)  if args.contestazioni else None,
        output_name           = args.output,
    )
    print(f"\n✅  Report generato: {out}\n")


if __name__ == "__main__":
    main()
