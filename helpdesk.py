# -*- coding: utf-8 -*-
"""
Help Desk INAIL - Registro interventi
Registra gli interventi di help desk (utente, richiesta, soluzione, note, ticket),
con filtri, conteggio, elenco utenti che si costruisce da solo ed esportazione
in Excel/PDF.
"""
import base64
import io
import json
import math
import os
import queue
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import unicodedata
import uuid
import webbrowser
import zipfile
from collections import Counter
from datetime import date, datetime, timedelta, timezone

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from printer_images import IMAGES as PRINTER_IMAGES
except ImportError:
    PRINTER_IMAGES = {}

try:
    from desk_images import IMAGES as DESK_IMAGES        # foto dei dispositivi sulle scrivanie
except ImportError:
    DESK_IMAGES = {}

try:
    import verbali as VB          # verbali Word (trasferimento, consegna, assegnazione, censimento)
except ImportError:
    VB = None

try:
    import codici as CD           # lettura di codici a barre e QR (fotocamera, foto, appunti)
except ImportError:
    CD = None

try:
    import servicenow as SN       # importazione dei ticket di ServiceNow come interventi
except ImportError:
    SN = None

try:
    import corridoio3d as C3D     # visite 3D: passeggiata nel piano e edificio intero
except ImportError:
    C3D = None

APP_NAME = "Help Desk INAIL"

# Colori presi dal logo INAIL
NAVY = "#002D5C"
NAVY_DARK = "#001f40"
GREEN = "#2c954b"
GREEN_DARK = "#237a3c"
RED = "#d12233"
RED_DARK = "#a91b29"
LIGHT = "#f2f9fd"
WHITE = "#ffffff"
GREY = "#5b6b7d"
BORDER = "#c5d3e0"
TEXT = "#1c2733"
BTN_BG = "#dde8f3"
BTN_BG_ACTIVE = "#c9daec"

# le persone si ordinano per COGNOME: la colonna «ordine» (quando c'è) è «Cognome Nome» come nella
# lista INAIL; altrimenti si ordina per il nome come è scritto
SORT_UTENTE = "COALESCE(NULLIF(u.ordine, ''), u.nome) COLLATE NOCASE"

SEDI_TOSCANA = ["Arezzo", "Carrara", "DR Toscana", "Firenze", "Grosseto", "Livorno", "Lucca",
                "Massa", "Pisa", "Pistoia", "Prato", "Siena"]

# (nome, colore a schermo, colore leggibile su carta bianca)
COLORI_TONER = [("Nero", "#1c1c1c", "#1c1c1c"), ("Azzurro", "#1fa3dc", "#0d7fb0"),
                ("Magenta", "#d6247f", "#c0176c"), ("Giallo", "#f2c500", "#a67c00"),
                ("Vaschetta", "#8a96a3", "#4f5d6b")]
TONER_NOMI = [c[0] for c in COLORI_TONER]
ROOM_CODE_MAX = 24                # lunghezza massima del codice di una stanza
PIANTA_PX = 1600                  # lato lungo della pianta mandata al telefono
PANORAMA_PX = 2600                # lato lungo dei panorami mandati al telefono
BADGE_GIALLO = "#f2c500"          # bollino giallo (scorta in esaurimento)
BADGE_GIALLO_TESTO = "#8a5200"    # testo leggibile per la scorta bassa
TONER_TESTO = {c[0]: c[2] for c in COLORI_TONER}

# i due modelli di Via delle Porte Nuove (nome come scritto sul fronte della stampante)
MODELLI_STAMPANTI = ["bizhub C250i", "bizhub C3350i"]

# nomi usati in passato per sbaglio -> nome giusto (il programma propone di correggerli una volta sola)
MODELLI_DA_CORREGGERE = [("bizhub C3351i", "bizhub C3350i"), ("bizhub C360i", "bizhub C250i")]

# stato di un intervento dell'utente e di una richiesta al tecnico Konica
STATI_INTERVENTO = ["Aperto", "In attesa", "Risolto"]
STATI_APERTI = ("Aperto", "In attesa")
STATO_COLORI = {"Aperto": RED, "In attesa": "#b36b00"}
FILTRO_STATO = ["Tutti", "Da richiamare (aperti e in attesa)", "Aperto", "In attesa", "Risolto"]
STATI_TECNICO = ["Richiesto", "Eseguito"]
SOGLIA_SCORTA = 1   # avviso quando di un toner ne resta questo numero (o meno)

IMG_EXTS = (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tif", ".tiff", ".heic")
MAX_IMG_PX = 2200   # le foto JPEG più grandi vengono ridotte a questa misura
THUMB_PX = 72

# (chiave, titolo, acceso all'inizio, larghezza PDF in mm, larghezza Excel)
USER_COLUMNS = [
    ("nome", "Utente", True, 55, 26), ("matricola", "Matricola", True, 28, 14),
    ("ufficio", "Ufficio", False, 45, 26), ("sede", "Sede", True, 30, 16),
    ("telefono", "Telefono", True, 35, 18), ("email", "Email", False, 55, 30),
    ("seriale_pc", "Seriale PC", True, 45, 22),
    ("imei", "IMEI cellulare", True, 45, 20), ("dotazione", "Dotazione", False, 70, 40),
    ("n", "N. interventi", False, 22, 14),
    ("ultimo", "Ultimo intervento", False, 32, 18),
]


def floor_label(n):
    if n == 0:
        return "Piano terra"
    if n == -1:
        return "Seminterrato"
    if n < -1:
        return "Interrato %d" % (-n)
    return "%d° piano" % n


def scorta_status(quantita, soglia):
    if quantita <= 0:
        return "Finito"
    if quantita <= soglia:
        return "In esaurimento"
    return "OK"


def model_key(modello):
    """Il modello scritto in modo un po' diverso (spazi, maiuscole) conta come lo stesso."""
    return " ".join((modello or "").split()).casefold()


def natural_key(text):
    """Ordine "naturale": T2 viene prima di T12."""
    return [int(t) if t.isdigit() else t.casefold() for t in re.split(r"(\d+)", text or "")]


def clean_room_code(text):
    return " ".join((text or "").split())[:ROOM_CODE_MAX]


def clean_serial(text):
    """Numero di serie: senza spazi in testa/coda o doppi, al massimo 40 caratteri."""
    return " ".join(str(text or "").split())[:40]


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def clean_email(text):
    """Email in minuscolo e senza spazi, al massimo 120 caratteri."""
    return "".join(str(text or "").split()).lower()[:120]


def mail_to(address):
    """Apre il programma di posta con un messaggio nuovo per questo indirizzo."""
    address = clean_email(address)
    if not EMAIL_RE.match(address):
        return False
    try:
        webbrowser.open("mailto:" + address)
    except Exception:
        return False
    return True


# apparecchi che una persona può avere in dotazione (l'elenco è libero: si può scrivere altro)
TIPI_DOTAZIONE_PERSONA = ["Notebook", "PC fisso", "Monitor", "Cellulare", "Smartphone",
                          "Telefono Teams", "Tablet", "Scheda SIM", "Cuffie", "Telefono fisso",
                          "Token / smart card", "Stampante personale", "Altro"]


def clean_dotazione(value):
    """Elenco apparecchi di una persona: [{"tipo", "modello", "seriale"}]. Accetta anche il testo
    JSON salvato nel database; scarta le righe vuote e quelle non valide."""
    if isinstance(value, (str, bytes)):
        try:
            value = json.loads(value or "[]")
        except ValueError:
            value = []
    out = []
    for r in value if isinstance(value, (list, tuple)) else []:
        if not isinstance(r, dict):
            continue
        row = {"tipo": " ".join(str(r.get("tipo") or "").split())[:40],
               "modello": " ".join(str(r.get("modello") or "").split())[:80],
               "seriale": clean_serial(r.get("seriale"))}
        if any(row.values()):
            out.append(row)
    return out


def merge_dotazione(old, new):
    """Aggiunge alla dotazione `old` le righe di `new` senza toccare quello che c'è già: una riga
    dello stesso tipo (stesso apparecchio) resta com'è e riceve solo il seriale o il modello che le
    mancano. Non cancella e non sovrascrive niente."""
    rows = clean_dotazione(old)

    def ident(r):
        return _row_kind(r) or r["tipo"].casefold()
    for r in clean_dotazione(new):
        same = next((x for x in rows if ident(x) == ident(r)), None)
        if same is None:
            rows.append(dict(r))
            continue
        for k in ("seriale", "modello"):
            if not same[k] and r[k]:
                same[k] = r[k]
    return rows


def dotazione_text(items, sep="; "):
    """Riassunto in una riga: «Smartphone Samsung A54 (IMEI 3567…); Monitor Dell»."""
    parts = []
    for r in clean_dotazione(items):
        txt = " ".join(x for x in (r["tipo"], r["modello"]) if x)
        if r["seriale"]:
            txt = "%s (%s)" % (txt, r["seriale"]) if txt else r["seriale"]
        parts.append(txt)
    return sep.join(parts)


def parse_room_codes(text):
    """Un codice per riga (vanno bene anche virgole e punti e virgola). Restituisce
    (codici validi senza doppioni, codici scartati perché troppo lunghi)."""
    seen, good, bad = set(), [], []
    for raw in re.split(r"[\r\n;,\t]+", text or ""):
        code = " ".join(raw.split())
        if not code:
            continue
        if len(code) > ROOM_CODE_MAX:
            bad.append(code)
        elif code.casefold() not in seen:
            seen.add(code.casefold())
            good.append(code)
    return good, bad


_DESK_WORDS = r"(?:postazion[ei]|posti|scrivani[ae]|desks?)"
_ROOM_PAIR = (
    # T12; 3   T12<tab>3   T12: 3   T12 = 3   T12 - 3   (anche "3 postazioni")
    re.compile(r"^(?P<c>[^;\t]+?)\s*(?:[;\t:=]|\s[-\u2013\u2014]\s)\s*(?P<n>\d{1,3})\s*(?:" + _DESK_WORDS
               + r")?\s*$", re.I),
    # T12 (3)   T12 (3 postazioni)
    re.compile(r"^(?P<c>[^;\t()]+?)\s*\(\s*(?P<n>\d{1,3})\s*(?:" + _DESK_WORDS + r")?\s*\)\s*$", re.I),
)


def parse_room_lines(text):
    """Una stanza per riga, con le postazioni se scritte ("T12; 3", "T12 - 3", "T12 (3)" o due
    colonne copiate da Excel). Senza numero: solo il codice (più codici sulla stessa riga si
    separano con virgola o punto e virgola). Restituisce ([(codice, postazioni o None)], scartati)."""
    order, desks, bad = [], {}, []

    def add(code, n):
        code = " ".join(code.split())
        if not code:
            return
        if len(code) > ROOM_CODE_MAX:
            bad.append(code)
            return
        k = code.casefold()
        if k not in desks:
            order.append(code)
            desks[k] = n
        elif n is not None:
            desks[k] = n

    for line in re.split(r"[\r\n]+", text or ""):
        line = line.strip()
        if not line:
            continue
        m = None
        for rx in _ROOM_PAIR:
            m = rx.match(line)
            if m:
                break
        if m:
            add(m.group("c"), int(m.group("n")))
        else:
            for part in re.split(r"[;,\t]+", line):
                add(part, None)
    return [(c, desks[c.casefold()]) for c in order], bad


def room_label(r):
    """Scritta sulla pianta e negli elenchi: "T12 · 3" (codice e numero di persone)."""
    return "%s · %d" % (r["codice"], r.get("n_persone") or 0)


def room_occupancy(r):
    """('occupate 2/3', False) se le scrivanie totali sono indicate, altrimenti ('', False).
    Il secondo valore è True quando le persone sono più delle scrivanie."""
    tot = r.get("scrivanie") or 0
    n = r.get("n_persone") or 0
    if tot <= 0:
        return "", False
    return "occupate %d/%d" % (n, tot), n > tot


def days_ago(data_ora, today=None):
    """Giorni di calendario passati da 'AAAA-MM-GG hh:mm' (0 = oggi); None se manca la data."""
    try:
        d = datetime.strptime((data_ora or "")[:10], "%Y-%m-%d").date()
    except ValueError:
        return None
    return max(0, ((today or datetime.now().date()) - d).days)


def ago_short(n):
    """Per lo spazio stretto sotto la stampante: mai / oggi / ieri / 12 gg fa."""
    if n is None:
        return "mai"
    return "oggi" if n == 0 else "ieri" if n == 1 else "%d gg fa" % n


def ago_long(n):
    """Frase completa: cambiato 12 giorni fa."""
    if n is None:
        return "mai registrato"
    return "cambiato oggi" if n == 0 else "cambiato ieri" if n == 1 else "cambiato %d giorni fa" % n


TONER_GAP_MIN = 3    # due cambi a meno di 3 giorni uno dall'altro non contano come "durata"
TONER_BASSO = 20     # sotto questa percentuale la barretta ha il contorno rosso


def _day_no(data_ora):
    try:
        return datetime.strptime((data_ora or "")[:10], "%Y-%m-%d").date().toordinal()
    except ValueError:
        return None


def toner_gaps(dates):
    """Giorni tra un cambio e il successivo (stesso colore, stessa stampante), senza doppioni
    e senza intervalli troppo corti."""
    d = sorted({n for n in (_day_no(x) for x in dates) if n is not None})
    return [b - a for a, b in zip(d, d[1:]) if b - a >= TONER_GAP_MIN]


def median_int(values):
    s = sorted(values)
    n, h = len(s), len(s) // 2
    return s[h] if n % 2 else (s[h - 1] + s[h] + 1) // 2


def toner_estimate(own_dates, pool_gaps, today=None):
    """STIMA del toner che resta, in percentuale: dai giorni dall'ultimo cambio e da quanto sono
    durati i cambi precedenti (della stampante, o di quelle dello stesso modello). Non e' il
    livello vero letto dalla stampante."""
    own = [d for d in own_dates if d]
    last = max(own) if own else None
    days = days_ago(last, today) if last else None
    g, src = toner_gaps(own), "stampante"
    if not g:
        g, src = list(pool_gaps), "modello"
    if days is None or not g:
        return {"pct": None, "days": days}
    life = median_int(g)
    pct = max(0, min(100, int(math.floor(100 * (1 - days / life) + 0.5))))
    return {"pct": pct, "life": life, "days": days, "src": src, "n": len(g)}


def toner_estimate_text(e):
    if e.get("pct") is None:
        return ("stima non disponibile (servono almeno due cambi di questo colore, di questa "
                "stampante o di una dello stesso modello)")
    return "stima: circa %d%% - di solito dura %d giorni (%s)" % (
        e["pct"], e["life"], "dai cambi di questa stampante" if e["src"] == "stampante"
        else "dai cambi di stampanti dello stesso modello")


def compute_toner_estimates(printers, toner_rows, today=None):
    """printers = [(id, modello)], toner_rows = [(id stampante, colore, data_ora)].
    Restituisce {id stampante: {colore: stima}}."""
    dates, model = {}, {}
    for sid, colore, d in toner_rows:
        dates.setdefault((sid, colore), []).append(d)
    for sid, modello in printers:
        model[sid] = model_key(modello)
    out = {}
    for sid, _m in printers:
        per = {}
        for colore in TONER_NOMI:
            pool = []
            if model[sid]:
                for oid, om in printers:
                    if oid != sid and model[oid] == model[sid]:
                        pool += toner_gaps(dates.get((oid, colore), []))
            per[colore] = toner_estimate(dates.get((sid, colore), []), pool, today)
        out[sid] = per
    return out


def tec_date(m):
    """Data in cui è venuto il tecnico, già scritta gg/mm/aaaa (vuota se non è ancora venuto)."""
    if not m.get("data_tecnico"):
        return ""
    try:
        return fmt_date(datetime.strptime(m["data_tecnico"], "%Y-%m-%d"))
    except ValueError:
        return m["data_tecnico"]


def printer_image_key(model):
    """Sceglie quale foto usare in base al modello scritto."""
    m = (model or "").casefold()
    if any(x in m for x in ("3351", "3350", "3320")):
        return "c3350i"        # la piccola (A4)
    return "c250i"             # la grande (A3)

# --------------------------------------------------------------------------
# Utilità
# --------------------------------------------------------------------------
def resource_path(name):
    """Percorso di un file incluso nell'app (anche dentro il .exe)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def data_dir():
    """Cartella dove restano database e impostazioni."""
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    path = os.path.join(base, "HelpDeskINAIL")
    os.makedirs(path, exist_ok=True)
    return path


def parse_date(text):
    text = (text or "").strip().replace("-", "/").replace(".", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def parse_datetime(text):
    text = (text or "").strip().replace("-", "/").replace(".", "/")
    for fmt in ("%d/%m/%Y %H:%M", "%d/%m/%y %H:%M", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def fmt_dt(iso):
    try:
        return datetime.strptime(iso, "%Y-%m-%d %H:%M").strftime("%d/%m/%Y %H:%M")
    except ValueError:
        return iso


def fmt_date(d):
    return d.strftime("%d/%m/%Y")


def normalize_name(name):
    name = re.sub(r"\s+", " ", (name or "").strip())
    if name and (name.islower() or name.isupper()):
        name = name.title()
    return name


def one_line(text, limit=110):
    text = " ".join((text or "").split())
    return text if len(text) <= limit else text[: limit - 1] + "…"


def matches_all(terms, haystack):
    haystack = haystack.casefold()
    return all(t in haystack for t in terms)


def now_utc_iso():
    n = datetime.now(timezone.utc)
    return n.strftime("%Y-%m-%dT%H:%M:%S.") + "%03dZ" % (n.microsecond // 1000)


def attachments_root():
    path = os.path.join(data_dir(), "allegati")
    os.makedirs(path, exist_ok=True)
    return path


def attachment_dir(iid):
    return os.path.join(attachments_root(), str(iid))


def bolle_foto_root():
    """Foto delle bolle dei toner (incluse nel backup insieme agli allegati)."""
    path = os.path.join(attachments_root(), "bolle")
    os.makedirs(path, exist_ok=True)
    return path


def staging_dir():
    """Cartella temporanea per le immagini incollate prima di salvare l'intervento."""
    return os.path.join(attachments_root(), "_temp")


# ---------------------------------------------------------------------------
# Verbali Word: impostazioni e cartelle
# ---------------------------------------------------------------------------
SETTINGS_DEFAULT = {"tecnico": "", "sede_citta": "Firenze", "indirizzo": "Via Delle Porte Nuove, 61",
                    "cap": "50144", "cartella_verbali": "", "mappa_pos": "",
                    "modelli_rinomina_chiesto": ""}


def settings_path():
    return os.path.join(data_dir(), "impostazioni.json")


def load_settings():
    s = dict(SETTINGS_DEFAULT)
    try:
        with open(settings_path(), "r", encoding="utf-8") as f:
            saved = json.load(f)
        if isinstance(saved, dict):
            s.update({k: str(v) for k, v in saved.items() if k in SETTINGS_DEFAULT})
    except (OSError, ValueError):
        pass
    return s


def save_settings(s):
    try:
        with open(settings_path(), "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


# ---------------------------------------------------------------------------
# Mappa della Toscana: contorno semplificato (disegnato a mano da punti di riferimento, NON è un
# confine ufficiale) e coordinate dei comuni. Le sedi si riconoscono dal nome; quelle che il
# programma non conosce si posizionano a mano (la posizione resta in impostazioni.json).
# ---------------------------------------------------------------------------
# --- TOSCANA-DATI-INIZIO (generato dallo stesso elenco di docs/toscana.js: non modificare a mano) ---
TOSC_K, TOSC_S, TOSC_LON0, TOSC_LAT1, TOSC_W, TOSC_H = 0.72537, 492.0, 9.65, 44.55, 1000, 1181
TOSC_OUTLINE = [(10.03, 44.055), (10.045, 44.03), (10.1, 43.99), (10.15, 43.97),
    (10.18, 43.95), (10.2, 43.92), (10.23, 43.87), (10.26, 43.82), (10.28, 43.75),
    (10.27, 43.68), (10.29, 43.62), (10.31, 43.57), (10.3, 43.54), (10.33, 43.49),
    (10.37, 43.43), (10.4, 43.4), (10.44, 43.35), (10.47, 43.31), (10.5, 43.25),
    (10.53, 43.18), (10.54, 43.12), (10.55, 43.04), (10.5, 42.99), (10.49, 42.96),
    (10.52, 42.93), (10.56, 42.92), (10.62, 42.97), (10.7, 42.95), (10.76, 42.92),
    (10.78, 42.89), (10.77, 42.84), (10.74, 42.8), (10.8, 42.78), (10.87, 42.77),
    (10.93, 42.72), (10.97, 42.69), (11.03, 42.64), (11.07, 42.6), (11.1, 42.57),
    (11.13, 42.55), (11.17, 42.53), (11.19, 42.5), (11.13, 42.48), (11.1, 42.45),
    (11.08, 42.42), (11.1, 42.39), (11.17, 42.375), (11.22, 42.385), (11.25, 42.4),
    (11.29, 42.415), (11.38, 42.385), (11.47, 42.355), (11.52, 42.39), (11.57, 42.43),
    (11.62, 42.5), (11.69, 42.56), (11.74, 42.62), (11.78, 42.7), (11.82, 42.77),
    (11.88, 42.78), (11.95, 42.82), (11.97, 42.88), (12.03, 42.93), (12.03, 42.97),
    (12.06, 43.02), (12.1, 43.07), (12.05, 43.12), (12.03, 43.2), (12.05, 43.26),
    (12.1, 43.31), (12.12, 43.36), (12.13, 43.42), (12.12, 43.47), (12.15, 43.53),
    (12.17, 43.58), (12.2, 43.62), (12.25, 43.65), (12.33, 43.65), (12.38, 43.7),
    (12.37, 43.76), (12.3, 43.8), (12.22, 43.83), (12.15, 43.84), (12.05, 43.8),
    (11.98, 43.82), (11.9, 43.83), (11.83, 43.86), (11.78, 43.87), (11.7, 43.9),
    (11.68, 43.94), (11.62, 43.98), (11.66, 44.04), (11.7, 44.09), (11.66, 44.14),
    (11.57, 44.16), (11.5, 44.17), (11.46, 44.2), (11.38, 44.22), (11.3, 44.2), (11.24, 44.17),
    (11.17, 44.14), (11.1, 44.12), (11.0, 44.08), (10.9, 44.15), (10.82, 44.17),
    (10.75, 44.15), (10.68, 44.17), (10.62, 44.15), (10.55, 44.16), (10.5, 44.2),
    (10.45, 44.25), (10.35, 44.26), (10.28, 44.3), (10.22, 44.31), (10.15, 44.33),
    (10.12, 44.36), (10.05, 44.4), (10.0, 44.43), (9.93, 44.46), (9.86, 44.43), (9.8, 44.4),
    (9.78, 44.35), (9.82, 44.3), (9.88, 44.25), (9.93, 44.2), (9.97, 44.16), (9.99, 44.13),
    (10.0, 44.09)]
TOSC_ISLANDS = {
    'Elba': [(10.11, 42.8), (10.17, 42.82), (10.26, 42.83), (10.34, 42.83), (10.4, 42.86),     (10.45, 42.84), (10.45, 42.81), (10.42, 42.78), (10.41, 42.73),     (10.34, 42.73), (10.27, 42.75), (10.19, 42.74), (10.12, 42.76)],
    'Capraia': [(9.865, 43.05), (9.862, 43.076), (9.852, 43.097), (9.838, 43.108),     (9.822, 43.108), (9.808, 43.097), (9.798, 43.076), (9.795, 43.05),     (9.798, 43.024), (9.808, 43.003), (9.822, 42.992), (9.838, 42.992),     (9.852, 43.003), (9.862, 43.024)],
    'Gorgona': [(9.914, 43.43), (9.91, 43.441), (9.9, 43.446), (9.89, 43.441), (9.886, 43.43),     (9.89, 43.419), (9.9, 43.414), (9.91, 43.419)],
    'Pianosa': [(10.1, 42.585), (10.094, 42.595), (10.08, 42.599), (10.066, 42.595),     (10.06, 42.585), (10.066, 42.575), (10.08, 42.571), (10.094, 42.575)],
    'Montecristo': [(10.33, 42.335), (10.324, 42.349), (10.31, 42.355), (10.296, 42.349),     (10.29, 42.335), (10.296, 42.321), (10.31, 42.315), (10.324, 42.321)],
    'Giglio': [(10.95, 42.36), (10.947, 42.382), (10.937, 42.399), (10.923, 42.409),     (10.907, 42.409), (10.893, 42.399), (10.883, 42.382), (10.88, 42.36),     (10.883, 42.338), (10.893, 42.321), (10.907, 42.311), (10.923, 42.311),     (10.937, 42.321), (10.947, 42.338)],
    'Giannutri': [(11.125, 42.255), (11.118, 42.262), (11.1, 42.265), (11.082, 42.262),     (11.075, 42.255), (11.082, 42.248), (11.1, 42.245), (11.118, 42.248)]}
TOSC_COMUNI = {'Firenze': (43.7696, 11.2558, 1), 'Pisa': (43.7228, 10.4017, 1),
    'Siena': (43.3188, 11.3308, 1), 'Arezzo': (43.4633, 11.8797, 1),
    'Livorno': (43.5485, 10.3106, 1), 'Lucca': (43.843, 10.5027, 1),
    'Grosseto': (42.7635, 11.1124, 1), 'Massa': (44.0354, 10.1393, 1),
    'Carrara': (44.0793, 10.0977, 0), 'Pistoia': (43.9331, 10.9175, 1),
    'Prato': (43.8777, 11.1022, 1), 'Empoli': (43.7192, 10.9456, 0),
    'Viareggio': (43.8681, 10.2511, 0), 'Piombino': (42.9254, 10.5266, 0),
    'Cecina': (43.3085, 10.5169, 0), 'Pontedera': (43.6619, 10.6313, 0),
    'Montecatini Terme': (43.8836, 10.7727, 0), 'Sansepolcro': (43.571, 12.14, 0),
    'Cortona': (43.2756, 11.9855, 0), 'Montepulciano': (43.0988, 11.7826, 0),
    'Poggibonsi': (43.4693, 11.1472, 0), "Colle di Val d'Elsa": (43.4228, 11.1197, 0),
    'San Gimignano': (43.4677, 11.0432, 0), 'Volterra': (43.4002, 10.8605, 0),
    'Follonica': (42.9233, 10.7606, 0), 'Orbetello': (42.4423, 11.214, 0),
    'Porto Santo Stefano': (42.4372, 11.118, 0), 'Portoferraio': (42.8143, 10.3263, 0),
    'Massa Marittima': (43.0505, 10.8933, 0),
    'Castiglione della Pescaia': (42.7647, 10.8808, 0), 'Aulla': (44.2108, 9.9689, 0),
    'Pontremoli': (44.3764, 9.8806, 0), 'Fivizzano': (44.2373, 10.1288, 0),
    'Castelnuovo di Garfagnana': (44.11, 10.4097, 0), 'Barga': (44.0772, 10.4778, 0),
    'Borgo a Mozzano': (43.9993, 10.5421, 0), 'Pescia': (43.9022, 10.689, 0),
    'Pietrasanta': (43.9588, 10.2282, 0), 'Forte dei Marmi': (43.9626, 10.1711, 0),
    'Camaiore': (43.9378, 10.3052, 0), 'San Miniato': (43.6823, 10.8505, 0),
    'Cascina': (43.6767, 10.5547, 0), 'Pontassieve': (43.7667, 11.4333, 0),
    'Sesto Fiorentino': (43.8326, 11.1999, 0), 'Scandicci': (43.7547, 11.1856, 0),
    'Bagno a Ripoli': (43.7428, 11.3036, 0), 'Figline Valdarno': (43.6175, 11.4672, 0),
    'Borgo San Lorenzo': (43.9548, 11.3872, 0), 'Barberino di Mugello': (44.0, 11.2167, 0),
    'Firenzuola': (44.1228, 11.3789, 0), 'Montevarchi': (43.5236, 11.5681, 0),
    'San Giovanni Valdarno': (43.5708, 11.5276, 0), 'Bibbiena': (43.697, 11.816, 0),
    'Poppi': (43.7194, 11.7672, 0), 'Sinalunga': (43.2088, 11.7347, 0),
    'Chiusi': (43.0172, 11.9494, 0), 'Abbadia San Salvatore': (42.8793, 11.6786, 0),
    'Radicofani': (42.8999, 11.7679, 0), 'Pitigliano': (42.6336, 11.6683, 0),
    'Manciano': (42.587, 11.5167, 0), 'Capalbio': (42.4167, 11.3667, 0),
    'Scansano': (42.6819, 11.3253, 0), 'Roccastrada': (42.9988, 11.1608, 0),
    'Rosignano Marittimo': (43.4044, 10.4752, 0), 'Castagneto Carducci': (43.1633, 10.6053, 0),
    'San Vincenzo': (43.1, 10.539, 0), 'Campiglia Marittima': (43.0636, 10.6097, 0),
    'Suvereto': (43.0742, 10.6753, 0), 'Capraia Isola': (43.05, 9.8333, 0),
    'Marciana Marina': (42.8, 10.1964, 0), 'Porto Azzurro': (42.764, 10.3965, 0),
    'Rio Marina': (42.8167, 10.4283, 0), 'Capoliveri': (42.7466, 10.3797, 0),
    'Isola del Giglio': (42.3647, 10.9156, 0), 'Ponsacco': (43.6264, 10.6203, 0),
    'Altopascio': (43.8142, 10.6717, 0), 'Capannori': (43.8419, 10.5698, 0),
    'Fucecchio': (43.7289, 10.8074, 0), 'Castelfiorentino': (43.6073, 10.9646, 0),
    'Certaldo': (43.547, 11.04, 0), 'Montelupo Fiorentino': (43.7333, 11.0167, 0),
    'Greve in Chianti': (43.5855, 11.3162, 0), 'Impruneta': (43.6972, 11.2528, 0),
    'Fiesole': (43.8056, 11.2938, 0), 'Reggello': (43.6864, 11.5336, 0),
    'Rufina': (43.8131, 11.48, 0), 'Vicchio': (43.9179, 11.4504, 0),
    'Marradi': (44.0731, 11.6228, 0), 'Abetone': (44.1433, 10.6553, 0),
    'San Marcello Piteglio': (44.0554, 10.7959, 0), 'Quarrata': (43.85, 10.98, 0),
    'Monsummano Terme': (43.87, 10.82, 0), 'Vernio': (44.0633, 11.1511, 0),
    'Montemurlo': (43.9199, 11.0153, 0), 'Campi Bisenzio': (43.8167, 11.1333, 0),
    'Piancastagnaio': (42.845, 11.702, 0), 'Castel del Piano': (42.8944, 11.5333, 0),
    'Arcidosso': (42.85, 11.5333, 0), 'Sarteano': (43.0, 11.8667, 0),
    'Cetona': (43.0, 11.9, 0), 'Montalcino': (43.0575, 11.49, 0),
    'Pienza': (43.0766, 11.6786, 0), 'Chianciano Terme': (43.0617, 11.8219, 0),
    'Monte San Savino': (43.3333, 11.7333, 0), 'Castiglion Fiorentino': (43.3333, 11.9167, 0),
    'Anghiari': (43.54, 12.06, 0), 'Pieve Santo Stefano': (43.6667, 12.0333, 0),
    'Stia': (43.8, 11.7, 0), 'Sorano': (42.6833, 11.7167, 0), 'Gavorrano': (42.9167, 10.9, 0),
    'Scarlino': (42.9167, 10.8667, 0), 'Pomarance': (43.2967, 10.8718, 0),
    'Peccioli': (43.55, 10.7167, 0), "Santa Croce sull'Arno": (43.7167, 10.7667, 0),
    'San Giuliano Terme': (43.75, 10.4333, 0), 'Massarosa': (43.8667, 10.3333, 0),
    'Seravezza': (43.9961, 10.2292, 0), 'Bagni di Lucca': (44.0167, 10.5833, 0),
    'Camporgiano': (44.15, 10.3833, 0), 'Villafranca in Lunigiana': (44.3, 9.95, 0),
    'Licciana Nardi': (44.2, 10.0, 0), 'Albinia': (42.5, 11.19, 0),
    'Talamone': (42.55, 11.13, 0), 'Cavo': (42.85, 10.4167, 0)}
TOSC_ALIAS = {'dr toscana': 'Firenze', 'direzione regionale': 'Firenze',
    'direzione regionale toscana': 'Firenze', 'massa carrara': 'Massa',
    'massa-carrara': 'Massa', 'mc': 'Massa', 'elba': 'Portoferraio',
    "isola d'elba": 'Portoferraio', 'valdarno': 'Montevarchi', 'mugello': 'Borgo San Lorenzo',
    'versilia': 'Viareggio', 'lunigiana': 'Aulla', 'garfagnana': 'Castelnuovo di Garfagnana',
    'val di cornia': 'Piombino', 'argentario': 'Porto Santo Stefano',
    'chianti': 'Greve in Chianti', 'valdichiana': 'Sinalunga', 'casentino': 'Poppi',
    'amiata': 'Abbadia San Salvatore'}
# --- TOSCANA-DATI-FINE ---

TOSC_PERIODI = [("tutto", "Sempre"), ("anno", "Quest'anno"), ("mese", "Questo mese"),
                ("30g", "Ultimi 30 giorni")]


def tosc_norm(s):
    """Nome senza accenni, maiuscole e punteggiatura (per confrontare i nomi delle sedi)."""
    s = unicodedata.normalize("NFKD", str(s or ""))
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = re.sub(r"[’`´]", "'", s)
    return re.sub(r"[^a-z0-9']+", " ", s).strip()


_TOSC_INDEX = []


def _tosc_index():
    if not _TOSC_INDEX:
        idx = {tosc_norm(n): n for n in TOSC_COMUNI}
        for k, v in TOSC_ALIAS.items():
            idx.setdefault(tosc_norm(k), v)
        _TOSC_INDEX.append((idx, sorted(idx, key=lambda k: (-len(k), k))))
    return _TOSC_INDEX[0]


def tosc_find(name):
    """Il comune di cui parla il nome della sede (Firenze, «Sede di Pisa», «DR Toscana»...)."""
    n = tosc_norm(name)
    if not n:
        return None
    idx, keys = _tosc_index()
    if n in idx:
        return idx[n]
    pad = " " + n + " "
    for k in keys:
        if " " + k + " " in pad:
            return idx[k]
    return None


def tosc_project(lon, lat):
    return ((lon - TOSC_LON0) * TOSC_K * TOSC_S, (TOSC_LAT1 - lat) * TOSC_S)


def tosc_unproject(x, y):
    """Da punto della mappa a (latitudine, longitudine)."""
    return (TOSC_LAT1 - y / TOSC_S, x / (TOSC_K * TOSC_S) + TOSC_LON0)


def tosc_position(name, overrides=None):
    """(lat, lon, da dove) oppure None se la sede non si riconosce."""
    k = tosc_norm(name)
    if not k:
        return None
    if overrides and k in overrides:
        return (overrides[k][0], overrides[k][1], "manuale")
    c = tosc_find(name)
    if c:
        return (TOSC_COMUNI[c][0], TOSC_COMUNI[c][1], c)
    return None


def load_map_pos(settings):
    """Le posizioni scelte a mano: {nome normalizzato: (lat, lon)}."""
    try:
        d = json.loads(settings.get("mappa_pos") or "{}")
    except (ValueError, AttributeError):
        return {}
    out = {}
    if isinstance(d, dict):
        for k, v in d.items():
            try:
                lat, lon = round(float(v[0]), 4), round(float(v[1]), 4)
            except (TypeError, ValueError, IndexError, KeyError):
                continue
            if 41.5 < lat < 45.0 and 9.0 < lon < 13.0:
                out[str(k)] = (lat, lon)
    return out


def dump_map_pos(d):
    return json.dumps({k: [v[0], v[1]] for k, v in d.items()}, ensure_ascii=False, sort_keys=True)


def tosc_period_from(kind, today=None):
    today = today or date.today()
    if kind == "anno":
        return today.replace(month=1, day=1).isoformat()
    if kind == "mese":
        return today.replace(day=1).isoformat()
    if kind == "30g":
        return (today - timedelta(days=29)).isoformat()
    return ""


def tosc_records(db, periodo="tutto", overrides=None, today=None):
    """Per ogni sede: interventi (nel periodo), da seguire (sempre), utenti, stampanti e posizione.
    Le sedi dell'elenco di base compaiono anche a zero, così si vede dove sono."""
    dal = tosc_period_from(periodo, today)
    recs = {}

    def rec(raw, peso=1.0):
        raw = (raw or "").strip()
        k = tosc_norm(raw)
        if not k:
            return None
        r = recs.get(k)
        if r is None:
            r = recs[k] = {"key": k, "raws": {}, "interventi": 0, "da_seguire": 0, "utenti": 0,
                           "stampanti": 0}
        r["raws"][raw] = r["raws"].get(raw, 0) + peso
        return r

    for s in SEDI_TOSCANA:
        rec(s, 0.5)
    senza = {"interventi": 0, "da_seguire": 0}
    for row in db.conn.execute("SELECT sede, stato, data_ora FROM interventi"):
        r = rec(row["sede"]) or senza
        if not dal or (row["data_ora"] or "")[:10] >= dal:
            r["interventi"] += 1
        if row["stato"] in STATI_APERTI:
            r["da_seguire"] += 1
    for row in db.conn.execute("SELECT sede FROM utenti"):
        r = rec(row["sede"])
        if r:
            r["utenti"] += 1
    for row in db.conn.execute("SELECT e.sede AS sede, COUNT(s.id) AS n FROM edifici e "
                               "LEFT JOIN stampanti s ON s.edificio_id = e.id GROUP BY e.id"):
        r = rec(row["sede"])
        if r:
            r["stampanti"] += row["n"]
    out = []
    for r in recs.values():
        raws = sorted(r["raws"].items(), key=lambda kv: (-kv[1], kv[0]))
        r["raws"] = [name for name, _ in raws]
        r["nome"] = r["raws"][0]
        pos = tosc_position(r["nome"], overrides)
        r["lat"], r["lon"], r["via"] = pos if pos else (None, None, "")
        out.append(r)
    out.sort(key=lambda r: (-r["interventi"], -r["da_seguire"], r["nome"].casefold()))
    return {"sedi": out, "senza_sede": senza}


def tosc_pin_radius(n):
    return min(46.0, 18 + 3.5 * math.sqrt(n))


def tosc_separate(pins, gap=2.0, rounds=60):
    """Allontana i pallini che si sovrapporrebbero (Carrara e Massa sono a pochi chilometri)."""
    for _ in range(rounds):
        moved = False
        for i in range(len(pins)):
            for j in range(i + 1, len(pins)):
                a, b = pins[i], pins[j]
                dx, dy = b["x"] - a["x"], b["y"] - a["y"]
                d = math.hypot(dx, dy)
                need = a["r"] + b["r"] + gap
                if d < need - 0.01:
                    if d < 1e-6:
                        dx, dy, d = 1.0, 0.0, 1.0
                    push = (need - d) / 2
                    ux, uy = dx / d, dy / d
                    a["x"] -= ux * push
                    a["y"] -= uy * push
                    b["x"] += ux * push
                    b["y"] += uy * push
                    moved = True
        if not moved:
            break


def tosc_pins(recs):
    """Le sedi con posizione, unite quando cadono nello stesso punto (Firenze e DR Toscana)."""
    groups = {}
    for r in recs:
        if r["lat"] is not None:
            groups.setdefault((round(r["lat"], 4), round(r["lon"], 4)), []).append(r)
    pins = []
    for (lat, lon), rs in groups.items():
        rs = sorted(rs, key=lambda r: (-r["interventi"], r["nome"].casefold()))
        n = sum(r["interventi"] for r in rs)
        x, y = tosc_project(lon, lat)
        pins.append({"x": x, "y": y, "tx": x, "ty": y, "r": tosc_pin_radius(n), "n": n,
                     "rosso": sum(r["da_seguire"] for r in rs),
                     "label": " + ".join(r["nome"] for r in rs), "sedi": [r["key"] for r in rs]})
    pins.sort(key=lambda p: (-p["n"], p["label"].casefold()))
    tosc_separate(pins)
    return pins


def tosc_place_labels(pins, fs=28.0):
    """Mette i nomi sotto, sopra, a destra o a sinistra dei pallini senza sovrapporli."""
    def over(a, b):
        return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]
    circles = [(p["x"] - p["r"], p["y"] - p["r"], p["x"] + p["r"], p["y"] + p["r"]) for p in pins]
    placed, out = [], []
    for i, p in enumerate(pins):
        tw = len(p["label"]) * fs * 0.56
        cands = [("middle", p["x"], p["y"] + p["r"] + fs * 0.95),
                 ("middle", p["x"], p["y"] - p["r"] - fs * 0.3),
                 ("start", p["x"] + p["r"] + 6, p["y"] + fs * 0.35),
                 ("end", p["x"] - p["r"] - 6, p["y"] + fs * 0.35)]
        first, chosen = None, None
        for anchor, tx, ty in cands:
            x0 = tx - tw / 2 if anchor == "middle" else (tx if anchor == "start" else tx - tw)
            shift = 0
            if x0 < 2:
                shift = 2 - x0
            elif x0 + tw > TOSC_W - 2:
                shift = (TOSC_W - 2) - (x0 + tw)
            x0, tx = x0 + shift, tx + shift
            box = (x0, ty - fs * 0.8, x0 + tw, ty + fs * 0.25)
            cand = {"anchor": anchor, "lx": tx, "ly": ty, "box": box}
            if first is None:
                first = cand
            if any(over(box, b) for b in placed) or any(
                    over(box, c) for j, c in enumerate(circles) if j != i):
                continue
            chosen = cand
            break
        chosen = chosen or first
        placed.append(chosen["box"])
        out.append(chosen)
    return out


def tosc_scene(pins, scale=1.0, fs=28.0):
    """Tutto quello che serve per disegnare la mappa (poligoni e pallini), già in pixel."""
    def flat(pts):
        res = []
        for lon, lat in pts:
            x, y = tosc_project(lon, lat)
            res += [round(x * scale, 1), round(y * scale, 1)]
        return res
    labels = tosc_place_labels(pins, fs)
    return {"w": int(TOSC_W * scale), "h": int(TOSC_H * scale), "land": flat(TOSC_OUTLINE),
            "islands": [flat(v) for v in TOSC_ISLANDS.values()],
            "pins": [{"cx": p["x"] * scale, "cy": p["y"] * scale, "r": p["r"] * scale, "n": p["n"],
                      "rosso": p["rosso"], "label": p["label"], "sedi": p["sedi"],
                      "lx": lb["lx"] * scale, "ly": lb["ly"] * scale, "anchor": lb["anchor"]}
                     for p, lb in zip(pins, labels)]}


def tosc_pin_at(pins, x, y, scale=1.0, margin=14.0):
    """Il pallino più vicino al punto toccato (x, y in pixel), se è abbastanza vicino."""
    best, bd = None, None
    for i, p in enumerate(pins):
        d = math.hypot(x / scale - p["x"], y / scale - p["y"])
        if d <= p["r"] + margin and (bd is None or d < bd):
            best, bd = i, d
    return best


def templates_dir():
    """Cartella dei modelli Word dei verbali (si possono modificare a mano)."""
    return os.path.join(data_dir(), "Modelli verbali")


def saved_signature_path():
    """Immagine della firma del tecnico (scelta in «Dati per i verbali»). Sta tra gli allegati,
    quindi finisce anche nel backup."""
    return os.path.join(attachments_root(), "firma_tecnico.png")


def saved_signature():
    p = saved_signature_path()
    return p if os.path.isfile(p) else None


def save_signature_image(src):
    """Prepara l'immagine scelta (ritaglia i margini, toglie il fondo bianco) e la salva come
    «la mia firma». Solleva VB.VerbaleError se l'immagine non va bene."""
    if os.path.normcase(os.path.abspath(src)) == os.path.normcase(
            os.path.abspath(saved_signature_path())):
        return saved_signature_path()
    tmp = VB.prepara_firma(src, staging_dir_firme(), "firma_tecnico_nuova")
    try:
        shutil.copyfile(tmp, saved_signature_path())
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass
    return saved_signature_path()


def staging_dir_firme():
    path = os.path.join(attachments_root(), "_temp")
    os.makedirs(path, exist_ok=True)
    return path


def stanze_foto_root():
    """Foto e panorami delle stanze (inclusi nel backup)."""
    path = os.path.join(attachments_root(), "stanze")
    os.makedirs(path, exist_ok=True)
    return path


def piante_root():
    """Foto delle piante dei piani (incluse nel backup)."""
    path = os.path.join(attachments_root(), "piante")
    os.makedirs(path, exist_ok=True)
    return path


def pianta_uid(edificio_uid, piano):
    """Identificativo fisso di una pianta: telefono e PC lo costruiscono allo stesso modo."""
    return "pi:%s:%d" % (edificio_uid, piano)


def verbali_foto_root():
    """Qui restano le foto allegate ai verbali (sono incluse nel backup)."""
    path = os.path.join(attachments_root(), "verbali")
    os.makedirs(path, exist_ok=True)
    return path


def parse_iso_date(text):
    try:
        return date.fromisoformat((text or "")[:10])
    except ValueError:
        return None


def verbale_stato_word(r):
    """Da creare (mai fatto), Da rifare (dati cambiati dopo il Word) o Creato."""
    if not r.get("word_il"):
        return "Da creare"
    return "Da rifare" if (r.get("modificato") or "") > r["word_il"] else "Creato"


def verbale_firme_text(r):
    """Chi ha firmato (col dito o con l'immagine) e chi firmerà a mano sul foglio (cartacea)."""
    try:
        dati = json.loads(r.get("dati") or "{}")
    except ValueError:
        dati = {}
    firme = dati.get("_firme") or {}
    cart = dati.get("_cartaceo") or {}
    nomi = []
    for k in ("utente", "tecnico"):
        if firme.get(k):
            nomi.append(k)
        elif cart.get(k):
            nomi.append("%s (cartacea)" % k)
    return " + ".join(nomi)


def riassunto_verbale(tipo, d):
    """(soggetto, dettaglio) mostrati nell'elenco dei verbali."""
    def join(*parts):
        return " · ".join(p.strip() for p in parts if p and p.strip())

    g = lambda k: (d.get(k) or "").strip()
    sn = ("S/N " + g("seriale")) if g("seriale") else ""
    if tipo == "trasferimento":
        rotta = " → ".join(x for x in (g("provenienza"), g("destinazione")) if x)
        return g("assegnatario") or g("tipologia"), join(g("tipologia"), g("marca_modello"), sn, rotta)
    if tipo == "presa_in_carico":
        ticket = join(("INC " + g("inc")) if g("inc") else "", ("RIC " + g("ric")) if g("ric") else "")
        return g("utente"), join(g("tipologia"), g("marca_modello"), sn, ticket)
    if tipo == "assegnazione":
        voci = [join(x.get("tipo", ""), x.get("modello", "")) for x in d.get("dotazione") or []]
        return g("utente"), ", ".join(v for v in voci if v)
    if tipo == "censimento":
        n = len(d.get("apparati") or [])
        return g("oggetto"), "%d apparecchi" % n + ((" · " + g("tecnico")) if g("tecnico") else "")
    return "", ""


# ---------------------------------------------------------------------------------------------
# Verbali in serie: un Word per ogni persona di una stanza o di un piano
# ---------------------------------------------------------------------------------------------
SERIE_TIPI = ("assegnazione", "presa_in_carico")      # i verbali che hanno senso «per ogni persona»
SERIE_COSA = (("pc", "Il PC (notebook)"), ("cellulare", "Il cellulare"))
_SERIE_ORDINE = {"laptop": 0, "monitor": 1, "mobile": 2, "teams": 3}


def _code_key(text):
    return re.sub(r"[\s-]", "", str(text or "")).casefold()


def serie_dotazione(person):
    """Gli oggetti da scrivere nel verbale «Assegnazione materiale» di una persona, presi dalla sua
    scheda: [{"tipo", "modello", "codice"}]. Il PC (seriale del PC) e il cellulare (IMEI) ci sono se
    scritti; la presa di rete non è un oggetto assegnato, quindi non c'è. Vuoto = niente da assegnare."""
    pc, imei = clean_serial(person.get("seriale_pc")), clean_serial(person.get("imei"))
    items, pc_ok, imei_ok = [], False, False
    for r in clean_dotazione(person.get("dotazione")):
        kind = _row_kind(r)
        if kind == "presa":
            continue
        codice = r["seriale"]
        if kind == "laptop" and not codice and pc and not pc_ok:
            codice = pc
        if kind == "mobile" and not codice and imei and not imei_ok:
            codice = imei
        if codice and pc and _code_key(codice) == _code_key(pc):
            pc_ok = True
        if codice and imei and _code_key(codice) == _code_key(imei):
            imei_ok = True
        items.append((_SERIE_ORDINE.get(kind, 4), {"tipo": r["tipo"] or r["modello"] or "Apparecchio",
                                                  "modello": r["modello"] if r["tipo"] else "",
                                                  "codice": codice}))
    if pc and not pc_ok:
        items.append((0, {"tipo": "Notebook aziendale", "modello": "", "codice": pc}))
    if imei and not imei_ok:
        items.append((2, {"tipo": "Smartphone aziendale", "modello": "", "codice": imei}))
    items.sort(key=lambda x: x[0])
    return [it for _, it in items]


def serie_dati(tipo, person, data, opz=None, titolo=""):
    """I dati del verbale di una persona (come li darebbe il modulo) oppure (None, motivo) se per
    questa persona non si può fare. opz: {"cosa": "pc"|"cellulare", "motivo": ...}."""
    opz = opz or {}
    nome = " ".join(str(person.get("nome") or "").split())
    mat = " ".join(str(person.get("matricola") or "").split())
    if tipo == "assegnazione":
        dot = serie_dotazione(person)
        if not dot:
            return None, "niente in dotazione"
        return {"data": data, "titolo": titolo, "utente": nome, "matricola": mat,
                "dotazione": dot, "accessori": []}, ""
    if tipo == "presa_in_carico":
        dot = clean_dotazione(person.get("dotazione"))
        info = desk_phone_info(dot, person.get("imei"))
        if opz.get("cosa") == "cellulare":
            ser = clean_serial(person.get("imei")) or next(
                (r["seriale"] for r in dot if _row_kind(r) == "mobile" and r["seriale"]), "")
            if not ser:
                return None, "senza IMEI del cellulare"
            modello = ("Samsung Galaxy " + info["label"]) if info["mobile"] in (
                "s24", "s24plus", "s25") else (info["label"] or "")
            tipologia = "Smartphone"
        else:
            ser = clean_serial(person.get("seriale_pc")) or next(
                (r["seriale"] for r in dot if _row_kind(r) == "laptop" and r["seriale"]), "")
            if not ser:
                return None, "senza seriale del PC"
            modello, tipologia = info["laptop_label"], "Notebook"
        motivo = (opz.get("motivo") or "").strip()
        return {"data": data, "tipologia": tipologia, "marca_modello": modello, "seriale": ser,
                "utente": "%s (matricola %s)" % (nome, mat) if mat else nome, "inc": "", "ric": "",
                "motivi": [motivo] if motivo else [], "altro": "", "materiale": []}, ""
    return None, "tipo di verbale non adatto"


def plan_verbali_serie(persone, tipo, data, opz=None, titoli=None):
    """Prepara i verbali in serie senza scrivere niente. persone: come da db.persone_of();
    titoli: {nome in minuscolo: «dott.»…}. Le persone sono in ordine di cognome.
    Restituisce {"verbali": [{"disp", "d", "persona"}], "saltati": [(disp, motivo)]}."""
    seen, rows = set(), []
    for p in persone:
        key = p.get("utente_id") or p.get("id") or p.get("nome")
        if key in seen:
            continue
        seen.add(key)
        rows.append(p)
    rows.sort(key=lambda p: ((p.get("ordine") or p.get("nome") or "").casefold(),
                             (p.get("nome") or "").casefold()))
    out = {"verbali": [], "saltati": []}
    for p in rows:
        disp = p.get("ordine") or p.get("nome") or ""
        titolo = (titoli or {}).get(str(p.get("nome") or "").casefold(), "")
        d, why = serie_dati(tipo, p, data, opz, titolo)
        if d is None:
            out["saltati"].append((disp, why))
        else:
            out["verbali"].append({"disp": disp, "d": d, "persona": p})
    return out


def unique_path(folder, stem, ext):
    path = os.path.join(folder, stem + ext)
    n = 2
    while os.path.exists(path):
        path = os.path.join(folder, "%s_%d%s" % (stem, n, ext))
        n += 1
    return path


def store_image(src, dest_dir):
    """Copia una foto nella cartella dell'intervento. Le JPEG molto grandi vengono ridotte
    (lato lungo massimo 2200 px) per non riempire il disco e restare sotto i limiti dei ticket."""
    os.makedirs(dest_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(src))
    ext = ext.lower()
    target = unique_path(dest_dir, stem, ext)
    if ext in (".jpg", ".jpeg"):
        try:
            from PIL import Image, ImageOps
            with Image.open(src) as im:
                if max(im.size) > MAX_IMG_PX:
                    im = ImageOps.exif_transpose(im)
                    im.thumbnail((MAX_IMG_PX, MAX_IMG_PX))
                    im.convert("RGB").save(target, "JPEG", quality=88)
                    return os.path.basename(target)
        except Exception:
            pass
    shutil.copy2(src, target)
    return os.path.basename(target)


def make_thumb_b64(path, size=THUMB_PX):
    try:
        from PIL import Image, ImageOps
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im)
            im.thumbnail((size, size))
            buf = io.BytesIO()
            im.convert("RGB").save(buf, "PNG")
        return base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        return None


def grab_clipboard_images(stage_dir):
    """Prende dagli appunti uno screenshot/immagine copiata (o dei file immagine copiati)."""
    try:
        from PIL import ImageGrab
        data = ImageGrab.grabclipboard()
    except Exception:
        return []
    if data is None:
        return []
    if isinstance(data, list):
        return [f for f in data if os.path.isfile(f) and os.path.splitext(f)[1].lower() in IMG_EXTS]
    if hasattr(data, "save"):
        os.makedirs(stage_dir, exist_ok=True)
        path = os.path.join(stage_dir, "incollata_%s.png" % datetime.now().strftime("%Y%m%d_%H%M%S_%f"))
        data.save(path, "PNG")
        return [path]
    return []


def open_path(path):
    try:
        os.startfile(path)  # solo Windows
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------
# Database
# --------------------------------------------------------------------------
# la stanza (codice) in cui sta una stampante, se la stanza c'è ancora
ROOM_OF_PRINTER = ("(SELECT r.codice FROM stanze r WHERE r.uid = s.stanza_uid "
                   "AND s.stanza_uid <> '' AND r.eliminata = 0 "
                   "AND r.edificio_id = s.edificio_id LIMIT 1) AS stanza_codice, ")

SCHEMA = """
CREATE TABLE IF NOT EXISTS utenti (
    id   INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT NOT NULL UNIQUE COLLATE NOCASE
);
CREATE TABLE IF NOT EXISTS interventi (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    data_ora   TEXT NOT NULL,
    utente_id  INTEGER NOT NULL REFERENCES utenti(id),
    richiesta  TEXT NOT NULL DEFAULT '',
    soluzione  TEXT NOT NULL DEFAULT '',
    note       TEXT NOT NULL DEFAULT '',
    ticket     TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_int_data   ON interventi(data_ora);
CREATE INDEX IF NOT EXISTS idx_int_utente ON interventi(utente_id);

CREATE TABLE IF NOT EXISTS edifici (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    nome       TEXT NOT NULL UNIQUE COLLATE NOCASE,
    sede       TEXT NOT NULL DEFAULT '',
    piano_min  INTEGER NOT NULL DEFAULT 0,
    piano_max  INTEGER NOT NULL DEFAULT 2
);
CREATE TABLE IF NOT EXISTS stampanti (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    edificio_id INTEGER NOT NULL REFERENCES edifici(id),
    piano       INTEGER NOT NULL DEFAULT 0,
    nome        TEXT NOT NULL DEFAULT '',
    modello     TEXT NOT NULL DEFAULT '',
    ip          TEXT NOT NULL DEFAULT '',
    note        TEXT NOT NULL DEFAULT '',
    seriale     TEXT NOT NULL DEFAULT '',
    fuori_servizio INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS stampanti_unite (
    da_uid TEXT PRIMARY KEY,
    a_uid  TEXT NOT NULL,
    quando TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS utenti_uniti (
    da_key  TEXT PRIMARY KEY,
    da_nome TEXT NOT NULL,
    a_nome  TEXT NOT NULL,
    quando  TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS toner (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    stampante_id INTEGER NOT NULL REFERENCES stampanti(id) ON DELETE CASCADE,
    colore       TEXT NOT NULL,
    data_ora     TEXT NOT NULL,
    note         TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_toner_stampante ON toner(stampante_id);
CREATE INDEX IF NOT EXISTS idx_toner_data      ON toner(data_ora);

CREATE TABLE IF NOT EXISTS allegati (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    intervento_id INTEGER NOT NULL REFERENCES interventi(id) ON DELETE CASCADE,
    nome          TEXT NOT NULL,
    file          TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_all_int ON allegati(intervento_id);

CREATE TABLE IF NOT EXISTS manutenzioni (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    stampante_id INTEGER NOT NULL REFERENCES stampanti(id) ON DELETE CASCADE,
    data_ora     TEXT NOT NULL,
    data_tecnico TEXT NOT NULL DEFAULT '',
    ticket       TEXT NOT NULL DEFAULT '',
    tecnico      TEXT NOT NULL DEFAULT '',
    problema     TEXT NOT NULL DEFAULT '',
    intervento   TEXT NOT NULL DEFAULT '',
    stato        TEXT NOT NULL DEFAULT 'Eseguito',
    uid          TEXT NOT NULL DEFAULT '',
    modificato   TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_man_stampante ON manutenzioni(stampante_id);
CREATE INDEX IF NOT EXISTS idx_man_data      ON manutenzioni(data_ora);

CREATE TABLE IF NOT EXISTS scorta (
    modello    TEXT NOT NULL COLLATE NOCASE,
    colore     TEXT NOT NULL,
    quantita   INTEGER NOT NULL DEFAULT 0,
    soglia     INTEGER NOT NULL DEFAULT 1,
    modificato TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (modello, colore)
);

CREATE TABLE IF NOT EXISTS bolle (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    uid                  TEXT NOT NULL UNIQUE,
    codice               TEXT NOT NULL DEFAULT '',
    numero               TEXT NOT NULL DEFAULT '',
    toner_codice         TEXT NOT NULL DEFAULT '',
    colore               TEXT NOT NULL,
    edificio_id          INTEGER REFERENCES edifici(id) ON DELETE CASCADE,
    piano                INTEGER NOT NULL DEFAULT 0,
    stampante_id         INTEGER REFERENCES stampanti(id) ON DELETE SET NULL,
    modello              TEXT NOT NULL DEFAULT '',
    seriale_bolla        TEXT NOT NULL DEFAULT '',
    data_arrivo          TEXT NOT NULL,
    stato                TEXT NOT NULL DEFAULT 'da_montare',
    data_montaggio       TEXT NOT NULL DEFAULT '',
    stampante_montata_id INTEGER REFERENCES stampanti(id) ON DELETE SET NULL,
    seriale_letto        TEXT NOT NULL DEFAULT '',
    toner_uid            TEXT NOT NULL DEFAULT '',
    note                 TEXT NOT NULL DEFAULT '',
    foto                 TEXT NOT NULL DEFAULT '',
    scorta_ok            INTEGER NOT NULL DEFAULT 0,
    modificato           TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_bolle_stato ON bolle(stato);

CREATE TABLE IF NOT EXISTS stanze (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    uid         TEXT NOT NULL DEFAULT '',
    edificio_id INTEGER NOT NULL REFERENCES edifici(id) ON DELETE CASCADE,
    piano       INTEGER NOT NULL DEFAULT 0,
    codice      TEXT NOT NULL,
    scrivanie   INTEGER NOT NULL DEFAULT 0,
    note        TEXT NOT NULL DEFAULT '',
    come        TEXT NOT NULL DEFAULT '',
    px          REAL,
    py          REAL,
    modificato  TEXT NOT NULL DEFAULT '',
    eliminata   INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_stanze_edificio ON stanze(edificio_id, piano);
CREATE INDEX IF NOT EXISTS idx_stanze_uid ON stanze(uid);

CREATE TABLE IF NOT EXISTS piante (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    uid         TEXT NOT NULL DEFAULT '',
    edificio_id INTEGER NOT NULL REFERENCES edifici(id) ON DELETE CASCADE,
    piano       INTEGER NOT NULL,
    file        TEXT NOT NULL DEFAULT '',
    modificato  TEXT NOT NULL DEFAULT '',
    UNIQUE (edificio_id, piano)
);

CREATE TABLE IF NOT EXISTS stanze_foto (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    uid        TEXT NOT NULL DEFAULT '',
    stanza_id  INTEGER NOT NULL REFERENCES stanze(id) ON DELETE CASCADE,
    nome       TEXT NOT NULL,
    tipo       TEXT NOT NULL DEFAULT 'foto',
    file       TEXT NOT NULL,
    modificato TEXT NOT NULL DEFAULT '',
    eliminata  INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS idx_stfoto_stanza ON stanze_foto(stanza_id);

CREATE TABLE IF NOT EXISTS postazioni (
    utente_id  INTEGER PRIMARY KEY REFERENCES utenti(id) ON DELETE CASCADE,
    stanza_id  INTEGER REFERENCES stanze(id) ON DELETE SET NULL,
    numero     INTEGER NOT NULL DEFAULT 0,
    modificato TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_post_stanza ON postazioni(stanza_id);

CREATE TABLE IF NOT EXISTS verbali (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    uid        TEXT NOT NULL DEFAULT '',
    tipo       TEXT NOT NULL,
    data       TEXT NOT NULL DEFAULT '',
    soggetto   TEXT NOT NULL DEFAULT '',
    dettaglio  TEXT NOT NULL DEFAULT '',
    dati       TEXT NOT NULL DEFAULT '{}',
    file       TEXT NOT NULL DEFAULT '',
    n_foto     INTEGER NOT NULL DEFAULT 0,
    creato     TEXT NOT NULL DEFAULT '',
    modificato TEXT NOT NULL DEFAULT '',
    word_il    TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_verbali_data ON verbali(data);
"""

SELECT_INTERVENTI = (
    "SELECT i.id, i.data_ora, u.nome AS utente, i.sede, i.richiesta, i.soluzione, "
    "i.note, i.ticket, i.stato, u.telefono, u.seriale_pc, u.imei, u.matricola, "
    "(SELECT COUNT(*) FROM allegati a WHERE a.intervento_id = i.id) AS n_foto "
    "FROM interventi i JOIN utenti u ON u.id = i.utente_id"
)

SELECT_TONER = (
    "SELECT t.id, t.data_ora, t.colore, t.note, s.id AS stampante_id, s.nome AS stampante, "
    "s.modello, s.piano, e.id AS edificio_id, e.nome AS edificio, e.sede "
    "FROM toner t JOIN stampanti s ON s.id = t.stampante_id "
    "JOIN edifici e ON e.id = s.edificio_id"
)

SELECT_MANUT = (
    "SELECT m.id, m.uid, m.data_ora, m.data_tecnico, m.ticket, m.tecnico, m.problema, "
    "m.intervento, m.stato, m.modificato, s.id AS stampante_id, s.nome AS stampante, s.modello, "
    "s.piano, e.id AS edificio_id, e.nome AS edificio, e.sede "
    "FROM manutenzioni m JOIN stampanti s ON s.id = m.stampante_id "
    "JOIN edifici e ON e.id = s.edificio_id"
)


class Database:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        self._migrate()
        self._seed()
        self.conn.commit()

    def close(self):
        self.conn.close()

    def _columns(self, table):
        return {r["name"] for r in self.conn.execute("PRAGMA table_info(%s)" % table)}

    def _migrate(self):
        """Aggiunge le colonne nuove ai database creati dalle versioni precedenti."""
        for col in ("sede", "telefono", "seriale_pc", "imei", "matricola"):
            if col not in self._columns("utenti"):
                self.conn.execute(
                    "ALTER TABLE utenti ADD COLUMN %s TEXT NOT NULL DEFAULT ''" % col)
        if "sede" not in self._columns("interventi"):
            self.conn.execute("ALTER TABLE interventi ADD COLUMN sede TEXT NOT NULL DEFAULT ''")
        if "modificato" not in self._columns("utenti"):
            self.conn.execute("ALTER TABLE utenti ADD COLUMN modificato TEXT NOT NULL DEFAULT ''")
        for col in ("ufficio", "dotazione", "ordine", "email"):
            if col not in self._columns("utenti"):
                self.conn.execute(
                    "ALTER TABLE utenti ADD COLUMN %s TEXT NOT NULL DEFAULT ''" % col)
        for col in ("uid", "modificato"):
            if col not in self._columns("interventi"):
                self.conn.execute(
                    "ALTER TABLE interventi ADD COLUMN %s TEXT NOT NULL DEFAULT ''" % col)
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_int_uid ON interventi(uid)")
        for table in ("edifici", "stampanti", "toner"):
            if "uid" not in self._columns(table):
                self.conn.execute("ALTER TABLE %s ADD COLUMN uid TEXT NOT NULL DEFAULT ''" % table)
        if "seriale" not in self._columns("stampanti"):
            self.conn.execute("ALTER TABLE stampanti ADD COLUMN seriale TEXT NOT NULL DEFAULT ''")
        if "fuori_servizio" not in self._columns("stampanti"):
            self.conn.execute(
                "ALTER TABLE stampanti ADD COLUMN fuori_servizio INTEGER NOT NULL DEFAULT 0")
        if "stanza_uid" not in self._columns("stampanti"):
            self.conn.execute("ALTER TABLE stampanti ADD COLUMN stanza_uid TEXT NOT NULL DEFAULT ''")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_stampanti_uid ON stampanti(uid)")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_toner_uid ON toner(uid)")
        if "stato" not in self._columns("interventi"):
            # gli interventi già registrati sono tutti da considerare chiusi
            self.conn.execute(
                "ALTER TABLE interventi ADD COLUMN stato TEXT NOT NULL DEFAULT 'Risolto'")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_man_uid ON manutenzioni(uid)")
        if "word_il" not in self._columns("verbali"):
            self.conn.execute("ALTER TABLE verbali ADD COLUMN word_il TEXT NOT NULL DEFAULT ''")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_verbali_uid ON verbali(uid)")
        self.conn.commit()

    def _seed(self):
        """Alla prima apertura prepara l'edificio di Firenze (si può cambiare o eliminare)."""
        version = self.conn.execute("PRAGMA user_version").fetchone()[0]
        if version >= 1:
            return
        n = self.conn.execute("SELECT COUNT(*) FROM edifici").fetchone()[0]
        if n == 0:
            self.conn.execute(
                "INSERT INTO edifici(nome, sede, piano_min, piano_max) VALUES (?,?,?,?)",
                ("Sede di Firenze - Via delle Porte Nuove", "Firenze", 0, 5))
        self.conn.execute("PRAGMA user_version = 1")

    # --- utenti ---
    def users(self):
        rows = self.conn.execute("SELECT nome FROM utenti ORDER BY %s" % SORT_UTENTE.replace("u.", ""))
        return [r["nome"] for r in rows]

    def users_with_counts(self):
        rows = self.conn.execute(
            "SELECT u.id, u.nome, u.matricola, u.sede, u.telefono, u.seriale_pc, u.imei, "
            "u.modificato, u.ufficio, u.dotazione, u.ordine, u.email, "
            "COUNT(i.id) AS n, "
            "MAX(i.data_ora) AS ultimo "
            "FROM utenti u LEFT JOIN interventi i ON i.utente_id = u.id "
            "GROUP BY u.id ORDER BY " + SORT_UTENTE
        )
        return [self._user_dict(r) for r in rows]

    @staticmethod
    def _user_dict(row):
        """Riga utente come dizionario, con la dotazione già letta come elenco."""
        d = dict(row)
        if "dotazione" in d:
            d["dotazione"] = clean_dotazione(d["dotazione"])
        return d

    def set_user_extra(self, user_id, ufficio=None, dotazione=None, modificato=None):
        """Imposta ufficio e/o dotazione di una persona (None = lascia com'è). Con valori uguali
        a quelli già salvati non cambia nulla, nemmeno la data di modifica."""
        row = self.conn.execute("SELECT ufficio, dotazione FROM utenti WHERE id = ?",
                                (user_id,)).fetchone()
        if not row:
            return False
        new_uff = row["ufficio"] if ufficio is None else " ".join(str(ufficio).split())[:120]
        old_dot = clean_dotazione(row["dotazione"])
        new_dot = old_dot if dotazione is None else clean_dotazione(dotazione)
        if new_uff == row["ufficio"] and new_dot == old_dot:
            return False
        with self.conn:
            self.conn.execute("UPDATE utenti SET ufficio = ?, dotazione = ?, modificato = ? "
                              "WHERE id = ?", (new_uff, json.dumps(new_dot, ensure_ascii=False)
                                               if new_dot else "", modificato or now_utc_iso(),
                                               user_id))
        return True

    def set_user_codes(self, user_id, seriale_pc=None, imei=None, matricola=None, modificato=None):
        """Scrive solo i codici indicati (None = lascia com'è): seriale del PC, IMEI, matricola.
        Restituisce True se qualcosa è cambiato (con valori uguali non cambia nemmeno la data)."""
        row = self.conn.execute("SELECT seriale_pc, imei, matricola FROM utenti WHERE id = ?",
                                (user_id,)).fetchone()
        if not row:
            return False
        new = {"seriale_pc": row["seriale_pc"] if seriale_pc is None else clean_serial(seriale_pc),
               "imei": row["imei"] if imei is None else " ".join(str(imei).split()),
               "matricola": row["matricola"] if matricola is None else " ".join(str(matricola).split())}
        if all(new[k] == row[k] for k in new):
            return False
        with self.conn:
            self.conn.execute("UPDATE utenti SET seriale_pc = ?, imei = ?, matricola = ?, "
                              "modificato = ? WHERE id = ?",
                              (new["seriale_pc"], new["imei"], new["matricola"],
                               modificato or now_utc_iso(), user_id))
        return True

    def set_user_basic(self, user_id, sede=None, telefono=None, modificato=None):
        """Imposta sede e/o telefono di una persona (None = lascia com'è). Con valori uguali a
        quelli già salvati non cambia nulla, nemmeno la data di modifica."""
        row = self.conn.execute("SELECT sede, telefono FROM utenti WHERE id = ?",
                                (user_id,)).fetchone()
        if not row:
            return False
        new_sede = row["sede"] if sede is None else " ".join(str(sede).split())
        new_tel = row["telefono"] if telefono is None else " ".join(str(telefono).split())
        if new_sede == row["sede"] and new_tel == row["telefono"]:
            return False
        with self.conn:
            self.conn.execute("UPDATE utenti SET sede = ?, telefono = ?, modificato = ? "
                              "WHERE id = ?", (new_sede, new_tel, modificato or now_utc_iso(),
                                               user_id))
        return True

    def set_user_email(self, user_id, email, modificato=None):
        """Imposta l'email di una persona (vuoto = la toglie). Con un valore uguale a quello già
        salvato non cambia nulla, nemmeno la data di modifica."""
        row = self.conn.execute("SELECT email FROM utenti WHERE id = ?", (user_id,)).fetchone()
        if not row:
            return False
        new = clean_email(email)
        if new == row["email"]:
            return False
        with self.conn:
            self.conn.execute("UPDATE utenti SET email = ?, modificato = ? WHERE id = ?",
                              (new, modificato or now_utc_iso(), user_id))
        return True

    def set_user_ordine(self, user_id, ordine, modificato=None):
        """Imposta il «Cognome Nome» con cui si ordina la persona (vuoto = si ordina per il nome).
        Con un valore uguale a quello già salvato non cambia nulla, nemmeno la data di modifica."""
        row = self.conn.execute("SELECT nome, ordine FROM utenti WHERE id = ?",
                                (user_id,)).fetchone()
        if not row:
            return False
        new = " ".join(str(ordine or "").split())[:120]
        if key_name(new) == key_name(row["nome"]):
            new = ""                            # uguale al nome: non serve
        if new == row["ordine"]:
            return False
        with self.conn:
            self.conn.execute("UPDATE utenti SET ordine = ?, modificato = ? WHERE id = ?",
                              (new, modificato or now_utc_iso(), user_id))
        return True

    def users_basic(self):
        """Tutte le persone con i dati che servono a riconoscerle (per l'importazione da Excel)."""
        return [dict(r) for r in self.conn.execute(
            "SELECT id, nome, matricola, sede, telefono, ufficio, ordine, seriale_pc, email, "
            "imei, dotazione FROM utenti ORDER BY id")]

    def uffici(self):
        """Gli uffici già scritti (senza doppioni), per il completamento automatico."""
        found = {}
        for r in self.conn.execute("SELECT DISTINCT ufficio FROM utenti WHERE ufficio <> ''"):
            found.setdefault(r[0].casefold(), r[0])
        return sorted(found.values(), key=str.casefold)

    def find_user(self, name):
        row = self.conn.execute(
            "SELECT id, nome, matricola, sede, telefono, seriale_pc, imei, modificato, "
            "ufficio, dotazione, ordine, email FROM utenti WHERE nome = ?",
            (normalize_name(name),)
        ).fetchone()
        return self._user_dict(row) if row else None

    def find_user_like(self, name):
        """La persona con questo nome; riconosce anche «Rossi Mario» per «Mario Rossi» (stesse
        parole in ordine diverso), purché ce ne sia una sola."""
        exact = self.find_user(name)
        if exact:
            return exact
        words = sorted(key_name(name).split())
        if len(words) < 2:
            return None
        hits = [r["id"] for r in self.conn.execute("SELECT id, nome FROM utenti")
                if sorted(key_name(r["nome"]).split()) == words]
        if len(hits) != 1:
            return None
        row = self.conn.execute(
            "SELECT id, nome, matricola, sede, telefono, seriale_pc, imei, modificato, "
            "ufficio, dotazione, ordine, email FROM utenti WHERE id = ?", (hits[0],)).fetchone()
        return self._user_dict(row)

    def find_user_by_matricola(self, matricola):
        m = (matricola or "").strip()
        if not m:
            return None
        row = self.conn.execute(
            "SELECT id, nome, matricola, sede, telefono, seriale_pc, imei, modificato, "
            "ufficio, dotazione, ordine, email FROM utenti WHERE matricola = ? COLLATE NOCASE "
            "ORDER BY nome LIMIT 1", (m,)
        ).fetchone()
        return self._user_dict(row) if row else None

    def sedi(self):
        found = list(SEDI_TOSCANA)
        for q in ("SELECT DISTINCT sede FROM utenti", "SELECT DISTINCT sede FROM interventi",
                  "SELECT DISTINCT sede FROM edifici"):
            found += [r[0].strip() for r in self.conn.execute(q) if r[0] and r[0].strip()]
        unique = {}
        for v in found:
            unique.setdefault(v.casefold(), v)
        return sorted(unique.values(), key=str.casefold)

    def _get_or_create_user(self, name, sede="", telefono="", seriale_pc="", imei="",
                            matricola="", update=True, modificato=None):
        name = normalize_name(name)
        row = self.conn.execute(
            "SELECT id, sede, telefono, seriale_pc, imei, matricola FROM utenti WHERE nome = ?",
            (name,)).fetchone()
        new = (sede, telefono, seriale_pc, imei, matricola)
        if row:
            if update and tuple(row[k] for k in ("sede", "telefono", "seriale_pc", "imei",
                                                 "matricola")) != new:
                self.conn.execute(
                    "UPDATE utenti SET sede=?, telefono=?, seriale_pc=?, imei=?, matricola=?, "
                    "modificato=? WHERE id=?", new + (modificato or now_utc_iso(), row["id"]))
            return row["id"], False
        cur = self.conn.execute(
            "INSERT INTO utenti(nome, sede, telefono, seriale_pc, imei, matricola, modificato) "
            "VALUES (?,?,?,?,?,?,?)", (name,) + new + (modificato or now_utc_iso(),))
        # una persona nuova con il nome di una già unita/rinominata è davvero nuova: l'unione cade
        self.conn.execute("DELETE FROM utenti_uniti WHERE da_key = ?", (key_name(name),))
        return cur.lastrowid, True

    def add_user(self, name, sede="", telefono="", seriale_pc="", imei="", matricola="",
                 modificato=None):
        with self.conn:
            return self._get_or_create_user(name, sede, telefono, seriale_pc, imei, matricola,
                                            update=False, modificato=modificato)

    def update_user(self, user_id, name, sede="", telefono="", seriale_pc="", imei="",
                    matricola="", modificato=None):
        old = self.conn.execute("SELECT nome FROM utenti WHERE id = ?", (user_id,)).fetchone()
        with self.conn:
            if old and sorted(key_name(old["nome"]).split()) != sorted(key_name(name).split()):
                self.conn.execute("UPDATE utenti SET ordine = '' WHERE id = ?", (user_id,))
            self.conn.execute(
                "UPDATE utenti SET nome=?, sede=?, telefono=?, seriale_pc=?, imei=?, matricola=?, "
                "modificato=? WHERE id=?", (normalize_name(name), sede, telefono, seriale_pc, imei,
                                            matricola, modificato or now_utc_iso(), user_id))
            if old:
                self._note_user_rename(old["nome"], normalize_name(name))

    # ----- persone unite o rinominate: il telefono deve saperlo, altrimenti ne farebbe tornare una -----
    def _note_user_rename(self, old, new):
        """Ricorda «il nome vecchio ora è questo nome nuovo» (da dentro una transazione)."""
        ko, kn = key_name(old), key_name(new)
        if ko == kn:
            return
        for r in self.conn.execute("SELECT da_key, a_nome FROM utenti_uniti").fetchall():
            if key_name(r["a_nome"]) == ko:             # chi era finito nel nome vecchio segue
                self.conn.execute("UPDATE utenti_uniti SET a_nome = ? WHERE da_key = ?",
                                  (new, r["da_key"]))
        self.conn.execute("DELETE FROM utenti_uniti WHERE da_key = ?", (kn,))
        self.conn.execute("INSERT OR REPLACE INTO utenti_uniti(da_key, da_nome, a_nome, quando) "
                          "VALUES (?,?,?,?)", (ko, normalize_name(old), normalize_name(new),
                                               now_utc_iso()))

    def user_merges(self):
        """Le unioni e i cambi di nome fatti («nome vecchio» -> «nome che resta»), da dire al telefono."""
        return [{"da": r["da_nome"], "a": r["a_nome"]} for r in self.conn.execute(
            "SELECT da_nome, a_nome FROM utenti_uniti ORDER BY quando, da_key")]

    def resolve_user_name(self, name):
        """Il nome con cui una persona esiste oggi: se «name» è di una persona unita a un'altra (o
        rinominata) e non c'è più nessuno con questo nome, il nome di quella che è rimasta. Così
        quello che arriva ancora dal telefono col nome vecchio finisce sulla persona giusta."""
        name = normalize_name(name)
        if not name or self.find_user(name):
            return name
        cur, hops = name, 0
        while hops < 8:
            row = self.conn.execute("SELECT a_nome FROM utenti_uniti WHERE da_key = ?",
                                    (key_name(cur),)).fetchone()
            if not row:
                break
            cur, hops = row["a_nome"], hops + 1
            if self.find_user(cur):
                return cur
        return name

    def merge_users(self, source_id, target_id):
        """Unisce due persone che sono la stessa: gli interventi e la postazione di «source» passano
        a «target» e «source» si toglie. «target» tiene i suoi dati; dove gli mancano (matricola,
        sede, telefono, ufficio, email, seriali, dotazione) prende quelli di «source». Il telefono
        lo viene a sapere dal file di scambio. Restituisce {"interventi": n, "presi": [campi],
        "diversi": [campi con dati diversi: restano quelli di target]}."""
        if source_id == target_id:
            raise ValueError("Sono la stessa persona.")
        cols = ("nome, sede, telefono, seriale_pc, imei, matricola, ufficio, dotazione, ordine, "
                "email")
        src = self.conn.execute("SELECT %s FROM utenti WHERE id = ?" % cols, (source_id,)).fetchone()
        tgt = self.conn.execute("SELECT %s FROM utenti WHERE id = ?" % cols, (target_id,)).fetchone()
        if not src or not tgt:
            raise ValueError("Persona non trovata.")
        new, presi, diversi = {}, [], []
        for f in ("sede", "telefono", "seriale_pc", "imei", "matricola", "ufficio", "email"):
            a, b = (tgt[f] or "").strip(), (src[f] or "").strip()
            if b and not a:
                new[f] = b
                presi.append(f)
            elif a and b and (key_name(a) != key_name(b) if f != "telefono"
                              else _phone_key(a) != _phone_key(b)):
                diversi.append(f)
        dot_t, dot_s = clean_dotazione(tgt["dotazione"]), clean_dotazione(src["dotazione"])
        dot = merge_dotazione(dot_t, dot_s)
        if dot != dot_t:
            new["dotazione"] = json.dumps(dot, ensure_ascii=False)
            presi.append("dotazione")
        if not tgt["ordine"] and src["ordine"] and key_name(src["ordine"]) != key_name(tgt["nome"]):
            new["ordine"] = src["ordine"]
        if self.posto_of_user(source_id):
            if self.posto_of_user(target_id):
                self.rimuovi_persona(source_id)       # il bersaglio ha già la sua postazione
            else:
                with self.conn:                       # la postazione passa al bersaglio
                    self.conn.execute("DELETE FROM postazioni WHERE utente_id = ?", (target_id,))
                    self.conn.execute("UPDATE postazioni SET utente_id = ?, modificato = ? "
                                      "WHERE utente_id = ?",
                                      (target_id, now_utc_iso(), source_id))
        with self.conn:
            n = self.conn.execute(
                "UPDATE interventi SET utente_id = ? WHERE utente_id = ?", (target_id, source_id)
            ).rowcount
            if new:
                sets = ", ".join("%s = ?" % k for k in new)
                self.conn.execute("UPDATE utenti SET %s, modificato = ? WHERE id = ?" % sets,
                                  tuple(new.values()) + (now_utc_iso(), target_id))
            self.conn.execute("DELETE FROM utenti WHERE id = ?", (source_id,))
            self._note_user_rename(src["nome"], tgt["nome"])
        return {"interventi": n, "presi": presi, "diversi": diversi}

    def delete_user(self, user_id):
        n = self.conn.execute(
            "SELECT COUNT(*) FROM interventi WHERE utente_id = ?", (user_id,)
        ).fetchone()[0]
        if n:
            return False
        self.rimuovi_persona(user_id)       # libera la postazione e rinumera la stanza
        with self.conn:
            self.conn.execute("DELETE FROM utenti WHERE id = ?", (user_id,))
        return True

    # --- interventi ---
    def add_intervento(self, data_ora, utente, richiesta, soluzione, note, ticket,
                       sede="", telefono="", seriale_pc="", imei="", matricola="",
                       stato="Risolto"):
        with self.conn:
            uid, created = self._get_or_create_user(utente, sede, telefono, seriale_pc, imei,
                                                    matricola)
            cur = self.conn.execute(
                "INSERT INTO interventi(data_ora, utente_id, sede, richiesta, soluzione, note, "
                "ticket, stato, modificato) VALUES (?,?,?,?,?,?,?,?,?)",
                (data_ora, uid, sede, richiesta, soluzione, note, ticket, stato, now_utc_iso()),
            )
        return cur.lastrowid, created

    def update_intervento(self, iid, data_ora, utente, richiesta, soluzione, note, ticket,
                          sede="", telefono="", seriale_pc="", imei="", matricola="",
                          stato="Risolto"):
        with self.conn:
            uid, created = self._get_or_create_user(utente, sede, telefono, seriale_pc, imei,
                                                    matricola)
            self.conn.execute(
                "UPDATE interventi SET data_ora=?, utente_id=?, sede=?, richiesta=?, soluzione=?, "
                "note=?, ticket=?, stato=?, modificato=? WHERE id=?",
                (data_ora, uid, sede, richiesta, soluzione, note, ticket, stato, now_utc_iso(),
                 iid),
            )
        return created

    def interventi_per_import(self):
        """Gli interventi già scritti, per riconoscere quelli che l'importazione ha già messo."""
        return [dict(r) for r in self.conn.execute(
            "SELECT id, ticket, soluzione, note, utente_id FROM interventi")]

    def update_intervento_import(self, iid, soluzione=None, utente=None, togli_dedotto=False):
        """Dopo un'importazione: scrive la soluzione (solo se è vuota) e/o cambia la persona di un
        intervento importato con la persona «dedotta» (e toglie la dicitura che lo diceva)."""
        row = self.conn.execute("SELECT soluzione, note FROM interventi WHERE id = ?",
                                (iid,)).fetchone()
        if not row:
            return False
        sol, note, uid = row["soluzione"], row["note"], None
        if soluzione and not (sol or "").strip():
            sol = soluzione
        with self.conn:
            if utente:
                uid, _ = self._get_or_create_user(utente, update=False)
            if togli_dedotto:
                note = re.sub(r"( - utente dedotto da [^\n]*| - utente non riconosciuto)(?=\n|$)",
                              "", note or "")
            if uid is not None:
                self.conn.execute("UPDATE interventi SET utente_id=?, soluzione=?, note=?, "
                                  "modificato=? WHERE id=?", (uid, sol, note, now_utc_iso(), iid))
            else:
                self.conn.execute("UPDATE interventi SET soluzione=?, note=?, modificato=? "
                                  "WHERE id=?", (sol, note, now_utc_iso(), iid))
        return True

    def set_stato(self, iid, stato):
        with self.conn:
            self.conn.execute("UPDATE interventi SET stato=?, modificato=? WHERE id=?",
                              (stato, now_utc_iso(), iid))

    def count_aperti(self):
        return self.conn.execute(
            "SELECT COUNT(*) FROM interventi WHERE stato IN ('Aperto','In attesa')").fetchone()[0]

    def delete_intervento(self, iid):
        with self.conn:
            self.conn.execute("DELETE FROM interventi WHERE id = ?", (iid,))

    # --- scambio con l'app del telefono ---
    def find_intervento_by_uid(self, uid):
        row = self.conn.execute(
            "SELECT id, modificato, stato FROM interventi WHERE uid = ?", (uid,)).fetchone()
        return dict(row) if row else None

    def save_phone_intervento(self, existing_id, uid, data_ora, utente, sede, richiesta,
                              soluzione, note, ticket, modificato, stato="Risolto"):
        """Inserisce (o aggiorna) un intervento arrivato dal telefono. Restituisce l'id."""
        with self.conn:
            user_id, _ = self._get_or_create_user(utente, update=False)
            if existing_id:
                self.conn.execute(
                    "UPDATE interventi SET data_ora=?, utente_id=?, sede=?, richiesta=?, "
                    "soluzione=?, note=?, ticket=?, stato=?, modificato=? WHERE id=?",
                    (data_ora, user_id, sede, richiesta, soluzione, note, ticket, stato,
                     modificato, existing_id))
                return existing_id
            cur = self.conn.execute(
                "INSERT INTO interventi(data_ora, utente_id, sede, richiesta, soluzione, note, "
                "ticket, stato, uid, modificato) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (data_ora, user_id, sede, richiesta, soluzione, note, ticket, stato, uid,
                 modificato))
            return cur.lastrowid

    def ensure_sync_ids(self):
        """Dà un identificativo (uid) a tutto ciò che nasce sul PC, così il telefono e il PC
        si riconoscono a vicenda e non creano doppioni."""
        now = now_utc_iso()
        with self.conn:
            self.conn.execute("UPDATE interventi SET modificato = ? WHERE modificato = ''", (now,))
            self.conn.execute("UPDATE manutenzioni SET modificato = ? WHERE modificato = ''",
                              (now,))
            for table, prefix in (("interventi", "pc-"), ("edifici", "ed-"),
                                  ("stampanti", "st-"), ("toner", "to-"),
                                  ("manutenzioni", "mt-")):
                for r in self.conn.execute("SELECT id FROM %s WHERE uid = ''" % table).fetchall():
                    self.conn.execute("UPDATE %s SET uid = ? WHERE id = ?" % table,
                                      (prefix + uuid.uuid4().hex, r["id"]))

    def find_stampante_by_uid(self, uid):
        """La stampante con questo uid; se è stata unita a un'altra (doppione tolto) segue l'unione,
        così i cambi toner e gli interventi che arrivano ancora col vecchio uid finiscono sulla giusta."""
        row = self.conn.execute("SELECT id FROM stampanti WHERE uid = ?", (uid,)).fetchone()
        hops = 0
        while row is None and uid and hops < 8:
            al = self.conn.execute("SELECT a_uid FROM stampanti_unite WHERE da_uid = ?",
                                   (uid,)).fetchone()
            if not al:
                break
            uid, hops = al["a_uid"], hops + 1
            row = self.conn.execute("SELECT id FROM stampanti WHERE uid = ?", (uid,)).fetchone()
        return row["id"] if row else None

    def toner_uid_exists(self, uid):
        return self.conn.execute("SELECT 1 FROM toner WHERE uid = ?", (uid,)).fetchone() is not None

    def add_toner_from_phone(self, stampante_id, uid, colore, data_ora, note):
        """Registra un cambio toner del telefono; se la scorta è tracciata la scala di uno.
        Restituisce la riga di scorta aggiornata (o None)."""
        with self.conn:
            self.conn.execute(
                "INSERT INTO toner(stampante_id, colore, data_ora, note, uid) VALUES (?,?,?,?,?)",
                (stampante_id, colore, data_ora, note, uid))
            return self._scorta_scale(self._modello_of(stampante_id), colore, -1)

    # --- foto allegate agli interventi ---
    def allegati_of(self, iid):
        rows = self.conn.execute(
            "SELECT id, nome, file FROM allegati WHERE intervento_id = ? ORDER BY id", (iid,))
        return [dict(r) for r in rows]

    def add_allegato(self, iid, nome, file):
        with self.conn:
            self.conn.execute(
                "INSERT INTO allegati(intervento_id, nome, file) VALUES (?,?,?)", (iid, nome, file))

    def delete_allegato(self, aid):
        with self.conn:
            self.conn.execute("DELETE FROM allegati WHERE id = ?", (aid,))

    def search(self, utente="", sede="", dal=None, al=None, ticket_mode="tutti", ticket="",
               testo="", stato=""):
        sql = SELECT_INTERVENTI + " WHERE 1=1"
        params = []
        if dal:
            sql += " AND substr(i.data_ora,1,10) >= ?"
            params.append(dal.isoformat())
        if al:
            sql += " AND substr(i.data_ora,1,10) <= ?"
            params.append(al.isoformat())
        if ticket_mode == "con":
            sql += " AND i.ticket <> ''"
        elif ticket_mode == "senza":
            sql += " AND i.ticket = ''"
        if stato == "aperti":
            sql += " AND i.stato IN ('Aperto','In attesa')"
        elif stato in STATI_INTERVENTO:
            sql += " AND i.stato = ?"
            params.append(stato)
        sql += " ORDER BY i.data_ora DESC, i.id DESC"
        rows = [dict(r) for r in self.conn.execute(sql, params)]

        terms = utente.casefold().split()
        if terms:
            rows = [r for r in rows if matches_all(terms, r["utente"] + " " + r["matricola"])]
        s = sede.strip().casefold()
        if s:
            rows = [r for r in rows if s in r["sede"].casefold()]
        t = ticket.strip().casefold()
        if t:
            rows = [r for r in rows if t in r["ticket"].casefold()]
        terms = testo.casefold().split()
        if terms:
            rows = [
                r for r in rows
                if matches_all(terms, r["richiesta"] + " " + r["soluzione"] + " " + r["note"])
            ]
        return rows

    def stats(self):
        today = date.today()
        c = self.conn
        tot = c.execute("SELECT COUNT(*) FROM interventi").fetchone()[0]
        oggi = c.execute(
            "SELECT COUNT(*) FROM interventi WHERE substr(data_ora,1,10)=?", (today.isoformat(),)
        ).fetchone()[0]
        mese = c.execute(
            "SELECT COUNT(*) FROM interventi WHERE substr(data_ora,1,7)=?",
            (today.strftime("%Y-%m"),),
        ).fetchone()[0]
        return oggi, mese, tot

    # --- edifici ---
    def edifici(self):
        rows = self.conn.execute("SELECT * FROM edifici ORDER BY nome COLLATE NOCASE")
        return [dict(r) for r in rows]

    def add_edificio(self, nome, sede, piano_min, piano_max):
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO edifici(nome, sede, piano_min, piano_max) VALUES (?,?,?,?)",
                (nome, sede, piano_min, piano_max))
        return cur.lastrowid

    def update_edificio(self, eid, nome, sede, piano_min, piano_max):
        with self.conn:
            self.conn.execute(
                "UPDATE edifici SET nome=?, sede=?, piano_min=?, piano_max=? WHERE id=?",
                (nome, sede, piano_min, piano_max, eid))

    def edificio_stats(self, eid):
        n_s = self.conn.execute(
            "SELECT COUNT(*) FROM stampanti WHERE edificio_id=?", (eid,)).fetchone()[0]
        n_t = self.conn.execute(
            "SELECT COUNT(*) FROM toner t JOIN stampanti s ON s.id=t.stampante_id "
            "WHERE s.edificio_id=?", (eid,)).fetchone()[0]
        return n_s, n_t

    def delete_edificio(self, eid):
        for r in self.stanze_of(eid):
            self.delete_stanza(r["id"])
        for pr in self.conn.execute("SELECT piano FROM piante WHERE edificio_id=?", (eid,)).fetchall():
            self.delete_pianta(eid, pr["piano"])
        with self.conn:
            self.conn.execute("DELETE FROM stampanti WHERE edificio_id=?", (eid,))
            self.conn.execute("DELETE FROM edifici WHERE id=?", (eid,))

    # --- stampanti ---
    def printers_of(self, edificio_id):
        rows = self.conn.execute(
            "SELECT s.id, s.edificio_id, s.piano, s.nome, s.modello, s.ip, s.note, s.seriale, "
            "s.fuori_servizio, s.stanza_uid, "
            + ROOM_OF_PRINTER +
            "(SELECT t.colore || '|' || t.data_ora FROM toner t WHERE t.stampante_id = s.id "
            " ORDER BY t.data_ora DESC, t.id DESC LIMIT 1) AS ultimo, "
            "(SELECT COUNT(*) FROM manutenzioni m WHERE m.stampante_id = s.id "
            " AND m.stato = 'Richiesto') AS man_richieste "
            "FROM stampanti s WHERE s.edificio_id = ? ORDER BY s.piano DESC, s.id",
            (edificio_id,))
        out = [dict(r) for r in rows]
        # ultimo cambio di ogni colore e scorta bassa del modello (per i bollini e i "giorni fa")
        ultimi = {}
        for r in self.conn.execute(
                "SELECT t.stampante_id AS sid, t.colore AS colore, MAX(t.data_ora) AS d FROM toner t "
                "JOIN stampanti s ON s.id = t.stampante_id WHERE s.edificio_id = ? "
                "GROUP BY t.stampante_id, t.colore", (edificio_id,)):
            ultimi.setdefault(r["sid"], {})[r["colore"]] = r["d"]
        scorta = [r for r in self.scorta_all() if scorta_status(r["quantita"], r["soglia"]) != "OK"]
        stime = self.toner_estimates()
        for p in out:
            p["ultimi"] = ultimi.get(p["id"], {})
            p["stima"] = stime.get(p["id"], {})
            mk = model_key(p["modello"])
            p["scorta_bassa"] = [(r["colore"], r["quantita"]) for r in scorta
                                 if mk and model_key(r["modello"]) == mk]
        return out

    def toner_estimates(self):
        """Stima del toner che resta per ogni stampante e colore: {id stampante: {colore: stima}}."""
        printers = [(r["id"], r["modello"]) for r in self.conn.execute(
            "SELECT id, modello FROM stampanti")]
        rows = [(r["stampante_id"], r["colore"], r["data_ora"]) for r in self.conn.execute(
            "SELECT stampante_id, colore, data_ora FROM toner")]
        return compute_toner_estimates(printers, rows)

    def all_printers(self):
        rows = self.conn.execute(
            "SELECT s.id, s.piano, s.nome, s.modello, e.nome AS edificio FROM stampanti s "
            "JOIN edifici e ON e.id = s.edificio_id "
            "ORDER BY e.nome COLLATE NOCASE, s.piano DESC, s.id")
        return [dict(r) for r in rows]

    def get_stampante(self, sid):
        row = self.conn.execute(
            "SELECT s.id, s.edificio_id, s.piano, s.nome, s.modello, s.ip, s.note, s.seriale, "
            "s.fuori_servizio, s.stanza_uid, "
            + ROOM_OF_PRINTER +
            "e.nome AS edificio, e.sede FROM stampanti s JOIN edifici e ON e.id = s.edificio_id "
            "WHERE s.id = ?", (sid,)).fetchone()
        return dict(row) if row else None

    def models(self):
        rows = self.conn.execute(
            "SELECT DISTINCT modello FROM stampanti WHERE modello <> '' ORDER BY modello")
        return [r[0] for r in rows]

    def count_model(self, modello):
        """Quante stampanti e quante voci di scorta toner portano questo modello."""
        key = model_key(modello)
        n_p = sum(1 for r in self.conn.execute("SELECT modello FROM stampanti")
                  if model_key(r["modello"]) == key)
        n_s = sum(1 for r in self.conn.execute("SELECT modello FROM scorta")
                  if model_key(r["modello"]) == key)
        return n_p, n_s

    def rename_model(self, old, new):
        """Cambia il nome del modello su tutte le stampanti e nella scorta toner.
        Se il nome nuovo c'è già nella scorta, le quantità si sommano.
        Restituisce (stampanti cambiate, voci di scorta cambiate)."""
        new = " ".join((new or "").split())
        key = model_key(old)
        if not new or not key:
            return 0, 0
        n_p = n_s = 0
        with self.conn:
            for r in self.conn.execute("SELECT id, modello FROM stampanti").fetchall():
                if model_key(r["modello"]) == key:
                    self.conn.execute("UPDATE stampanti SET modello = ? WHERE id = ?",
                                      (new, r["id"]))
                    n_p += 1
            rows = [r for r in self.conn.execute(
                "SELECT modello, colore, quantita, soglia FROM scorta").fetchall()
                if model_key(r["modello"]) == key]
            for r in rows:
                self.conn.execute("DELETE FROM scorta WHERE modello = ? AND colore = ?",
                                  (r["modello"], r["colore"]))
                t = self.conn.execute(
                    "SELECT quantita, soglia FROM scorta WHERE modello = ? AND colore = ?",
                    (new, r["colore"])).fetchone()
                if t:
                    self.conn.execute(
                        "UPDATE scorta SET quantita = ?, soglia = ?, modificato = ? "
                        "WHERE modello = ? AND colore = ?",
                        (t["quantita"] + r["quantita"], max(t["soglia"], r["soglia"]),
                         now_utc_iso(), new, r["colore"]))
                else:
                    self.conn.execute(
                        "INSERT INTO scorta(modello, colore, quantita, soglia, modificato) "
                        "VALUES (?,?,?,?,?)",
                        (new, r["colore"], r["quantita"], r["soglia"], now_utc_iso()))
                n_s += 1
        return n_p, n_s

    def add_stampante(self, edificio_id, piano, nome, modello, ip, note, seriale="",
                      fuori_servizio=False, stanza_uid=""):
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO stampanti(edificio_id, piano, nome, modello, ip, note, seriale, "
                "fuori_servizio, stanza_uid) VALUES (?,?,?,?,?,?,?,?,?)",
                (edificio_id, piano, nome, modello, ip, note, clean_serial(seriale),
                 1 if fuori_servizio else 0, stanza_uid or ""))
        return cur.lastrowid

    def update_stampante(self, sid, piano, nome, modello, ip, note, seriale=None,
                         fuori_servizio=None, stanza_uid=None):
        """Numero di serie, stato "fuori servizio" e stanza, se non li passi, restano quelli che c'erano."""
        with self.conn:
            self.conn.execute(
                "UPDATE stampanti SET piano=?, nome=?, modello=?, ip=?, note=? WHERE id=?",
                (piano, nome, modello, ip, note, sid))
            if seriale is not None:
                self.conn.execute("UPDATE stampanti SET seriale=? WHERE id=?",
                                  (clean_serial(seriale), sid))
            if fuori_servizio is not None:
                self.conn.execute("UPDATE stampanti SET fuori_servizio=? WHERE id=?",
                                  (1 if fuori_servizio else 0, sid))
            if stanza_uid is not None:
                self.conn.execute("UPDATE stampanti SET stanza_uid=? WHERE id=?",
                                  (stanza_uid or "", sid))

    def printers_in_room(self, stanza_uid):
        """Le stampanti collegate a una stanza (per la scheda della stanza)."""
        if not stanza_uid:
            return []
        rows = self.conn.execute(
            "SELECT s.id, s.nome, s.modello, s.seriale, s.fuori_servizio FROM stampanti s "
            "WHERE s.stanza_uid = ? ORDER BY s.id", (stanza_uid,))
        return [dict(r) for r in rows]

    def printer_with_serial(self, seriale, escludi_id=None):
        """Un'altra stampante con lo stesso numero di serie (per avvisare di un doppione)."""
        s = clean_serial(seriale)
        if not s:
            return None
        row = self.conn.execute(
            "SELECT s.id, s.nome, s.modello, s.piano, e.nome AS edificio FROM stampanti s "
            "JOIN edifici e ON e.id = s.edificio_id WHERE s.seriale = ? COLLATE NOCASE "
            "AND (? IS NULL OR s.id <> ?) ORDER BY s.id LIMIT 1", (s, escludi_id, escludi_id)).fetchone()
        return dict(row) if row else None

    def printer_merge_proposals(self, edificio_id):
        """Stampanti SENZA numero di serie che potrebbero essere la stessa di una CON numero di serie
        (stesso edificio, stesso piano, stesso modello): tipico quando il file con i seriali ha
        aggiunto quelle nuove accanto alle vecchie. Per ognuna: le candidate e, se è chiaro, quella
        suggerita (stessa stanza, stesso nome, l'unica del piano o il nome che somiglia di più)."""
        rows = self.printers_of(edificio_id)
        news = [r for r in rows if (r["seriale"] or "").strip()]
        olds = [r for r in rows if not (r["seriale"] or "").strip()]

        def compat(o, n):
            return n["piano"] == o["piano"] and (not model_key(o["modello"])
                                                  or model_key(o["modello"]) == model_key(n["modello"]))

        def words(s):
            skip = {"stampante", "bizhub", "konica", "minolta", "di", "del", "della", "la", "il",
                    "lo", "a", "e", "in"}
            return {w for w in re.findall(r"[0-9a-zà-ÿ]+", (s or "").casefold())
                    if (len(w) >= 2 or w.isdigit()) and w not in skip}

        out = []
        for o in olds:
            plaus = [n for n in news if compat(o, n)]
            if not plaus:
                continue
            same = [x for x in olds if any(compat(x, n) for n in plaus) and x["piano"] == o["piano"]
                    and (not model_key(x["modello"]) or model_key(x["modello"]) == model_key(o["modello"]))]
            pick = None
            by_room = [n for n in plaus if o.get("stanza_uid") and n.get("stanza_uid") == o["stanza_uid"]]
            by_name = [n for n in plaus if o["nome"].strip() and
                       n["nome"].strip().casefold() == o["nome"].strip().casefold()]
            if len(by_room) == 1:
                pick = by_room[0]
            elif len(by_name) == 1:
                pick = by_name[0]
            elif len(same) == 1 and len(plaus) == 1:
                pick = plaus[0]
            else:
                sc = sorted(((len(words(o["nome"]) & words(n["nome"])), n["id"], n) for n in plaus),
                            key=lambda t: (-t[0], t[1]))
                if sc and sc[0][0] > 0 and (len(sc) == 1 or sc[0][0] > sc[1][0]):
                    pick = sc[0][2]
            n_t = self.conn.execute("SELECT COUNT(*) FROM toner WHERE stampante_id = ?",
                                    (o["id"],)).fetchone()[0]
            n_m = self.conn.execute("SELECT COUNT(*) FROM manutenzioni WHERE stampante_id = ?",
                                    (o["id"],)).fetchone()[0]
            rest = [n for n in news if n not in plaus]
            out.append({"old": o, "plausibili": plaus, "altre": rest,
                        "suggerita": pick["id"] if pick else None, "n_toner": n_t, "n_manut": n_m})
        return out

    def merge_printers(self, old_id, keep_id):
        """Unisce due stampanti che sono la stessa: tutto lo storico di «old» (cambi toner,
        interventi del tecnico, bolle) passa a «keep», poi «old» si toglie. «keep» tiene nome,
        seriale e stanza; IP, note, modello e stanza si prendono dalla vecchia solo se mancano
        alla nuova; se una delle due era «fuori servizio» resta tale. Il telefono lo viene a sapere
        dal file di scambio (stampanti_unite). Restituisce {"toner": n, "manutenzioni": n}."""
        if old_id == keep_id:
            raise ValueError("Sono la stessa stampante.")
        self.ensure_sync_ids()
        old, keep = self.get_stampante(old_id), self.get_stampante(keep_id)
        if not old or not keep:
            raise ValueError("Stampante non trovata.")
        o_uid = self.conn.execute("SELECT uid FROM stampanti WHERE id = ?", (old_id,)).fetchone()["uid"]
        k_uid = self.conn.execute("SELECT uid FROM stampanti WHERE id = ?", (keep_id,)).fetchone()["uid"]
        with self.conn:
            n_t = self.conn.execute("UPDATE toner SET stampante_id = ? WHERE stampante_id = ?",
                                    (keep_id, old_id)).rowcount
            n_m = self.conn.execute("UPDATE manutenzioni SET stampante_id = ? WHERE stampante_id = ?",
                                    (keep_id, old_id)).rowcount
            self.conn.execute("UPDATE bolle SET stampante_id = ? WHERE stampante_id = ?",
                              (keep_id, old_id))
            self.conn.execute("UPDATE bolle SET stampante_montata_id = ? WHERE stampante_montata_id = ?",
                              (keep_id, old_id))
            ip = keep["ip"] or old["ip"]
            note = keep["note"]
            if old["note"] and old["note"] not in (note or ""):
                note = ("%s | %s" % (note, old["note"])) if note else old["note"]
            self.conn.execute(
                "UPDATE stampanti SET ip = ?, note = ?, modello = ?, fuori_servizio = ?, "
                "stanza_uid = ? WHERE id = ?",
                (ip, note, keep["modello"] or old["modello"],
                 1 if (keep["fuori_servizio"] or old["fuori_servizio"]) else 0,
                 keep.get("stanza_uid") or old.get("stanza_uid") or "", keep_id))
            self.conn.execute("UPDATE stampanti_unite SET a_uid = ? WHERE a_uid = ?", (k_uid, o_uid))
            self.conn.execute("INSERT OR REPLACE INTO stampanti_unite(da_uid, a_uid, quando) "
                              "VALUES (?,?,?)", (o_uid, k_uid, now_utc_iso()))
            self.conn.execute("DELETE FROM stampanti WHERE id = ?", (old_id,))
        return {"toner": n_t, "manutenzioni": n_m}

    def printer_merges(self):
        """Le unioni fatte (uid vecchio -> uid nuovo), da dire al telefono."""
        return [{"da": r["da_uid"], "a": r["a_uid"]} for r in self.conn.execute(
            "SELECT da_uid, a_uid FROM stampanti_unite ORDER BY quando, da_uid")]

    def delete_stampante(self, sid):
        with self.conn:
            self.conn.execute("DELETE FROM stampanti WHERE id=?", (sid,))

    def _modello_of(self, stampante_id):
        row = self.conn.execute("SELECT modello FROM stampanti WHERE id=?",
                                (stampante_id,)).fetchone()
        return row["modello"] if row else ""

    # --- cambi toner ---
    def add_toner(self, stampante_id, colori, data_ora, note=""):
        """Registra i cambi toner e scala la scorta (se la tieni). Restituisce l'elenco dei toner
        che, dopo questo cambio, sono in esaurimento o finiti."""
        low = []
        with self.conn:
            modello = self._modello_of(stampante_id)
            for colore in colori:
                self.conn.execute(
                    "INSERT INTO toner(stampante_id, colore, data_ora, note) VALUES (?,?,?,?)",
                    (stampante_id, colore, data_ora, note))
                row = self._scorta_scale(modello, colore, -1)
                if row and scorta_status(row["quantita"], row["soglia"]) != "OK":
                    low.append(row)
        return low

    def delete_toner(self, tid):
        """Elimina un cambio toner; se la scorta è tracciata, il pezzo torna in magazzino."""
        with self.conn:
            r = self.conn.execute(
                "SELECT t.colore, s.modello FROM toner t JOIN stampanti s ON s.id = t.stampante_id "
                "WHERE t.id = ?", (tid,)).fetchone()
            self.conn.execute("DELETE FROM toner WHERE id=?", (tid,))
            if r:
                self._scorta_scale(r["modello"], r["colore"], +1)

    # --- scorta toner (magazzino) ---
    def scorta_all(self):
        rows = self.conn.execute(
            "SELECT modello, colore, quantita, soglia FROM scorta "
            "ORDER BY modello COLLATE NOCASE, colore")
        return [dict(r) for r in rows]

    def scorta_get(self, modello, colore):
        row = self.conn.execute(
            "SELECT modello, colore, quantita, soglia FROM scorta WHERE modello = ? AND colore = ?",
            (modello, colore)).fetchone()
        return dict(row) if row else None

    def set_scorta(self, modello, colore, quantita, soglia):
        modello = " ".join((modello or "").split())
        quantita, soglia = max(0, int(quantita)), max(0, int(soglia))
        with self.conn:
            if self.scorta_get(modello, colore):
                self.conn.execute(
                    "UPDATE scorta SET quantita=?, soglia=?, modificato=? "
                    "WHERE modello = ? AND colore = ?",
                    (quantita, soglia, now_utc_iso(), modello, colore))
            else:
                self.conn.execute(
                    "INSERT INTO scorta(modello, colore, quantita, soglia, modificato) "
                    "VALUES (?,?,?,?,?)", (modello, colore, quantita, soglia, now_utc_iso()))

    def add_scorta(self, modello, colore, n):
        """Toner ricevuti: li aggiunge al magazzino (se non era tracciato, comincia da qui)."""
        modello = " ".join((modello or "").split())
        row = self.scorta_get(modello, colore)
        if row:
            self.set_scorta(modello, colore, row["quantita"] + n, row["soglia"])
        else:
            self.set_scorta(modello, colore, n, SOGLIA_SCORTA)

    def delete_scorta(self, modello, colore):
        with self.conn:
            self.conn.execute("DELETE FROM scorta WHERE modello = ? AND colore = ?",
                              (modello, colore))

    def _scorta_scale(self, modello, colore, delta):
        """Cambia la quantità (mai sotto zero) solo se questo toner è tracciato."""
        row = self.scorta_get(modello, colore)
        if not row:
            return None
        row["quantita"] = max(0, row["quantita"] + delta)
        self.conn.execute(
            "UPDATE scorta SET quantita=?, modificato=? WHERE modello = ? AND colore = ?",
            (row["quantita"], now_utc_iso(), modello, colore))
        return row

    def scorta_bassa(self):
        return [r for r in self.scorta_all()
                if scorta_status(r["quantita"], r["soglia"]) != "OK"]

    # --- bolle dei toner (arrivi e montaggi registrati con il telefono) ---
    def find_bolla_by_uid(self, uid):
        r = self.conn.execute("SELECT * FROM bolle WHERE uid = ?", (uid,)).fetchone()
        return dict(r) if r else None

    def bolle_all(self, stato=None):
        """Le bolle con i nomi di edificio e stampanti; stato: 'da_montare', 'montato' o None (tutte)."""
        sql = ("SELECT b.*, COALESCE(e.nome, '?') AS edificio, COALESCE(s.nome, '') AS stampante_nome, "
               "COALESCE(m.nome, '') AS montata_nome, COALESCE(m.modello, '') AS montata_modello "
               "FROM bolle b LEFT JOIN edifici e ON e.id = b.edificio_id "
               "LEFT JOIN stampanti s ON s.id = b.stampante_id "
               "LEFT JOIN stampanti m ON m.id = b.stampante_montata_id")
        params = ()
        if stato:
            sql += " WHERE b.stato = ?"
            params = (stato,)
        sql += " ORDER BY b.data_arrivo DESC, b.id DESC"
        return [dict(r) for r in self.conn.execute(sql, params)]

    def bolle_inventario(self):
        """Toner arrivati e ancora da montare, divisi per edificio e piano: una riga per piano con i
        colori contati."""
        rows = {}
        for b in self.bolle_all("da_montare"):
            k = (b["edificio"], b["piano"])
            r = rows.setdefault(k, {"edificio": b["edificio"], "piano": b["piano"],
                                    "colori": {c: 0 for c in TONER_NOMI}, "totale": 0})
            if b["colore"] in r["colori"]:
                r["colori"][b["colore"]] += 1
            r["totale"] += 1
        return [rows[k] for k in sorted(rows, key=lambda k: (k[0].casefold(), -k[1]))]

    def save_phone_bolla(self, existing, uid, f):
        """Salva una bolla arrivata dal telefono (nuova o aggiornata). La scorta del modello sale di uno
        quando il modello è noto e la bolla non è ancora stata contata (se quel toner non è tracciato,
        la bolla conta come già gestita)."""
        with self.conn:
            cols = ("codice", "numero", "toner_codice", "colore", "edificio_id", "piano", "stampante_id",
                    "modello", "seriale_bolla", "data_arrivo", "stato", "data_montaggio",
                    "stampante_montata_id", "seriale_letto", "toner_uid", "note", "foto", "modificato")
            vals = [f[c] for c in cols]
            if existing:
                self.conn.execute("UPDATE bolle SET " + ", ".join("%s=?" % c for c in cols)
                                  + " WHERE id = ?", vals + [existing["id"]])
                bid, contata = existing["id"], existing["scorta_ok"]
            else:
                cur = self.conn.execute(
                    "INSERT INTO bolle(uid, " + ", ".join(cols) + ") VALUES (?" + ",?" * len(cols) + ")",
                    [uid] + vals)
                bid, contata = cur.lastrowid, 0
            if not contata and f["modello"]:
                self._scorta_scale(f["modello"], f["colore"], +1)
                self.conn.execute("UPDATE bolle SET scorta_ok = 1 WHERE id = ?", (bid,))
        return bid

    # --- piani e stanze ---
    ROOM_COLS = ("s.id, s.uid, s.edificio_id, s.piano, s.codice, s.scrivanie, s.note, s.come, "
                 "s.px, s.py, s.modificato, s.eliminata")

    def stanze_of(self, edificio_id, piano=None):
        """Le stanze dell'edificio (quelle eliminate non compaiono)."""
        sql = ("SELECT " + self.ROOM_COLS + ", "
               "(SELECT COUNT(*) FROM postazioni p WHERE p.stanza_id = s.id) AS n_persone, "
               "(SELECT COUNT(*) FROM stanze_foto f WHERE f.stanza_id = s.id AND f.eliminata = 0) "
               "AS n_foto FROM stanze s WHERE s.edificio_id = ? AND s.eliminata = 0")
        params = [edificio_id]
        if piano is not None:
            sql += " AND s.piano = ?"
            params.append(piano)
        rows = [dict(r) for r in self.conn.execute(sql, params)]
        rows.sort(key=lambda r: (-r["piano"], natural_key(r["codice"])))
        return rows

    def get_stanza(self, sid, anche_eliminata=False):
        row = self.conn.execute(
            "SELECT " + self.ROOM_COLS + ", e.nome AS edificio, e.uid AS edificio_uid, "
            "(SELECT COUNT(*) FROM postazioni p WHERE p.stanza_id = s.id) AS n_persone, "
            "(SELECT COUNT(*) FROM stanze_foto f WHERE f.stanza_id = s.id AND f.eliminata = 0) "
            "AS n_foto FROM stanze s JOIN edifici e ON e.id = s.edificio_id WHERE s.id = ?",
            (sid,)).fetchone()
        if not row or (row["eliminata"] and not anche_eliminata):
            return None
        return dict(row)

    def find_stanza_by_uid(self, uid, anche_eliminata=False):
        row = self.conn.execute("SELECT id FROM stanze WHERE uid = ?", (uid,)).fetchone()
        return self.get_stanza(row["id"], anche_eliminata) if row else None

    def find_stanza_by_code(self, edificio_id, piano, codice):
        row = self.conn.execute(
            "SELECT id FROM stanze WHERE edificio_id = ? AND piano = ? AND codice = ? COLLATE NOCASE "
            "AND eliminata = 0", (edificio_id, piano, clean_room_code(codice))).fetchone()
        return self.get_stanza(row["id"]) if row else None

    def add_stanza(self, edificio_id, piano, codice, scrivanie=0, note="", come="", px=None,
                   py=None, uid=None, modificato=None):
        codice = clean_room_code(codice)
        if not codice:
            raise ValueError("Scrivi il codice della stanza.")
        if self.find_stanza_by_code(edificio_id, piano, codice):
            raise ValueError("La stanza %s esiste già su questo piano." % codice)
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO stanze(uid, edificio_id, piano, codice, scrivanie, note, come, px, py, "
                "modificato) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (uid or "ro-" + uuid.uuid4().hex, edificio_id, piano, codice,
                 max(0, int(scrivanie or 0)), note or "", come or "", px, py,
                 modificato or now_utc_iso()))
        return cur.lastrowid

    def add_stanze_bulk(self, edificio_id, piano, testo):
        """Incolla più stanze (una per riga, anche con le postazioni: "T12; 3"). Una stanza già
        presente non si duplica; se la riga porta un numero di postazioni diverso lo aggiorna.
        Restituisce {"nuove", "esistenti", "aggiornate", "scartate", "scrivanie"}."""
        rooms, bad = parse_room_lines(testo)
        res = {"nuove": [], "esistenti": [], "aggiornate": [], "scartate": bad, "scrivanie": {}}
        for code, n in rooms:
            cur = self.find_stanza_by_code(edificio_id, piano, code)
            if cur:
                if n is not None and n != (cur["scrivanie"] or 0):
                    self.update_stanza(cur["id"], cur["codice"], n, cur["note"], cur["come"])
                    res["aggiornate"].append(code)
                else:
                    res["esistenti"].append(code)
            else:
                self.add_stanza(edificio_id, piano, code, n or 0)
                res["nuove"].append(code)
            if n is not None:
                res["scrivanie"][code] = n
        return res

    def update_stanza(self, sid, codice, scrivanie, note, come, modificato=None):
        cur = self.get_stanza(sid)
        codice = clean_room_code(codice)
        if not cur or not codice:
            raise ValueError("Scrivi il codice della stanza.")
        other = self.find_stanza_by_code(cur["edificio_id"], cur["piano"], codice)
        if other and other["id"] != sid:
            raise ValueError("La stanza %s esiste già su questo piano." % codice)
        with self.conn:
            self.conn.execute(
                "UPDATE stanze SET codice=?, scrivanie=?, note=?, come=?, modificato=? WHERE id=?",
                (codice, max(0, int(scrivanie or 0)), note or "", come or "",
                 modificato or now_utc_iso(), sid))

    def set_stanza_pos(self, sid, px, py, modificato=None):
        """Posizione della stanza sulla pianta (frazioni 0..1); None, None la toglie."""
        if px is not None:
            px, py = min(1.0, max(0.0, float(px))), min(1.0, max(0.0, float(py)))
        with self.conn:
            self.conn.execute("UPDATE stanze SET px=?, py=?, modificato=? WHERE id=?",
                              (px, py, modificato or now_utc_iso(), sid))

    def delete_stanza(self, sid, modificato=None):
        """Toglie la stanza (resta un segno "eliminata" così anche il telefono la toglie), libera
        le sue postazioni (le persone restano negli utenti) e cancella le foto."""
        now = modificato or now_utc_iso()
        for f in self.foto_stanza(sid):
            self._remove_file(os.path.join(stanze_foto_root(), *f["file"].split("/")))
        with self.conn:
            self.conn.execute("UPDATE stanze_foto SET eliminata = 1, modificato = ? "
                              "WHERE stanza_id = ?", (now, sid))
            self.conn.execute(
                "UPDATE postazioni SET stanza_id = NULL, numero = 0, modificato = ? "
                "WHERE stanza_id = ?", (now, sid))
            self.conn.execute("UPDATE stanze SET eliminata = 1, modificato = ? WHERE id = ?",
                              (now, sid))

    @staticmethod
    def _remove_file(path):
        try:
            os.remove(path)
        except OSError:
            pass

    # persone e postazioni: le postazioni le gestisce il programma (1, 2, 3...), non l'utente
    def persone_of(self, stanza_id):
        rows = self.conn.execute(
            "SELECT p.numero, p.modificato AS posto_modificato, u.id AS utente_id, u.nome, "
            "u.matricola, u.telefono, u.seriale_pc, u.imei, u.sede, u.ufficio, u.dotazione, "
            "u.ordine, u.email FROM postazioni p "
            "JOIN utenti u ON u.id = p.utente_id WHERE p.stanza_id = ? "
            "ORDER BY p.numero, " + SORT_UTENTE, (stanza_id,))
        return [self._user_dict(r) for r in rows]

    def posto_of_user(self, utente_id):
        row = self.conn.execute(
            "SELECT p.numero, s.id AS stanza_id, s.uid AS stanza_uid, s.codice, s.piano, "
            "e.id AS edificio_id, e.nome AS edificio FROM postazioni p "
            "JOIN stanze s ON s.id = p.stanza_id JOIN edifici e ON e.id = s.edificio_id "
            "WHERE p.utente_id = ?", (utente_id,)).fetchone()
        return dict(row) if row else None

    def persone_piano(self, edificio_id, piano):
        rows = self.conn.execute(
            "SELECT s.codice, s.id AS stanza_id, p.numero, u.id AS utente_id, u.nome, "
            "u.telefono, u.seriale_pc, u.imei, u.ordine FROM postazioni p "
            "JOIN stanze s ON s.id = p.stanza_id JOIN utenti u ON u.id = p.utente_id "
            "WHERE s.edificio_id = ? AND s.piano = ?", (edificio_id, piano))
        out = [dict(r) for r in rows]
        out.sort(key=lambda r: (natural_key(r["codice"]), r["numero"],
                                (r["ordine"] or r["nome"]).casefold()))
        return out

    def _renumber(self, stanza_id, now, bump=True):
        """Postazioni di nuovo consecutive: 1, 2, 3... (senza buchi)."""
        rows = self.conn.execute(
            "SELECT p.utente_id, p.numero FROM postazioni p JOIN utenti u ON u.id = p.utente_id "
            "WHERE p.stanza_id = ? ORDER BY p.numero, u.nome COLLATE NOCASE", (stanza_id,)).fetchall()
        for i, r in enumerate(rows, 1):
            if r["numero"] != i:
                if bump:
                    self.conn.execute("UPDATE postazioni SET numero = ?, modificato = ? "
                                      "WHERE utente_id = ?", (i, now, r["utente_id"]))
                else:
                    self.conn.execute("UPDATE postazioni SET numero = ? WHERE utente_id = ?",
                                      (i, r["utente_id"]))

    def assegna_persona(self, utente_id, stanza_id, modificato=None):
        """Mette la persona nella stanza, alla prima postazione libera. Se sedeva in un'altra
        stanza la toglie da lì (e rinumera). Restituisce {"numero", "gia", "da", "sopra"}:
        "da" = la stanza di prima (se è stata spostata), "sopra" = ora ci sono più persone che
        scrivanie (solo se le scrivanie totali sono indicate)."""
        now = modificato or now_utc_iso()
        room = self.get_stanza(stanza_id)
        if not room:
            raise ValueError("Questa stanza non c'è più.")
        old = self.posto_of_user(utente_id)
        if old and old["stanza_id"] == stanza_id:
            return {"numero": old["numero"], "gia": True, "da": None, "sopra": False}
        presa = self.togli_presa(utente_id, now) if old else ""
        with self.conn:
            if old:
                self.conn.execute("UPDATE postazioni SET stanza_id = NULL, numero = 0, modificato = ? "
                                  "WHERE utente_id = ?", (now, utente_id))
                self._renumber(old["stanza_id"], now)
            n = self.conn.execute("SELECT COUNT(*) FROM postazioni WHERE stanza_id = ?",
                                  (stanza_id,)).fetchone()[0] + 1
            self.conn.execute(
                "INSERT INTO postazioni(utente_id, stanza_id, numero, modificato) VALUES (?,?,?,?) "
                "ON CONFLICT(utente_id) DO UPDATE SET stanza_id = excluded.stanza_id, "
                "numero = excluded.numero, modificato = excluded.modificato",
                (utente_id, stanza_id, n, now))
        tot = room["scrivanie"] or 0
        return {"numero": n, "gia": False, "da": old, "sopra": bool(tot and n > tot),
                "presa_tolta": presa}

    def rimuovi_persona(self, utente_id, modificato=None):
        """Libera la postazione (la riga resta con "nessuna stanza" così il telefono lo sa)."""
        now = modificato or now_utc_iso()
        old = self.posto_of_user(utente_id)
        if not old:
            return False
        self.togli_presa(utente_id, now)
        with self.conn:
            self.conn.execute("UPDATE postazioni SET stanza_id = NULL, numero = 0, modificato = ? "
                              "WHERE utente_id = ?", (now, utente_id))
            self._renumber(old["stanza_id"], now)
        return True

    def togli_presa(self, utente_id, modificato=None):
        """La presa di rete è della scrivania, non della persona: quando la persona cambia scrivania
        il codice della presa di prima si toglie dalla dotazione. Restituisce il codice tolto ('' se
        non c'era)."""
        row = self.conn.execute("SELECT dotazione FROM utenti WHERE id = ?", (utente_id,)).fetchone()
        if not row:
            return ""
        dot = clean_dotazione(row["dotazione"])
        old = next((r for r in dot if _row_kind(r) == "presa"), None)
        if not old:
            return ""
        self.set_user_extra(utente_id, dotazione=[r for r in dot if _row_kind(r) != "presa"],
                            modificato=modificato)
        return old["seriale"] or old["modello"]

    def apply_stanza(self, sid, codice, scrivanie, note, come, px, py, modificato, eliminata=False):
        """Aggiorna una stanza con i dati arrivati dall'altro dispositivo (senza toccare "adesso").
        eliminata=True: l'altro dispositivo l'ha tolta (si liberano le postazioni e le foto)."""
        cur = self.get_stanza(sid, anche_eliminata=True)
        if eliminata:
            if not cur["eliminata"]:
                self.delete_stanza(sid, modificato)
            return
        codice = clean_room_code(codice) or cur["codice"]
        other = self.find_stanza_by_code(cur["edificio_id"], cur["piano"], codice)
        if other and other["id"] != sid:
            codice = cur["codice"]
        with self.conn:
            self.conn.execute(
                "UPDATE stanze SET codice=?, scrivanie=?, note=?, come=?, px=?, py=?, modificato=?, "
                "eliminata=0 WHERE id=?", (codice, max(0, int(scrivanie or 0)), note or "",
                                           come or "", px, py, modificato, sid))

    def apply_posto(self, utente_id, stanza_id, numero, modificato):
        """Postazione arrivata dall'altro dispositivo: vince la più recente. stanza_id None =
        la persona non è in nessuna stanza. Restituisce le stanze toccate (da rinumerare)."""
        old = self.conn.execute("SELECT stanza_id, modificato FROM postazioni WHERE utente_id = ?",
                                (utente_id,)).fetchone()
        if old and (old["modificato"] or "") >= modificato:
            return None
        with self.conn:
            self.conn.execute(
                "INSERT INTO postazioni(utente_id, stanza_id, numero, modificato) VALUES (?,?,?,?) "
                "ON CONFLICT(utente_id) DO UPDATE SET stanza_id = excluded.stanza_id, "
                "numero = excluded.numero, modificato = excluded.modificato",
                (utente_id, stanza_id, numero if stanza_id else 0, modificato))
        return {x for x in (old["stanza_id"] if old else None, stanza_id) if x}

    def renumber_rooms(self, ids):
        """Dopo uno scambio: postazioni di nuovo consecutive (senza cambiare le date)."""
        with self.conn:
            for sid in ids:
                self._renumber(sid, now_utc_iso(), bump=False)

    # foto e panorami della stanza
    def foto_stanza(self, stanza_id):
        rows = self.conn.execute(
            "SELECT id, uid, nome, tipo, file, modificato FROM stanze_foto WHERE stanza_id = ? "
            "AND eliminata = 0 ORDER BY id", (stanza_id,))
        return [dict(r) for r in rows]

    def add_foto_stanza(self, stanza_id, nome, tipo, file, uid=None, modificato=None):
        tipo = tipo if tipo in ("foto", "panorama") else "foto"
        now = modificato or now_utc_iso()
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO stanze_foto(uid, stanza_id, nome, tipo, file, modificato) "
                "VALUES (?,?,?,?,?,?)", (uid or "sf-" + uuid.uuid4().hex, stanza_id, nome, tipo,
                                         file, now))
            self.conn.execute("UPDATE stanze SET modificato = ? WHERE id = ?", (now, stanza_id))
        return cur.lastrowid

    def set_tipo_foto_stanza(self, fid, tipo):
        if tipo in ("foto", "panorama"):
            with self.conn:
                self.conn.execute("UPDATE stanze_foto SET tipo = ?, modificato = ? WHERE id = ?",
                                  (tipo, now_utc_iso(), fid))

    def delete_foto_stanza(self, fid, modificato=None):
        row = self.conn.execute("SELECT file, stanza_id FROM stanze_foto WHERE id = ?",
                                (fid,)).fetchone()
        if not row:
            return
        now = modificato or now_utc_iso()
        self._remove_file(os.path.join(stanze_foto_root(), *row["file"].split("/")))
        with self.conn:
            self.conn.execute("UPDATE stanze_foto SET eliminata = 1, modificato = ? WHERE id = ?",
                              (now, fid))
            self.conn.execute("UPDATE stanze SET modificato = ? WHERE id = ?",
                              (now, row["stanza_id"]))

    # pianta del piano (una foto per edificio e piano)
    def pianta_of(self, edificio_id, piano):
        row = self.conn.execute(
            "SELECT id, uid, edificio_id, piano, file, modificato FROM piante "
            "WHERE edificio_id = ? AND piano = ?", (edificio_id, piano)).fetchone()
        return dict(row) if row else None

    def set_pianta(self, edificio_id, piano, file, modificato=None):
        e = self.conn.execute("SELECT uid FROM edifici WHERE id = ?", (edificio_id,)).fetchone()
        if not e:
            raise ValueError("Questo edificio non c'è più.")
        if not e["uid"]:
            self.ensure_sync_ids()
            e = self.conn.execute("SELECT uid FROM edifici WHERE id = ?", (edificio_id,)).fetchone()
        old = self.pianta_of(edificio_id, piano)
        with self.conn:
            self.conn.execute(
                "INSERT INTO piante(uid, edificio_id, piano, file, modificato) VALUES (?,?,?,?,?) "
                "ON CONFLICT(edificio_id, piano) DO UPDATE SET file = excluded.file, "
                "modificato = excluded.modificato",
                (pianta_uid(e["uid"], piano), edificio_id, piano, file,
                 modificato or now_utc_iso()))
        if old and old["file"] and old["file"] != file:
            self._remove_file(os.path.join(piante_root(), *old["file"].split("/")))

    def delete_pianta(self, edificio_id, piano):
        old = self.pianta_of(edificio_id, piano)
        if not old:
            return
        if old["file"]:
            self._remove_file(os.path.join(piante_root(), *old["file"].split("/")))
        with self.conn:
            self.conn.execute("DELETE FROM piante WHERE id = ?", (old["id"],))

    # --- interventi del tecnico Konica sulle stampanti ---
    def manutenzioni_search(self, edificio_id=None, stampante_id=None, stato="", dal=None, al=None,
                            testo=""):
        sql = SELECT_MANUT + " WHERE 1=1"
        params = []
        if edificio_id:
            sql += " AND e.id = ?"
            params.append(edificio_id)
        if stampante_id:
            sql += " AND s.id = ?"
            params.append(stampante_id)
        if stato in STATI_TECNICO:
            sql += " AND m.stato = ?"
            params.append(stato)
        if dal:
            sql += " AND substr(m.data_ora,1,10) >= ?"
            params.append(dal.isoformat())
        if al:
            sql += " AND substr(m.data_ora,1,10) <= ?"
            params.append(al.isoformat())
        sql += " ORDER BY m.data_ora DESC, m.id DESC"
        rows = [dict(r) for r in self.conn.execute(sql, params)]
        terms = testo.casefold().split()
        if terms:
            rows = [r for r in rows if matches_all(terms, " ".join((
                r["problema"], r["intervento"], r["tecnico"], r["ticket"], r["stampante"],
                r["modello"])))]
        return rows

    def manutenzioni_of(self, stampante_id):
        return self.manutenzioni_search(stampante_id=stampante_id)

    def get_manutenzione(self, mid):
        row = self.conn.execute(SELECT_MANUT + " WHERE m.id = ?", (mid,)).fetchone()
        return dict(row) if row else None

    def add_manutenzione(self, stampante_id, data_ora, data_tecnico, ticket, tecnico, problema,
                         intervento, stato, uid="", modificato=None):
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO manutenzioni(stampante_id, data_ora, data_tecnico, ticket, tecnico, "
                "problema, intervento, stato, uid, modificato) VALUES (?,?,?,?,?,?,?,?,?,?)",
                (stampante_id, data_ora, data_tecnico, ticket, tecnico, problema, intervento,
                 stato, uid, modificato or now_utc_iso()))
        return cur.lastrowid

    def update_manutenzione(self, mid, stampante_id, data_ora, data_tecnico, ticket, tecnico,
                            problema, intervento, stato, modificato=None):
        with self.conn:
            self.conn.execute(
                "UPDATE manutenzioni SET stampante_id=?, data_ora=?, data_tecnico=?, ticket=?, "
                "tecnico=?, problema=?, intervento=?, stato=?, modificato=? WHERE id=?",
                (stampante_id, data_ora, data_tecnico, ticket, tecnico, problema, intervento,
                 stato, modificato or now_utc_iso(), mid))

    def set_stato_manutenzione(self, mid, stato, data_tecnico=None):
        with self.conn:
            if data_tecnico is not None:
                self.conn.execute(
                    "UPDATE manutenzioni SET stato=?, data_tecnico=?, modificato=? WHERE id=?",
                    (stato, data_tecnico, now_utc_iso(), mid))
            else:
                self.conn.execute("UPDATE manutenzioni SET stato=?, modificato=? WHERE id=?",
                                  (stato, now_utc_iso(), mid))

    def delete_manutenzione(self, mid):
        with self.conn:
            self.conn.execute("DELETE FROM manutenzioni WHERE id=?", (mid,))

    def find_manutenzione_by_uid(self, uid):
        row = self.conn.execute("SELECT id, modificato FROM manutenzioni WHERE uid = ?",
                                (uid,)).fetchone()
        return dict(row) if row else None

    def count_manut_richieste(self):
        return self.conn.execute(
            "SELECT COUNT(*) FROM manutenzioni WHERE stato = 'Richiesto'").fetchone()[0]

    def toner_of(self, stampante_id):
        rows = self.conn.execute(
            "SELECT id, data_ora, colore, note FROM toner WHERE stampante_id = ? "
            "ORDER BY data_ora DESC, id DESC", (stampante_id,))
        return [dict(r) for r in rows]

    def search_toner(self, edificio_id=None, stampante_id=None, colore="", dal=None, al=None):
        sql = SELECT_TONER + " WHERE 1=1"
        params = []
        if edificio_id:
            sql += " AND e.id = ?"
            params.append(edificio_id)
        if stampante_id:
            sql += " AND s.id = ?"
            params.append(stampante_id)
        if colore:
            sql += " AND t.colore = ?"
            params.append(colore)
        if dal:
            sql += " AND substr(t.data_ora,1,10) >= ?"
            params.append(dal.isoformat())
        if al:
            sql += " AND substr(t.data_ora,1,10) <= ?"
            params.append(al.isoformat())
        sql += " ORDER BY t.data_ora DESC, t.id DESC"
        return [dict(r) for r in self.conn.execute(sql, params)]

    # ---------------- verbali Word ----------------
    def add_verbale(self, uid, tipo, data_iso, soggetto, dettaglio, dati, file, n_foto,
                    modificato=None, word_il=""):
        now = now_utc_iso()
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO verbali (uid, tipo, data, soggetto, dettaglio, dati, file, n_foto, "
                "creato, modificato, word_il) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (uid, tipo, data_iso, soggetto, dettaglio, json.dumps(dati, ensure_ascii=False),
                 file, n_foto, now, modificato or now, word_il))
        return cur.lastrowid

    def update_verbale(self, vid, data_iso, soggetto, dettaglio, dati, file, n_foto,
                       modificato=None, word_il=None):
        """word_il = quando e' stato fatto il Word (None = non cambia)."""
        sql = ("UPDATE verbali SET data=?, soggetto=?, dettaglio=?, dati=?, file=?, n_foto=?, "
               "modificato=?")
        params = [data_iso, soggetto, dettaglio, json.dumps(dati, ensure_ascii=False), file,
                  n_foto, modificato or now_utc_iso()]
        if word_il is not None:
            sql += ", word_il=?"
            params.append(word_il)
        with self.conn:
            self.conn.execute(sql + " WHERE id=?", params + [vid])

    def find_verbale_by_uid(self, uid):
        row = self.conn.execute("SELECT * FROM verbali WHERE uid = ?", (uid,)).fetchone()
        return dict(row) if row else None

    def get_verbale(self, vid):
        row = self.conn.execute("SELECT * FROM verbali WHERE id = ?", (vid,)).fetchone()
        return dict(row) if row else None

    def search_verbali(self, tipo="", testo=""):
        sql = "SELECT * FROM verbali WHERE 1=1"
        params = []
        if tipo:
            sql += " AND tipo = ?"
            params.append(tipo)
        sql += " ORDER BY data DESC, id DESC"
        rows = [dict(r) for r in self.conn.execute(sql, params)]
        terms = (testo or "").casefold().split()
        if terms:
            rows = [r for r in rows if matches_all(
                terms, " ".join((r["soggetto"], r["dettaglio"], r["file"])))]
        return rows

    def delete_verbale(self, vid):
        with self.conn:
            self.conn.execute("DELETE FROM verbali WHERE id = ?", (vid,))

    def backup_to(self, dest_path):
        dest = sqlite3.connect(dest_path)
        try:
            self.conn.backup(dest)
        finally:
            dest.close()


# --------------------------------------------------------------------------
# Esportazione Excel / PDF
# --------------------------------------------------------------------------
def _excel_sheet(ws, headers, rows, widths, text_cols=()):
    from openpyxl.styles import Alignment, Font, PatternFill

    ws.append(headers)
    for r in rows:
        ws.append(r)
    head_fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.column in text_cols:
                cell.number_format = "@"
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions


def _excel_summary(wb, title, filtri_desc, total_label, total, tables):
    """Foglio 'Riepilogo': titolo, filtri, totale e una o più tabelle di conteggi."""
    from openpyxl.styles import Font, PatternFill

    head_fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
    rs = wb.create_sheet("Riepilogo")
    rs["A1"] = title
    rs["A1"].font = Font(bold=True, size=14, color=NAVY.lstrip("#"))
    rs["A2"] = "Generato il " + datetime.now().strftime("%d/%m/%Y %H:%M")
    rs["A3"] = "Filtri: " + filtri_desc
    rs["A4"] = total_label
    rs["B4"] = total
    rs["A4"].font = rs["B4"].font = Font(bold=True)
    n = 6
    for headers, lines in tables:
        for i, h in enumerate(headers, start=1):
            c = rs.cell(row=n, column=i, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
        n += 1
        for line in lines:
            for i, v in enumerate(line, start=1):
                rs.cell(row=n, column=i, value=v)
            n += 1
        n += 1
    rs.column_dimensions["A"].width = 40
    for col in "BCDEF":
        rs.column_dimensions[col].width = 14


def export_excel(path, rows, filtri_desc):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Interventi"
    _excel_sheet(
        ws,
        ["Data e ora", "Utente", "Matricola", "Sede", "N. ticket", "Stato",
         "Per cosa mi hanno chiamato", "Come ho risolto", "Note"],
        [[fmt_dt(r["data_ora"]), r["utente"], r["matricola"], r["sede"], r["ticket"], r["stato"],
          r["richiesta"], r["soluzione"], r["note"]] for r in rows],
        [17, 24, 12, 14, 14, 11, 50, 50, 35], text_cols=(3, 5))
    per_user = [[nome, cnt] for nome, cnt in Counter(r["utente"] for r in rows).most_common()]
    per_sede = [[(sede or "(non indicata)"), cnt]
                for sede, cnt in Counter(r["sede"] for r in rows).most_common()]
    cnt_stato = Counter(r["stato"] for r in rows)
    per_stato = [[s, cnt_stato[s]] for s in STATI_INTERVENTO]
    _excel_summary(wb, "Riepilogo interventi", filtri_desc, "Totale interventi", len(rows),
                   [(["Stato", "Interventi"], per_stato), (["Utente", "Interventi"], per_user),
                    (["Sede", "Interventi"], per_sede)])
    wb.save(path)


def export_toner_excel(path, rows, filtri_desc):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Cambi toner"
    _excel_sheet(
        ws,
        ["Data e ora", "Sede", "Edificio", "Piano", "Stampante", "Modello", "Colore", "Note"],
        [[fmt_dt(r["data_ora"]), r["sede"], r["edificio"], floor_label(r["piano"]),
          r["stampante"], r["modello"], r["colore"], r["note"]] for r in rows],
        [17, 14, 34, 14, 28, 16, 12, 40])
    counts = Counter()
    names = {}
    for r in rows:
        key = (r["stampante_id"])
        names[key] = "%s - %s - %s" % (r["edificio"], floor_label(r["piano"]),
                                       r["stampante"] or r["modello"] or "Stampante")
        counts[(key, r["colore"])] += 1
    lines = []
    for key in sorted(names, key=lambda k: names[k].casefold()):
        per_color = [counts[(key, c)] for c in TONER_NOMI]
        lines.append([names[key]] + per_color + [sum(per_color)])
    _excel_summary(wb, "Riepilogo cambi toner", filtri_desc, "Totale cambi toner", len(rows),
                   [(["Stampante"] + TONER_NOMI + ["Totale"], lines)])
    wb.save(path)


def export_manut_excel(path, rows, filtri_desc):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Interventi tecnico"
    _excel_sheet(
        ws,
        ["Richiesta del", "Sede", "Edificio", "Piano", "Stampante", "Modello", "N. ticket",
         "Stato", "Intervento del tecnico", "Tecnico", "Problema segnalato",
         "Intervento svolto"],
        [[fmt_dt(r["data_ora"]), r["sede"], r["edificio"], floor_label(r["piano"]),
          r["stampante"], r["modello"], r["ticket"], r["stato"],
          fmt_date(datetime.strptime(r["data_tecnico"], "%Y-%m-%d")) if r["data_tecnico"] else "",
          r["tecnico"], r["problema"], r["intervento"]] for r in rows],
        [17, 14, 30, 14, 26, 16, 14, 11, 16, 20, 45, 45], text_cols=(7,))
    names, counts = {}, Counter()
    for r in rows:
        names[r["stampante_id"]] = "%s - %s - %s" % (
            r["edificio"], floor_label(r["piano"]), r["stampante"] or r["modello"] or "Stampante")
        counts[(r["stampante_id"], r["stato"])] += 1
    lines = []
    for key in sorted(names, key=lambda k: names[k].casefold()):
        per = [counts[(key, s)] for s in STATI_TECNICO]
        lines.append([names[key]] + per + [sum(per)])
    _excel_summary(wb, "Riepilogo interventi del tecnico", filtri_desc, "Totale interventi",
                   len(rows), [(["Stampante"] + STATI_TECNICO + ["Totale"], lines)])
    wb.save(path)


def _clean(v):
    return str(v).strip() if v is not None else ""


def safe_photo_name(name):
    name = os.path.basename(_clean(name)) or "foto.jpg"
    name = re.sub(r"[^\w.\- ]", "_", name)
    if os.path.splitext(name)[1].lower() not in IMG_EXTS:
        name += ".jpg"
    return name


def import_phone_toner(db, items, res):
    """Cambi toner registrati sul telefono. Quelli già presenti (stesso uid) non si duplicano;
    se la stampante non esiste più sul PC il cambio viene saltato."""
    for t in items:
        uid = _clean(t.get("uid"))
        colore = _clean(t.get("colore"))
        data_ora = _clean(t.get("data_ora"))
        try:
            datetime.strptime(data_ora, "%Y-%m-%d %H:%M")
        except ValueError:
            uid = ""
        if not uid or colore not in TONER_NOMI:
            res["toner_saltati"] += 1
            continue
        if db.toner_uid_exists(uid):
            res["toner_gia_presenti"] += 1
            continue
        sid = db.find_stampante_by_uid(_clean(t.get("stampante_uid")))
        if not sid:
            res["toner_saltati"] += 1
            continue
        row = db.add_toner_from_phone(sid, uid, colore, data_ora, _clean(t.get("note")))
        res["toner_nuovi"] += 1
        if row and scorta_status(row["quantita"], row["soglia"]) != "OK":
            tag = "%s %s: ne restano %d" % (row["colore"], row["modello"], row["quantita"])
            if tag not in res["scorta_bassa"]:
                res["scorta_bassa"].append(tag)


def _valid_dt(text):
    try:
        datetime.strptime(text, "%Y-%m-%d %H:%M")
        return True
    except ValueError:
        return False


def import_phone_bolle(db, items, res):
    """Bolle dei toner registrate sul telefono (arrivo ed eventuale montaggio). Quelle già presenti si
    aggiornano solo se sul telefono sono state modificate dopo; la foto si salva una volta sola.
    All'arrivo la scorta del modello sale di uno (la stessa cosa di «Ho ricevuto toner»)."""
    for k in ("bolle_nuove", "bolle_aggiornate", "bolle_gia_presenti", "bolle_saltate", "bolle_foto"):
        res.setdefault(k, 0)
    ed = {r["uid"]: r["id"] for r in db.conn.execute("SELECT id, uid FROM edifici")}
    for it in items:
        if not isinstance(it, dict):
            continue
        uid = re.sub(r"[^\w\-]", "_", _clean(it.get("uid")))
        eid = ed.get(_clean(it.get("edificio_uid")))
        colore, arrivo = _clean(it.get("colore")), _clean(it.get("data_arrivo"))
        try:
            piano = int(it.get("piano"))
        except (TypeError, ValueError):
            piano = None
        if not uid or eid is None or piano is None or colore not in TONER_NOMI or not _valid_dt(arrivo):
            res["bolle_saltate"] += 1
            continue
        mod = _clean(it.get("modificato")) or now_utc_iso()
        cur = db.find_bolla_by_uid(uid)
        if cur and not mod > cur["modificato"]:
            res["bolle_gia_presenti"] += 1
            continue
        data_m = _clean(it.get("data_montaggio"))
        montato = _clean(it.get("stato")) == "montato" and _valid_dt(data_m)
        foto = cur["foto"] if cur else ""
        payload = _clean(it.get("foto"))
        if payload.startswith("data:") and not foto:
            name = uid + ".jpg"
            if _write_data_url(payload, os.path.join(bolle_foto_root(), name), 12 * 1024 * 1024):
                foto = name
                res["bolle_foto"] += 1
        db.save_phone_bolla(cur, uid, {
            "codice": _clean(it.get("codice")), "numero": _clean(it.get("numero")),
            "toner_codice": _clean(it.get("toner_codice")), "colore": colore, "edificio_id": eid,
            "piano": piano, "stampante_id": db.find_stampante_by_uid(_clean(it.get("stampante_uid"))),
            "modello": " ".join(_clean(it.get("modello")).split()),
            "seriale_bolla": _clean(it.get("seriale_bolla")), "data_arrivo": arrivo,
            "stato": "montato" if montato else "da_montare", "data_montaggio": data_m if montato else "",
            "stampante_montata_id": db.find_stampante_by_uid(
                _clean(it.get("stampante_montata_uid"))) if montato else None,
            "seriale_letto": _clean(it.get("seriale_letto")) if montato else "",
            "toner_uid": _clean(it.get("toner_uid")) if montato else "", "note": _clean(it.get("note")),
            "foto": foto, "modificato": mod})
        res["bolle_aggiornate" if cur else "bolle_nuove"] += 1


def bolle_for_phone(db, with_photos):
    """Le bolle per il telefono (serve a ripristinarle se il telefono viene azzerato)."""
    out, n_foto = [], 0
    rows = db.conn.execute(
        "SELECT b.*, e.uid AS e_uid, s.uid AS s_uid, m.uid AS m_uid FROM bolle b "
        "JOIN edifici e ON e.id = b.edificio_id LEFT JOIN stampanti s ON s.id = b.stampante_id "
        "LEFT JOIN stampanti m ON m.id = b.stampante_montata_id ORDER BY b.id").fetchall()
    for r in rows:
        item = {"uid": r["uid"], "codice": r["codice"], "numero": r["numero"],
                "toner_codice": r["toner_codice"], "colore": r["colore"], "edificio_uid": r["e_uid"],
                "piano": r["piano"], "stampante_uid": r["s_uid"] or "", "modello": r["modello"],
                "seriale_bolla": r["seriale_bolla"], "data_arrivo": r["data_arrivo"], "stato": r["stato"],
                "data_montaggio": r["data_montaggio"], "stampante_montata_uid": r["m_uid"] or "",
                "seriale_letto": r["seriale_letto"], "toner_uid": r["toner_uid"], "note": r["note"],
                "modificato": r["modificato"]}
        if with_photos and r["foto"]:
            path = os.path.join(bolle_foto_root(), r["foto"])
            data = photo_data_url(path) if os.path.isfile(path) else None
            if data:
                item["foto"] = data
                n_foto += 1
        out.append(item)
    return out, n_foto


def import_phone_manutenzioni(db, items, res):
    """Interventi del tecnico Konica registrati sul telefono (stessa regola degli interventi:
    se ci sono già si aggiornano solo se sul telefono sono stati modificati dopo)."""
    for m in items:
        uid = _clean(m.get("uid"))
        data_ora = _clean(m.get("data_ora"))
        try:
            datetime.strptime(data_ora, "%Y-%m-%d %H:%M")
        except ValueError:
            uid = ""
        data_t = _clean(m.get("data_tecnico"))
        if data_t:
            try:
                datetime.strptime(data_t, "%Y-%m-%d")
            except ValueError:
                data_t = ""
        stato = _clean(m.get("stato"))
        if stato not in STATI_TECNICO:
            stato = "Eseguito" if data_t else "Richiesto"
        sid = db.find_stampante_by_uid(_clean(m.get("stampante_uid")))
        if not uid or not sid:
            res["manut_saltate"] += 1
            continue
        mod = _clean(m.get("modificato")) or now_utc_iso()
        existing = db.find_manutenzione_by_uid(uid)
        if existing and not mod > existing["modificato"]:
            res["manut_gia_presenti"] += 1
            continue
        args = (sid, data_ora, data_t, _clean(m.get("ticket")), _clean(m.get("tecnico")),
                _clean(m.get("problema")), _clean(m.get("intervento")), stato)
        if existing:
            db.update_manutenzione(existing["id"], *args, modificato=mod)
            res["manut_aggiornate"] += 1
        else:
            db.add_manutenzione(*args, uid=uid, modificato=mod)
            res["manut_nuove"] += 1


def import_phone_data(db, data):
    """Unisce nel database i dati ricevuti dal telefono (o da un backup del telefono).
    Utenti: vince la modifica più recente. Interventi: quelli già presenti non si duplicano
    (si aggiornano solo se sul telefono sono stati modificati dopo)."""
    if isinstance(data, dict) and data.get("tipo") == ELENCO_TIPO:
        raise ValueError("Questo è un elenco di stanze e persone: usa File › Importa elenco "
                         "stanze e persone.")
    if not isinstance(data, dict) or data.get("tipo") not in ("helpdesk-telefono", "helpdesk-pc"):
        raise ValueError("Questo non è un file dati dell'app Help Desk.")
    res = {"utenti_nuovi": 0, "utenti_aggiornati": 0, "utenti_uniti": 0, "interventi_nuovi": 0,
           "interventi_aggiornati": 0, "interventi_gia_presenti": 0, "foto": 0, "saltati": 0,
           "toner_nuovi": 0, "toner_gia_presenti": 0, "toner_saltati": 0, "scorta_bassa": [],
           "manut_nuove": 0, "manut_aggiornate": 0, "manut_gia_presenti": 0, "manut_saltate": 0,
           "verbali_nuovi": 0, "verbali_aggiornati": 0, "verbali_gia_presenti": 0,
           "verbali_saltati": 0, "verbali_foto": 0, "stanze_nuove": 0, "stanze_aggiornate": 0,
           "stanze_gia_presenti": 0, "stanze_saltate": 0, "piante": 0, "stanze_foto": 0,
           "postazioni": 0, "postazioni_saltate": 0, "bolle_nuove": 0, "bolle_aggiornate": 0,
           "bolle_gia_presenti": 0, "bolle_saltate": 0, "bolle_foto": 0}

    for u in data.get("utenti") or []:
        nome = normalize_name(_clean(u.get("nome")))
        if not nome:
            continue
        canon = db.resolve_user_name(nome)
        if canon != nome:
            # persona unita a un'altra (o rinominata) sul PC: il telefono ne ha ancora il vecchio
            # record. Non la faccio tornare; a quella rimasta aggiungo solo quello che le manca.
            res["utenti_uniti"] += 1
            tgt = db.find_user(canon)
            if tgt:
                ch = [db.set_user_basic(
                          tgt["id"], sede=None if tgt["sede"] else normalize_name(_clean(u.get("sede"))),
                          telefono=None if tgt["telefono"] else _clean(u.get("telefono"))),
                      db.set_user_codes(
                          tgt["id"],
                          seriale_pc=None if tgt["seriale_pc"] else _clean(u.get("seriale_pc")),
                          imei=None if tgt["imei"] else _clean(u.get("imei")),
                          matricola=None if tgt["matricola"] else _clean(u.get("matricola")))]
                if ch[0] or ch[1]:
                    res["utenti_aggiornati"] += 1
            continue
        cur = db.find_user(nome)
        # un file di una versione vecchia dell'app non ha la matricola: in quel caso si tiene
        # quella già salvata
        matricola = _clean(u.get("matricola")) if "matricola" in u else (
            cur["matricola"] if cur else "")
        fields = (normalize_name(_clean(u.get("sede"))), _clean(u.get("telefono")),
                  _clean(u.get("seriale_pc")), _clean(u.get("imei")), matricola)
        mod = _clean(u.get("modificato"))
        # ufficio e dotazione: un file di una versione vecchia non li ha, e in quel caso si
        # tengono quelli già salvati
        uff = " ".join(_clean(u.get("ufficio")).split()) if "ufficio" in u else None
        dot = clean_dotazione(u.get("dotazione")) if "dotazione" in u else None
        # anche l'email: un file di una versione vecchia non ce l'ha, e in quel caso si tiene
        mail = clean_email(u.get("email")) if "email" in u else None
        if mail is not None and not EMAIL_RE.match(mail):
            mail = ""
        if not cur:
            uid, _ = db.add_user(nome, *fields, modificato=mod or now_utc_iso())
            db.set_user_extra(uid, uff, dot, modificato=mod or now_utc_iso())
            if mail:
                db.set_user_email(uid, mail, modificato=mod or now_utc_iso())
            res["utenti_nuovi"] += 1
        elif mod and mod > (cur.get("modificato") or ""):
            same_extra = (uff is None or uff == cur["ufficio"]) and \
                         (dot is None or dot == cur["dotazione"]) and \
                         (mail is None or mail == cur["email"])
            same_fields = fields == (cur["sede"], cur["telefono"], cur["seriale_pc"],
                                     cur["imei"], cur["matricola"])
            if not (same_fields and same_extra):
                if not same_fields:
                    db.update_user(cur["id"], cur["nome"], *fields, modificato=mod)
                db.set_user_extra(cur["id"], uff, dot, modificato=mod)
                if mail is not None:
                    db.set_user_email(cur["id"], mail, modificato=mod)
                res["utenti_aggiornati"] += 1

    if data.get("tipo") != "helpdesk-telefono":
        return res
    import_phone_toner(db, data.get("toner") or [], res)
    import_phone_bolle(db, data.get("bolle") or [], res)
    import_phone_manutenzioni(db, data.get("manutenzioni") or [], res)
    import_phone_verbali(db, data.get("verbali") or [], res)
    import_phone_stanze(db, data, res)
    for it in data.get("interventi") or []:
        uid = _clean(it.get("uid"))
        utente = db.resolve_user_name(_clean(it.get("utente")))     # persona unita: quella rimasta
        try:
            datetime.strptime(_clean(it.get("data_ora")), "%Y-%m-%d %H:%M")
        except ValueError:
            uid = ""
        if not uid or not utente:
            res["saltati"] += 1
            continue
        mod = _clean(it.get("modificato")) or now_utc_iso()
        existing = db.find_intervento_by_uid(uid)
        if existing and not mod > existing["modificato"]:
            res["interventi_gia_presenti"] += 1
            continue
        stato = _clean(it.get("stato"))
        if stato not in STATI_INTERVENTO:
            stato = existing["stato"] if existing else "Risolto"
        iid = db.save_phone_intervento(
            existing["id"] if existing else None, uid, _clean(it.get("data_ora")), utente,
            normalize_name(_clean(it.get("sede"))), _clean(it.get("richiesta")),
            _clean(it.get("soluzione")), _clean(it.get("note")), _clean(it.get("ticket")), mod,
            stato)
        res["interventi_aggiornati" if existing else "interventi_nuovi"] += 1

        have = {a["nome"] for a in db.allegati_of(iid)}
        for f in it.get("foto") or []:
            nome = safe_photo_name(f.get("nome"))
            payload = _clean(f.get("data"))
            if nome in have or "," not in payload:
                continue
            try:
                raw = base64.b64decode(payload.split(",", 1)[1])
            except ValueError:
                continue
            stage = os.path.join(staging_dir(), "import_%d" % iid)
            os.makedirs(stage, exist_ok=True)
            tmp = os.path.join(stage, nome)
            with open(tmp, "wb") as fh:
                fh.write(raw)
            try:
                stored = store_image(tmp, attachment_dir(iid))
                db.add_allegato(iid, stored, "%s/%s" % (iid, stored))
                res["foto"] += 1
                have.add(nome)
                have.add(stored)
            finally:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
    return res


def clean_verbale_dati(dati):
    """Pulisce i dati di un verbale arrivato dal telefono (testi, elenchi, date gg/mm/aaaa)."""
    out = {}
    for k, v in (dati or {}).items():
        if not isinstance(k, str) or k.startswith("_"):
            continue
        if k in ("motivi", "materiale", "accessori"):
            out[k] = [_clean(x) for x in (v if isinstance(v, list) else []) if _clean(x)]
        elif k == "apparati":
            out[k] = [{c: _clean(r.get(c)) for c in ("marca", "modello", "codice")}
                      for r in (v if isinstance(v, list) else []) if isinstance(r, dict)]
        elif k == "dotazione":
            out[k] = [{c: _clean(r.get(c)) for c in ("tipo", "modello", "codice", "iccid", "numero")}
                      for r in (v if isinstance(v, list) else []) if isinstance(r, dict)]
        else:
            out[k] = _clean(v)
    for k in ("data", "data_sped"):
        m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", out.get(k, ""))
        if m:
            out[k] = "%s/%s/%s" % (m.group(3), m.group(2), m.group(1))
    return out


def _write_data_url(payload, path, max_bytes):
    """Salva su file un'immagine in formato data URL. True se riuscito."""
    payload = _clean(payload)
    if "," not in payload:
        return False
    try:
        raw = base64.b64decode(payload.split(",", 1)[1])
    except ValueError:
        return False
    if not raw or len(raw) > max_bytes:
        return False
    with open(path, "wb") as fh:
        fh.write(raw)
    return True


def import_phone_verbali(db, items, res):
    """Verbali fatti sul telefono (dati, foto, firme). Un verbale gia' presente si aggiorna solo
    se sul telefono e' stato modificato dopo."""
    for it in items:
        if not isinstance(it, dict):
            continue
        uid = re.sub(r"[^A-Za-z0-9_-]", "", _clean(it.get("uid")))[:40]
        tipo = _clean(it.get("tipo"))
        if VB is None or tipo not in VB.TIPI or not uid or not isinstance(it.get("dati"), dict):
            res["verbali_saltati"] += 1
            continue
        mod = _clean(it.get("modificato")) or now_utc_iso()
        existing = db.find_verbale_by_uid(uid)
        if existing and not mod > existing["modificato"]:
            res["verbali_gia_presenti"] += 1
            continue
        dati = clean_verbale_dati(it["dati"])
        folder = os.path.join(verbali_foto_root(), uid)
        os.makedirs(folder, exist_ok=True)
        keep, foto_names, firme = set(), [], {}
        for n, f in enumerate(it.get("foto") or [], 1):
            if not isinstance(f, dict):
                continue
            name = "foto_%02d_%s" % (n, safe_photo_name(f.get("nome")))
            if _write_data_url(f.get("data"), os.path.join(folder, name), 15 * 1024 * 1024):
                foto_names.append(name)
                keep.add(name)
                res["verbali_foto"] += 1
        for chi in ("tecnico", "utente"):
            src = (it.get("firme") or {}).get(chi) if isinstance(it.get("firme"), dict) else None
            name = "firma_%s.png" % chi
            if src and _write_data_url(src, os.path.join(folder, name), 3 * 1024 * 1024):
                firme[chi] = name
                keep.add(name)
        cart_src = it.get("cartaceo") if isinstance(it.get("cartaceo"), dict) else {}
        # firma a mano sul foglio: vale solo se per quel posto non c'è già una firma (l'immagine vince)
        cartaceo = {chi: True for chi in ("tecnico", "utente")
                    if cart_src.get(chi) is True and chi not in firme}
        for name in os.listdir(folder):          # tolgo i file che sul telefono non ci sono piu'
            if name not in keep:
                try:
                    os.remove(os.path.join(folder, name))
                except OSError:
                    pass
        dati["_foto"] = foto_names
        dati["_firme"] = firme
        dati["_cartaceo"] = cartaceo
        soggetto, dettaglio = riassunto_verbale(tipo, dati)
        data_iso = (parse_date(dati.get("data", "")) or date.today()).isoformat()
        if existing:
            db.update_verbale(existing["id"], data_iso, soggetto, dettaglio, dati,
                              existing["file"], len(foto_names), modificato=mod)
            res["verbali_aggiornati"] += 1
        else:
            db.add_verbale(uid, tipo, data_iso, soggetto, dettaglio, dati, "", len(foto_names),
                           modificato=mod)
            res["verbali_nuovi"] += 1


PHONE_PHOTO_PX = 1400   # lato lungo massimo delle foto mandate al telefono


def photo_data_url(path, px=None):
    """Foto -> testo (data URL) per il file del telefono, ridotta per non appesantirlo."""
    px = px or PHONE_PHOTO_PX
    try:
        from PIL import Image, ImageOps
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im)
            im.thumbnail((px, px))
            buf = io.BytesIO()
            im.convert("RGB").save(buf, "JPEG", quality=78)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        pass
    mime = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png",
            ".gif": "image/gif", ".webp": "image/webp"}.get(os.path.splitext(path)[1].lower())
    if not mime:
        return None
    try:
        with open(path, "rb") as fh:
            return "data:%s;base64,%s" % (mime, base64.b64encode(fh.read()).decode("ascii"))
    except OSError:
        return None


# ---------------------------------------------------------------------------------------------
# Importazione delle persone da una «Lista persone» in Excel (esportata dal sistema INAIL)
# ---------------------------------------------------------------------------------------------
PEOPLE_COLUMNS = {"nominativo": "nome", "cellulare": "cellulare", "telefono_fisso": "fisso",
                  "ufficio": "ufficio", "email": "email", "username": "matricola"}


def read_people_xlsx(path):
    """Legge una lista persone in Excel (colonne Nominativo, Cellulare, Telefono_Fisso, Ufficio,
    Email, Username). Restituisce (titolo del foglio, [persone]); ogni persona è un dizionario con
    nome, cellulare, fisso, ufficio, email, matricola. Non scrive niente."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        raise ValueError("non riesco ad aprire il file Excel (%s)" % e)
    for ws in wb.worksheets:
        cols, people = None, []
        for row in ws.iter_rows(values_only=True):
            if cols is None:
                found = {}
                for i, v in enumerate(row):
                    k = PEOPLE_COLUMNS.get(_clean(v).casefold())
                    if k and k not in found:
                        found[k] = i
                if "nome" in found:
                    cols = found
                continue

            def cell(k, limit):
                i = cols.get(k)
                return " ".join(_clean(row[i]).split())[:limit] if i is not None and i < len(row) \
                    else ""
            nome = cell("nome", 120)
            if not nome:
                continue
            people.append({"nome": nome, "cellulare": cell("cellulare", 40),
                           "fisso": cell("fisso", 40), "ufficio": cell("ufficio", 120),
                           "email": cell("email", 120).lower(),
                           "matricola": cell("matricola", 30).replace(" ", "").upper()})
        if cols is not None:
            return ws.title, people
    raise ValueError("non sembra una lista di persone: manca la colonna «Nominativo»")


def guess_people_sede(title, filename=""):
    """La sede che si capisce dal nome del foglio (per esempio «DIREZIONE REGIONALE» = DR Toscana);
    vuoto se non si capisce."""
    t = " ".join(("%s %s" % (title, filename)).upper().split())
    if "DIREZIONE REGIONALE" in t or re.search(r"\bDR\b", t):
        return "DR Toscana"
    for s in SEDI_TOSCANA:
        if s != "DR Toscana" and s.upper() in t:
            return s
    return ""


def _phone_key(text):
    d = re.sub(r"\D", "", text or "")
    return d[2:] if d.startswith("39") and len(d) > 10 else d


def plan_people_import(users, lists, update_sede=False):
    """Prepara l'importazione senza scrivere niente. users: le persone dell'app (db.users_basic());
    lists: [{"sede": ..., "persone": [...]}] come da read_people_xlsx. Riconosce la persona prima
    dalla matricola, poi dal nome (anche «Rossi Mario» per «Mario Rossi»). Dove c'è già qualcosa
    scritto non lo cambia (a parte la sede, solo se update_sede è vero): scrive matricola,
    cellulare, ufficio e sede dove mancano e aggiunge le persone nuove."""
    by_mat, by_words, recs = {}, {}, []

    def words(name):
        return " ".join(sorted(key_name(name).split()))

    def index(rec):
        if rec["matricola"]:
            by_mat.setdefault(rec["matricola"].casefold(), rec)
        by_words.setdefault(words(rec["nome"]), []).append(rec)

    for u in users:
        rec = {"id": u["id"], "nome": u["nome"], "matricola": u.get("matricola") or "",
               "sede": u.get("sede") or "", "telefono": u.get("telefono") or "",
               "ufficio": u.get("ufficio") or "", "ordine": u.get("ordine") or "",
               "email": u.get("email") or "", "new": False, "set": {}}
        recs.append(rec)
        index(rec)
    st = {"persone": 0, "nuove": 0, "presenti": 0, "aggiornate": 0, "invariate": 0,
          "matricole": 0, "telefoni": 0, "sedi_scritte": 0, "sedi_cambiate": 0, "uffici": 0,
          "telefono_diverso": 0, "sede_diversa": 0, "senza_telefono": 0, "email_nel_file": 0,
          "email": 0, "ordinate": 0,
          "doppi": 0, "conflitti": []}
    seen = set()
    for lst in lists:
        sede = " ".join(str(lst.get("sede") or "").split())
        for p in lst.get("persone") or []:
            st["persone"] += 1
            nome = normalize_name(p.get("nome"))
            if not nome:
                continue
            mat = " ".join(str(p.get("matricola") or "").split())
            tel = p.get("cellulare") or p.get("fisso") or ""
            uff = p.get("ufficio") or ""
            mail = clean_email(p.get("email"))
            if not EMAIL_RE.match(mail):
                mail = ""
            if mail:
                st["email_nel_file"] += 1
            if not tel:
                st["senza_telefono"] += 1
            rec = by_mat.get(mat.casefold()) if mat else None
            if rec is None:
                hits = by_words.get(words(nome), [])
                if len(hits) > 1:
                    st["conflitti"].append("%s: ci sono più persone con questo nome" % nome)
                    continue
                if hits:
                    rec = hits[0]
                    if mat and rec["matricola"] and rec["matricola"].casefold() != mat.casefold():
                        st["conflitti"].append(
                            "%s: stesso nome ma matricola diversa (nell'app %s, nel file %s)"
                            % (nome, rec["matricola"], mat))
                        continue
            if rec is not None:
                if id(rec) in seen:
                    st["doppi"] += 1
                    continue
                seen.add(id(rec))
                st["presenti"] += 1
                ch = rec["set"]
                if mat and not rec["matricola"]:
                    ch["matricola"] = mat
                    st["matricole"] += 1
                if tel:
                    if not rec["telefono"]:
                        ch["telefono"] = tel
                        st["telefoni"] += 1
                    elif _phone_key(rec["telefono"]) != _phone_key(tel):
                        st["telefono_diverso"] += 1
                if uff and not rec["ufficio"]:
                    ch["ufficio"] = uff
                    st["uffici"] += 1
                if mail and not rec["email"]:
                    ch["email"] = mail
                    st["email"] += 1
                # per ordinare per cognome: il «Cognome Nome» del file (vuoto se il nome è già così)
                ordine = "" if key_name(nome) == key_name(rec["nome"]) else nome
                if ordine != rec["ordine"]:
                    ch["ordine"] = ordine
                    st["ordinate"] += 1
                if sede and rec["sede"] != sede:
                    if not rec["sede"]:
                        ch["sede"] = sede
                        st["sedi_scritte"] += 1
                    elif update_sede:
                        ch["sede"] = sede
                        st["sedi_cambiate"] += 1
                    else:
                        st["sede_diversa"] += 1
                if ch:
                    st["aggiornate"] += 1
                else:
                    st["invariate"] += 1
                continue
            rec = {"id": None, "nome": nome, "matricola": mat, "sede": sede, "telefono": tel,
                   "ufficio": uff, "ordine": "", "email": mail, "new": True, "set": {}}
            recs.append(rec)
            index(rec)
            seen.add(id(rec))
            st["nuove"] += 1
    return {"actions": [r for r in recs if r["new"] or r["set"]], "stats": st}


def apply_people_plan(db, plan):
    """Scrive nel database quello che ha preparato plan_people_import."""
    for rec in plan["actions"]:
        if rec["new"]:
            uid, _ = db.add_user(rec["nome"], rec["sede"], rec["telefono"], "", "",
                                 rec["matricola"])
            if rec["ufficio"]:
                db.set_user_extra(uid, ufficio=rec["ufficio"])
            if rec["email"]:
                db.set_user_email(uid, rec["email"])
            continue
        ch = rec["set"]
        db.set_user_basic(rec["id"], sede=ch.get("sede"), telefono=ch.get("telefono"))
        if ch.get("matricola"):
            db.set_user_codes(rec["id"], matricola=ch["matricola"])
        if ch.get("ufficio"):
            db.set_user_extra(rec["id"], ufficio=ch["ufficio"])
        if ch.get("email"):
            db.set_user_email(rec["id"], ch["email"])
        if "ordine" in ch:
            db.set_user_ordine(rec["id"], ch["ordine"])


# ---------------------------------------------------------------------------------------------
# Seriali dei PC da un file Excel: colonne «Nominativo» e «PC» (il nome del PC nel sistema)
# ---------------------------------------------------------------------------------------------
PC_NAME_RE = re.compile(r"^\s*\d{5}\s*-\s*([A-Za-z0-9]{5,15})\s*$")


def parse_pc_serial(value):
    """Il numero di serie dal nome del PC: «60016-AB12CDE» -> «AB12CDE» (le cinque cifre prima del
    trattino sono il codice dell'ente, non fanno parte del seriale). Un nome che non ha questa
    forma (per esempio «21000PC0001», un nome casuale o «PC non trovato») non è un seriale: vuoto."""
    m = PC_NAME_RE.match(str(value or ""))
    return m.group(1).upper() if m else ""


def read_pc_xlsx(path):
    """Legge un Excel con una colonna di nomi (Nominativo, Nome, Utente...) e una dei PC (PC, PC
    (displayName), Seriale, S/N). Restituisce {nome: [valori]} (una persona può avere più righe)."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        raise ValueError("non riesco ad aprire il file Excel (%s)" % e)
    for ws in wb.worksheets:
        cols, found = None, {}
        for row in ws.iter_rows(values_only=True):
            if cols is None:
                heads = [_clean(v).casefold() for v in row]
                ni = next((i for i, v in enumerate(heads)
                           if v in ("nominativo", "nome", "utente", "dipendente")), None)
                pi = next((i for i, v in enumerate(heads) if i != ni and v.startswith(
                    ("pc", "seriale", "s/n", "sn", "serial"))), None)
                if ni is not None and pi is not None:
                    cols = (ni, pi)
                continue
            ni, pi = cols
            nome = " ".join(_clean(row[ni]).split()) if ni < len(row) else ""
            if nome:
                val = " ".join(_clean(row[pi]).split()) if pi < len(row) else ""
                found.setdefault(nome, []).append(val)
        if cols is not None:
            return found
    raise ValueError("non sembra un elenco di PC: servono le colonne «Nominativo» e «PC»")


def merge_pc_entries(parts):
    """Più file con le stesse persone: i valori si sommano (senza ripetere quelli uguali)."""
    out = {}
    for part in parts:
        for nome, vals in part.items():
            cur = out.setdefault(nome, [])
            for v in vals:
                if v not in cur:
                    cur.append(v)
    return out


def plan_pc_serials(users, entries):
    """Prepara l'assegnazione dei seriali senza scrivere niente. users: db.users_basic(); entries:
    {nome: [nomi dei PC]}. Chi è nel file ma non nell'app NON si aggiunge. Chi ha già un seriale
    scritto non viene cambiato. Se una persona ha più PC si scrive il primo seriale valido."""
    by_words = {}
    for u in users:      # anche col «Cognome Nome» della lista INAIL (campo «ordine»), se c'è
        for name in {u["nome"], u.get("ordine") or u["nome"]}:
            lst = by_words.setdefault(" ".join(sorted(key_name(name).split())), [])
            if u not in lst:
                lst.append(u)
    st = {"nel_file": len(entries), "scritti": 0, "gia_uguali": 0, "diversi": [], "non_in_app": 0,
          "ambigui": [], "senza_pc": 0, "nome_non_seriale": [], "piu_pc": [], "prefisso_tolto": 0}
    def disp(u):                 # nei rapporti: «Cognome Nome», così gli elenchi sono per cognome
        return u.get("ordine") or u["nome"]
    actions = []
    def bare(text):              # «60016-AB12CDE» scritto a mano vale «AB12CDE»
        text = clean_serial(text or "")
        return parse_pc_serial(text) or text

    used = {bare(u.get("seriale_pc")).replace(" ", "").upper(): u["nome"]
            for u in users if bare(u.get("seriale_pc"))}                  # seriali già dati
    for nome, vals in entries.items():
        hits = by_words.get(" ".join(sorted(key_name(normalize_name(nome)).split())), [])
        if not hits:
            st["non_in_app"] += 1
            continue
        if len(hits) > 1:
            st["ambigui"].append("%s: nell'app ci sono %d persone che corrispondono (%s): "
                                 "scrivi il seriale a mano" % (
                                     normalize_name(nome), len(hits),
                                     ", ".join(disp(x) for x in hits)))
            continue
        u = hits[0]
        serials = []
        for v in vals:
            sn = parse_pc_serial(v)
            if sn and sn not in serials:
                serials.append(sn)
        if not serials:
            real = [v for v in vals if v and "non trovato" not in v.casefold()]
            if real:
                st["nome_non_seriale"].append("%s: %s" % (disp(u), ", ".join(real)))
            else:
                st["senza_pc"] += 1
            continue
        if len(serials) > 1:
            st["piu_pc"].append("%s: %s" % (disp(u), ", ".join(serials)))
        sn = serials[0]
        cur = clean_serial(u.get("seriale_pc") or "")
        if not cur:
            if sn in used:
                st["ambigui"].append("%s: seriale %s già dato a %s" % (disp(u), sn, used[sn]))
                continue
            used[sn] = disp(u)
            actions.append({"id": u["id"], "nome": u["nome"], "seriale": sn})
            st["scritti"] += 1
        elif bare(cur).replace(" ", "").casefold() == sn.casefold():
            if cur != bare(cur):         # scritto col codice dell'ente davanti: lo tolgo
                actions.append({"id": u["id"], "nome": u["nome"], "seriale": sn})
                st["prefisso_tolto"] += 1
            else:
                st["gia_uguali"] += 1
        else:
            st["diversi"].append("%s: nell'app %s, nel file %s" % (disp(u), cur, sn))
    return {"actions": actions, "stats": st}


def apply_pc_plan(db, plan):
    for a in plan["actions"]:
        db.set_user_codes(a["id"], seriale_pc=a["seriale"])


# ---------------------------------------------------------------------------------------------
# Persone doppie: riconoscimento (da proporre, non da fare da soli) e scelta di quale tenere
# ---------------------------------------------------------------------------------------------
def _damerau(a, b):
    """Quante lettere cambiano tra due parole (aggiunta, tolta, cambiata o due scambiate = 1)."""
    if a == b:
        return 0
    prev2, prev = None, list(range(len(b) + 1))
    for i in range(1, len(a) + 1):
        cur = [i] + [0] * len(b)
        for j in range(1, len(b) + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                cur[j] = min(cur[j], prev2[j - 2] + 1)
        prev2, prev = prev, cur
    return prev[len(b)]


def _typo_kind(x, y):
    """Due parole che differiscono poco: «sure» se sembra un errore di battitura (due lettere
    scambiate, una lettera in più in mezzo), «maybe» se potrebbero essere due nomi veri
    (Marco/Mario, Claudia/Claudio, Luca/Lucia), None se sono troppo diverse."""
    if x == y or min(len(x), len(y)) < 4:
        return None
    if len(x) == len(y):
        diff = [i for i in range(len(x)) if x[i] != y[i]]
        if len(diff) == 2 and diff[1] == diff[0] + 1 and x[diff[0]] == y[diff[1]] \
                and x[diff[1]] == y[diff[0]]:
            return "sure" if len(x) >= 4 else None                       # lettere scambiate
        if len(diff) == 1:
            return "sure" if len(x) >= 7 and diff[0] < len(x) - 1 else "maybe"
        return None
    if abs(len(x) - len(y)) == 1:
        short, long_ = (x, y) if len(x) < len(y) else (y, x)
        for i in range(len(long_)):
            if long_[:i] + long_[i + 1:] == short:
                return "sure" if len(long_) >= 7 and i < len(long_) - 1 else "maybe"
    return None


def _dup_dot_count(u):
    return (len(clean_dotazione(u.get("dotazione"))) + sum(
        1 for k in ("seriale_pc", "imei", "matricola", "telefono", "ufficio", "email") if u.get(k)))


def pick_duplicate_keeper(members):
    """Quale persona tenere tra le doppie: quella scritta «Cognome Nome» (si ordina per cognome) e
    scarto «Nome Cognome»; il «Cognome Nome» di riferimento è quello preso dalle liste INAIL
    (campo ordine). Se non c'è modo di saperlo, quella con più dati (matricola, interventi...).
    Restituisce (id da tenere, sicuro: True se è stata la forma del nome a deciderlo)."""
    ref = next((m["ordine"] for m in members if m.get("ordine")), "")
    refk = key_name(ref)
    ref_words = sorted(refk.split())

    def form(m):
        k = key_name(m["nome"])
        if not refk:
            return 1
        if k == refk:
            return 2                              # scritta come nella lista INAIL: Cognome Nome
        if m.get("ordine") or sorted(k.split()) == ref_words:
            return 0                              # stesse parole in ordine diverso: Nome Cognome
        return 1                                  # scritta in un altro modo (errore di battitura?)

    ranked = sorted(members, key=lambda m: (form(m), bool(m.get("matricola")), m.get("n") or 0,
                                            _dup_dot_count(m), -(m.get("id") or 0)), reverse=True)
    top = form(ranked[0])
    sure = top == 2 and all(form(m) < 2 for m in ranked[1:])
    return ranked[0]["id"], sure


def find_duplicate_users(users):
    """Gruppi di persone che sembrano la stessa. users: come db.users_with_counts() (id, nome, ordine,
    matricola, n...). Riconosce: stessa matricola; stesse parole in ordine diverso («Mario Rossi» e
    «Rossi Mario»); una parola scritta quasi uguale («Claudia» e «Caludia»). Più incerti (da
    controllare): una parola simile ma diversa, o una parola in più. Restituisce
    [{"members": [...], "why": [testi], "sure": bool, "likely": bool, "keep": id,
    "keep_sure": bool}] in ordine di cognome. Non scrive niente."""
    from collections import Counter
    n = len(users)
    parent = list(range(n))
    notes = {}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(i, j, why, sure):
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[rj] = ri
            notes[ri] = notes.get(ri, []) + notes.pop(rj, [])
            ri = find(i)
        notes.setdefault(ri, []).append((why, sure))

    words = [key_name(u["nome"]).split() for u in users]
    mats = [" ".join(str(u.get("matricola") or "").split()).casefold() for u in users]
    by_mat = {}
    for i, m in enumerate(mats):
        if m:
            by_mat.setdefault(m, []).append(i)
    for m, idx in by_mat.items():
        for i in idx[1:]:
            union(idx[0], i, "stessa matricola", True)
    for i in range(n):
        for j in range(i + 1, n):
            wa, wb = words[i], words[j]
            if len(wa) < 2 or len(wb) < 2:
                continue
            ca, cb = Counter(wa), Counter(wb)
            da, db = list((ca - cb).elements()), list((cb - ca).elements())
            diverse = bool(mats[i] and mats[j] and mats[i] != mats[j])
            if not da and not db:
                if diverse:
                    union(i, j, "stesso nome e cognome ma matricola diversa: controlla", False)
                else:
                    union(i, j, "stesso nome e cognome", True)
            elif diverse:
                continue                       # due matricole diverse: sono persone diverse
            elif len(da) == 1 and len(db) == 1:
                x, y = da[0], db[0]
                kind = _typo_kind(x, y)
                if kind == "sure":
                    union(i, j, "nome scritto quasi uguale", True)
                elif kind == "maybe" or (_damerau(x, y) <= 3 and min(len(x), len(y)) >= 5
                                         and x[0] == y[0]):
                    union(i, j, "nome simile: controlla", False)
            elif len(da) + len(db) == 1:
                union(i, j, "un nome ha una parola in più: controlla", False)
    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)
    out = []
    for root, idx in groups.items():
        if len(idx) < 2:
            continue
        members = sorted((users[i] for i in idx),
                         key=lambda u: ((u.get("ordine") or u["nome"]).casefold(), u["nome"].casefold()))
        why, sure = [], True
        for t, ok in notes.get(root, []):
            if t not in why:
                why.append(t)
            sure = sure and ok
        keep, keep_sure = pick_duplicate_keeper(members)
        # «probabile»: una sola ha la matricola (viene dalla lista INAIL) e le altre no
        likely = (not sure) and sum(1 for m in members if m.get("matricola")) == 1
        out.append({"members": members, "why": why, "sure": sure, "likely": likely, "keep": keep,
                    "keep_sure": keep_sure})
    out.sort(key=lambda g: ((g["members"][0].get("ordine") or g["members"][0]["nome"]).casefold()))
    return out


ELENCO_TIPO = "helpdesk-elenco-stanze"


def import_rooms_list(db, data, edificio_id):
    """Carica nell'edificio un elenco di stanze con le persone (nome e ufficio) e le stampanti
    (modello, seriale, stanza). Si può rifare quante volte si vuole senza creare doppioni e senza
    sovrascrivere quello che hai già scritto a mano:
    - una stanza che c'è già (stesso piano e codice) resta com'è; le scrivanie si scrivono solo se
      erano vuote;
    - una persona che c'è già (anche «Rossi Mario» per «Mario Rossi») resta com'è; l'ufficio e la
      dotazione (facoltativa nel file: [{"tipo", "modello", "seriale"}]) si scrivono solo se erano
      vuoti; chi siede già in un'altra stanza non viene spostato;
    - una stampante con lo stesso numero di serie non si duplica: le si collega la stanza (se non
      ne aveva) e, se serve, il modello e lo stato «fuori servizio»;
    - una stampante vecchia SENZA seriale (stesso piano e modello, e sicuramente quella: l'unica,
      o stesso nome, o stessa stanza) non si duplica: prende il seriale e il nome del file e tiene
      tutto lo storico. Se non è chiaro quale sia, ne resta una nuova e la si unisce a mano con
      «Unisci doppioni stampanti»."""
    if not isinstance(data, dict) or data.get("tipo") != ELENCO_TIPO:
        raise ValueError("Questo non è un elenco di stanze e persone dell'app Help Desk.")
    edificio = db.conn.execute("SELECT id, sede FROM edifici WHERE id = ?", (edificio_id,)).fetchone()
    if not edificio:
        raise ValueError("Scegli prima un edificio.")
    res = {"stanze_nuove": 0, "stanze_esistenti": 0, "stanze_aggiornate": 0,
           "persone_nuove": 0, "persone_esistenti": 0, "persone_assegnate": 0,
           "persone_gia_in_stanza": 0, "persone_altrove": [], "uffici_scritti": 0,
           "matricole_scritte": 0, "dotazioni_scritte": 0, "persone_senza_stanza": 0,
           "stampanti_nuove": 0, "stampanti_aggiornate": 0, "stampanti_gia_presenti": 0,
           "stampanti_adottate": 0, "stampanti_da_unire": 0, "scartati": []}

    def to_int(v, default=0):
        try:
            return int(v)
        except (TypeError, ValueError):
            return default

    def fill_person(uid, user, p, uff):
        """Ufficio, matricola e dotazione di una persona, solo dove mancano (non sovrascrive)."""
        if uff and not (user or {}).get("ufficio") and db.set_user_extra(uid, uff):
            res["uffici_scritti"] += 1
        mat = " ".join(_clean(p.get("matricola")).split())
        if mat and not (user or {}).get("matricola") and db.set_user_codes(uid, matricola=mat):
            res["matricole_scritte"] += 1
        dot = clean_dotazione(p.get("dotazione"))
        if dot:
            cur = db.conn.execute("SELECT dotazione FROM utenti WHERE id = ?", (uid,)).fetchone()
            if db.set_user_extra(uid, dotazione=merge_dotazione(cur["dotazione"], dot)):
                res["dotazioni_scritte"] += 1

    for r in data.get("stanze") or []:
        if not isinstance(r, dict):
            continue
        piano, codice = to_int(r.get("piano")), clean_room_code(_clean(r.get("codice")))
        if not codice:
            res["scartati"].append("stanza senza codice (piano %d)" % piano)
            continue
        if len(" ".join(_clean(r.get("codice")).split())) > ROOM_CODE_MAX:
            res["scartati"].append("stanza %s...: codice più lungo di %d caratteri"
                                   % (codice[:12], ROOM_CODE_MAX))
            continue
        scriv = max(0, to_int(r.get("scrivanie")))
        room = db.find_stanza_by_code(edificio_id, piano, codice)
        if not room:
            rid = db.add_stanza(edificio_id, piano, codice, scriv, _clean(r.get("note")),
                                _clean(r.get("come")))
            res["stanze_nuove"] += 1
        else:
            rid = room["id"]
            res["stanze_esistenti"] += 1
            if scriv and not room["scrivanie"]:
                db.update_stanza(rid, room["codice"], scriv, room["note"], room["come"])
                res["stanze_aggiornate"] += 1
        for p in r.get("persone") or []:
            p = p if isinstance(p, dict) else {"nome": p}
            nome = normalize_name(_clean(p.get("nome")))
            if not nome:
                continue
            uff = " ".join((_clean(p.get("ufficio")) or _clean(r.get("ufficio"))).split())
            user = db.find_user_like(nome)
            if user:
                uid = user["id"]
                res["persone_esistenti"] += 1
            else:
                uid, _ = db.add_user(nome, normalize_name(edificio["sede"]))
                res["persone_nuove"] += 1
            fill_person(uid, user, p, uff)
            posto = db.posto_of_user(uid)
            if posto and posto["stanza_id"] == rid:
                res["persone_gia_in_stanza"] += 1
            elif posto:
                res["persone_altrove"].append("%s: già in %s (%s), non spostato"
                                              % (nome, posto["codice"], floor_label(posto["piano"])))
            else:
                db.assegna_persona(uid, rid)
                res["persone_assegnate"] += 1

    for p in data.get("persone_senza_stanza") or []:       # persone che non siedono in nessuna stanza
        p = p if isinstance(p, dict) else {"nome": p}
        nome = normalize_name(_clean(p.get("nome")))
        if not nome:
            continue
        user = db.find_user_like(nome)
        if user:
            uid = user["id"]
        else:
            uid, _ = db.add_user(nome, normalize_name(edificio["sede"]))
            res["persone_senza_stanza"] += 1
        fill_person(uid, user, p, " ".join(_clean(p.get("ufficio")).split()))

    in_file = Counter()                  # quante stampanti con seriale ci sono nel file, per piano e modello
    for pr in data.get("stampanti") or []:
        if isinstance(pr, dict) and clean_serial(pr.get("seriale")):
            in_file[(to_int(pr.get("piano")), model_key(" ".join(_clean(pr.get("modello")).split())))] += 1
    adopted = set()

    def old_twin(piano, nome, modello, stanza_uid):
        """La vecchia stampante (senza seriale) che è sicuramente questa: stesso piano e modello, e
        o è l'unica da una parte e dall'altra, o ha lo stesso nome, o è nella stessa stanza.
        Se non è chiaro non indovina: ne resta una nuova e la si unisce a mano."""
        mk = model_key(modello)
        if not mk:
            return None, 0
        olds = [x for x in db.printers_of(edificio_id)
                if not (x["seriale"] or "").strip() and x["id"] not in adopted
                and x["piano"] == piano and model_key(x["modello"]) == mk]
        if not olds:
            return None, 0
        by_room = [x for x in olds if stanza_uid and x.get("stanza_uid") == stanza_uid]
        by_name = [x for x in olds if nome and x["nome"].strip().casefold() == nome.casefold()]
        for grp in (by_room, by_name):
            if len(grp) == 1:
                return grp[0], len(olds)
        if len(olds) == 1 and in_file[(piano, mk)] == 1:
            return olds[0], 1
        return None, len(olds)

    for pr in data.get("stampanti") or []:
        if not isinstance(pr, dict):
            continue
        piano = to_int(pr.get("piano"))
        nome, modello = " ".join(_clean(pr.get("nome")).split()), " ".join(_clean(pr.get("modello")).split())
        seriale = clean_serial(pr.get("seriale"))
        if not (nome or modello or seriale):
            continue
        stanza_uid = ""
        if _clean(pr.get("stanza")):
            sr = db.find_stanza_by_code(edificio_id, piano, _clean(pr.get("stanza")))
            if sr:
                stanza_uid = sr["uid"]
            else:
                res["scartati"].append("stampante %s: la stanza %s non esiste al piano %d"
                                       % (seriale or nome or modello, _clean(pr.get("stanza")), piano))
        fuori = bool(pr.get("fuori_servizio"))
        cur = None
        if seriale:
            other = db.printer_with_serial(seriale)
            cur = db.get_stampante(other["id"]) if other else None
            if cur is None:
                twin, n_old = old_twin(piano, nome, modello, stanza_uid)
                if twin:                   # è la vecchia stampante: prende il seriale e il nome del file
                    adopted.add(twin["id"])
                    cur0 = db.get_stampante(twin["id"])
                    db.update_stampante(cur0["id"], cur0["piano"], nome or cur0["nome"],
                                        modello or cur0["modello"], cur0["ip"], cur0["note"],
                                        seriale=seriale, fuori_servizio=True if fuori else None,
                                        stanza_uid=stanza_uid if (stanza_uid and not cur0.get("stanza_uid")) else None)
                    res["stampanti_adottate"] += 1
                    continue
                if n_old:
                    res["stampanti_da_unire"] += 1
        else:
            for x in db.printers_of(edificio_id):
                if x["piano"] == piano and x["nome"].casefold() == nome.casefold() \
                        and model_key(x["modello"]) == model_key(modello):
                    cur = db.get_stampante(x["id"])
                    break
        if cur is None:
            db.add_stampante(edificio_id, piano, nome, modello, _clean(pr.get("ip")),
                             _clean(pr.get("note")), seriale=seriale, fuori_servizio=fuori,
                             stanza_uid=stanza_uid)
            res["stampanti_nuove"] += 1
            continue
        new_stanza = stanza_uid if (stanza_uid and not cur.get("stanza_uid")) else None
        new_modello = modello if (modello and not cur["modello"]) else cur["modello"]
        new_fuori = True if (fuori and not cur["fuori_servizio"]) else None
        if new_stanza is None and new_fuori is None and new_modello == cur["modello"]:
            res["stampanti_gia_presenti"] += 1
            continue
        db.update_stampante(cur["id"], cur["piano"], cur["nome"], new_modello, cur["ip"],
                            cur["note"], fuori_servizio=new_fuori, stanza_uid=new_stanza)
        res["stampanti_aggiornate"] += 1
    return res


def build_phone_export(db, with_photos=False, settings=None):
    """Dati del PC per il telefono: utenti e interventi (con le foto, se richiesto).
    Gli interventi nati sul PC ricevono un identificativo (uid), così quando il telefono li
    rimanda indietro il programma li riconosce e non li duplica."""
    now = now_utc_iso()
    db.ensure_sync_ids()
    users = [{k: u[k] for k in ("nome", "matricola", "sede", "telefono", "seriale_pc", "imei",
                                "ufficio", "dotazione", "ordine", "email", "modificato")}
             for u in db.users_with_counts()]
    rows = db.conn.execute(
        "SELECT i.id, i.uid, i.data_ora, u.nome AS utente, i.sede, i.richiesta, i.soluzione, "
        "i.note, i.ticket, i.stato, i.modificato FROM interventi i "
        "JOIN utenti u ON u.id = i.utente_id ORDER BY i.data_ora, i.id").fetchall()
    items, n_foto = [], 0
    for r in rows:
        item = {k: r[k] for k in ("uid", "data_ora", "utente", "sede", "richiesta", "soluzione",
                                  "note", "ticket", "stato", "modificato")}
        item["foto"] = []
        if with_photos:
            for a in db.allegati_of(r["id"]):
                p = os.path.join(attachments_root(), a["file"])
                data = photo_data_url(p) if os.path.isfile(p) else None
                if data:
                    item["foto"].append({"nome": a["nome"], "data": data})
        n_foto += len(item["foto"])
        items.append(item)
    edifici = [{k: r[k] for k in ("uid", "nome", "sede", "piano_min", "piano_max")}
               for r in db.conn.execute("SELECT * FROM edifici ORDER BY id").fetchall()]
    stampanti = [{k: r[k] for k in ("uid", "edificio_uid", "piano", "nome", "modello", "ip", "note",
                                    "seriale", "fuori_servizio", "stanza_uid")}
                 for r in db.conn.execute(
                     "SELECT s.uid, e.uid AS edificio_uid, s.piano, s.nome, s.modello, s.ip, "
                     "s.note, s.seriale, s.fuori_servizio, s.stanza_uid FROM stampanti s "
                     "JOIN edifici e ON e.id = s.edificio_id ORDER BY s.id").fetchall()]
    toner = [{k: r[k] for k in ("uid", "stampante_uid", "colore", "data_ora", "note")}
             for r in db.conn.execute(
                 "SELECT t.uid, s.uid AS stampante_uid, t.colore, t.data_ora, t.note "
                 "FROM toner t JOIN stampanti s ON s.id = t.stampante_id ORDER BY t.id").fetchall()]
    manutenzioni = [{k: r[k] for k in ("uid", "stampante_uid", "data_ora", "data_tecnico",
                                       "ticket", "tecnico", "problema", "intervento", "stato",
                                       "modificato")}
                    for r in db.conn.execute(
                        "SELECT m.uid, s.uid AS stampante_uid, m.data_ora, m.data_tecnico, "
                        "m.ticket, m.tecnico, m.problema, m.intervento, m.stato, m.modificato "
                        "FROM manutenzioni m JOIN stampanti s ON s.id = m.stampante_id "
                        "ORDER BY m.id").fetchall()]
    verbali, n_foto_v = verbali_for_phone(db, with_photos)
    piani, n_foto_s = stanze_for_phone(db, with_photos)
    bolle, n_foto_b = bolle_for_phone(db, with_photos)
    out = {"tipo": "helpdesk-pc", "versione": 7, "creato": now, "utenti": users,
           "interventi": items, "edifici": edifici, "stampanti": stampanti,
           "stampanti_unite": db.printer_merges(), "utenti_uniti": db.user_merges(),
           "toner": toner, "bolle": bolle, "manutenzioni": manutenzioni, "scorta": db.scorta_all(),
           "scorta_aggiornata": now, "verbali": verbali,
           "impostazioni_verbali": {k: _clean((settings or {}).get(k))
                                    for k in ("tecnico", "sede_citta", "indirizzo", "cap")}}
    out.update(piani)
    return out, n_foto + n_foto_v + n_foto_s + n_foto_b


FOTO_STANZA_PX = 1600
MARK_H = 24


def _reduced_jpeg(src, dest, max_px):
    """Legge un'immagine, la raddrizza, la riduce (lato lungo max_px) e la salva come JPEG."""
    try:
        from PIL import Image, ImageOps
        with Image.open(src) as im:
            im = ImageOps.exif_transpose(im)
            im.thumbnail((max_px, max_px))
            im.convert("RGB").save(dest, "JPEG", quality=85)
    except Exception:
        raise ValueError("Non riesco a leggere questa immagine: %s" % os.path.basename(src))


def save_room_photo(db, stanza_id, src, tipo="foto"):
    """Aggiunge alla stanza una foto o un panorama (ridotti: 1600 px, il panorama 2600 px)."""
    room = db.get_stanza(stanza_id)
    if not room:
        raise ValueError("Questa stanza non c'è più.")
    tipo = tipo if tipo in ("foto", "panorama") else "foto"
    folder = os.path.join(stanze_foto_root(), room["uid"])
    os.makedirs(folder, exist_ok=True)
    stem = os.path.splitext(safe_photo_name(os.path.basename(src)))[0] or tipo
    dest = unique_path(folder, stem, ".jpg")
    _reduced_jpeg(src, dest, PANORAMA_PX if tipo == "panorama" else FOTO_STANZA_PX)
    name = os.path.basename(dest)
    return db.add_foto_stanza(stanza_id, name, tipo, "%s/%s" % (room["uid"], name))


def save_plan_photo(db, edificio_id, piano, src):
    """Salva (o sostituisce) la foto della pianta di un piano."""
    folder = os.path.join(piante_root(), "%d_%d" % (edificio_id, piano))
    os.makedirs(folder, exist_ok=True)
    name = "pianta_%s.jpg" % datetime.now().strftime("%Y%m%d%H%M%S%f")
    _reduced_jpeg(src, os.path.join(folder, name), PIANTA_PX)
    db.set_pianta(edificio_id, piano, "%d_%d/%s" % (edificio_id, piano, name))


def plan_image_b64(path, max_w):
    """La foto della pianta, della larghezza giusta, come PNG in base64 per Tk: (b64, larghezza, altezza)."""
    try:
        from PIL import Image, ImageOps
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im).convert("RGB")
            k = min(1.0, float(max_w) / im.width)
            if k < 1.0:
                im = im.resize((max(1, int(im.width * k)), max(1, int(im.height * k))))
            buf = io.BytesIO()
            im.save(buf, "PNG")
        return base64.b64encode(buf.getvalue()).decode("ascii"), im.width, im.height
    except Exception:
        return None


def marker_width(label):
    return max(46, 8 * len(label) + 14)


def plan_markers(rooms, width, height):
    """I segni delle stanze sulla foto della pianta (solo quelle già segnate), con le misure in pixel."""
    out = []
    for r in rooms:
        if r.get("px") is None or r.get("py") is None:
            continue
        label = room_label(r)
        out.append({"id": r["id"], "label": label, "x": r["px"] * width, "y": r["py"] * height,
                    "w": marker_width(label), "h": MARK_H, "over": room_occupancy(r)[1]})
    return out


def plan_hit(markers, x, y):
    """L'id della stanza il cui segno si trova nel punto (x, y), oppure None."""
    for m in reversed(markers):
        if abs(x - m["x"]) <= m["w"] / 2.0 and abs(y - m["y"]) <= m["h"] / 2.0:
            return m["id"]
    return None


def plan_click_fraction(x, y, width, height):
    """Il punto cliccato sulla foto come frazione 0..1 (None se è fuori dalla foto)."""
    if width <= 0 or height <= 0 or not (0 <= x <= width and 0 <= y <= height):
        return None
    return round(x / float(width), 4), round(y / float(height), 4)


def grid_tiles(rooms, width, tile_w=120, tile_h=62, gap=10, pad=10):
    """Lo schema del piano: una casella per stanza, in ordine, a righe. Restituisce (caselle, altezza)."""
    per_row = max(1, int((width - 2 * pad + gap) // (tile_w + gap)))
    tiles = []
    for i, r in enumerate(rooms):
        x0 = pad + (i % per_row) * (tile_w + gap)
        y0 = pad + (i // per_row) * (tile_h + gap)
        occ, over = room_occupancy(r)
        tiles.append({"id": r["id"], "label": room_label(r), "occ": occ, "over": over,
                      "x0": x0, "y0": y0, "x1": x0 + tile_w, "y1": y0 + tile_h})
    rows = (len(rooms) + per_row - 1) // per_row
    height = pad * 2 + rows * tile_h + max(0, rows - 1) * gap if rows else pad * 2 + 40
    return tiles, height


def tile_hit(tiles, x, y):
    for t in tiles:
        if t["x0"] <= x <= t["x1"] and t["y0"] <= y <= t["y1"]:
            return t["id"]
    return None


# ---- vista a scrivanie: una scrivania disegnata per ogni postazione, col PC e il suo S/N ----
DESK_W, DESK_H, DESK_GAP, DESK_PAD, DESK_LABEL_W = 150, 170, 10, 10, 104


# Il telefono che si vede sulla scrivania viene dalla dotazione della persona: un «Cellulare» (o
# «Smartphone») Samsung Galaxy S24, S24 Plus o S25 e/o un «Telefono Teams» (AudioCodes) da scrivania.
CELL_MODELLI = ("S24", "S24 Plus", "S25")
# nome normalizzato -> (tipo di disegno, nome mostrato); S4, S4 Plus e S5 sono i nomi scritti per
# errore all'inizio: valgono come S24, S24 Plus e S25
CELL_KEYS = {"s24": ("s24", "S24"), "s24plus": ("s24plus", "S24 Plus"), "s25": ("s25", "S25"),
             "s4": ("s24", "S24"), "s4plus": ("s24plus", "S24 Plus"), "s5": ("s25", "S25")}
TEAMS_TIPO, TEAMS_MODELLO = "Telefono Teams", "AudioCodes C435HD"
PRESA_TIPO = "Presa di rete"      # riga della dotazione: seriale = codice della presa, modello = porta dello switch
LAPTOP_MODELLI = {"13": "Dell 13 pollici", "15": "Dell 15 pollici"}


def _norm_model(txt):
    t = str(txt or "").lower().replace("+", " plus ")
    for w in ("samsung", "galaxy"):
        t = t.replace(w, " ")
    return "".join(ch for ch in t if ch.isalnum())


def _row_kind(r):
    """'presa' (presa di rete a muro), 'teams' (telefono Teams da scrivania), 'mobile' (cellulare /
    smartphone), 'laptop' (notebook / portatile), 'monitor' (il monitor AOC) oppure ''."""
    tipo = str(r.get("tipo") or "").lower()
    both = tipo + " " + str(r.get("modello") or "").lower()
    if tipo.strip().startswith("presa"):
        return "presa"
    if "teams" in both or "audiocodes" in both:
        return "teams"
    if "cellulare" in tipo or "smartphone" in tipo:
        return "mobile"
    if "notebook" in tipo or "portatile" in tipo or "laptop" in tipo:
        return "laptop"
    if "monitor" in tipo:
        return "monitor"
    return ""


def _laptop_size(modello):
    m = re.search(r"(?<!\d)(13|15)(?!\d)", str(modello or ""))
    return m.group(1) if m else ""


def desk_phone_info(dotazione, imei=""):
    """Cosa disegnare sulla scrivania: {"mobile": None|"s24"|"s24plus"|"s25"|"altro", "label": nome del
    modello, "teams": bool, "laptop": None|"13"|"15"|"altro", "laptop_label": modello del portatile}.
    Chi ha un IMEI ma nessun modello indicato ha il cellulare classico (S24)."""
    mobile, label, teams, laptop, llabel = None, "", False, None, ""
    for r in clean_dotazione(dotazione):
        k = _row_kind(r)
        if k == "teams":
            teams = True
        elif k == "laptop" and laptop is None and r["modello"]:
            laptop, llabel = _laptop_size(r["modello"]) or "altro", r["modello"]
        elif k == "mobile" and mobile is None:
            n = _norm_model(r["modello"])
            mobile, label = CELL_KEYS.get(n, ("altro", r["modello"]))
    if mobile is None and str(imei or "").strip():
        mobile = "s24"
    return {"mobile": mobile, "label": label, "teams": teams, "laptop": laptop,
            "laptop_label": llabel}


def with_desk_phones(dotazione, mobile="", teams=False, laptop=None):
    """La dotazione con il cellulare scelto (S24, S24 Plus, S25: '' = non indicato, None = lascia
    com'è), il telefono Teams da scrivania e il portatile Dell ('13' o '15' pollici, '' = non
    indicato, None = lascia com'è); il resto della dotazione e i seriali già scritti restano."""
    rows = clean_dotazione(dotazione)
    old_mob = next((r for r in rows if _row_kind(r) == "mobile"), None)
    old_teams = next((r for r in rows if _row_kind(r) == "teams"), None)
    laps = [r for r in rows if _row_kind(r) == "laptop"]
    old_lap = laps[0] if laps else None
    out = [r for r in rows if _row_kind(r) in ("", "monitor", "presa")] + laps[1:]   # monitor AOC e presa restano
    if laptop is None:
        out += laps[:1]
    elif laptop in LAPTOP_MODELLI:
        out.append({"tipo": old_lap["tipo"] if old_lap else "Notebook",
                    "modello": LAPTOP_MODELLI[laptop], "seriale": old_lap["seriale"] if old_lap else ""})
    elif old_lap and old_lap["seriale"]:       # «non indicato»: il seriale scritto non si perde
        out.append(dict(old_lap, modello=""))
    if mobile is None:
        out += [r for r in rows if _row_kind(r) == "mobile"]
    elif mobile:
        out.append({"tipo": old_mob["tipo"] if old_mob else "Cellulare", "modello": mobile,
                    "seriale": old_mob["seriale"] if old_mob else ""})
    if teams:
        out.append(old_teams or {"tipo": TEAMS_TIPO, "modello": TEAMS_MODELLO, "seriale": ""})
    return clean_dotazione(out)


# ---- i dispositivi di una scrivania: PC, monitor AOC, cellulare (IMEI) e telefono Teams ----
# Ognuno ha il suo bollino (verde = codice scritto, rosso = manca): si clicca e si scrive il codice a
# mano oppure si legge con la fotocamera. Posizione del bollino nella scena 130 x 66.
DESK_BOLLINI_POS = {"pc": (36, 50), "monitor": (65, 19), "mobile": (113, 50), "teams": (16, 58),
                    "presa": (113, 33)}
DESK_BIG = 5                    # la scrivania a pagina intera: la scena ingrandita 5 volte
DESK_BOLLINO_R = 18             # raggio del bollino a pagina intera (px)


def desk_devices(person):
    """I dispositivi della scrivania di una persona (dizionario con seriale_pc, imei, dotazione):
    [{"key": "pc"|"monitor"|"mobile"|"teams"|"presa", "kind": "serial"|"imei"|"presa", "name", "sub",
    "label", "value", "pos"}]. Il PC, il monitor e la presa di rete ci sono sempre; il cellulare solo
    se c'è sulla scrivania (modello scelto o IMEI scritto), il telefono Teams solo se è stato aggiunto."""
    dot = clean_dotazione(person.get("dotazione"))
    ph = desk_phone_info(dot, person.get("imei"))
    mon = next((r for r in dot if _row_kind(r) == "monitor"), None)
    tm = next((r for r in dot if _row_kind(r) == "teams"), None)
    out = [{"key": "pc", "kind": "serial", "name": "Computer", "sub": ph["laptop_label"],
            "label": "S/N del computer", "value": (person.get("seriale_pc") or "").strip()},
           {"key": "monitor", "kind": "serial", "name": "Monitor AOC",
            "sub": mon["modello"] if mon and mon["modello"] and mon["modello"].lower() != "aoc" else "",
            "label": "S/N del monitor AOC", "value": mon["seriale"] if mon else ""}]
    if ph["mobile"]:
        out.append({"key": "mobile", "kind": "imei", "name": "Cellulare",
                    "sub": ph["label"] if ph["mobile"] == "altro" else "Samsung Galaxy " + ph["label"],
                    "label": "IMEI del cellulare", "value": str(person.get("imei") or "").strip()})
    if ph["teams"]:
        out.append({"key": "teams", "kind": "serial", "name": "Telefono Teams", "sub": TEAMS_MODELLO,
                    "label": "S/N del telefono Teams", "value": tm["seriale"] if tm else ""})
    presa = next((r for r in dot if _row_kind(r) == "presa"), None)
    out.append({"key": "presa", "kind": "presa", "name": PRESA_TIPO,
                "sub": presa["modello"] if presa else "",
                "label": "Codice della presa di rete", "value": presa["seriale"] if presa else "",
                "port": presa["modello"] if presa else ""})
    for d in out:
        d["pos"] = DESK_BOLLINI_POS[d["key"]]
    return out


def with_device_serial(dotazione, key, value, port=None):
    """La dotazione con il seriale del monitor ("monitor") o del telefono Teams ("teams") cambiato;
    se la riga non c'è si crea (Monitor AOC, Telefono Teams AudioCodes). Per la presa di rete
    ("presa") `value` è il codice della presa e `port` la porta dello switch (None = lascia com'è);
    senza codice e senza porta la riga sparisce. Il resto non si tocca."""
    rows = [dict(r) for r in clean_dotazione(dotazione)]
    if key == "presa":
        row = next((r for r in rows if _row_kind(r) == "presa"), None)
        code = clean_serial(value)
        porta = (" ".join(str(port or "").split())[:80] if port is not None
                 else (row["modello"] if row else ""))
        rows = [r for r in rows if _row_kind(r) != "presa"]
        if code or porta:
            rows.append({"tipo": PRESA_TIPO, "modello": porta, "seriale": code})
        return clean_dotazione(rows)
    kind = "teams" if key == "teams" else "monitor"
    row = next((r for r in rows if _row_kind(r) == kind), None)
    if row is not None:
        row["seriale"] = clean_serial(value)
    elif value:
        rows.append({"tipo": TEAMS_TIPO, "modello": TEAMS_MODELLO, "seriale": clean_serial(value)}
                    if kind == "teams" else {"tipo": "Monitor", "modello": "AOC",
                                             "seriale": clean_serial(value)})
    return clean_dotazione(rows)


def _luhn_ok(s):
    total = 0
    for i, ch in enumerate(reversed(s)):
        d = int(ch)
        if i % 2 == 1:
            d = d * 2 - 9 if d > 4 else d * 2
        total += d
    return total % 10 == 0


PRESA_MAX = 30       # lunghezza massima del codice di una presa di rete


def check_device_code(kind, text):
    """Controlla un codice scritto a mano o letto con la fotocamera. Restituisce (valore, errore):
    il S/N senza spazi e in maiuscolo (da 4 a 40 caratteri), l'IMEI di 15 cifre con la cifra di
    controllo giusta; testo vuoto = codice cancellato ("", ""). Se non va: (None, spiegazione)."""
    raw = "".join(ch for ch in str(text or "") if ch >= " " and ch != "\x7f")
    if kind == "imei":
        t = re.sub(r"[\s\-]", "", re.sub(r"^\s*imei[\s:\-]*", "", raw, flags=re.I))
        if not t:
            return "", ""
        if not re.fullmatch(r"\d{15}", t):
            return None, ("L'IMEI ha 15 cifre: ne hai scritte %d. Controlla di non aver saltato un "
                          "numero." % len(t))
        if not _luhn_ok(t):
            return None, ("Questo IMEI non torna con la cifra di controllo: c'è un numero sbagliato. "
                          "Riguardalo, oppure leggilo con la fotocamera.")
        return t, ""
    if kind == "presa":
        t = " ".join(raw.split()).upper()
        if not t:
            return "", ""
        if len(t) > PRESA_MAX:
            return None, "Il codice della presa può avere al massimo %d caratteri (ne hai scritti %d)." % (
                PRESA_MAX, len(t))
        if not re.fullmatch(r"[A-Z0-9][A-Z0-9 \-./_]*", t):
            return None, ("Il codice della presa può avere lettere, numeri e i segni - . / _ "
                          "(per esempio A12-34).")
        return t, ""
    t = re.sub(r"\s+", "", raw).upper()
    if not t:
        return "", ""
    if not 4 <= len(t) <= 40:
        return None, "Il S/N deve avere da 4 a 40 caratteri (ne hai scritti %d)." % len(t)
    return t, ""


def set_device_code(db, utente_id, key, value, dotazione=None, port=None):
    """Scrive il codice di un dispositivo: S/N del PC o IMEI nei dati della persona, S/N del monitor,
    del telefono Teams o codice della presa di rete (con la porta dello switch, `port`) nella sua
    dotazione (`dotazione` = quella attuale della persona). Restituisce True se qualcosa è cambiato."""
    if key == "pc":
        return db.set_user_codes(utente_id, seriale_pc=value)
    if key == "mobile":
        return db.set_user_codes(utente_id, imei=value)
    return db.set_user_extra(utente_id, dotazione=with_device_serial(dotazione, key, value, port))


def desk_bollini(devices, cx, y, scale=DESK_BIG):
    """Dove cadono i bollini nella scena ingrandita (centrata in cx, bordo alto a y):
    [(chiave, x, y)]."""
    return [(d["key"], cx - 65 * scale + d["pos"][0] * scale, y + d["pos"][1] * scale)
            for d in devices]


def desk_bollino_hit(devices, cx, y, px, py, scale=DESK_BIG, r=DESK_BOLLINO_R):
    """La chiave del dispositivo il cui bollino è sotto il clic (px, py), oppure None."""
    best, dist = None, (r + 4) ** 2
    for key, bx, by in desk_bollini(devices, cx, y, scale):
        d = (bx - px) ** 2 + (by - py) ** 2
        if d <= dist:
            best, dist = key, d
    return best


def desk_cards(room, people):
    """Le scrivanie di una stanza: prima chi ci siede (in ordine di postazione), poi quelle libere
    (solo se le «scrivanie totali» sono indicate). Ogni scrivania è un dizionario."""
    cards = []
    for p in people:
        cards.append({"kind": "occ", "numero": p["numero"], "utente_id": p["utente_id"],
                      "nome": p["nome"], "sn": (p.get("seriale_pc") or "").strip(),
                      "extra": len([r for r in clean_dotazione(p.get("dotazione"))
                                    if not _row_kind(r)]),
                      "phones": desk_phone_info(p.get("dotazione"), p.get("imei"))})
    libere = max(0, (room.get("scrivanie") or 0) - len(cards))
    for _ in range(libere):
        cards.append({"kind": "free", "numero": len(cards) + 1, "utente_id": None, "nome": "",
                      "sn": "", "extra": 0,
                      "phones": {"mobile": None, "label": "", "teams": False, "laptop": None,
                                 "laptop_label": ""}})
    return cards


def desk_layout(rooms, people_by_room, width, y0=0, show_label=True):
    """Dispone le scrivanie: una riga per stanza (come i piani nella vista Edificio), con la
    scritta della stanza a sinistra e le scrivanie che vanno a capo. Restituisce
    {"rows": [...], "height": ...}; ogni riga ha "desks" con le coordinate x0, y0, x1, y1."""
    label_w = DESK_LABEL_W if show_label else 0
    x_start = DESK_PAD + label_w + (DESK_PAD if show_label else 0)
    per_line = max(1, int((width - x_start - DESK_PAD + DESK_GAP) // (DESK_W + DESK_GAP)))
    rows, y = [], y0
    for r in rooms:
        cards = desk_cards(r, people_by_room.get(r["id"], []))
        lines = max(1, (len(cards) + per_line - 1) // per_line)
        h = (DESK_PAD * 2 + lines * DESK_H + (lines - 1) * DESK_GAP) if cards else 64
        for i, c in enumerate(cards):
            c["x0"] = x_start + (i % per_line) * (DESK_W + DESK_GAP)
            c["y0"] = y + DESK_PAD + (i // per_line) * (DESK_H + DESK_GAP)
            c["x1"], c["y1"] = c["x0"] + DESK_W, c["y0"] + DESK_H
        occ, over = room_occupancy(r)
        rows.append({"room_id": r["id"], "codice": r["codice"], "occ": occ, "over": over,
                     "n": len(cards), "x0": 0, "y0": y, "x1": width, "y1": y + h,
                     "label_x1": DESK_PAD + label_w, "desks": cards})
        y += h
    return {"rows": rows, "height": y - y0, "per_line": per_line}


def desk_hit(layout, x, y):
    """Cosa c'è sotto il clic: ("occ", id stanza, id utente), ("free", id stanza, None),
    ("room", id stanza, None) per la scritta o lo spazio vuoto della riga, oppure None."""
    for row in layout["rows"]:
        if not (row["y0"] <= y < row["y1"]):
            continue
        for d in row["desks"]:
            if d["x0"] <= x <= d["x1"] and d["y0"] <= y <= d["y1"]:
                return d["kind"], row["room_id"], d.get("utente_id")
        return "room", row["room_id"], None
    return None


def _mix(color, bg, a):
    """Schiarisce un colore #rrggbb verso lo sfondo (per le scrivanie libere)."""
    c1 = [int(color[k:k + 2], 16) for k in (1, 3, 5)]
    c2 = [int(bg[k:k + 2], 16) for k in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x * (1 - a) + y * a) for x, y in zip(c1, c2))


# misure (larghezza, altezza) delle foto nella scena 130 x 66; le stesse sono in desk_images.py
DESK_IMG_SIZE = {"lap13": (26, 18), "lap15": (32, 22), "s24": (10, 20), "s24plus": (10, 22),
                 "s25": (10, 20), "teams": (32, 24)}
_DESK_PHOTOS = {}


def desk_photo(key):
    """La foto (PhotoImage) «key» di desk_images, per esempio «s24_s» (scena) o «s24_m»
    (anteprima); None se manca."""
    if key not in _DESK_PHOTOS:
        data = DESK_IMAGES.get(key)
        try:
            _DESK_PHOTOS[key] = tk.PhotoImage(data=data) if data else None
        except Exception:
            _DESK_PHOTOS[key] = None
    return _DESK_PHOTOS[key]


def desk_photo_keys(phones):
    """Le foto che servono per una scrivania: {"laptop": chiave|None, "mobile": ..., "teams": ...}."""
    ph = phones or {}
    return {"laptop": {"13": "lap13", "15": "lap15"}.get(ph.get("laptop")),
            "mobile": ph["mobile"] if ph.get("mobile") in ("s24", "s24plus", "s25") else None,
            "teams": "teams" if ph.get("teams") else None}


def draw_desk_scene(c, cx, y, tags, phones=None, faded=False, bg="#f6f2e6", scale=1, big=False):
    """La scrivania vista da davanti, come nelle foto: pannello azzurro, monitor con la webcam,
    tastiera, mouse e, con le FOTO dei dispositivi scelti, il portatile Dell (13 o 15 pollici),
    il cellulare Galaxy (S24, S24 Plus, S25) a destra e il telefono Teams AudioCodes a sinistra.
    Occupa 130 x 66 px (per `scale`) attorno a cx. Senza le foto (file mancante) disegna dei
    rettangoli. scale > 1: la scena ingrandita (con le foto «_l»); big: la scrivania a pagina
    intera, con la scritta AOC sul monitor e sempre un PC (coperchio generico se non c'è il Dell)."""
    ph = phones or {}
    k = scale
    ox = cx - 65 * k
    suf = "_l" if k > 1 else "_s"

    def col(h):
        return _mix(h, bg, 0.62) if faded else h

    def rect(x, yy, w, h, fill):
        c.create_rectangle(ox + x * k, y + yy * k, ox + (x + w) * k, y + (yy + h) * k,
                           fill=col(fill), outline=col(fill), tags=tags)

    def oval(x, yy, w, h, fill):
        c.create_oval(ox + x * k, y + yy * k, ox + (x + w) * k, y + (yy + h) * k, fill=col(fill),
                      outline=col(fill), tags=tags)

    def photo(key, x, yy, fallback):
        """La foto «key» con l'angolo in alto a sinistra in (x, yy); se manca un rettangolo."""
        w, h = DESK_IMG_SIZE[key]
        img = desk_photo(key + suf)
        if img is None:
            rect(x, yy, w, h, fallback)
            return
        c.create_image(ox + x * k, y + yy * k, image=img, anchor="nw",
                       tags=tuple(tags) + ("dkimg:" + key,))
    rect(0, 40, 130, 26, "#ecd3a8")          # piano della scrivania
    rect(2, 28, 126, 14, "#9db8e6")          # pannello azzurro dietro il monitor
    rect(54, 41, 22, 3, "#15181d")           # base del monitor
    rect(62, 35, 6, 6, "#15181d")            # piedistallo
    rect(42, 6, 46, 30, "#15181d")           # monitor
    rect(44.5, 8.5, 41, 22, "#2b313b")       # schermo spento
    rect(60, 3, 10, 3.5, "#0b0d10")          # webcam
    rect(44, 48, 40, 8, "#1d2026")           # tastiera
    for i in range(3):
        rect(47, 50 + i * 2.2, 34, 0.9, "#3a3f49")
    oval(88.5, 47, 7, 9, "#1d2026")          # mouse
    if big:
        c.create_text(ox + 65 * k, y + 34.3 * k, text="AOC", fill="#c9ced6",
                      font=("Segoe UI", 9, "bold"), tags=tags)
        rect(108, 29, 10, 9, "#f4f6f8")      # presa di rete a muro: placca con la porta
        rect(111, 32.5, 4, 3, "#2b313b")
        rect(112.2, 31.6, 1.6, 1, "#2b313b")
    keys = desk_photo_keys(ph)
    if keys["laptop"]:                       # portatile Dell chiuso, visto dall'alto (centro a x 27)
        w, h = DESK_IMG_SIZE[keys["laptop"]]
        photo(keys["laptop"], 27 - w / 2, 63 - h, "#a9b2bd")
    elif ph.get("laptop") == "altro" or big:  # un altro modello (o PC senza modello): coperchio generico
        rect(13, 44, 28, 19, "#a9b2bd")
        if big and ph.get("laptop") != "altro":
            c.create_text(ox + 21 * k, y + 59 * k, text="PC", fill="#5b6b7d",
                          font=("Segoe UI", 12, "bold"), tags=tags)
    if keys["teams"]:                        # telefono Teams (AudioCodes), davanti, a sinistra
        photo("teams", 0, 66 - DESK_IMG_SIZE["teams"][1], "#15181d")
    if keys["mobile"]:                       # cellulare Samsung Galaxy (centro a x 113)
        w, h = DESK_IMG_SIZE[keys["mobile"]]
        photo(keys["mobile"], 113 - w / 2, 62 - h, "#14171c")
    elif ph.get("mobile") == "altro":        # un altro modello: cellulare generico
        rect(108.5, 45, 9, 17, "#14171c")
        rect(109.5, 46.5, 7, 14, "#2b313b")


def draw_desk_big(c, phones, devices, cx, y):
    """La scrivania a pagina intera: la scena ingrandita 5 volte con un bollino su ogni dispositivo
    (verde con la spunta = codice scritto, rosso col + = da inserire)."""
    draw_desk_scene(c, cx, y, ("dkbig",), phones=phones, scale=DESK_BIG, big=True)
    for key, bx, by in desk_bollini(devices, cx, y):
        ok = bool(next(d for d in devices if d["key"] == key)["value"])
        c.create_oval(bx - DESK_BOLLINO_R, by - DESK_BOLLINO_R, bx + DESK_BOLLINO_R,
                      by + DESK_BOLLINO_R, fill=GREEN if ok else RED, outline=WHITE, width=3,
                      tags=("bollino", "bol:" + key))
        c.create_text(bx, by, text="\u2713" if ok else "+", fill=WHITE,
                      font=("Segoe UI", 16, "bold"), tags=("bollino", "bol:" + key))


def draw_desk_rows(c, layout, tag_prefix="dk", cursor=True):
    """Disegna nel Canvas le righe di desk_layout: scritta della stanza e scrivanie."""
    for idx, row in enumerate(layout["rows"]):
        top, bottom = row["y0"], row["y1"]
        c.create_rectangle(0, top, row["x1"], bottom, fill="#efe7cf" if idx % 2 == 0 else "#e7ddc0",
                           outline="#cfc3a3")
        if row["label_x1"] > DESK_PAD:
            c.create_rectangle(0, top, row["label_x1"], bottom, fill=RED if row["over"] else NAVY,
                               outline=RED if row["over"] else NAVY)
            c.create_text(row["label_x1"] / 2, (top + bottom) / 2 - (8 if row["occ"] else 0),
                          text=row["codice"], fill=WHITE, font=("Segoe UI", 12, "bold"),
                          width=row["label_x1"] - 10)
            if row["occ"]:
                c.create_text(row["label_x1"] / 2, (top + bottom) / 2 + 14, text=row["occ"],
                              fill=WHITE, font=("Segoe UI", 8))
        if not row["desks"]:
            c.create_text(row["label_x1"] + 14, (top + bottom) / 2, anchor="w", fill=GREY,
                          font=("Segoe UI", 9),
                          text="Nessuna postazione: aggiungi una persona o indica le scrivanie totali.")
        for d in row["desks"]:
            free = d["kind"] == "free"
            tag = "%s%d_%d" % (tag_prefix, row["room_id"], d["numero"])
            c.create_rectangle(d["x0"], d["y0"], d["x1"], d["y1"], fill="#f6f2e6" if free else WHITE,
                               outline="#cfc3a3" if free else "#b9ae8c", width=1,
                               dash=(4, 3) if free else (), tags=(tag,))
            cx = (d["x0"] + d["x1"]) / 2
            c.create_text(cx, d["y0"] + 6, anchor="n", text="Postazione %d" % d["numero"],
                          fill=GREY, font=("Segoe UI", 8), tags=(tag,))
            draw_desk_scene(c, cx, d["y0"] + 22, (tag,), phones=d.get("phones"), faded=free)
            if free:
                c.create_text(cx, d["y0"] + 96, anchor="n", text="Libera", fill=GREY,
                              font=("Segoe UI", 9, "italic"), tags=(tag,))
            else:
                c.create_text(cx, d["y0"] + 92, anchor="n", text=d["nome"], fill=NAVY,
                              width=DESK_W - 10, justify="center", font=("Segoe UI", 9, "bold"),
                              tags=(tag,))
                if d["sn"]:
                    c.create_text(cx, d["y0"] + 114, anchor="n", text="S/N " + d["sn"], fill=TEXT,
                                  width=DESK_W - 8, justify="center", font=("Segoe UI", 8, "bold"),
                                  tags=(tag,))
                else:
                    c.create_text(cx, d["y0"] + 114, anchor="n", text="S/N non inserito", fill=GREY,
                                  font=("Segoe UI", 8, "italic"), tags=(tag,))
                ph = d.get("phones") or {}
                righe = []
                if ph.get("laptop_label"):
                    righe.append(ph["laptop_label"])
                tel = " · ".join(x for x in ((("Cell. " + ph["label"]) if ph.get("label") else ""),
                                             "Telefono Teams" if ph.get("teams") else "") if x)
                if tel:
                    righe.append(tel)
                if d["extra"]:
                    righe.append("+ %d %s" % (d["extra"], "apparecchio" if d["extra"] == 1
                                              else "apparecchi"))
                if righe:
                    c.create_text(cx, d["y0"] + 128, anchor="n", text="\n".join(righe), fill=GREY,
                                  width=DESK_W - 8, justify="center", font=("Segoe UI", 8),
                                  tags=(tag,))
            if cursor:
                c.tag_bind(tag, "<Enter>", lambda e, c=c: c.configure(cursor="hand2"))
                c.tag_bind(tag, "<Leave>", lambda e, c=c: c.configure(cursor=""))


def stanze_for_phone(db, with_photos):
    """Stanze, piante dei piani, foto/panorami delle stanze e postazioni per il telefono.
    Restituisce (dizionario da aggiungere al file, numero di foto)."""
    db.ensure_sync_ids()
    stanze, foto, piante, n_foto = [], [], [], 0
    for e in db.conn.execute("SELECT id, uid FROM edifici ORDER BY id").fetchall():
        for r in db.conn.execute("SELECT id FROM stanze WHERE edificio_id = ? ORDER BY id",
                                 (e["id"],)).fetchall():
            r = db.get_stanza(r["id"], anche_eliminata=True)
            stanze.append({"uid": r["uid"], "edificio_uid": e["uid"], "piano": r["piano"],
                           "codice": r["codice"], "scrivanie": r["scrivanie"], "note": r["note"],
                           "come": r["come"], "px": r["px"], "py": r["py"],
                           "eliminata": bool(r["eliminata"]), "modificato": r["modificato"]})
            for f in db.conn.execute(
                    "SELECT uid, nome, tipo, file, modificato, eliminata FROM stanze_foto "
                    "WHERE stanza_id = ? ORDER BY id", (r["id"],)).fetchall():
                item = {"uid": f["uid"], "stanza_uid": r["uid"], "nome": f["nome"],
                        "tipo": f["tipo"], "data": "", "eliminata": bool(f["eliminata"]),
                        "modificato": f["modificato"]}
                path = os.path.join(stanze_foto_root(), *f["file"].split("/"))
                if with_photos and not f["eliminata"] and os.path.isfile(path):
                    data = photo_data_url(path, PANORAMA_PX if f["tipo"] == "panorama" else None)
                    if data:
                        item["data"] = data
                        n_foto += 1
                foto.append(item)
    for r in db.conn.execute(
            "SELECT p.uid, p.piano, p.file, p.modificato, e.uid AS edificio_uid FROM piante p "
            "JOIN edifici e ON e.id = p.edificio_id ORDER BY p.id").fetchall():
        item = {"uid": r["uid"], "edificio_uid": r["edificio_uid"], "piano": r["piano"],
                "foto": "", "modificato": r["modificato"]}
        path = os.path.join(piante_root(), *r["file"].split("/")) if r["file"] else ""
        if with_photos and path and os.path.isfile(path):
            data = photo_data_url(path, PIANTA_PX)
            if data:
                item["foto"] = data
                n_foto += 1
        piante.append(item)
    postazioni = [{"utente": r["nome"], "stanza_uid": r["stanza_uid"] or "",
                   "postazione": r["numero"], "modificato": r["modificato"]}
                  for r in db.conn.execute(
                      "SELECT u.nome, s.uid AS stanza_uid, p.numero, p.modificato FROM postazioni p "
                      "JOIN utenti u ON u.id = p.utente_id LEFT JOIN stanze s ON s.id = p.stanza_id "
                      "ORDER BY u.nome COLLATE NOCASE").fetchall()]
    return ({"stanze": stanze, "stanze_foto": foto, "piante": piante, "postazioni": postazioni},
            n_foto)


def import_phone_stanze(db, data, res):
    """Stanze, piante, foto e postazioni arrivate dal telefono. Vince la modifica più recente.
    Se la stessa stanza (stesso edificio, piano e codice) è stata creata due volte con
    identificativi diversi, resta una sola stanza: quella già presente."""
    for k in ("stanze_nuove", "stanze_aggiornate", "stanze_gia_presenti", "stanze_saltate",
              "piante", "stanze_foto", "postazioni", "postazioni_saltate"):
        res.setdefault(k, 0)
    ed = {r["uid"]: r["id"] for r in db.conn.execute("SELECT id, uid FROM edifici")}
    room_of = {}                      # uid in arrivo -> id della stanza qui

    def to_float(v):
        try:
            return None if v is None else min(1.0, max(0.0, float(v)))
        except (TypeError, ValueError):
            return None

    for it in data.get("stanze") or []:
        if not isinstance(it, dict):
            continue
        uid, eid = _clean(it.get("uid")), ed.get(_clean(it.get("edificio_uid")))
        try:
            piano = int(it.get("piano"))
        except (TypeError, ValueError):
            piano = None
        codice = clean_room_code(_clean(it.get("codice")))
        if not uid or eid is None or piano is None or not codice:
            res["stanze_saltate"] += 1
            continue
        mod = _clean(it.get("modificato")) or now_utc_iso()
        px, py = to_float(it.get("px")), to_float(it.get("py"))
        if px is None or py is None:
            px = py = None
        try:
            scriv = max(0, int(it.get("scrivanie") or 0))
        except (TypeError, ValueError):
            scriv = 0
        gone = bool(it.get("eliminata"))
        ex = db.find_stanza_by_uid(uid, anche_eliminata=True) or \
            db.find_stanza_by_code(eid, piano, codice)
        if not ex:
            if gone:                      # eliminata altrove e qui non c'era: niente da fare
                continue
            room_of[uid] = db.add_stanza(eid, piano, codice, scriv, _clean(it.get("note")),
                                         _clean(it.get("come")), px, py, uid=uid, modificato=mod)
            res["stanze_nuove"] += 1
        else:
            if not ex["eliminata"]:
                room_of[uid] = ex["id"]
            if mod > (ex["modificato"] or ""):
                db.apply_stanza(ex["id"], codice, scriv, _clean(it.get("note")),
                                _clean(it.get("come")), px, py, mod, eliminata=gone)
                res["stanze_aggiornate"] += 1
                if gone:
                    room_of.pop(uid, None)
                else:
                    room_of[uid] = ex["id"]
            else:
                res["stanze_gia_presenti"] += 1

    def room_id(uid):
        if uid in room_of:
            return room_of[uid]
        r = db.find_stanza_by_uid(uid)
        return r["id"] if r else None

    for it in data.get("piante") or []:
        if not isinstance(it, dict):
            continue
        eid = ed.get(_clean(it.get("edificio_uid")))
        try:
            piano = int(it.get("piano"))
        except (TypeError, ValueError):
            continue
        payload, mod = _clean(it.get("foto")), _clean(it.get("modificato")) or now_utc_iso()
        cur = db.pianta_of(eid, piano) if eid is not None else None
        if eid is None or "," not in payload or (cur and not mod > (cur["modificato"] or "")):
            continue
        folder = os.path.join(piante_root(), "%d_%d" % (eid, piano))
        os.makedirs(folder, exist_ok=True)
        name = "pianta_%s.jpg" % re.sub(r"\D", "", mod)[:17]
        if _write_data_url(payload, os.path.join(folder, name), 25 * 1024 * 1024):
            db.set_pianta(eid, piano, "%d_%d/%s" % (eid, piano, name), modificato=mod)
            res["piante"] += 1

    for it in data.get("stanze_foto") or []:
        if not isinstance(it, dict):
            continue
        uid, payload = _clean(it.get("uid")), _clean(it.get("data"))
        have = db.conn.execute("SELECT id, tipo, modificato, eliminata FROM stanze_foto "
                               "WHERE uid = ?", (uid,)).fetchone() if uid else None
        mod = _clean(it.get("modificato")) or now_utc_iso()
        tipo = it.get("tipo") if it.get("tipo") in ("foto", "panorama") else "foto"
        if it.get("eliminata"):
            if have and not have["eliminata"]:
                db.delete_foto_stanza(have["id"], mod)
            continue
        sid = room_id(_clean(it.get("stanza_uid")))
        if sid is None or not uid or "," not in payload:
            continue
        if have:
            if not have["eliminata"] and tipo != have["tipo"] and mod > (have["modificato"] or ""):
                db.set_tipo_foto_stanza(have["id"], tipo)
            continue
        room = db.get_stanza(sid)
        folder = os.path.join(stanze_foto_root(), room["uid"])
        os.makedirs(folder, exist_ok=True)
        name = os.path.basename(unique_path(folder, os.path.splitext(
            safe_photo_name(it.get("nome")))[0], ".jpg"))
        if _write_data_url(payload, os.path.join(folder, name), 25 * 1024 * 1024):
            db.add_foto_stanza(sid, name, tipo, "%s/%s" % (room["uid"], name), uid=uid,
                               modificato=mod)
            res["stanze_foto"] += 1

    touched = set()
    for it in data.get("postazioni") or []:
        if not isinstance(it, dict):
            continue
        u = db.find_user(db.resolve_user_name(_clean(it.get("utente"))))
        suid = _clean(it.get("stanza_uid"))
        sid = room_id(suid) if suid else None
        if not u or (suid and sid is None):
            res["postazioni_saltate"] += 1
            continue
        try:
            numero = max(0, int(it.get("postazione") or 0))
        except (TypeError, ValueError):
            numero = 0
        t = db.apply_posto(u["id"], sid, numero, _clean(it.get("modificato")) or now_utc_iso())
        if t is not None:
            touched |= t
            res["postazioni"] += 1
    db.renumber_rooms(touched)


def verbali_for_phone(db, with_photos):
    """I verbali (anche quelli fatti sul PC) con foto e firme, per farli firmare sul telefono."""
    out, n_foto = [], 0
    for r in db.search_verbali():
        try:
            dati = json.loads(r["dati"] or "{}")
        except ValueError:
            continue
        folder = os.path.join(verbali_foto_root(), r["uid"])
        foto_names, firme_names = dati.pop("_foto", []), dati.pop("_firme", {})
        cart_names = dati.pop("_cartaceo", {})
        item = {"uid": r["uid"], "tipo": r["tipo"], "dati": dati, "foto": [], "firme": {},
                "cartaceo": {k: True for k, v in (cart_names or {}).items()
                             if v and k not in (firme_names or {})},
                "modificato": r["modificato"], "creato": r["creato"]}
        if with_photos:
            for name in foto_names:
                p = os.path.join(folder, name)
                data = photo_data_url(p) if os.path.isfile(p) else None
                if data:
                    item["foto"].append({"nome": name, "data": data})
            n_foto += len(item["foto"])
        for chi, name in (firme_names or {}).items():
            p = os.path.join(folder, name)
            if os.path.isfile(p):
                with open(p, "rb") as fh:
                    item["firme"][chi] = "data:image/png;base64," + \
                        base64.b64encode(fh.read()).decode("ascii")
        out.append(item)
    return out, n_foto


def user_cell(u, key):
    if key == "n":
        return u["n"]
    if key == "dotazione":
        return dotazione_text(u.get("dotazione"))
    if key == "ultimo":
        return fmt_dt(u["ultimo"]) if u.get("ultimo") else ""
    return u.get(key, "") or ""


def export_users_excel(path, users, keys, filtri_desc):
    from openpyxl import Workbook

    cols = [c for c in USER_COLUMNS if c[0] in keys]
    wb = Workbook()
    ws = wb.active
    ws.title = "Utenti"
    text_cols = tuple(i for i, c in enumerate(cols, start=1)
                      if c[0] in ("telefono", "seriale_pc", "imei"))
    _excel_sheet(ws, [c[1] for c in cols],
                 [[user_cell(u, c[0]) for c in cols] for u in users],
                 [c[4] for c in cols], text_cols=text_cols)
    _excel_summary(wb, "Estrazione dati utenti", filtri_desc, "Utenti esportati", len(users), [])
    wb.save(path)


def export_users_pdf(path, users, keys, filtri_desc):
    cols = [c for c in USER_COLUMNS if c[0] in keys]
    total = sum(c[3] for c in cols)
    widths = [c[3] * 267.0 / total for c in cols]
    build_pdf(path, "Estrazione dati utenti", filtri_desc, "Utenti",
              [c[1] for c in cols],
              [[str(user_cell(u, c[0])) for c in cols] for u in users],
              widths, "Nessun utente per i filtri selezionati.")


def build_pdf(path, titolo, filtri_desc, conteggio, headers, rows, col_widths_mm, empty_text):
    """PDF A4 orizzontale con logo INAIL a sinistra, tricolore e tabella.
    Ogni cella è un testo semplice oppure ("html", testo) per un testo già formattato."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    from xml.sax.saxutils import escape

    navy = colors.HexColor(NAVY)
    base = getSampleStyleSheet()
    st_title = ParagraphStyle("t", parent=base["Title"], fontName="Helvetica-Bold",
                              fontSize=17, textColor=navy, alignment=0, leading=21)
    st_info = ParagraphStyle("i", parent=base["Normal"], fontSize=9,
                             textColor=colors.HexColor(GREY), leading=12)
    st_head = ParagraphStyle("h", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=8.5,
                             textColor=colors.white, leading=10)
    st_cell = ParagraphStyle("c", parent=base["Normal"], fontSize=8, leading=10)

    def para(cell, style=st_cell):
        if isinstance(cell, tuple):
            return Paragraph(cell[1], style)
        text = (cell or "").replace("’", "'").replace("‘", "'")
        return Paragraph(escape(text).replace("\n", "<br/>"), style)

    doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=12 * mm, bottomMargin=14 * mm, title=titolo)
    width = landscape(A4)[0] - 30 * mm
    story = []

    try:
        iw, ih = ImageReader(resource_path("logo.png")).getSize()
        lw = 48 * mm
        left = Image(resource_path("logo.png"), width=lw, height=lw * ih / iw)
    except Exception:
        left = Paragraph("<b>INAIL</b>", st_title)
    right = [
        Paragraph(escape(titolo), st_title),
        Paragraph("%s: <b>%d</b> &nbsp;|&nbsp; Filtri: %s &nbsp;|&nbsp; Generato il %s"
                  % (escape(conteggio), len(rows), escape(filtri_desc),
                     datetime.now().strftime("%d/%m/%Y %H:%M")), st_info),
    ]
    head = Table([[left, right]], colWidths=[62 * mm, width - 62 * mm])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                              ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(head)
    story.append(Spacer(1, 3 * mm))
    bar = Table([["", "", ""]], colWidths=[width / 3] * 3, rowHeights=[1.6 * mm])
    bar.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 0), colors.HexColor(GREEN)),
        ("BACKGROUND", (1, 0), (1, 0), colors.HexColor("#dfe9f2")),
        ("BACKGROUND", (2, 0), (2, 0), colors.HexColor(RED)),
    ]))
    story.append(bar)
    story.append(Spacer(1, 5 * mm))

    if not rows:
        story.append(Paragraph(escape(empty_text), st_info))
    else:
        data = [[para(h, st_head) for h in headers]]
        for r in rows:
            data.append([para(c) for c in r])
        fixed = sum(col_widths_mm[:-1]) * mm
        cols = [w * mm for w in col_widths_mm[:-1]] + [width - fixed]
        table = Table(data, colWidths=cols, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), navy),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor(LIGHT)]),
            ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor(BORDER)),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]))
        story.append(table)

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor(GREY))
        canvas.drawString(15 * mm, 8 * mm, "Help Desk INAIL - %s" % titolo)
        canvas.drawRightString(landscape(A4)[0] - 15 * mm, 8 * mm, "Pagina %d" % doc_.page)
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)


def stato_html(stato, colori):
    from xml.sax.saxutils import escape
    color = colori.get(stato)
    if not color:
        return stato
    return ("html", '<font color="%s"><b>%s</b></font>' % (color, escape(stato)))


def export_pdf(path, rows, filtri_desc):
    build_pdf(
        path, "Registro interventi Help Desk", filtri_desc, "Interventi",
        ["Data e ora", "Utente", "Matricola", "Sede", "Ticket", "Stato",
         "Per cosa mi hanno chiamato", "Come ho risolto", "Note"],
        [[fmt_dt(r["data_ora"]), r["utente"], r["matricola"], r["sede"], r["ticket"],
          stato_html(r["stato"], STATO_COLORI), r["richiesta"], r["soluzione"], r["note"]]
         for r in rows],
        [23, 28, 20, 18, 16, 18, 53, 53, 38], "Nessun intervento per i filtri selezionati.")


def export_manut_pdf(path, rows, filtri_desc):
    from xml.sax.saxutils import escape

    cells = []
    for r in rows:
        dove = "%s\n%s - %s" % (r["stampante"] or r["modello"] or "Stampante", r["edificio"],
                                floor_label(r["piano"]))
        if r["stampante"] and r["modello"]:
            dove = "%s (%s)\n%s - %s" % (r["stampante"], r["modello"], r["edificio"],
                                         floor_label(r["piano"]))
        tec = []
        if r["data_tecnico"]:
            tec.append(fmt_date(datetime.strptime(r["data_tecnico"], "%Y-%m-%d")))
        if r["tecnico"]:
            tec.append(r["tecnico"])
        stato = ("html", '<font color="%s"><b>%s</b></font>'
                 % (RED if r["stato"] == "Richiesto" else GREEN_DARK, escape(r["stato"])))
        cells.append([fmt_dt(r["data_ora"]), dove, r["ticket"], stato, "\n".join(tec),
                      r["problema"], r["intervento"]])
    build_pdf(
        path, "Interventi del tecnico sulle stampanti", filtri_desc, "Interventi",
        ["Richiesta del", "Stampante", "Ticket", "Stato", "Tecnico",
         "Problema segnalato", "Intervento svolto"],
        cells, [24, 56, 20, 20, 32, 55, 60],
        "Nessun intervento del tecnico per i filtri selezionati.")


def export_toner_pdf(path, rows, filtri_desc):
    from xml.sax.saxutils import escape

    cells = []
    for r in rows:
        colore = ("html", '<font color="%s"><b>%s</b></font>'
                  % (TONER_TESTO.get(r["colore"], TEXT), escape(r["colore"])))
        cells.append([fmt_dt(r["data_ora"]), r["sede"], r["edificio"], floor_label(r["piano"]),
                      r["stampante"], r["modello"], colore, r["note"]])
    build_pdf(
        path, "Registro cambi toner", filtri_desc, "Cambi toner",
        ["Data e ora", "Sede", "Edificio", "Piano", "Stampante", "Modello", "Colore", "Note"],
        cells, [24, 22, 40, 22, 40, 28, 20, 71], "Nessun cambio toner per i filtri selezionati.")


# --------------------------------------------------------------------------
# Interfaccia
# --------------------------------------------------------------------------
NAV_KEYS = {"Up", "Down", "Left", "Right", "Return", "Escape", "Tab", "Home", "End", "Prior",
            "Next", "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"}


def key_name(name):
    return " ".join((name or "").split()).casefold()


def bind_autocomplete(combo, var, get_values):
    """Restringe l'elenco del menu a tendina mentre si scrive (nome o cognome)."""
    def on_key(event):
        if event.keysym in NAV_KEYS:
            return
        terms = var.get().casefold().split()
        values = get_values()
        combo["values"] = [v for v in values if matches_all(terms, v)] if terms else values
    combo.bind("<KeyRelease>", on_key, add="+")


def make_text(parent, height):
    t = tk.Text(parent, height=height, wrap="word", font=("Segoe UI", 10), bg=WHITE, fg=TEXT,
                relief="flat", highlightthickness=1, highlightbackground=BORDER,
                highlightcolor=NAVY, padx=6, pady=5, undo=True)

    def focus_next(event):
        event.widget.tk_focusNext().focus()
        return "break"

    def focus_prev(event):
        event.widget.tk_focusPrev().focus()
        return "break"

    t.bind("<Tab>", focus_next)
    t.bind("<Shift-Tab>", focus_prev)
    return t


def period_dates(kind):
    today = date.today()
    if kind == "oggi":
        return today, today
    if kind == "7g":
        return today - timedelta(days=6), today
    if kind == "mese":
        return today.replace(day=1), today
    if kind == "anno":
        return today.replace(month=1, day=1), today
    return None, None


class ModalDialog(tk.Toplevel):
    def __init__(self, parent, title, geometry=None):
        super().__init__(parent)
        self.title(title)
        self.configure(bg=LIGHT)
        self.transient(parent)
        if geometry:
            self.geometry(geometry)
        self.bind("<Escape>", lambda e: self.destroy())
        self.after(50, self._grab)

    def _grab(self):
        try:
            self.grab_set()
        except tk.TclError:
            pass


class UserDialog(ModalDialog):
    """Scheda di una persona: nome, ufficio, sede, dati della postazione e dotazione."""

    FIELDS = (("Nome e cognome *", "nome"), ("Ufficio", "ufficio"),
              ("Matricola aziendale", "matricola"),
              ("Sede", "sede"), ("Telefono", "telefono"), ("Email", "email"),
              ("Seriale del PC", "seriale_pc"), ("IMEI cellulare aziendale", "imei"))

    def __init__(self, parent, user, sedi, on_save, seat=None, on_room=None, uffici=(),
                 defaults=None):
        super().__init__(parent, "Nuovo utente" if user is None else "Scheda utente")
        self.resizable(False, False)
        self.on_save = on_save
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        data = dict(defaults or {})
        data.update(user or {})
        uffici = list(uffici)
        self.vars = {k: tk.StringVar(value=data.get(k, "") or "") for _, k in self.FIELDS}
        for i, (label, key) in enumerate(self.FIELDS):
            ttk.Label(body, text=label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 14))
            if key == "sede":
                w = ttk.Combobox(body, textvariable=self.vars[key], values=sedi, width=34)
                bind_autocomplete(w, self.vars[key], lambda: sedi)
            elif key == "ufficio":
                w = ttk.Combobox(body, textvariable=self.vars[key], values=uffici, width=34)
                bind_autocomplete(w, self.vars[key], lambda: uffici)
            else:
                w = ttk.Entry(body, textvariable=self.vars[key], width=36)
            w.grid(row=i, column=1, sticky="ew", pady=5)
            if key in ("seriale_pc", "imei"):
                scan_button(body, self.vars[key], "imei" if key == "imei" else "serial", entry=w).grid(
                    row=i, column=2, sticky="w", padx=(8, 0))
            if key == "email":
                ttk.Button(body, text="Scrivi mail", command=self.write_mail).grid(
                    row=i, column=2, sticky="w", padx=(8, 0))
            if i == 0:
                w.focus_set()
        nrow = len(self.FIELDS)
        self._build_dotazione(body, nrow, data.get("dotazione"))
        nrow += 1
        if user:
            ttk.Label(body, text="Dove si siede").grid(row=nrow, column=0, sticky="w", pady=5,
                                                       padx=(0, 14))
            fr = ttk.Frame(body)
            fr.grid(row=nrow, column=1, sticky="w", pady=5)
            ttk.Label(fr, text=seat or "Nessuna stanza assegnata (si assegna da Stampanti › "
                                       "Piani e stanze).").pack(side="left")
            if seat and on_room:
                ttk.Button(fr, text="Apri la stanza", command=lambda: (self.destroy(), on_room())
                           ).pack(side="left", padx=(10, 0))
            nrow += 1
        ttk.Label(body, style="Hint.TLabel", wraplength=420, justify="left",
                  text="Se cambi il nome con quello di un altro utente già in lista, i due "
                       "utenti vengono uniti.").grid(row=nrow, column=0,
                                                     columnspan=2, sticky="w", pady=(8, 0))
        btns = ttk.Frame(body)
        btns.grid(row=nrow + 1, column=0, columnspan=3, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))
        self.bind("<Return>", lambda e: self.save())

    def write_mail(self):
        address = clean_email(self.vars["email"].get())
        if not address:
            messagebox.showinfo(APP_NAME, "Scrivi prima l'email di questa persona.", parent=self)
        elif not mail_to(address):
            messagebox.showwarning(APP_NAME, "Questa email non sembra valida: %s" % address,
                                   parent=self)

    # ----- dotazione: elenco libero di apparecchi (tipo, modello, seriale o IMEI) -----
    def _build_dotazione(self, body, row, items):
        self.dotazione = clean_dotazione(items)
        self.dot_idx = None                      # riga che si sta modificando (None = nuova)
        self.v_dt, self.v_dm, self.v_ds = tk.StringVar(), tk.StringVar(), tk.StringVar()
        ttk.Label(body, text="Dotazione").grid(row=row, column=0, sticky="nw", pady=5, padx=(0, 14))
        box = ttk.Frame(body)
        box.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
        self.tree_d = ttk.Treeview(box, columns=("tipo", "modello", "seriale"), show="headings",
                                   height=3, selectmode="browse")
        for col, text, w in (("tipo", "Apparecchio", 120), ("modello", "Modello", 140),
                             ("seriale", "Seriale / IMEI", 160)):
            self.tree_d.heading(col, text=text)
            self.tree_d.column(col, width=w, anchor="w")
        self.tree_d.grid(row=0, column=0, columnspan=4, sticky="ew")
        self.tree_d.bind("<<TreeviewSelect>>", lambda e: self.dot_pick())
        ttk.Combobox(box, textvariable=self.v_dt, values=TIPI_DOTAZIONE_PERSONA, width=16).grid(
            row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(box, textvariable=self.v_dm, width=16).grid(row=1, column=1, sticky="w",
                                                              padx=(6, 0), pady=(6, 0))
        e_s = ttk.Entry(box, textvariable=self.v_ds, width=20)
        e_s.grid(row=1, column=2, sticky="w", padx=(6, 0), pady=(6, 0))
        scan_button(box, self.v_ds, "any", entry=e_s).grid(row=1, column=3, sticky="w",
                                                           padx=(6, 0), pady=(6, 0))
        bar = ttk.Frame(box)
        bar.grid(row=2, column=0, columnspan=4, sticky="w", pady=(6, 0))
        self.btn_dot = ttk.Button(bar, text="Aggiungi alla dotazione", command=self.dot_add)
        self.btn_dot.pack(side="left")
        ttk.Button(bar, text="Togli la riga scelta", command=self.dot_remove).pack(
            side="left", padx=(8, 0))
        ttk.Label(box, style="Hint.TLabel", wraplength=420, justify="left",
                  text="Per cambiare una riga: sceglila nell'elenco, correggi i campi e premi "
                       "«Salva la riga».").grid(row=3, column=0, columnspan=4, sticky="w",
                                               pady=(4, 0))
        self.dot_fill()

    def dot_fill(self):
        self.tree_d.delete(*self.tree_d.get_children())
        for i, r in enumerate(self.dotazione):
            self.tree_d.insert("", "end", iid=str(i), values=(r["tipo"], r["modello"], r["seriale"]))

    def dot_clear(self):
        self.dot_idx = None
        for v in (self.v_dt, self.v_dm, self.v_ds):
            v.set("")
        self.btn_dot.configure(text="Aggiungi alla dotazione")

    def dot_pick(self):
        sel = self.tree_d.selection()
        if not sel:
            return
        self.dot_idx = int(sel[0])
        r = self.dotazione[self.dot_idx]
        self.v_dt.set(r["tipo"])
        self.v_dm.set(r["modello"])
        self.v_ds.set(r["seriale"])
        self.btn_dot.configure(text="Salva la riga")

    def dot_add(self):
        rows = clean_dotazione([{"tipo": self.v_dt.get(), "modello": self.v_dm.get(),
                                 "seriale": self.v_ds.get()}])
        if not rows:
            messagebox.showinfo(APP_NAME, "Scrivi almeno il tipo, il modello o il seriale "
                                          "dell'apparecchio.", parent=self)
            return
        if self.dot_idx is not None and self.dot_idx < len(self.dotazione):
            self.dotazione[self.dot_idx] = rows[0]
        else:
            self.dotazione.append(rows[0])
        self.dot_clear()
        self.dot_fill()

    def dot_remove(self):
        sel = self.tree_d.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Scegli prima una riga della dotazione.", parent=self)
            return
        del self.dotazione[int(sel[0])]
        self.dot_clear()
        self.dot_fill()

    def save(self):
        data = {k: v.get() for k, v in self.vars.items()}
        pending = clean_dotazione([{"tipo": self.v_dt.get(), "modello": self.v_dm.get(),
                                    "seriale": self.v_ds.get()}])
        dot = list(self.dotazione)
        if pending and self.dot_idx is None:      # riga scritta ma non ancora aggiunta: non si perde
            dot += pending
        elif pending and self.dot_idx < len(dot):
            dot[self.dot_idx] = pending[0]
        data["dotazione"] = dot
        if self.on_save(data):
            self.destroy()


class BuildingDialog(ModalDialog):
    def __init__(self, parent, building, sedi, on_save):
        super().__init__(parent, "Nuovo edificio" if building is None else "Modifica edificio")
        self.resizable(False, False)
        self.on_save = on_save
        b = building or {}
        self.v_nome = tk.StringVar(value=b.get("nome", ""))
        self.v_sede = tk.StringVar(value=b.get("sede", ""))
        self.v_min = tk.StringVar(value=str(b.get("piano_min", 0)))
        self.v_max = tk.StringVar(value=str(b.get("piano_max", 5)))
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Nome edificio *").grid(row=0, column=0, sticky="w", pady=5,
                                                     padx=(0, 14))
        e = ttk.Entry(body, textvariable=self.v_nome, width=40)
        e.grid(row=0, column=1, sticky="ew", pady=5)
        e.focus_set()
        ttk.Label(body, text="Sede").grid(row=1, column=0, sticky="w", pady=5, padx=(0, 14))
        cb = ttk.Combobox(body, textvariable=self.v_sede, values=sedi, width=38)
        cb.grid(row=1, column=1, sticky="ew", pady=5)
        bind_autocomplete(cb, self.v_sede, lambda: sedi)
        ttk.Label(body, text="Piano più basso").grid(row=2, column=0, sticky="w", pady=5,
                                                     padx=(0, 14))
        ttk.Spinbox(body, from_=-3, to=30, textvariable=self.v_min, width=6).grid(
            row=2, column=1, sticky="w", pady=5)
        ttk.Label(body, text="Piano più alto").grid(row=3, column=0, sticky="w", pady=5,
                                                    padx=(0, 14))
        ttk.Spinbox(body, from_=0, to=30, textvariable=self.v_max, width=6).grid(
            row=3, column=1, sticky="w", pady=5)
        ttk.Label(body, style="Hint.TLabel", wraplength=380, justify="left",
                  text="0 = piano terra, -1 = seminterrato. Esempio: dal piano 0 al piano 5 "
                       "sono 6 livelli.").grid(row=4, column=0, columnspan=2, sticky="w",
                                               pady=(6, 0))
        btns = ttk.Frame(body)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))

    def save(self):
        try:
            pmin = int(self.v_min.get())
            pmax = int(self.v_max.get())
        except ValueError:
            messagebox.showwarning(APP_NAME, "I piani devono essere numeri interi.", parent=self)
            return
        data = {"nome": self.v_nome.get(), "sede": self.v_sede.get(),
                "piano_min": pmin, "piano_max": pmax}
        if self.on_save(data):
            self.destroy()


class MergePrintersDialog(ModalDialog):
    """Stampanti doppie: a sinistra le vecchie (senza numero di serie), a destra la nuova (col numero
    di serie) con cui unirle. Lo storico (cambi toner, interventi del tecnico, bolle) passa alla nuova,
    la vecchia si toglie. Il programma propone da solo i casi chiari; gli altri li scegli tu."""

    NESSUNA = "(non unire)"

    def __init__(self, parent, building, proposals, on_merge):
        super().__init__(parent, "Unisci doppioni stampanti")
        self.on_merge = on_merge
        self.rows = []                           # (id della vecchia, variabile, {etichetta: id nuova})
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Edificio: %s" % building["nome"],
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(body, text="A sinistra le stampanti vecchie (senza numero di serie), a destra quella "
                             "nuova (col numero di serie) con cui unirle. Lo storico (cambi toner, "
                             "interventi del tecnico, bolle) passa alla nuova e la vecchia si toglie; "
                             "la nuova tiene nome, seriale e stanza. Dove il programma non è sicuro "
                             "non sceglie: scegli tu, o lascia «(non unire)».",
                  style="Hint.TLabel", wraplength=820, justify="left").pack(anchor="w", pady=(2, 10))
        grid = ttk.Frame(body)
        grid.pack(fill="x")
        ttk.Label(grid, text="Stampante vecchia", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, sticky="w", padx=(0, 16))
        ttk.Label(grid, text="Unisci con la nuova", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=1, sticky="w")
        for i, pr in enumerate(proposals[:30], start=1):
            o = pr["old"]
            info = "%s - %s\n%s: %d cambi toner, %d interventi del tecnico" % (
                o["nome"] or o["modello"] or "senza nome", floor_label(o["piano"]),
                o["modello"] or "modello non indicato", pr["n_toner"], pr["n_manut"])
            ttk.Label(grid, text=info, justify="left").grid(row=i, column=0, sticky="w",
                                                            padx=(0, 16), pady=5)
            labels = {}
            for n in pr["plausibili"] + pr["altre"]:
                labels["%s - S/N %s - %s" % (n["nome"] or n["modello"] or "senza nome",
                                             n["seriale"], floor_label(n["piano"]))] = n["id"]
            sugg = next((k for k, v in labels.items() if v == pr["suggerita"]), self.NESSUNA)
            var = tk.StringVar(value=sugg)
            ttk.Combobox(grid, textvariable=var, values=[self.NESSUNA] + list(labels),
                         state="readonly", width=58).grid(row=i, column=1, sticky="w", pady=5)
            self.rows.append((o["id"], var, labels))
        if len(proposals) > 30:
            ttk.Label(body, text="Ce ne sono altre %d: unisci queste e riapri questa finestra." %
                      (len(proposals) - 30), style="Hint.TLabel").pack(anchor="w", pady=(6, 0))
        bf = ttk.Frame(body)
        bf.pack(fill="x", pady=(14, 0))
        ttk.Button(bf, text="Unisci quelle scelte", style="Save.TButton",
                   command=self.go).pack(side="left")
        ttk.Button(bf, text="Chiudi", command=self.destroy).pack(side="left", padx=(8, 0))

    def pairs(self):
        out = []
        for old_id, var, labels in self.rows:
            keep = labels.get(var.get())
            if keep:
                out.append((old_id, keep))
        return out

    def go(self):
        pairs = self.pairs()
        if not pairs:
            messagebox.showinfo(APP_NAME, "Non hai scelto nessuna unione.", parent=self)
            return
        if not messagebox.askyesno(
                APP_NAME, "Unisco %d %s?\n\nLo storico passa alla stampante nuova e quella vecchia "
                          "viene tolta. Non si può annullare (se vuoi, prima fai File › Salva una "
                          "copia di backup)." % (len(pairs), "stampante" if len(pairs) == 1 else "stampanti"),
                parent=self):
            return
        try:
            n_t, n_m = self.on_merge(pairs)
        except (ValueError, sqlite3.Error) as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a unirle:\n%s" % e, parent=self)
            return
        messagebox.showinfo(APP_NAME, "Fatto: %d %s unite.\nSpostati %d cambi toner e %d interventi del "
                                      "tecnico.\n\nAl prossimo «Esporta dati per il telefono» anche il "
                                      "telefono toglie i doppioni." % (
                                          len(pairs), "stampante" if len(pairs) == 1 else "stampanti",
                                          n_t, n_m), parent=self)
        self.destroy()


class PrinterEditDialog(ModalDialog):
    def __init__(self, parent, printer, building, floor, models, on_save, rooms_fn=None):
        super().__init__(parent, "Nuova stampante" if printer is None else "Modifica stampante")
        self.resizable(False, False)
        self.on_save = on_save
        self.has_rooms = rooms_fn is not None    # senza elenco stanze, quella salvata non si tocca
        self.rooms_fn = rooms_fn or (lambda piano: [])
        self.room_uids = {}                      # codice della stanza -> uid, per il piano scelto
        self.floors = list(range(building["piano_max"], building["piano_min"] - 1, -1))
        self.labels = {floor_label(n): n for n in self.floors}
        p = printer or {}
        if "piano" in p:
            start = p["piano"]
        elif floor is not None:
            start = floor
        else:
            start = 0 if building["piano_min"] <= 0 <= building["piano_max"] else building["piano_min"]
        self.v_piano = tk.StringVar(value=floor_label(start))
        self.v_nome = tk.StringVar(value=p.get("nome", ""))
        self.v_modello = tk.StringVar(value=p.get("modello", models[0] if models else ""))
        self.v_ip = tk.StringVar(value=p.get("ip", ""))
        self.v_note = tk.StringVar(value=p.get("note", ""))
        self.v_seriale = tk.StringVar(value=p.get("seriale", ""))
        self.v_fuori = tk.BooleanVar(value=bool(p.get("fuori_servizio")))
        self.v_stanza = tk.StringVar()
        self._stanza_uid0 = p.get("stanza_uid", "") or ""

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Piano").grid(row=0, column=0, sticky="w", pady=5, padx=(0, 14))
        cb_piano = ttk.Combobox(body, textvariable=self.v_piano, state="readonly", width=24,
                                values=[floor_label(n) for n in self.floors])
        cb_piano.grid(row=0, column=1, sticky="w", pady=5)
        cb_piano.bind("<<ComboboxSelected>>", lambda e: self.fill_rooms())
        ttk.Label(body, text="Stanza").grid(row=1, column=0, sticky="w", pady=5, padx=(0, 14))
        self.cb_stanza = ttk.Combobox(body, textvariable=self.v_stanza, state="readonly", width=24,
                                      values=[])
        self.cb_stanza.grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(body, text="Nome o posizione").grid(row=2, column=0, sticky="w", pady=5,
                                                     padx=(0, 14))
        e = ttk.Entry(body, textvariable=self.v_nome, width=40)
        e.grid(row=2, column=1, sticky="ew", pady=5)
        e.focus_set()
        ttk.Label(body, text="Modello").grid(row=3, column=0, sticky="w", pady=5, padx=(0, 14))
        ttk.Combobox(body, textvariable=self.v_modello, values=models, width=38).grid(
            row=3, column=1, sticky="ew", pady=5)
        ttk.Label(body, text="Indirizzo IP").grid(row=4, column=0, sticky="w", pady=5,
                                                  padx=(0, 14))
        ttk.Entry(body, textvariable=self.v_ip, width=24).grid(row=4, column=1, sticky="w",
                                                               pady=5)
        ttk.Label(body, text="Numero di serie").grid(row=5, column=0, sticky="w", pady=5, padx=(0, 14))
        sf = ttk.Frame(body)
        sf.grid(row=5, column=1, sticky="ew", pady=5)
        e_ser = ttk.Entry(sf, textvariable=self.v_seriale, width=30)
        e_ser.pack(side="left")
        scan_button(sf, self.v_seriale, "serial", entry=e_ser).pack(side="left", padx=(6, 0))
        ttk.Label(body, text="Note").grid(row=6, column=0, sticky="w", pady=5, padx=(0, 14))
        ttk.Entry(body, textvariable=self.v_note, width=40).grid(row=6, column=1, sticky="ew",
                                                                 pady=5)
        ttk.Checkbutton(body, text="Fuori servizio (stampante rotta)",
                        variable=self.v_fuori).grid(row=7, column=1, sticky="w", pady=(4, 0))
        ttk.Label(body, style="Hint.TLabel", wraplength=400, justify="left",
                  text="Nome o posizione, ad esempio \"corridoio ovest\" o \"stanza 12\". "
                       "Il modello sceglie la foto nella grafica. Il numero di serie e' sull'etichetta "
                       "della stampante: puoi anche leggerlo con la fotocamera. "
                       "Con la spunta \"Fuori servizio\" la stampante resta in elenco ma con una "
                       "croce rossa, anche sul telefono.").grid(
            row=8, column=0, columnspan=2, sticky="w", pady=(6, 0))
        btns = ttk.Frame(body)
        btns.grid(row=9, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))
        self.fill_rooms(first=True)

    NO_ROOM = "(nessuna stanza)"

    def fill_rooms(self, first=False):
        """Elenco delle stanze del piano scelto. Se la stanza di prima non è su questo piano,
        si torna a «nessuna stanza»."""
        piano = self.labels.get(self.v_piano.get(), 0)
        rooms = list(self.rooms_fn(piano))
        self.room_uids = {r["codice"]: r["uid"] for r in rooms}
        self.cb_stanza.configure(values=[self.NO_ROOM] + [r["codice"] for r in rooms])
        keep = self._stanza_uid0 if first else self.room_uids.get(self.v_stanza.get(), "")
        code = next((c for c, u in self.room_uids.items() if keep and u == keep), "")
        self.v_stanza.set(code or self.NO_ROOM)

    def save(self):
        data = {"piano": self.labels.get(self.v_piano.get(), 0), "nome": self.v_nome.get(),
                "modello": self.v_modello.get(), "ip": self.v_ip.get(), "note": self.v_note.get(),
                "seriale": self.v_seriale.get(), "fuori_servizio": bool(self.v_fuori.get())}
        if self.has_rooms:
            data["stanza_uid"] = self.room_uids.get(self.v_stanza.get(), "")
        if self.on_save(data):
            self.destroy()


class PrinterDialog(ModalDialog):
    """Scheda di una stampante: dati, registrazione cambio toner e storico."""

    def __init__(self, app, printer_id):
        super().__init__(app, "Stampante", "880x720")
        self.minsize(780, 620)
        self.app = app
        self.printer_id = printer_id
        self.printer = None
        self.colors = {n: tk.BooleanVar(value=False) for n in TONER_NOMI}
        self.var_dt = tk.StringVar()
        self.var_note = tk.StringVar()

        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(2, weight=1)

        top = ttk.Frame(body)
        top.grid(row=0, column=0, sticky="ew")
        self.img_label = tk.Label(top, bg=LIGHT, fg=GREY, text="")
        self.img_label.pack(side="left", padx=(0, 20))
        info = ttk.Frame(top)
        info.pack(side="left", anchor="n")
        self.lbl_name = ttk.Label(info, text="", font=("Segoe UI", 15, "bold"))
        self.lbl_name.pack(anchor="w")
        self.lbl_info = ttk.Label(info, text="", justify="left")
        self.lbl_info.pack(anchor="w", pady=(6, 0))
        bf = ttk.Frame(info)
        bf.pack(anchor="w", pady=(10, 0))
        ttk.Button(bf, text="Modifica stampante…", command=self.edit).pack(side="left")
        ttk.Button(bf, text="Elimina stampante", style="Danger.TButton",
                   command=self.delete).pack(side="left", padx=(8, 0))
        ttk.Button(info, text="Interventi del tecnico Konica…", style="Primary.TButton",
                   command=self.open_manut).pack(anchor="w", pady=(8, 0))

        form = ttk.LabelFrame(body, text="Registra un cambio toner o vaschetta", padding=10)
        form.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        r1 = ttk.Frame(form)
        r1.pack(fill="x")
        ttk.Label(r1, text="Cambiati:").pack(side="left")
        for nome, _, testo in COLORI_TONER:
            tk.Checkbutton(r1, text=nome, variable=self.colors[nome], bg=LIGHT, fg=testo,
                           activebackground=LIGHT, activeforeground=testo, selectcolor=WHITE,
                           font=("Segoe UI", 11, "bold")).pack(side="left", padx=(12, 0))
        r2 = ttk.Frame(form)
        r2.pack(fill="x", pady=(10, 0))
        ttk.Label(r2, text="Data e ora").pack(side="left")
        ttk.Entry(r2, textvariable=self.var_dt, width=18).pack(side="left", padx=(6, 4))
        ttk.Button(r2, text="Adesso", command=self.set_now).pack(side="left")
        ttk.Label(r2, text="Note").pack(side="left", padx=(16, 6))
        ttk.Entry(r2, textvariable=self.var_note, width=30).pack(side="left", fill="x",
                                                                 expand=True)
        r3 = ttk.Frame(form)
        r3.pack(fill="x", pady=(10, 0))
        ttk.Button(r3, text="Registra il cambio", style="Save.TButton",
                   command=self.register).pack(side="left")
        self.lbl_status = ttk.Label(r3, text="", foreground=GREEN, font=("Segoe UI", 10, "bold"))
        self.lbl_status.pack(side="left", padx=12)

        hist = ttk.Frame(body)
        hist.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        hist.columnconfigure(0, weight=1)
        hist.rowconfigure(1, weight=1)
        ttk.Label(hist, text="Storico cambi toner e vaschetta di questa stampante",
                  font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.tree = ttk.Treeview(hist, columns=("data", "colore", "note"), show="headings",
                                 selectmode="browse", height=6)
        for c, title, width in (("data", "Data e ora", 140), ("colore", "Colore", 100),
                                ("note", "Note", 420)):
            self.tree.heading(c, text=title)
            self.tree.column(c, width=width, anchor="w", stretch=(c == "note"))
        for nome in TONER_NOMI:
            self.tree.tag_configure(nome, foreground=TONER_TESTO[nome])
        sb = ttk.Scrollbar(hist, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        bottom = ttk.Frame(body)
        bottom.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bottom, text="Elimina il cambio selezionato", style="Danger.TButton",
                   command=self.delete_toner_selected).pack(side="left")
        ttk.Button(bottom, text="Chiudi", command=self.destroy).pack(side="right")

        self.set_now()
        self.refresh()

    def set_now(self):
        self.var_dt.set(datetime.now().strftime("%d/%m/%Y %H:%M"))

    def refresh(self):
        p = self.app.db.get_stampante(self.printer_id)
        if not p:
            self.destroy()
            return
        self.printer = p
        img = self.app.get_printer_image(p["modello"], "m")
        if img:
            self.img_label.configure(image=img, text="")
            self.img_label.image = img
        else:
            self.img_label.configure(image="", text="[foto non disponibile]")
        self.lbl_name.configure(text=p["nome"] or p["modello"] or "Stampante")
        lines = ["Modello: %s" % (p["modello"] or "-"),
                 "Dove: %s - %s%s" % (p["edificio"], floor_label(p["piano"]),
                                      " - stanza %s" % p["stanza_codice"]
                                      if p.get("stanza_codice") else "")]
        if p.get("fuori_servizio"):
            lines.insert(0, "FUORI SERVIZIO (stampante rotta)")
        if p.get("seriale"):
            lines.append("Numero di serie: %s" % p["seriale"])
        if p["ip"]:
            lines.append("IP: %s" % p["ip"])
        if p["note"]:
            lines.append("Note: %s" % p["note"])
        ultimi = {r["colore"]: r["data_ora"] for r in reversed(self.app.db.toner_of(self.printer_id))}
        lines.append("Ultimo cambio:")
        stime = self.app.db.toner_estimates().get(self.printer_id, {})
        for colore in TONER_NOMI:
            quando = ultimi.get(colore)
            lines.append("   %s: %s%s" % (colore, ago_long(days_ago(quando)),
                                          " (%s)" % fmt_dt(quando) if quando else ""))
            lines.append("      %s" % toner_estimate_text(stime.get(colore, {"pct": None})))
        bassa = [(r["colore"], r["quantita"]) for r in self.app.db.scorta_all()
                 if model_key(r["modello"]) == model_key(p["modello"])
                 and scorta_status(r["quantita"], r["soglia"]) != "OK"]
        if bassa and p["modello"]:
            lines.append("ATTENZIONE, scorta in esaurimento: %s" % ", ".join(
                "%s (ne restano %d)" % (col, q) for col, q in bassa))
        man = self.app.db.manutenzioni_of(self.printer_id)
        if man:
            richieste = sum(1 for m in man if m["stato"] == "Richiesto")
            tec = "Tecnico Konica: %d interventi (ultimo %s)" % (len(man), fmt_dt(man[0]["data_ora"])[:10])
            if richieste:
                tec += " - %d da eseguire" % richieste
            lines.append(tec)
        self.lbl_info.configure(text="\n".join(lines))
        self.tree.delete(*self.tree.get_children())
        for r in self.app.db.toner_of(self.printer_id):
            self.tree.insert("", "end", iid=str(r["id"]), tags=(r["colore"],),
                             values=(fmt_dt(r["data_ora"]), r["colore"], r["note"]))

    def register(self):
        colori = [n for n in TONER_NOMI if self.colors[n].get()]
        if not colori:
            messagebox.showinfo(APP_NAME, "Spunta almeno un colore di toner.", parent=self)
            return
        dt = parse_datetime(self.var_dt.get())
        if not dt:
            messagebox.showwarning(APP_NAME, "La data non è valida.\nScrivila così: "
                                   "gg/mm/aaaa oo:mm (esempio 13/11/2026 14:30).", parent=self)
            return
        low = self.app.db.add_toner(self.printer_id, colori, dt.strftime("%Y-%m-%d %H:%M"),
                                    self.var_note.get().strip())
        for v in self.colors.values():
            v.set(False)
        self.var_note.set("")
        self.lbl_status.configure(text="Registrato: %s - %s" % (", ".join(colori),
                                                               dt.strftime("%d/%m/%Y %H:%M")))
        self.set_now()
        self.refresh()
        self.app.after_toner_change()
        if low:
            messagebox.showwarning(
                APP_NAME, "Attenzione: la scorta sta finendo.\n\n%s\n\nRicordati di ordinarne."
                % "\n".join("- %s %s: ne restano %d" % (r["colore"], r["modello"], r["quantita"])
                            for r in low), parent=self)

    def open_manut(self):
        PrinterManutenzioniDialog(self.app, self.printer_id, parent=self)

    def delete_toner_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima un cambio dallo storico.", parent=self)
            return
        if messagebox.askyesno(APP_NAME, "Eliminare questo cambio toner?", parent=self):
            self.app.db.delete_toner(int(sel[0]))
            self.refresh()
            self.app.after_toner_change()

    def edit(self):
        building = next((b for b in self.app.buildings if b["id"] == self.printer["edificio_id"]),
                        None)
        if not building:
            return

        def on_save(data):
            ok = self.app.save_printer(self.printer_id, building["id"], data)
            if ok:
                self.refresh()
            return ok

        PrinterEditDialog(self, self.printer, building, None, self.app.models_list(), on_save,
                          rooms_fn=lambda piano: self.app.db.stanze_of(building["id"], piano))

    def delete(self):
        n = len(self.app.db.toner_of(self.printer_id))
        extra = ("\nVerranno eliminati anche i %d cambi toner registrati." % n) if n else ""
        n_m = len(self.app.db.manutenzioni_of(self.printer_id))
        if n_m:
            extra += "\nE i %d interventi del tecnico." % n_m
        if messagebox.askyesno(APP_NAME, "Eliminare questa stampante?%s" % extra,
                               icon="warning", parent=self):
            self.app.db.delete_stampante(self.printer_id)
            self.app.after_printers_change()
            self.destroy()


class ManutenzioneDialog(ModalDialog):
    """Richiesta o intervento del tecnico Konica su una stampante."""

    def __init__(self, parent, printers, current, on_save, lock_printer=False):
        is_new = not current.get("id")
        super().__init__(parent, "Nuovo intervento del tecnico" if is_new
                         else "Modifica intervento del tecnico")
        self.resizable(False, False)
        self.on_save = on_save
        self.printers = printers
        c = current
        label = next((l for l, i in printers.items() if i == c.get("stampante_id")), "")
        self.v_pr = tk.StringVar(value=label)
        self.v_dt = tk.StringVar(value=fmt_dt(c["data_ora"]) if c.get("data_ora")
                                 else datetime.now().strftime("%d/%m/%Y %H:%M"))
        self.v_dtec = tk.StringVar(
            value=fmt_date(datetime.strptime(c["data_tecnico"], "%Y-%m-%d"))
            if c.get("data_tecnico") else "")
        self.v_ticket = tk.StringVar(value=c.get("ticket", ""))
        self.v_tec = tk.StringVar(value=c.get("tecnico", ""))
        self.v_stato = tk.StringVar(value=c.get("stato", "Richiesto"))
        self.v_stato.trace_add("write", self._stato_changed)

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        ttk.Label(body, text="Stampante *").grid(row=0, column=0, sticky="w", pady=5, padx=(0, 14))
        ttk.Combobox(body, textvariable=self.v_pr, values=list(printers), width=62,
                     state="disabled" if lock_printer else "readonly").grid(
            row=0, column=1, sticky="ew", pady=5)
        ttk.Label(body, text="Data richiesta *").grid(row=1, column=0, sticky="w", pady=5,
                                                       padx=(0, 14))
        r1 = ttk.Frame(body)
        r1.grid(row=1, column=1, sticky="w", pady=5)
        ttk.Entry(r1, textvariable=self.v_dt, width=18).pack(side="left")
        ttk.Button(r1, text="Adesso", command=lambda: self.v_dt.set(
            datetime.now().strftime("%d/%m/%Y %H:%M"))).pack(side="left", padx=(6, 0))
        ttk.Label(r1, text="gg/mm/aaaa oo:mm", style="Hint.TLabel").pack(side="left", padx=(8, 0))
        ttk.Label(body, text="N. ticket").grid(row=2, column=0, sticky="w", pady=5, padx=(0, 14))
        r2 = ttk.Frame(body)
        r2.grid(row=2, column=1, sticky="w", pady=5)
        ttk.Entry(r2, textvariable=self.v_ticket, width=22).pack(side="left")
        ttk.Label(r2, text="Stato").pack(side="left", padx=(24, 6))
        ttk.Combobox(r2, textvariable=self.v_stato, values=STATI_TECNICO, state="readonly",
                     width=12).pack(side="left")
        ttk.Label(body, text="Il tecnico è venuto\nil giorno").grid(
            row=3, column=0, sticky="w", pady=5, padx=(0, 14))
        r3 = ttk.Frame(body)
        r3.grid(row=3, column=1, sticky="w", pady=5)
        ttk.Entry(r3, textvariable=self.v_dtec, width=14).pack(side="left")
        ttk.Label(r3, text="gg/mm/aaaa. Lasciala vuota se non è ancora venuto.",
                  style="Hint.TLabel").pack(side="left", padx=(8, 0))
        ttk.Label(body, text="Nome del tecnico").grid(row=4, column=0, sticky="w", pady=5,
                                                      padx=(0, 14))
        ttk.Entry(body, textvariable=self.v_tec, width=34).grid(row=4, column=1, sticky="w",
                                                                pady=5)
        ttk.Label(body, text="Problema segnalato *").grid(row=5, column=0, sticky="nw", pady=5,
                                                          padx=(0, 14))
        self.t_prob = make_text(body, 3)
        self.t_prob.grid(row=5, column=1, sticky="ew", pady=5)
        self.t_prob.insert("1.0", c.get("problema", ""))
        ttk.Label(body, text="Cosa ha fatto\nil tecnico").grid(row=6, column=0, sticky="nw",
                                                              pady=5, padx=(0, 14))
        self.t_int = make_text(body, 4)
        self.t_int.grid(row=6, column=1, sticky="ew", pady=5)
        self.t_int.insert("1.0", c.get("intervento", ""))
        btns = ttk.Frame(body)
        btns.grid(row=7, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))

    def _stato_changed(self, *_):
        if self.v_stato.get() == "Eseguito" and not self.v_dtec.get().strip():
            self.v_dtec.set(fmt_date(date.today()))

    def save(self):
        dt = parse_datetime(self.v_dt.get())
        if not dt:
            messagebox.showwarning(APP_NAME, "La data della richiesta non è valida.\nScrivila "
                                   "così: gg/mm/aaaa oo:mm (esempio 13/11/2026 14:30).",
                                   parent=self)
            return
        dtec = ""
        if self.v_dtec.get().strip():
            d = parse_date(self.v_dtec.get())
            if not d:
                messagebox.showwarning(APP_NAME, "La data in cui è venuto il tecnico non è "
                                       "valida.\nScrivila così: gg/mm/aaaa.", parent=self)
                return
            dtec = d.isoformat()
        pid = self.printers.get(self.v_pr.get())
        if not pid:
            messagebox.showwarning(APP_NAME, "Scegli la stampante.", parent=self)
            return
        problema = self.t_prob.get("1.0", "end").strip()
        if not problema:
            messagebox.showwarning(APP_NAME, "Scrivi qual è il problema segnalato.", parent=self)
            return
        data = {"stampante_id": pid, "data_ora": dt.strftime("%Y-%m-%d %H:%M"),
                "data_tecnico": dtec, "ticket": self.v_ticket.get(), "tecnico": self.v_tec.get(),
                "problema": problema, "intervento": self.t_int.get("1.0", "end").strip(),
                "stato": self.v_stato.get()}
        if self.on_save(data):
            self.destroy()


class PrinterManutenzioniDialog(ModalDialog):
    """Elenco degli interventi del tecnico Konica su una singola stampante."""

    def __init__(self, app, printer_id, parent=None):
        super().__init__(parent or app, "Interventi del tecnico", "1000x470")
        self.minsize(820, 360)
        self.app = app
        self.printer_id = printer_id
        self.rows = {}
        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(1, weight=1)
        self.lbl_title = ttk.Label(body, text="", font=("Segoe UI", 13, "bold"))
        self.lbl_title.grid(row=0, column=0, sticky="w", pady=(0, 8))
        cols = ("data", "ticket", "stato", "data_tec", "tecnico", "problema", "intervento")
        self.tree = ttk.Treeview(body, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("data", "Richiesta del", 125, False), ("ticket", "Ticket", 80, False),
                ("stato", "Stato", 80, False), ("data_tec", "Tecnico venuto il", 110, False),
                ("tecnico", "Tecnico", 120, False), ("problema", "Problema segnalato", 220, True),
                ("intervento", "Intervento svolto", 220, True)):
            self.tree.heading(c, text=title)
            self.tree.column(c, width=width, anchor="w", stretch=stretch)
        self.tree.tag_configure("Richiesto", foreground=RED)
        self.tree.tag_configure("Eseguito", foreground=GREEN_DARK)
        sb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")
        self.tree.bind("<Double-1>", lambda e: self.edit())
        bottom = ttk.Frame(body)
        bottom.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(bottom, text="Nuovo intervento del tecnico…", style="Save.TButton",
                   command=self.new).pack(side="left")
        ttk.Button(bottom, text="Modifica…", command=self.edit).pack(side="left", padx=(8, 0))
        ttk.Button(bottom, text="Segna come eseguito", command=self.done).pack(
            side="left", padx=(8, 0))
        ttk.Button(bottom, text="Elimina", style="Danger.TButton", command=self.delete).pack(
            side="left", padx=(8, 0))
        ttk.Button(bottom, text="Chiudi", command=self.destroy).pack(side="right")
        self.refresh()

    def refresh(self):
        p = self.app.db.get_stampante(self.printer_id)
        if not p:
            self.destroy()
            return
        self.lbl_title.configure(text="%s - %s, %s" % (
            p["nome"] or p["modello"] or "Stampante", p["edificio"], floor_label(p["piano"])))
        self.tree.delete(*self.tree.get_children())
        self.rows = {}
        for m in self.app.db.manutenzioni_of(self.printer_id):
            self.rows[str(m["id"])] = m
            self.tree.insert("", "end", iid=str(m["id"]), tags=(m["stato"],), values=(
                fmt_dt(m["data_ora"]), m["ticket"], m["stato"], tec_date(m), m["tecnico"],
                one_line(m["problema"], 60), one_line(m["intervento"], 60)))

    def selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima un intervento dall'elenco.",
                                parent=self)
            return None
        return self.rows.get(sel[0])

    def new(self):
        self.app.open_manut_editor(self.printer_id, None, after=self.refresh, lock=True,
                                   parent=self)

    def edit(self):
        m = self.selected()
        if m:
            self.app.open_manut_editor(self.printer_id, m, after=self.refresh, lock=True,
                                       parent=self)

    def done(self):
        m = self.selected()
        if m:
            if m["stato"] == "Eseguito":
                messagebox.showinfo(APP_NAME, "Questo intervento è già segnato come eseguito.",
                                    parent=self)
                return
            self.app.set_manut_done(m)
            self.refresh()

    def delete(self):
        m = self.selected()
        if m and messagebox.askyesno(
                APP_NAME, "Eliminare questo intervento del tecnico?\n\n%s - %s" % (
                    fmt_dt(m["data_ora"]), one_line(m["problema"], 70)), parent=self):
            self.app.db.delete_manutenzione(m["id"])
            self.app.after_manut_change()
            self.refresh()


class ScortaDialog(ModalDialog):
    """Toner ricevuti (si aggiungono al magazzino) oppure correzione di quantità e soglia."""

    def __init__(self, parent, models, mode, modello, colore, quantita, soglia, on_save):
        super().__init__(parent, "Toner ricevuti" if mode == "add" else "Modifica scorta")
        self.resizable(False, False)
        self.on_save = on_save
        self.mode = mode
        self.v_mod = tk.StringVar(value=modello or (models[0] if models else ""))
        self.v_col = tk.StringVar(value=colore or TONER_NOMI[0])
        self.v_q = tk.StringVar(value=str(quantita))
        self.v_s = tk.StringVar(value=str(soglia))
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Modello stampante").grid(row=0, column=0, sticky="w", pady=5,
                                                       padx=(0, 14))
        ttk.Combobox(body, textvariable=self.v_mod, values=models, width=30,
                     state="readonly" if (mode == "set" and modello) else "normal").grid(
            row=0, column=1, sticky="w", pady=5)
        ttk.Label(body, text="Toner / vaschetta").grid(row=1, column=0, sticky="w", pady=5,
                                                      padx=(0, 14))
        ttk.Combobox(body, textvariable=self.v_col, values=TONER_NOMI, width=16,
                     state="disabled" if (mode == "set" and colore) else "readonly").grid(
            row=1, column=1, sticky="w", pady=5)
        if mode == "add":
            ttk.Label(body, text="Quantità ricevuta").grid(row=2, column=0, sticky="w", pady=5,
                                                          padx=(0, 14))
            ttk.Spinbox(body, from_=1, to=500, textvariable=self.v_q, width=6).grid(
                row=2, column=1, sticky="w", pady=5)
            hint = ("Li aggiungo a quelli che hai già in magazzino. Se è il primo inserimento "
                    "per questo toner, parto da qui.")
        else:
            ttk.Label(body, text="Quantità in magazzino").grid(row=2, column=0, sticky="w", pady=5,
                                                              padx=(0, 14))
            ttk.Spinbox(body, from_=0, to=999, textvariable=self.v_q, width=6).grid(
                row=2, column=1, sticky="w", pady=5)
            ttk.Label(body, text="Avvisami quando\nne restano (o meno)").grid(
                row=3, column=0, sticky="w", pady=5, padx=(0, 14))
            ttk.Spinbox(body, from_=0, to=99, textvariable=self.v_s, width=6).grid(
                row=3, column=1, sticky="w", pady=5)
            hint = "Usalo per correggere la quantità dopo un conteggio in magazzino."
        ttk.Label(body, style="Hint.TLabel", wraplength=380, justify="left", text=hint).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))
        btns = ttk.Frame(body)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))

    def save(self):
        try:
            q = int(self.v_q.get())
            s = int(self.v_s.get())
        except ValueError:
            messagebox.showwarning(APP_NAME, "Le quantità devono essere numeri interi.",
                                   parent=self)
            return
        if q < 0 or s < 0 or (self.mode == "add" and q < 1):
            messagebox.showwarning(APP_NAME, "Controlla le quantità: non possono essere negative "
                                   "(e i toner ricevuti sono almeno 1).", parent=self)
            return
        data = {"modello": self.v_mod.get(), "colore": self.v_col.get(), "quantita": q,
                "soglia": s}
        if self.on_save(self.mode, data):
            self.destroy()


class ModelRenameDialog(ModalDialog):
    """Cambia il nome di un modello su tutte le stampanti e nella scorta toner."""

    def __init__(self, parent, models, count_fn, on_apply):
        super().__init__(parent, "Correggi il nome di un modello")
        self.resizable(False, False)
        self.count_fn = count_fn
        self.on_apply = on_apply
        self.v_old = tk.StringVar(value=models[0] if models else "")
        self.v_new = tk.StringVar(value="")
        self.v_info = tk.StringVar(value="")
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Nome da correggere").grid(row=0, column=0, sticky="w", pady=5,
                                                        padx=(0, 14))
        cb = ttk.Combobox(body, textvariable=self.v_old, values=models, width=30,
                          state="readonly")
        cb.grid(row=0, column=1, sticky="w", pady=5)
        cb.bind("<<ComboboxSelected>>", lambda e: self.update_info())
        ttk.Label(body, text="Nome giusto").grid(row=1, column=0, sticky="w", pady=5, padx=(0, 14))
        ttk.Entry(body, textvariable=self.v_new, width=32).grid(row=1, column=1, sticky="w",
                                                                pady=5)
        ttk.Label(body, style="Hint.TLabel", wraplength=380, justify="left",
                  textvariable=self.v_info).grid(row=2, column=0, columnspan=2, sticky="w",
                                                 pady=(6, 0))
        btns = ttk.Frame(body)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Cambia nome", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))
        self.update_info()

    def update_info(self):
        n_p, n_s = self.count_fn(self.v_old.get())
        self.v_info.set("Lo cambio su %d stampant%s e su %d voc%s della scorta toner. Se per una "
                        "sola stampante il modello è diverso, lascia stare qui e cambialo dalla "
                        "scheda di quella stampante." % (
                            n_p, "e" if n_p == 1 else "i", n_s, "e" if n_s == 1 else "i"))

    def save(self):
        old, new = self.v_old.get(), " ".join(self.v_new.get().split())
        if not old:
            messagebox.showwarning(APP_NAME, "Non c'è nessun modello da correggere.", parent=self)
            return
        if not new:
            messagebox.showwarning(APP_NAME, "Scrivi il nome giusto del modello.", parent=self)
            return
        if model_key(old) == model_key(new) and old == new:
            messagebox.showwarning(APP_NAME, "Il nome nuovo è uguale a quello vecchio.",
                                   parent=self)
            return
        if self.on_apply(old, new):
            self.destroy()


class PeopleImportDialog(ModalDialog):
    """Importa le persone da una o più liste Excel: per ogni file si sceglie la sede."""

    def __init__(self, parent, files, sedi, preview, on_apply):
        super().__init__(parent, "Importa persone da Excel")
        self.resizable(False, False)
        self.files, self.preview, self.on_apply = files, preview, on_apply
        self.v_sede = [tk.StringVar(value=f["sede"]) for f in files]
        self.v_upd = tk.BooleanVar(value=False)
        self.v_info = tk.StringVar(value="")
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Importa le persone dai file Excel",
                  font=("Segoe UI", 11, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(body, style="Hint.TLabel", wraplength=480, justify="left",
                  text="Per ogni persona scrivo matricola (Username), cellulare e sede, ma solo "
                       "dove nell'app mancano: quello che hai già scritto non viene toccato. "
                       "Le persone che non ci sono vengono aggiunte. Scegli la sede di ogni file."
                  ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(2, 10))
        for i, f in enumerate(files):
            ttk.Label(body, justify="left", text="%s\n%d persone, foglio «%s»"
                      % (f["name"], len(f["persone"]), f["title"])).grid(
                row=2 + i, column=0, sticky="w", pady=5, padx=(0, 14))
            cb = ttk.Combobox(body, textvariable=self.v_sede[i], values=sedi, width=18,
                              state="readonly")
            cb.grid(row=2 + i, column=1, sticky="w", pady=5)
            cb.bind("<<ComboboxSelected>>", lambda e: self.update_info())
        n = len(files)
        ttk.Checkbutton(body, text="Cambia la sede anche a chi è già nell'elenco",
                        variable=self.v_upd, command=self.update_info).grid(
            row=2 + n, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(body, style="Hint.TLabel", wraplength=480, justify="left",
                  textvariable=self.v_info).grid(row=3 + n, column=0, columnspan=2, sticky="w",
                                                 pady=(10, 0))
        btns = ttk.Frame(body)
        btns.grid(row=4 + n, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Importa", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))
        self.update_info()

    def current(self):
        return [{"name": f["name"], "sede": v.get(), "persone": f["persone"]}
                for f, v in zip(self.files, self.v_sede)]

    def update_info(self):
        self.v_info.set(self.preview(self.current(), self.v_upd.get()))

    def save(self):
        lists = self.current()
        for lst in lists:
            if not lst["sede"]:
                messagebox.showwarning(APP_NAME, "Scegli la sede del file «%s»." % lst["name"],
                                       parent=self)
                return
        if self.on_apply(lists, self.v_upd.get()):
            self.destroy()


class UserExportDialog(ModalDialog):
    """Scelta dei dati degli utenti da estrarre (IMEI, seriale PC, ...)."""

    def __init__(self, parent, count, scope, on_export):
        super().__init__(parent, "Esporta dati utenti")
        self.resizable(False, False)
        self.on_export = on_export
        self.vars = {c[0]: tk.BooleanVar(value=c[2]) for c in USER_COLUMNS}
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Quali dati vuoi estrarre?", font=("Segoe UI", 11, "bold")).pack(
            anchor="w")
        ttk.Label(body, style="Hint.TLabel", wraplength=380, justify="left",
                  text="Utenti inclusi: %d (%s)" % (count, scope)).pack(anchor="w", pady=(2, 8))
        for key, title, _, _, _ in USER_COLUMNS:
            ttk.Checkbutton(body, text=title, variable=self.vars[key]).pack(anchor="w", pady=2)
        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Esporta in PDF", style="Primary.TButton",
                   command=lambda: self.go("pdf")).pack(side="right", padx=(0, 8))
        ttk.Button(btns, text="Esporta in Excel", style="Primary.TButton",
                   command=lambda: self.go("xlsx")).pack(side="right", padx=(0, 8))

    def go(self, kind):
        keys = [k for k, v in self.vars.items() if v.get()]
        if not keys:
            messagebox.showinfo(APP_NAME, "Spunta almeno un dato da estrarre.", parent=self)
            return
        if self.on_export(kind, keys):
            self.destroy()


# ---------------------------------------------------------------------------
# Verbali Word: finestre
# ---------------------------------------------------------------------------
VERB_LIST_KEYS = ("materiale", "accessori")
PC_WORDS = ("notebook", "pc", "portatile", "computer", "desktop")
PHONE_WORDS = ("smartphone", "cellulare", "telefono", "tablet")
TITOLI_ELENCO = ["", "avv.", "dott.", "dott.ssa", "ing.", "arch.", "geom.", "rag.", "sig.",
                 "sig.ra", "prof.", "prof.ssa"]


class ScrollFrame(ttk.Frame):
    """Contenitore con barra di scorrimento verticale (si scorre anche con la rotellina)."""

    def __init__(self, parent):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, bg=LIGHT, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = ttk.Frame(self.canvas, padding=16)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._win, width=e.width))

    def wheel(self, event):
        if event.widget.winfo_class() in ("Treeview", "Text", "TCombobox", "Listbox"):
            return
        self.canvas.yview_scroll(int(-event.delta / 120), "units")


class PasteDialog(ModalDialog):
    """Finestra dove incollare un elenco (per esempio copiato da Excel)."""

    def __init__(self, parent, title, intro, on_ok):
        super().__init__(parent, title, "640x430")
        self.on_ok = on_ok
        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=intro, style="Hint.TLabel", wraplength=600, justify="left").pack(
            anchor="w")
        self.txt = make_text(body, 12)
        self.txt.pack(fill="both", expand=True, pady=8)
        self.txt.focus_set()
        btns = ttk.Frame(body)
        btns.pack(fill="x")
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Aggiungi all'elenco", style="Save.TButton",
                   command=self.ok).pack(side="right", padx=(0, 8))

    def ok(self):
        text = self.txt.get("1.0", "end").strip()
        if text:
            self.on_ok(text)
        self.destroy()


class RoomNewDialog(ModalDialog):
    """Una nuova stanza: codice (anche con lettere, per esempio T12) e scrivanie totali (facoltative)."""

    def __init__(self, parent, titolo, on_ok):
        super().__init__(parent, "Nuova stanza")
        self.resizable(False, False)
        self.on_ok = on_ok
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text=titolo, style="Hint.TLabel").grid(row=0, column=0, columnspan=2,
                                                                sticky="w", pady=(0, 8))
        ttk.Label(body, text="Codice della stanza *").grid(row=1, column=0, sticky="w", pady=5,
                                                           padx=(0, 14))
        self.v_cod = tk.StringVar()
        e = ttk.Entry(body, textvariable=self.v_cod, width=20)
        e.grid(row=1, column=1, sticky="w", pady=5)
        e.focus_set()
        ttk.Label(body, text="Scrivanie totali (facoltativo)").grid(row=2, column=0, sticky="w",
                                                                    pady=5, padx=(0, 14))
        self.v_scr = tk.StringVar()
        ttk.Entry(body, textvariable=self.v_scr, width=8).grid(row=2, column=1, sticky="w", pady=5)
        ttk.Label(body, style="Hint.TLabel", wraplength=360, justify="left",
                  text="Le postazioni (Postazione 1, 2, 3...) le numera il programma da solo "
                       "quando aggiungi le persone.").grid(row=3, column=0, columnspan=2,
                                                          sticky="w", pady=(6, 0))
        btns = ttk.Frame(body)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Crea la stanza", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))
        self.bind("<Return>", lambda e: self.save())

    def save(self):
        if self.on_ok(self.v_cod.get(), self.v_scr.get()):
            self.destroy()


class DeviceCodeDialog(ModalDialog):
    """Il codice di un dispositivo della scrivania (S/N del PC, del monitor AOC o del telefono Teams,
    IMEI del cellulare): si scrive a mano oppure si legge con la fotocamera."""

    def __init__(self, parent, device, person_name, on_save):
        presa = device["kind"] == "presa"
        super().__init__(parent, device["name"], "620x430" if presa else "620x360")
        self.resizable(False, False)
        self.device, self.on_save = device, on_save
        imei = device["kind"] == "imei"
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        ttk.Label(body, text=device["name"] + ("  ·  " + device["sub"] if device["sub"] else ""),
                  font=("Segoe UI", 14, "bold"), foreground=NAVY).grid(row=0, column=0,
                                                                       columnspan=2, sticky="w")
        ttk.Label(body, text=person_name, style="Hint.TLabel").grid(row=1, column=0, columnspan=2,
                                                                   sticky="w", pady=(0, 12))
        ttk.Label(body, text=device["label"], font=("Segoe UI", 9, "bold")).grid(
            row=2, column=0, columnspan=2, sticky="w")
        self.v = tk.StringVar(value=device["value"])
        self.entry = ttk.Entry(body, textvariable=self.v, width=34, font=("Consolas", 15, "bold"))
        self.entry.grid(row=3, column=0, sticky="ew", pady=(2, 2))
        self.btn_cam = scan_button(body, self.v, kind="imei" if imei else "any" if presa else "serial",
                                   entry=self.entry,
                                   on_code=lambda t: self.lbl_err.configure(text=""),
                                   title="Leggi l'IMEI" if imei else "Leggi: " + device["label"])
        self.btn_cam.grid(row=3, column=1, padx=(8, 0))
        ttk.Label(body, style="Hint.TLabel", wraplength=560, justify="left",
                  text=("L'IMEI sono 15 cifre: lo trovi sulla scatola o digitando *#06# sul cellulare."
                        if imei else
                        "Il codice della presa è scritto sulla placca a muro (per esempio A12-34). "
                        "Puoi scriverlo oppure leggerlo con la fotocamera se c'è un'etichetta con il "
                        "codice a barre o il QR." if presa else
                        "Il S/N è sull'etichetta dell'apparecchio (di solito sotto o dietro), vicino "
                        "al codice a barre. Puoi scriverlo oppure leggerlo con la fotocamera.")
                  ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))
        self.v_port = tk.StringVar(value=device.get("port", ""))
        if presa:
            ttk.Label(body, text="Porta dello switch (facoltativa)", font=("Segoe UI", 9, "bold")).grid(
                row=5, column=0, columnspan=2, sticky="w", pady=(10, 0))
            ttk.Entry(body, textvariable=self.v_port, width=34, font=("Consolas", 13)).grid(
                row=6, column=0, sticky="ew", pady=(2, 0))
        self.lbl_err = ttk.Label(body, foreground=RED, wraplength=560, justify="left")
        self.lbl_err.grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 0))
        bt = ttk.Frame(body)
        bt.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(16, 0))
        ttk.Button(bt, text="Salva", style="Save.TButton", command=self.save).pack(side="left")
        if device["value"]:
            ttk.Button(bt, text="Cancella il codice", command=self.clear).pack(side="left",
                                                                              padx=(8, 0))
        ttk.Button(bt, text="Annulla", command=self.destroy).pack(side="right")
        self.bind("<Return>", lambda e: self.save())
        try:
            self.entry.focus_set()
        except Exception:
            pass

    def save(self):
        value, err = check_device_code(self.device["kind"], self.v.get())
        if err:
            self.lbl_err.configure(text=err)
            return
        if not value:
            self.lbl_err.configure(text="Scrivi il codice oppure leggilo con la fotocamera.")
            return
        if self.device["kind"] == "presa":
            self.on_save(value, " ".join(self.v_port.get().split())[:80])
        else:
            self.on_save(value)
        self.destroy()

    def clear(self):
        if self.device["kind"] == "presa":
            self.on_save("", "")
        else:
            self.on_save("")
        self.destroy()


class DeskDialog(ModalDialog):
    """La scrivania a pagina intera: la scena in grande con un bollino su ogni dispositivo (PC, monitor
    AOC, cellulare, telefono Teams); cliccando un bollino si scrive il codice (S/N o IMEI) a mano o con
    la fotocamera. A destra i dati della persona, i modelli sulla scrivania e la dotazione."""

    NO_CELL = "Non indicato"
    SCENE_W, SCENE_H, SCENE_Y = 130 * DESK_BIG + 40, 66 * DESK_BIG + 24, 12

    def __init__(self, parent, app, room_id, utente_id, open_room=None, on_change=None):
        super().__init__(parent, "Postazione", "1240x780")
        self.minsize(960, 640)
        self.app, self.room_id, self.utente_id, self.open_room_cb = app, room_id, utente_id, open_room
        self.on_change = on_change
        self.devices = []
        self.after(60, self._maximize)
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(1, weight=1)
        head = ttk.Frame(body)
        head.grid(row=0, column=0, columnspan=2, sticky="ew")
        self.lbl_title = ttk.Label(head, font=("Segoe UI", 14, "bold"), foreground=NAVY)
        self.lbl_title.pack(anchor="w")
        self.lbl_where = ttk.Label(head, style="Hint.TLabel")
        self.lbl_where.pack(anchor="w", pady=(0, 8))
        # a sinistra: la scrivania in grande e i suoi dispositivi
        left = ttk.Frame(body)
        left.grid(row=1, column=0, sticky="nsew", padx=(0, 18))
        self.lbl_person = ttk.Label(left, font=("Segoe UI", 16, "bold"))
        self.lbl_person.pack(anchor="w")
        self.lbl_office = ttk.Label(left, style="Hint.TLabel")
        self.lbl_office.pack(anchor="w", pady=(0, 6))
        self.scene = tk.Canvas(left, width=self.SCENE_W, height=self.SCENE_H, bg="#f6f2e6",
                               highlightthickness=1, highlightbackground="#cfc3a3")
        self.scene.pack(anchor="w")
        self.scene.bind("<Button-1>", self.on_scene_click)
        self.scene.bind("<Motion>", self.on_scene_move)
        ttk.Label(left, style="Hint.TLabel",
                  text="Clicca un bollino per scrivere il codice o leggerlo con la fotocamera: "
                       "verde = codice scritto, rosso = da inserire.").pack(anchor="w", pady=(4, 6))
        self.dev_frame = ttk.Frame(left)
        self.dev_frame.pack(fill="x")
        # a destra: dati, modelli, dotazione
        right = ttk.Frame(body)
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)
        ttk.Label(right, text="Numero di serie del PC (S/N)", font=("Segoe UI", 9, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        self.v_sn = tk.StringVar()
        self.e_sn = ttk.Entry(right, textvariable=self.v_sn, state="readonly", width=24,
                              font=("Consolas", 13, "bold"))
        self.e_sn.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 2))
        ttk.Button(right, text="Copia S/N", command=self.copy_sn).grid(row=1, column=2, padx=(8, 0))
        self.lbl_copy = ttk.Label(right, style="Hint.TLabel")
        self.lbl_copy.grid(row=2, column=0, columnspan=3, sticky="w", pady=(0, 8))
        self.vals = {}
        for i, (key, text) in enumerate((("imei", "IMEI cellulare"), ("telefono", "Telefono"),
                                         ("matricola", "Matricola"), ("sede", "Sede"))):
            ttk.Label(right, text=text).grid(row=3 + i, column=0, sticky="w", pady=2, padx=(0, 14))
            self.vals[key] = ttk.Label(right, font=("Segoe UI", 10, "bold"))
            self.vals[key].grid(row=3 + i, column=1, columnspan=2, sticky="w", pady=2)
        pf = ttk.Frame(right)
        pf.grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))
        ttk.Label(pf, text="Cellulare sulla scrivania").pack(side="left")
        self.v_mob = tk.StringVar()
        self.cb_mob = ttk.Combobox(pf, textvariable=self.v_mob, state="readonly", width=16,
                                   values=[self.NO_CELL] + list(CELL_MODELLI))
        self.cb_mob.pack(side="left", padx=(8, 0))
        self.cb_mob.bind("<<ComboboxSelected>>", lambda e: self.apply_phones())
        pt = ttk.Frame(right)
        pt.grid(row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))
        self.v_teams = tk.BooleanVar()
        ttk.Checkbutton(pt, text="Telefono Teams (AudioCodes)", variable=self.v_teams,
                        command=self.apply_phones).pack(side="left")
        pl = ttk.Frame(right)
        pl.grid(row=9, column=0, columnspan=3, sticky="w", pady=(6, 0))
        ttk.Label(pl, text="Portatile").pack(side="left")
        self.v_lap = tk.StringVar()
        self.cb_lap = ttk.Combobox(pl, textvariable=self.v_lap, state="readonly", width=22,
                                   values=[self.NO_CELL] + list(LAPTOP_MODELLI.values()))
        self.cb_lap.pack(side="left", padx=(8, 0))
        self.cb_lap.bind("<<ComboboxSelected>>", lambda e: self.apply_phones())
        # le foto dei tre apparecchi scelti: cambiano quando cambi la scelta
        fr = ttk.Frame(right)
        fr.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        self.pic, self.pic_cap = {}, {}
        for slot in ("laptop", "mobile", "teams"):
            col = ttk.Frame(fr)
            col.pack(side="left", padx=(0, 18), anchor="s")
            self.pic[slot] = ttk.Label(col, anchor="center")
            self.pic[slot].pack()
            self.pic_cap[slot] = ttk.Label(col, style="Hint.TLabel", anchor="center")
            self.pic_cap[slot].pack()
        ttk.Label(right, text="Dotazione", font=("Segoe UI", 9, "bold")).grid(
            row=11, column=0, columnspan=3, sticky="w", pady=(10, 2))
        self.tree_d = ttk.Treeview(right, columns=("tipo", "modello", "seriale"), show="headings",
                                   height=5, selectmode="browse")
        for col, text, w in (("tipo", "Apparecchio", 110), ("modello", "Modello", 150),
                             ("seriale", "Seriale / IMEI", 170)):
            self.tree_d.heading(col, text=text)
            self.tree_d.column(col, width=w, anchor="w")
        self.tree_d.grid(row=12, column=0, columnspan=3, sticky="nsew")
        right.rowconfigure(12, weight=1)
        bt = ttk.Frame(right)
        bt.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        ttk.Button(bt, text="Modifica la persona…", style="Save.TButton", command=self.edit).pack(
            side="left")
        if open_room:
            ttk.Button(bt, text="Apri la stanza…", command=self.go_room).pack(side="left",
                                                                             padx=(8, 0))
        ttk.Button(bt, text="Chiudi", command=self.destroy).pack(side="right")
        self.refresh()

    def _maximize(self):
        """La scrivania si apre a pagina intera (su Windows «zoomed»; altrove se il sistema lo sa)."""
        try:
            self.state("zoomed")
        except Exception:
            try:
                self.attributes("-zoomed", True)
            except Exception:
                pass

    def scene_cx(self):
        return self.SCENE_W / 2

    def on_scene_click(self, event):
        key = desk_bollino_hit(self.devices, self.scene_cx(), self.SCENE_Y, event.x, event.y)
        if key:
            self.open_device(key)

    def on_scene_move(self, event):
        key = desk_bollino_hit(self.devices, self.scene_cx(), self.SCENE_Y, event.x, event.y)
        try:
            self.scene.configure(cursor="hand2" if key else "")
        except Exception:
            pass

    def open_device(self, key):
        p = self.person()
        d = next((x for x in self.devices if x["key"] == key), None)
        if not p or not d:
            return None
        return DeviceCodeDialog(self, d, p["nome"],
                                lambda value, port=None: self.save_device(key, value, port))

    def save_device(self, key, value, port=None):
        """Salva il codice di un dispositivo (per la presa di rete anche la porta dello switch) e
        ridisegna la scrivania."""
        p = self.person()
        if not p:
            return False
        changed = set_device_code(self.app.db, self.utente_id, key, value, p.get("dotazione"), port)
        self.app.refresh_users()
        self.refresh()
        if self.on_change:
            self.on_change()
        return changed

    def draw_devices(self, p, info):
        """Ridisegna la scena grande con i bollini e l'elenco dei dispositivi sotto."""
        self.devices = desk_devices(p)
        self.scene.delete("all")
        draw_desk_big(self.scene, info, self.devices, self.scene_cx(), self.SCENE_Y)
        for w in self.dev_frame.winfo_children():
            w.destroy()
        for d in self.devices:
            ok = bool(d["value"])
            row = ttk.Frame(self.dev_frame)
            row.pack(fill="x", pady=2)
            dot = tk.Canvas(row, width=26, height=26, highlightthickness=0, bg=LIGHT)
            dot.create_oval(2, 2, 24, 24, fill=GREEN if ok else RED, outline=GREEN if ok else RED)
            dot.create_text(13, 13, text="\u2713" if ok else "+", fill=WHITE,
                            font=("Segoe UI", 11, "bold"))
            dot.pack(side="left")
            ttk.Label(row, text=d["name"] + ("  ·  " + d["sub"] if d["sub"] else ""),
                      font=("Segoe UI", 10, "bold"), width=40).pack(side="left", padx=(8, 0))
            ttk.Label(row, text=d["value"] or ("IMEI mancante" if d["kind"] == "imei"
                                               else "Codice mancante" if d["kind"] == "presa"
                                               else "S/N mancante"),
                      font=("Consolas", 11, "bold") if ok else ("Segoe UI", 10, "italic"),
                      foreground=TEXT if ok else GREY).pack(side="left", padx=(8, 0))
            ttk.Button(row, text="Modifica…" if ok else "Inserisci…",
                       command=lambda k=d["key"]: self.open_device(k)).pack(side="right")

    def person(self):
        return next((p for p in self.app.db.persone_of(self.room_id)
                     if p["utente_id"] == self.utente_id), None)

    def refresh(self):
        room, p = self.app.db.get_stanza(self.room_id), self.person()
        if not room or not p:
            self.destroy()
            return
        self.lbl_title.configure(text="%s · Postazione %d" % (room["codice"], p["numero"]))
        self.lbl_where.configure(text="%s · %s" % (room["edificio"], floor_label(room["piano"])))
        self.lbl_person.configure(text=p["nome"])
        self.lbl_office.configure(text=p["ufficio"] or "Ufficio non indicato")
        self.v_sn.set(p["seriale_pc"] or "(non ancora inserito)")
        self.sn = p["seriale_pc"] or ""
        self.lbl_copy.configure(text="")
        for key, lab in self.vals.items():
            lab.configure(text=p.get(key) or "—")
        info = desk_phone_info(p.get("dotazione"), p.get("imei"))
        values = [self.NO_CELL] + list(CELL_MODELLI)
        if info["mobile"] == "altro":
            self.v_mob.set("Altro: %s" % info["label"])
            values.append(self.v_mob.get())
        else:
            self.v_mob.set(info["label"] if info["label"] in CELL_MODELLI else self.NO_CELL)
        self.cb_mob.configure(values=values)
        self.v_teams.set(info["teams"])
        lvals = [self.NO_CELL] + list(LAPTOP_MODELLI.values())
        if info["laptop"] == "altro":
            self.v_lap.set("Altro: %s" % info["laptop_label"])
            lvals.append(self.v_lap.get())
        else:
            self.v_lap.set(LAPTOP_MODELLI.get(info["laptop"], self.NO_CELL))
        self.cb_lap.configure(values=lvals)
        self.draw_devices(p, info)
        keys = desk_photo_keys(info)
        caps = {"laptop": (LAPTOP_MODELLI.get(info["laptop"]) or info["laptop_label"]
                           or "Portatile non indicato"),
                "mobile": (("Samsung Galaxy " + info["label"]) if keys["mobile"] else
                           (info["label"] or "Cellulare non indicato")),
                "teams": (TEAMS_MODELLO if info["teams"] else "Telefono Teams: no")}
        for slot in ("laptop", "mobile", "teams"):
            img = desk_photo(keys[slot] + "_m") if keys[slot] else None
            self.pic[slot].configure(image=img or "")
            self.pic[slot].image = img                  # il programma deve tenerla
            self.pic_cap[slot].configure(text=caps[slot])
        self.tree_d.delete(*self.tree_d.get_children())
        for i, r in enumerate(p.get("dotazione") or []):
            self.tree_d.insert("", "end", iid=str(i),
                               values=(r["tipo"], r["modello"], r["seriale"]))

    def apply_phones(self):
        """Cambia cellulare, portatile e/o telefono Teams: la scrivania cambia foto."""
        p = self.person()
        if not p:
            return
        choice = self.v_mob.get()
        mob = None if choice.startswith("Altro:") else ("" if choice == self.NO_CELL else choice)
        lchoice = self.v_lap.get()
        lap = (None if lchoice.startswith("Altro:") else
               next((k for k, v in LAPTOP_MODELLI.items() if v == lchoice), ""))
        new = with_desk_phones(p.get("dotazione"), mob, bool(self.v_teams.get()), lap)
        self.app.db.set_user_extra(self.utente_id, dotazione=new)
        self.app.refresh_users()
        self.refresh()
        if self.on_change:
            self.on_change()

    def copy_sn(self):
        if not self.sn:
            messagebox.showinfo(APP_NAME, "Il numero di serie di questo PC non è ancora stato "
                                          "inserito: scrivilo da «Modifica la persona…».",
                                parent=self)
            return
        self.clipboard_clear()
        self.clipboard_append(self.sn)
        self.lbl_copy.configure(text="Copiato negli appunti: %s" % self.sn)

    def edit(self):
        self.app.edit_person_by_id(self.utente_id, parent=self, after=self.refresh)

    def go_room(self):
        if self.open_room_cb:
            self.destroy()
            self.open_room_cb()


class RoomDialog(ModalDialog):
    """Scheda di una stanza: codice, scrivanie, come arrivarci, chi siede dove, foto e panorami."""

    def __init__(self, app, room_id):
        super().__init__(app, "Stanza", "1020x880")
        self.minsize(900, 700)
        self.app = app
        self.room_id = room_id
        room = app.db.get_stanza(room_id)
        self.v_codice = tk.StringVar(value=room["codice"])
        self.v_scriv = tk.StringVar(value=str(room["scrivanie"]) if room["scrivanie"] else "")
        self.v_persona = tk.StringVar()
        body = ttk.Frame(self, padding=14)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        self.lbl_head = ttk.Label(body, style="Hint.TLabel")
        self.lbl_head.grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(body, text="Codice stanza *").grid(row=1, column=0, sticky="w", pady=4, padx=(0, 12))
        ttk.Entry(body, textvariable=self.v_codice, width=18).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(body, text="Scrivanie totali").grid(row=2, column=0, sticky="w", pady=4, padx=(0, 12))
        fr = ttk.Frame(body)
        fr.grid(row=2, column=1, sticky="w", pady=4)
        ttk.Entry(fr, textvariable=self.v_scriv, width=8).pack(side="left")
        ttk.Label(fr, text="  (facoltativo: se lo indichi vedi «occupate 2/3» e ti avviso quando "
                           "sono troppe)", style="Hint.TLabel").pack(side="left")
        ttk.Label(body, text="Come arrivarci").grid(row=3, column=0, sticky="nw", pady=4, padx=(0, 12))
        self.txt_come = make_text(body, 2)
        self.txt_come.grid(row=3, column=1, sticky="ew", pady=4)
        self.txt_come.insert("1.0", room["come"])
        ttk.Label(body, text="Note").grid(row=4, column=0, sticky="nw", pady=4, padx=(0, 12))
        self.txt_note = make_text(body, 2)
        self.txt_note.grid(row=4, column=1, sticky="ew", pady=4)
        self.txt_note.insert("1.0", room["note"])
        ttk.Button(body, text="Salva la stanza", style="Save.TButton", command=self.save).grid(
            row=5, column=1, sticky="w", pady=(2, 8))

        ttk.Label(body, text="Chi siede qui", font=("Segoe UI", 10, "bold")).grid(
            row=6, column=0, sticky="w", pady=(6, 2))
        self.lbl_occ = ttk.Label(body, style="Hint.TLabel")
        self.lbl_occ.grid(row=6, column=1, sticky="w", pady=(6, 2))
        self.nb_p = ttk.Notebook(body)
        self.nb_p.grid(row=7, column=0, columnspan=2, sticky="ew")
        tab_d, tab_e = ttk.Frame(self.nb_p), ttk.Frame(self.nb_p)
        self.nb_p.add(tab_d, text="Scrivanie")
        self.nb_p.add(tab_e, text="Elenco")
        self.dcanvas = tk.Canvas(tab_d, bg="#efe7cf", height=176, highlightthickness=0)
        dsb = ttk.Scrollbar(tab_d, orient="vertical", command=self.dcanvas.yview)
        self.dcanvas.configure(yscrollcommand=dsb.set)
        self.dcanvas.pack(side="left", fill="both", expand=True)
        dsb.pack(side="right", fill="y")
        self.dcanvas.bind("<Button-1>", self.on_desk_click)
        self.dcanvas.bind("<Configure>", lambda e: self.draw_desks())
        self.dcanvas.bind("<MouseWheel>",
                          lambda e: self.dcanvas.yview_scroll(int(-e.delta / 120), "units"))
        self.desk_lay = {"rows": [], "height": 0}
        self.tree_p = ttk.Treeview(tab_e, columns=("posto", "nome", "ufficio", "tel", "pc", "imei",
                                                   "dot"),
                                   show="headings", height=5, selectmode="browse")
        for col, text, w in (("posto", "Postazione", 90), ("nome", "Persona", 150),
                             ("ufficio", "Ufficio", 150), ("tel", "Telefono", 90),
                             ("pc", "Seriale PC", 110), ("imei", "IMEI", 120),
                             ("dot", "Dotazione", 170)):
            self.tree_p.heading(col, text=text)
            self.tree_p.column(col, width=w, anchor="w")
        self.tree_p.pack(fill="x")
        self.tree_p.bind("<Double-1>", lambda e: self.edit_person())
        pr = ttk.Frame(body)
        pr.grid(row=8, column=0, columnspan=2, sticky="w", pady=(6, 8))
        pr1 = ttk.Frame(pr)
        pr1.pack(anchor="w")
        ttk.Button(pr1, text="Nuova persona…", style="Save.TButton",
                   command=self.new_person).pack(side="left")
        ttk.Button(pr1, text="Modifica la persona scelta…", command=self.edit_person).pack(
            side="left", padx=(8, 0))
        ttk.Button(pr1, text="Togli dalla stanza", command=self.remove_person).pack(
            side="left", padx=(8, 0))
        pr2 = ttk.Frame(pr)
        pr2.pack(anchor="w", pady=(6, 0))
        ttk.Label(pr2, text="Persona già in elenco:").pack(side="left")
        self.cb_persona = ttk.Combobox(pr2, textvariable=self.v_persona, width=34,
                                       values=app.users_cache)
        self.cb_persona.pack(side="left", padx=(8, 0))
        bind_autocomplete(self.cb_persona, self.v_persona, lambda: app.users_cache)
        ttk.Button(pr2, text="Aggiungi alla stanza", command=self.add_person).pack(
            side="left", padx=(8, 0))

        ttk.Label(body, text="Foto e panorami", font=("Segoe UI", 10, "bold")).grid(
            row=9, column=0, columnspan=2, sticky="w", pady=(6, 2))
        self.tree_f = ttk.Treeview(body, columns=("nome", "tipo"), show="headings", height=3,
                                   selectmode="browse")
        self.tree_f.heading("nome", text="File")
        self.tree_f.heading("tipo", text="Tipo")
        self.tree_f.column("nome", width=420, anchor="w")
        self.tree_f.column("tipo", width=110, anchor="w")
        self.tree_f.grid(row=10, column=0, columnspan=2, sticky="ew")
        self.tree_f.bind("<Double-1>", lambda e: self.open_photo())
        fb = ttk.Frame(body)
        fb.grid(row=11, column=0, columnspan=2, sticky="w", pady=(6, 8))
        ttk.Button(fb, text="Aggiungi foto…", command=lambda: self.add_photos("foto")).pack(
            side="left")
        ttk.Button(fb, text="Aggiungi panorama…",
                   command=lambda: self.add_photos("panorama")).pack(side="left", padx=(8, 0))
        ttk.Button(fb, text="Apri", command=self.open_photo).pack(side="left", padx=(8, 0))
        ttk.Button(fb, text="Cambia tipo (foto / panorama)", command=self.toggle_tipo).pack(
            side="left", padx=(8, 0))
        ttk.Button(fb, text="Togli", style="Danger.TButton", command=self.del_photo).pack(
            side="left", padx=(8, 0))
        ttk.Label(body, style="Hint.TLabel", wraplength=800, justify="left",
                  text="Le foto vengono ridotte (1600 px, i panorami 2600 px) per tenere leggero "
                       "il file di scambio col telefono. I video restano solo sul telefono.").grid(
            row=12, column=0, columnspan=2, sticky="w")

        ttk.Label(body, text="Posizione sulla pianta", font=("Segoe UI", 10, "bold")).grid(
            row=13, column=0, columnspan=2, sticky="w", pady=(8, 2))
        pf = ttk.Frame(body)
        pf.grid(row=14, column=0, columnspan=2, sticky="w")
        self.lbl_pos = ttk.Label(pf)
        self.lbl_pos.pack(side="left")
        ttk.Button(pf, text="Segna sulla pianta", command=self.mark_on_plan).pack(
            side="left", padx=(10, 0))
        bt = ttk.Frame(body)
        bt.grid(row=15, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        ttk.Button(bt, text="Elimina la stanza", style="Danger.TButton",
                   command=self.delete_room).pack(side="left")
        ttk.Button(bt, text="Chiudi", command=self.destroy).pack(side="right")
        self.refresh()

    def room(self):
        return self.app.db.get_stanza(self.room_id)

    def refresh(self):
        room = self.room()
        if not room:
            self.destroy()
            return
        db = self.app.db
        head = "%s · %s" % (room["edificio"], floor_label(room["piano"]))
        prn = db.printers_in_room(room["uid"])
        if prn:
            head += "\nStampanti in questa stanza: " + "; ".join(
                " ".join(x for x in (q["nome"] or q["modello"] or "senza nome",
                                     ("(" + q["seriale"] + ")") if q["seriale"] else "",
                                     "FUORI SERVIZIO" if q["fuori_servizio"] else "") if x)
                for q in prn)
        self.lbl_head.configure(text=head)
        occ, over = room_occupancy(room)
        self.lbl_occ.configure(text=occ or "scrivanie totali non indicate",
                               foreground=RED if over else GREY)
        self.tree_p.delete(*self.tree_p.get_children())
        for p in db.persone_of(self.room_id):
            self.tree_p.insert("", "end", iid=str(p["utente_id"]),
                               values=("Postazione %d" % p["numero"], p["nome"], p["ufficio"],
                                       p["telefono"], p["seriale_pc"], p["imei"],
                                       dotazione_text(p["dotazione"])))
        self.draw_desks()
        self.tree_f.delete(*self.tree_f.get_children())
        for f in db.foto_stanza(self.room_id):
            self.tree_f.insert("", "end", iid=str(f["id"]),
                               values=(f["nome"], "Panorama" if f["tipo"] == "panorama" else "Foto"))
        self.lbl_pos.configure(text="Segnata sulla pianta." if room["px"] is not None
                               else "Non ancora segnata sulla pianta.")

    def draw_desks(self):
        """Le scrivanie della stanza, disegnate: PC, nome e S/N. Un clic apre i dettagli."""
        room = self.room()
        c = self.dcanvas
        c.delete("all")
        if not room:
            return
        width = max(c.winfo_width(), 700)
        people = {room["id"]: self.app.db.persone_of(self.room_id)}
        self.desk_lay = desk_layout([room], people, width, y0=0, show_label=False)
        draw_desk_rows(c, self.desk_lay)
        c.configure(scrollregion=(0, 0, width, max(self.desk_lay["height"], 60)))

    def on_desk_click(self, event):
        c = self.dcanvas
        hit = desk_hit(self.desk_lay, c.canvasx(event.x), c.canvasy(event.y))
        if not hit:
            return
        kind, _, uid = hit
        if kind == "occ" and uid is not None:
            self.tree_p.selection_set(str(uid))
            DeskDialog(self, self.app, self.room_id, uid, on_change=self.refresh)
        elif kind == "free":
            self.cb_persona.focus_set()

    def save(self):
        txt = self.v_scriv.get().strip()
        if txt and not txt.isdigit():
            messagebox.showwarning(APP_NAME, "Le scrivanie totali devono essere un numero (oppure "
                                             "lascia vuoto).", parent=self)
            return False
        try:
            self.app.db.update_stanza(self.room_id, self.v_codice.get(), int(txt or 0),
                                      self.txt_note.get("1.0", "end").strip(),
                                      self.txt_come.get("1.0", "end").strip())
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e), parent=self)
            return False
        self.app.refresh_rooms()
        self.refresh()
        occ, over = room_occupancy(self.room())
        if over:
            messagebox.showwarning(APP_NAME, "Attenzione: in questa stanza ci sono più persone "
                                             "delle scrivanie (%s)." % occ, parent=self)
        return True

    def add_person(self):
        nome = normalize_name(self.v_persona.get())
        if not nome:
            messagebox.showinfo(APP_NAME, "Scrivi o scegli il nome della persona.", parent=self)
            return
        db = self.app.db
        user = db.find_user(nome)
        if not user:
            if not messagebox.askyesno(APP_NAME, "\"%s\" non è ancora tra gli utenti.\n"
                                                 "Vuoi crearlo adesso?" % nome, parent=self):
                return
            room = self.room()
            row = db.conn.execute("SELECT sede FROM edifici WHERE id = ?",
                                  (room["edificio_id"],)).fetchone()
            db.add_user(nome, row["sede"] if row else "", "", "", "", "")
            user = db.find_user(nome)
        self.assign_user(user["id"], user["nome"])

    def new_person(self):
        """Scheda completa per una persona nuova: nome, ufficio, matricola, PC, IMEI, dotazione.
        Quando la salvi, siede subito in questa stanza (alla prima postazione libera)."""
        room = self.room()
        row = self.app.db.conn.execute("SELECT sede FROM edifici WHERE id = ?",
                                       (room["edificio_id"],)).fetchone()

        def on_save(data):
            if not self.app.save_user_card(None, data):
                return False
            user = self.app.db.find_user(data["nome"])
            if user:
                self.assign_user(user["id"], user["nome"])
            return True
        UserDialog(self, None, self.app.sedi_cache, on_save, uffici=self.app.db.uffici(),
                   defaults={"sede": row["sede"] if row else ""})

    def edit_person(self):
        sel = self.tree_p.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima una persona dall'elenco.", parent=self)
            return
        user = next((u for u in self.app.users_data if str(u["id"]) == sel[0]), None)
        if not user:
            return
        posto = self.app.db.posto_of_user(user["id"])
        seat = "%s · Postazione %d" % (posto["codice"], posto["numero"]) if posto else None

        def on_save(data):
            if not self.app.save_user_card(user, data):
                return False
            self.refresh()
            return True
        UserDialog(self, user, self.app.sedi_cache, on_save, seat=seat,
                   uffici=self.app.db.uffici())

    def assign_user(self, utente_id, nome):
        db = self.app.db
        room = self.room()
        old = db.posto_of_user(utente_id)
        if old and old["stanza_id"] != self.room_id:
            if not messagebox.askyesno(
                    APP_NAME, "%s siede già in %s (%s, Postazione %d).\nVuoi spostarlo in %s?"
                              % (nome, old["codice"], floor_label(old["piano"]), old["numero"],
                                 room["codice"]), parent=self):
                return False
        res = db.assegna_persona(utente_id, self.room_id)
        self.v_persona.set("")
        self.app.refresh_users()
        self.refresh()
        if res.get("presa_tolta"):
            messagebox.showinfo(APP_NAME, "La presa di rete %s era della scrivania di prima: l'ho tolta. "
                                          "Scrivi quella della nuova scrivania." % res["presa_tolta"],
                                parent=self)
        if res["sopra"]:
            occ, _ = room_occupancy(self.room())
            messagebox.showwarning(APP_NAME, "Attenzione: ora in %s ci sono più persone delle "
                                             "scrivanie (%s)." % (room["codice"], occ), parent=self)
        return True

    def remove_person(self):
        sel = self.tree_p.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima una persona dall'elenco.", parent=self)
            return
        self.app.db.rimuovi_persona(int(sel[0]))
        self.app.refresh_users()
        self.refresh()

    def add_photos(self, tipo):
        paths = filedialog.askopenfilenames(
            parent=self, title="Scegli i panorami" if tipo == "panorama" else "Scegli le foto",
            filetypes=[("Immagini", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Tutti i file", "*.*")])
        bad = []
        for path in paths or ():
            try:
                save_room_photo(self.app.db, self.room_id, path, tipo)
            except ValueError as e:
                bad.append(str(e))
        self.refresh()
        self.app.refresh_rooms()
        if bad:
            messagebox.showwarning(APP_NAME, "\n".join(bad), parent=self)

    def selected_photo(self):
        sel = self.tree_f.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima una foto dall'elenco.", parent=self)
            return None
        return next((f for f in self.app.db.foto_stanza(self.room_id) if str(f["id"]) == sel[0]),
                    None)

    def open_photo(self):
        f = self.selected_photo()
        if f:
            try:
                os.startfile(os.path.join(stanze_foto_root(), *f["file"].split("/")))  # solo Windows
            except Exception:
                messagebox.showinfo(APP_NAME, "Non riesco ad aprire la foto da qui.", parent=self)

    def toggle_tipo(self):
        f = self.selected_photo()
        if f:
            self.app.db.set_tipo_foto_stanza(f["id"], "foto" if f["tipo"] == "panorama"
                                             else "panorama")
            self.refresh()

    def del_photo(self):
        f = self.selected_photo()
        if f and messagebox.askyesno(APP_NAME, "Togliere questa foto dalla stanza?", parent=self):
            self.app.db.delete_foto_stanza(f["id"])
            self.refresh()
            self.app.refresh_rooms()

    def mark_on_plan(self):
        rid = self.room_id
        self.destroy()
        self.app.start_marking(rid)

    def delete_room(self):
        room = self.room()
        n = room["n_persone"]
        msg = "Eliminare la stanza %s?" % room["codice"]
        if n:
            msg += "\nLe %d persone assegnate restano tra gli utenti ma senza stanza." % n
        msg += "\nLe foto della stanza vengono tolte."
        if messagebox.askyesno(APP_NAME, msg, icon="warning", parent=self):
            self.app.db.delete_stanza(self.room_id)
            self.app.refresh_users()
            self.app.refresh_rooms()
            self.destroy()


class MapDialog(ModalDialog):
    """Mappa della Toscana: dove sono le sedi e quanti interventi hanno (contorno semplificato)."""
    FS = 24.0
    SEA, LAND, LAND_LINE, OFF = "#dfeaf5", "#f4efe0", "#7a8a6a", "#8a99ab"

    def __init__(self, app):
        super().__init__(app, "Mappa della Toscana")
        self.app = app
        self.scale = max(0.4, min(0.7, (self.winfo_screenheight() - 170) / TOSC_H))
        self.geometry("%dx%d" % (int(TOSC_W * self.scale) + 600, int(TOSC_H * self.scale) + 50))
        self.overrides = load_map_pos(app.settings)
        self.var_per = tk.StringVar(value=TOSC_PERIODI[0][1])
        self.sel_keys, self.placing = set(), None
        self.recs, self.senza, self.pins, self.scene = [], {}, [], None
        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(body, bg=self.SEA, highlightthickness=1, highlightbackground=BORDER,
                                width=int(TOSC_W * self.scale), height=int(TOSC_H * self.scale))
        self.canvas.grid(row=0, column=0, sticky="n")
        self.canvas.bind("<Button-1>", self.on_click)
        right = ttk.Frame(body)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)
        top = ttk.Frame(right)
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="Interventi:").pack(side="left")
        cb = ttk.Combobox(top, textvariable=self.var_per, state="readonly", width=18,
                          values=[lab for _, lab in TOSC_PERIODI])
        cb.pack(side="left", padx=(6, 0))
        cb.bind("<<ComboboxSelected>>", lambda e: self.reload())
        self.lbl_info = ttk.Label(right, style="Hint.TLabel", wraplength=520, justify="left")
        self.lbl_info.grid(row=1, column=0, sticky="w", pady=(8, 6))
        tf = ttk.Frame(right)
        tf.grid(row=2, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)
        cols = ("sede", "interventi", "segui", "utenti", "stampanti", "dove")
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c, t, w, a in (("sede", "Sede", 150, "w"), ("interventi", "Interventi", 70, "center"),
                           ("segui", "Da seguire", 75, "center"), ("utenti", "Utenti", 55, "center"),
                           ("stampanti", "Stampanti", 70, "center"), ("dove", "Posizione", 110, "w")):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a, stretch=(c == "sede"))
        self.tree.tag_configure("rosso", foreground=RED)
        self.tree.tag_configure("nopos", foreground="#7a8794")
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.on_tree)
        self.tree.bind("<Double-1>", lambda e: self.show_in_registry(False))
        bt = ttk.Frame(right)
        bt.grid(row=3, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(bt, text="Mostra gli interventi", style="Primary.TButton",
                   command=lambda: self.show_in_registry(False)).pack(side="left")
        ttk.Button(bt, text="Solo quelli da seguire",
                   command=lambda: self.show_in_registry(True)).pack(side="left", padx=(8, 0))
        bt2 = ttk.Frame(right)
        bt2.grid(row=4, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(bt2, text="Posiziona sulla mappa", command=self.start_place).pack(side="left")
        ttk.Button(bt2, text="Togli la posizione scelta", command=self.clear_place).pack(
            side="left", padx=(8, 0))
        ttk.Button(bt2, text="Chiudi", command=self.destroy).pack(side="right")
        ttk.Label(right, style="Hint.TLabel", wraplength=520, justify="left",
                  text="Le sedi si riconoscono dal nome (Firenze, Pisa, «Sede di Siena»...). Se una "
                       "sede non ha la posizione, sceglila nell'elenco, premi «Posiziona sulla "
                       "mappa» e clicca il punto giusto. Il contorno della Toscana è semplificato."
                  ).grid(row=5, column=0, sticky="w", pady=(8, 0))
        self.reload()

    def period_kind(self):
        lab = self.var_per.get()
        return next((k for k, t in TOSC_PERIODI if t == lab), "tutto")

    def reload(self):
        res = tosc_records(self.app.db, self.period_kind(), self.overrides)
        self.recs, self.senza = res["sedi"], res["senza_sede"]
        self.pins = tosc_pins(self.recs)
        self.tree.delete(*self.tree.get_children())
        for r in self.recs:
            dove = "scelta da te" if r["via"] == "manuale" else (r["via"] or "non posizionata")
            tags = (("rosso",) if r["da_seguire"] else ()) + (() if r["lat"] is not None else ("nopos",))
            self.tree.insert("", "end", iid=r["key"], tags=tags,
                             values=(r["nome"], r["interventi"], r["da_seguire"], r["utenti"],
                                     r["stampanti"], dove))
        self.sel_keys = {k for k in self.sel_keys if any(r["key"] == k for r in self.recs)}
        self.draw()
        self.show_info()

    def draw(self):
        c = self.canvas
        c.delete("all")
        sc = self.scene = tosc_scene(self.pins, self.scale, self.FS)
        c.create_rectangle(0, 0, sc["w"], sc["h"], fill=self.SEA, outline="")
        c.create_polygon(*sc["land"], fill=self.LAND, outline=self.LAND_LINE)
        for isl in sc["islands"]:
            c.create_polygon(*isl, fill=self.LAND, outline=self.LAND_LINE)
        size = max(8, int(round(self.FS * self.scale * 0.72)))
        for p in sc["pins"]:
            sel = any(k in self.sel_keys for k in p["sedi"])
            cx, cy, r = p["cx"], p["cy"], p["r"]
            c.create_oval(cx - r, cy - r, cx + r, cy + r, fill=NAVY if p["n"] else self.OFF,
                          outline="#ffc20e" if sel else WHITE, width=4 if sel else 2)
            if p["n"]:
                c.create_text(cx, cy, text=str(p["n"]), fill=WHITE, font=("Segoe UI", size, "bold"))
            if p["rosso"]:
                b = 8 * self.scale + 3
                bx, by = cx + r * 0.85, cy - r * 0.85
                c.create_oval(bx - b, by - b, bx + b, by + b, fill=RED, outline=WHITE, width=2)
            anchor = {"middle": "s", "start": "sw", "end": "se"}[p["anchor"]]
            for dx, dy in ((-1, 0), (1, 0), (0, -1), (0, 1)):
                c.create_text(p["lx"] + dx, p["ly"] + dy, text=p["label"], anchor=anchor,
                              fill=WHITE, font=("Segoe UI", size, "bold"))
            c.create_text(p["lx"], p["ly"], text=p["label"], anchor=anchor, fill=TEXT,
                          font=("Segoe UI", size, "bold"))

    def selected_rec(self):
        return next((r for r in self.recs if r["key"] in self.sel_keys), None)

    def show_info(self):
        r = self.selected_rec()
        if self.placing:
            nome = next((x["nome"] for x in self.recs if x["key"] == self.placing), "")
            txt = "Clicca sulla mappa il punto dove si trova «%s»." % nome
        elif r:
            txt = ("%s: %d interventi nel periodo, %d da seguire (di qualsiasi data), %d utenti, "
                   "%d stampanti. Posizione: %s." % (
                       r["nome"], r["interventi"], r["da_seguire"], r["utenti"], r["stampanti"],
                       "scelta da te" if r["via"] == "manuale" else (r["via"] or "non indicata")))
        else:
            txt = ("Clicca su un pallino o scegli una sede nell'elenco. Il numero è quello degli "
                   "interventi nel periodo; il pallino rosso dice quanti sono ancora da seguire "
                   "(aperti o in attesa, di qualsiasi data).")
        if self.senza.get("interventi") or self.senza.get("da_seguire"):
            txt += "\nSenza sede indicata: %d interventi (%d da seguire)." % (
                self.senza["interventi"], self.senza["da_seguire"])
        self.lbl_info.configure(text=txt)

    def on_tree(self, _e=None):
        sel = self.tree.selection()
        self.sel_keys = {sel[0]} if sel else set()
        self.draw()
        self.show_info()

    def on_click(self, event):
        if self.placing:
            lat, lon = tosc_unproject(event.x / self.scale, event.y / self.scale)
            if not (41.9 < lat < 44.7 and 9.5 < lon < 12.6):
                messagebox.showinfo("Mappa della Toscana", "Clicca dentro la mappa.", parent=self)
                return
            self.overrides[self.placing] = (round(lat, 4), round(lon, 4))
            key, self.placing = self.placing, None
            self.canvas.configure(cursor="")
            self.save_overrides()
            self.sel_keys = {key}
            self.reload()
            if self.tree.exists(key):
                self.tree.selection_set(key)
            return
        i = tosc_pin_at(self.pins, event.x, event.y, self.scale)
        if i is None:
            return
        keys = self.pins[i]["sedi"]
        cur = self.selected_rec()
        # se il pallino ha più sedi, un altro clic passa alla successiva
        if cur and cur["key"] in keys and len(keys) > 1:
            key = keys[(keys.index(cur["key"]) + 1) % len(keys)]
        else:
            key = keys[0]
        self.sel_keys = {key}
        self.tree.selection_set(key)
        self.tree.see(key)
        self.draw()
        self.show_info()

    def save_overrides(self):
        self.app.settings["mappa_pos"] = dump_map_pos(self.overrides)
        if not save_settings(self.app.settings):
            messagebox.showwarning("Mappa della Toscana",
                                   "Non sono riuscito a salvare la posizione.", parent=self)

    def start_place(self):
        r = self.selected_rec()
        if not r:
            messagebox.showinfo("Mappa della Toscana", "Scegli prima una sede nell'elenco.",
                                parent=self)
            return
        self.placing = r["key"]
        self.canvas.configure(cursor="crosshair")
        self.show_info()

    def clear_place(self):
        r = self.selected_rec()
        if not r or r["key"] not in self.overrides:
            messagebox.showinfo("Mappa della Toscana",
                                "Questa sede non ha una posizione scelta da te.", parent=self)
            return
        del self.overrides[r["key"]]
        self.placing = None
        self.save_overrides()
        self.reload()

    def show_in_registry(self, only_open):
        r = self.selected_rec()
        if not r:
            messagebox.showinfo("Mappa della Toscana", "Scegli prima una sede.", parent=self)
            return
        self.app.show_sede_in_registry(r["nome"], "tutto" if only_open else self.period_kind(),
                                       only_open)
        self.destroy()


class ScanDialog(ModalDialog):
    """Legge un codice a barre (Code 128, Code 39) o un QR con la webcam. Se la webcam non c'è o non va:
    una foto scelta da file, un'immagine copiata negli appunti, oppure si scrive a mano.
    kind: "serial" (qualunque codice di almeno 4 caratteri), "imei" (solo un IMEI valido) o "any".
    multi: la webcam resta accesa e ogni codice nuovo viene passato a on_code (lettura a catena)."""
    W = 640

    def __init__(self, parent, kind="serial", multi=False, title="Leggi il codice", on_code=None,
                 on_manual=None, source=0):
        super().__init__(parent, title)
        self.resizable(False, False)
        self.bind("<Escape>", lambda e: self.close())
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.kind, self.multi, self.on_code, self.on_manual = kind, multi, on_code, on_manual
        self.source = source
        self.alive = True
        self.gen = 0
        self.lock = threading.Lock()
        self.frame, self.frame_no, self.shown_no = None, 0, 0
        self.events = queue.Queue()
        self.count = 0
        self.codes = []
        self._img = None
        self.var_msg = tk.StringVar()
        self.var_list = tk.StringVar()
        body = ttk.Frame(self, padding=12)
        body.pack(fill="both", expand=True)
        self.lbl_img = tk.Label(body, bg="#111111", fg="white", text="Avvio la fotocamera…",
                                width=80, height=18)
        self.lbl_img.pack()
        self.lbl_msg = ttk.Label(body, textvariable=self.var_msg, wraplength=self.W, justify="left")
        self.lbl_msg.pack(fill="x", pady=(8, 0))
        self.lbl_list = ttk.Label(body, textvariable=self.var_list, style="Hint.TLabel",
                                  wraplength=self.W, justify="left")
        self.lbl_list.pack(fill="x", pady=(4, 0))
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(10, 0))
        ttk.Button(row, text="Scegli una foto…", command=self.choose_photo).pack(side="left")
        ttk.Button(row, text="Incolla immagine", command=self.paste_photo).pack(side="left", padx=(8, 0))
        if isinstance(source, int):
            ttk.Button(row, text="Altra fotocamera", command=self.next_camera).pack(side="left", padx=(8, 0))
        self.btn_close = ttk.Button(row, text="Chiudi", command=self.close)
        self.btn_close.pack(side="right")
        ttk.Button(row, text="Scrivi a mano", command=self.manual).pack(side="right", padx=(0, 8))
        self.say("Avvio la fotocamera…")
        if CD is None or not CD.disponibile():
            self.say("In questa versione del programma la lettura dei codici non è disponibile. "
                     "Scrivi il codice a mano.", "err")
        else:
            self.start_camera()
        self.after(50, self.poll)

    # ----- messaggi -----
    def say(self, text, level=""):
        self.var_msg.set(text)
        try:
            self.lbl_msg.configure(foreground=RED if level == "err" else GREEN_DARK if level == "ok" else TEXT)
        except tk.TclError:
            pass

    def prompt(self):
        return ("Inquadra il codice a barre dell'IMEI (15 cifre)." if self.kind == "imei"
                else "Inquadra il codice a barre o il QR e tienilo fermo dentro il riquadro.")

    # ----- fotocamera (in due thread: uno prende i fotogrammi, uno li legge) -----
    def start_camera(self):
        self.gen += 1
        threading.Thread(target=self._capture, args=(self.source, self.gen), daemon=True).start()

    def _capture(self, source, gen):
        cam = None
        try:
            cam = CD.Camera(source)
            self.events.put((gen, "info", self.prompt()))
            threading.Thread(target=self._decode, args=(gen,), daemon=True).start()
            fails = 0
            while self.alive and self.gen == gen:
                fr = cam.read()
                if fr is None:
                    fails += 1
                    if fails > 40:
                        self.events.put((gen, "err", "La fotocamera non manda immagini: forse la sta usando un altro "
                                                     "programma, o Windows non le permette di usarla. "
                                                     "Scegli una foto oppure scrivi a mano."))
                        break
                    time.sleep(0.05)
                    continue
                fails = 0
                with self.lock:
                    self.frame, self.frame_no = fr, self.frame_no + 1
                time.sleep(1 / 30 if isinstance(source, str) else 0.004)
        except Exception as e:
            self.events.put((gen, "err", (str(e) or "La fotocamera non parte.") +
                             " Scegli una foto oppure scrivi a mano."))
        finally:
            if cam is not None:
                cam.close()

    def _decode(self, gen):
        rd = CD.LiveReader(self.kind, self.multi)
        last_no, last_hint = 0, ""
        while self.alive and self.gen == gen:
            with self.lock:
                fr, no = self.frame, self.frame_no
            if fr is None or no == last_no:
                time.sleep(0.01)
                continue
            last_no = no
            try:
                res = rd.read_frame(fr)
            except Exception:
                res = []
            t = rd.feed(res)
            if t is not None:
                self.events.put((gen, "code", t))
                if not self.multi:
                    return
            else:
                h = rd.hint()
                if h and h != last_hint:
                    last_hint = h
                    self.events.put((gen, "hint", h))

    def next_camera(self):
        if not isinstance(self.source, int):
            return
        self.source = (self.source + 1) % 4
        with self.lock:
            self.frame = None
        self.say("Provo la fotocamera numero %d…" % (self.source + 1))
        self.start_camera()

    # ----- foto da file o dagli appunti -----
    def choose_photo(self):
        f = filedialog.askopenfilename(
            parent=self, title="Scegli la foto del codice",
            filetypes=[("Immagini", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tif *.tiff"), ("Tutti i file", "*.*")])
        if f:
            self._photo_job(lambda: CD.load_image(f))

    def paste_photo(self):
        def load():
            try:
                from PIL import ImageGrab
                data = ImageGrab.grabclipboard()
            except Exception:
                data = None
            if isinstance(data, list):
                data = next((CD.load_image(p) for p in data if os.path.isfile(p)
                             and os.path.splitext(p)[1].lower() in IMG_EXTS), None)
            return data
        self._photo_job(load)

    def _photo_job(self, loader):
        if CD is None or not CD.disponibile():
            self.say("La lettura dei codici non è disponibile in questa versione.", "err")
            return
        self.say("Cerco il codice nella foto…")

        def run():
            try:
                im = loader()
                if im is None:
                    self.events.put((-1, "err", "Negli appunti non c'è nessuna immagine. Copia una foto o uno "
                                                "screenshot e riprova."))
                    return
                self.events.put((-1, "photo", CD.decode_image(im, self.kind)))
            except Exception as e:
                self.events.put((-1, "err", "Non riesco ad aprire questa immagine: %s" % e))
        threading.Thread(target=run, daemon=True).start()

    # ----- ciclo dell'interfaccia -----
    def poll(self):
        if not self.alive:
            return
        try:
            while self.alive:
                gen, kind, payload = self.events.get_nowait()
                if gen != self.gen and gen != -1:
                    continue
                if kind == "code":
                    self.accept(payload)
                elif kind == "photo":
                    if payload:
                        self.accept(payload["text"])
                    else:
                        self.say("Nella foto non trovo un IMEI valido (15 cifre). Riprova più vicino e dritto, "
                                 "oppure scrivilo a mano." if self.kind == "imei" else
                                 "Nella foto non trovo nessun codice. Riprova più vicino, con il codice dritto e "
                                 "ben illuminato, oppure scrivilo a mano.", "err")
                elif kind == "err":
                    self.say(payload, "err")
                else:
                    self.say(payload)
        except queue.Empty:
            pass
        if not self.alive:
            return
        self.draw()
        self.after(40, self.poll)

    def draw(self):
        with self.lock:
            fr, no = self.frame, self.frame_no
        if fr is None or no == self.shown_no:
            return
        self.shown_no = no
        try:
            self._img = tk.PhotoImage(data=CD.preview_png_b64(fr, self.W))
            self.lbl_img.configure(image=self._img, text="", width=self._img.width(), height=self._img.height())
        except Exception:
            pass

    def accept(self, text):
        try:
            self.bell()
        except Exception:
            pass
        cb = self.on_code
        if self.multi:
            self.count += 1
            self.codes.insert(0, text)
            self.var_list.set("Letti: %d · ultimo %s%s" % (self.count, text, (" · prima " + ", ".join(self.codes[1:4]))
                                                       if len(self.codes) > 1 else ""))
            self.btn_close.configure(text="Fine (%d)" % self.count)
            self.say("Letto: %s. Inquadra il prossimo codice." % text, "ok")
            if cb:
                cb(text)
            return
        self.close()
        if cb:
            cb(text)

    def manual(self):
        f = self.on_manual
        self.close()
        if f:
            f()

    def close(self):
        self.alive = False
        self.gen += 1
        try:
            self.destroy()
        except tk.TclError:
            pass


def scan_button(parent, var, kind="serial", multi=False, on_code=None, entry=None, kind_fn=None,
                title=None):
    """Il pulsante «Fotocamera…» da mettere accanto a un campo: apre ScanDialog e scrive nel campo (var).
    kind_fn: se serve, una funzione che dice di volta in volta cosa cercare ("imei" o "serial")."""
    def go():
        k = kind_fn() if kind_fn else kind

        def done(text):
            var.set(text)
            if on_code:
                on_code(text)

        def manual():
            try:
                entry.focus_set()
            except Exception:
                pass
        ScanDialog(parent.winfo_toplevel(), kind=k, multi=multi, on_code=done, on_manual=manual,
                   title=title or ("Leggi un codice dopo l'altro" if multi else "Leggi l'IMEI" if k == "imei"
                                   else "Leggi il codice"))
    return ttk.Button(parent, text="Fotocamera…", command=go)


class VerbaliSettingsDialog(ModalDialog):
    FIELDS = (("Nome del tecnico (compare nei censimenti)", "tecnico"),
              ("Sede (città)", "sede_citta"), ("Indirizzo della sede", "indirizzo"), ("CAP", "cap"))

    def __init__(self, parent, settings, on_save):
        super().__init__(parent, "Dati per i verbali")
        self.resizable(False, False)
        self.on_save = on_save
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        self.vars = {}
        for i, (label, key) in enumerate(self.FIELDS):
            ttk.Label(body, text=label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 14))
            self.vars[key] = tk.StringVar(value=settings.get(key, ""))
            ttk.Entry(body, textvariable=self.vars[key], width=40).grid(row=i, column=1,
                                                                        sticky="ew", pady=5)
        ttk.Label(body, style="Hint.TLabel", wraplength=480, justify="left",
                  text="Questi dati si propongono da soli nei nuovi verbali di censimento. "
                       "Li puoi comunque cambiare ogni volta.").grid(
            row=len(self.FIELDS), column=0, columnspan=2, sticky="w", pady=(8, 0))
        # immagine della firma: SOLO per il tecnico (per l'utente si firma col dito o a mano)
        sig = ttk.LabelFrame(body, text="La mia firma (immagine del tecnico)", padding=10)
        sig.grid(row=len(self.FIELDS) + 1, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.lbl_sig = ttk.Label(sig, text="", wraplength=440, justify="left")
        self.lbl_sig.pack(anchor="w")
        srow = ttk.Frame(sig)
        srow.pack(anchor="w", pady=(6, 0))
        ttk.Button(srow, text="Scegli l'immagine…", command=self.choose_signature).pack(side="left")
        ttk.Button(srow, text="Togli la firma salvata", command=self.remove_signature).pack(
            side="left", padx=(8, 0))
        ttk.Label(sig, style="Hint.TLabel", wraplength=440, justify="left",
                  text="Una foto o una scansione della tua firma su foglio bianco: ritaglio io i "
                       "margini e tolgo il fondo. Nel verbale la scegli con «Mia firma salvata». "
                       "Per l'utente e per le altre persone non si può caricare nessuna immagine: "
                       "firmano col dito sul telefono o a mano sul foglio.").pack(
            anchor="w", pady=(6, 0))
        self._show_signature()
        btns = ttk.Frame(body)
        btns.grid(row=len(self.FIELDS) + 2, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))

    def _show_signature(self):
        self.lbl_sig.configure(text="Firma salvata: sì (si usa con «Mia firma salvata»)."
                               if saved_signature() else "Nessuna firma salvata.")

    def choose_signature(self):
        path = pick_signature_image(self)
        if path:
            self._show_signature()

    def remove_signature(self):
        if saved_signature() and messagebox.askyesno(
                APP_NAME, "Togliere la firma salvata dal programma?\n\nI verbali già fatti "
                          "tengono la loro copia.", parent=self):
            try:
                os.remove(saved_signature_path())
            except OSError:
                pass
        self._show_signature()

    def save(self):
        self.on_save({k: v.get().strip() for k, v in self.vars.items()})
        self.destroy()


def pick_signature_image(parent):
    """Fa scegliere un'immagine di firma del tecnico e la salva come «la mia firma».
    Restituisce il percorso salvato oppure None."""
    f = filedialog.askopenfilename(
        parent=parent, title="Scegli l'immagine della tua firma",
        filetypes=[("Immagini", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tif *.tiff"),
                   ("Tutti i file", "*.*")])
    if not f:
        return None
    try:
        return save_signature_image(f)
    except VB.VerbaleError as e:
        messagebox.showerror(APP_NAME, str(e), parent=parent)
    except Exception as e:
        messagebox.showerror(APP_NAME, "Non riesco a usare questa immagine:\n%s" % e, parent=parent)
    return None


class FirmePanel(ttk.LabelFrame):
    """Riquadro delle firme di un verbale: per ogni posto firma si sceglie come si firma.
    Tecnico: nessuna / firma già presente / mia firma salvata (immagine) / a mano sul foglio.
    Utente e ogni altra persona: nessuna / firma già presente / a mano sul foglio (MAI immagine)."""

    LABELS = {"nessuna": "Nessuna firma (la riga resta da firmare)",
              "tenuta": "Tieni la firma già presente (fatta col dito sul telefono)",
              "salvata": "Metti la mia firma salvata (immagine)",
              "cartaceo": "Firmerà a mano sul foglio stampato (cartaceo)"}

    def __init__(self, master, tipo, esistenti, cartaceo, parent_win):
        super().__init__(master, text="Firme", padding=10)
        self.tipo, self.parent_win = tipo, parent_win
        self.vars, self.radios = {}, {}
        for chiave, etichetta, obbl in VB.FIRME_POSTI[tipo]:
            box = ttk.Frame(self)
            box.pack(fill="x", anchor="w", pady=(0, 8))
            ttk.Label(box, text="%s%s" % (etichetta, "" if obbl else " (facoltativa)"),
                      font=("Segoe UI", 10, "bold")).pack(anchor="w")
            if chiave in esistenti:
                start = "tenuta"
            elif chiave in cartaceo:
                start = "cartaceo"
            else:
                start = "nessuna"
            self.vars[chiave] = tk.StringVar(value=start)
            for modo in VB.modi_firma_ammessi(chiave):
                if modo == "tenuta" and chiave not in esistenti:
                    continue
                rb = ttk.Radiobutton(box, text=self.LABELS[modo], value=modo,
                                     variable=self.vars[chiave])
                rb.pack(anchor="w", padx=(12, 0))
                self.radios[(chiave, modo)] = rb
            if chiave == "tecnico":
                ttk.Button(box, text="Scegli l'immagine della mia firma…",
                           command=self.choose_signature).pack(anchor="w", padx=(12, 0), pady=(4, 0))
        self._refresh_saved()

    def _refresh_saved(self):
        rb = self.radios.get(("tecnico", "salvata"))
        if rb is not None:
            rb.configure(state="normal" if saved_signature() else "disabled")

    def choose_signature(self):
        if pick_signature_image(self.parent_win):
            self._refresh_saved()
            if "tecnico" in self.vars:
                self.vars["tecnico"].set("salvata")

    def scelte(self):
        return {k: v.get() for k, v in self.vars.items()}


class FirmeDialog(ModalDialog):
    """Cambia le firme di un verbale già nell'elenco."""

    def __init__(self, app, tipo, esistenti, cartaceo, on_save):
        super().__init__(app, "Firme del verbale - " + VB.TIPI[tipo][0])
        self.on_save = on_save
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        self.panel = FirmePanel(body, tipo, esistenti, cartaceo, self)
        self.panel.pack(fill="x")
        ttk.Label(body, style="Hint.TLabel", wraplength=520, justify="left",
                  text="Dopo il cambio premi «Crea il Word…» per rifare il documento. Con «a mano» "
                       "il Word ha la riga della firma vuota, con lo spazio per firmare sul foglio "
                       "stampato.").pack(anchor="w", pady=(8, 0))
        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(12, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))

    def save(self):
        if self.on_save(self.panel.scelte()) is not False:
            self.destroy()


class VerbaleDialog(ModalDialog):
    """Modulo per creare un verbale Word. I campi cambiano in base al tipo di verbale."""

    def __init__(self, app, tipo, ini, on_create):
        super().__init__(app, "Verbale - " + VB.TIPI[tipo][0])
        self.app, self.tipo, self.on_create, self.ini = app, tipo, on_create, ini or {}
        self.geometry("900x%d" % min(820, max(560, self.winfo_screenheight() - 110)))
        self.minsize(800, 520)
        self.v, self.t, self.checks = {}, {}, {}
        self.photos = [p for p in self.ini.get("foto", []) if os.path.isfile(p)]
        self.apparati = [dict(r) for r in self.ini.get("apparati", [])]
        self.dotazione = [dict(r) for r in self.ini.get("dotazione", [])]
        self.user_key = None
        self.hist = app.verbali_storico()      # quello che hai gia' scritto nei verbali precedenti
        self.dev_map = {}
        self._dot_auto = None
        self._sede_auto = {}
        self._marca_auto = ""          # marca rimasta scritta dal giro prima (non digitata da te)
        self._suggest_adders = {}
        self._auto_vals = {}           # campi messi da me per la persona scelta (si cambiano se cambi persona)
        self._dev_auto = None          # (tipologia, modello, seriale) messi da me
        self._owner_key = None         # la persona di cui sto mostrando gli apparati
        self._owner_devs = []
        self._dot_card_rows = []       # righe di dotazione che ho caricato io dalla scheda
        self._serial_all = []
        self.serial_cb = None
        self.dev_hint_lbl = None
        self.show_all = tk.BooleanVar(value=False)      # «mostra anche gli apparati degli altri»

        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)
        bar = ttk.Frame(outer, padding=(16, 10))
        bar.pack(side="bottom", fill="x")
        ttk.Button(bar, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Crea il Word…", style="Save.TButton", command=self.create).pack(
            side="right", padx=(0, 8))
        ttk.Label(bar, style="Hint.TLabel",
                  text="* = obbligatorio. Quello che lasci vuoto resta una riga da compilare a "
                       "mano.").pack(side="left")
        sf = self.sf = ScrollFrame(outer)
        sf.pack(fill="both", expand=True)
        self.bind("<MouseWheel>", sf.wheel)
        self.body = sf.inner
        self.body.columnconfigure(1, weight=1)
        self.row = 0
        self._entry("Data *", "data", 14, hint="gg/mm/aaaa",
                    default=fmt_date(date.today()))
        getattr(self, "_form_" + tipo)()
        self._photo_box()
        self.firme_panel = FirmePanel(self.body, tipo, dict(self.ini.get("firme_esistenti") or {}),
                                      set(self.ini.get("firme_cartaceo") or ()), self)
        self.firme_panel.grid(row=self.row, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self.row += 1
        self.firme_scelte = self.firme_panel.scelte()
        if self.user_key and self.v[self.user_key].get().strip():
            self._init_user()                  # aperto con la persona gia' scelta: compilo i suoi dati

    def _init_user(self):
        u = self._current_user()
        if not u:
            return
        for key, val in (("contatti", u["telefono"]), ("matricola", u["matricola"])):
            if key in self.v and val and self.v[key].get().strip() == val:
                self._auto_vals[key] = val      # messo da me: cambiando persona cambia anche lui
        self._user_picked()

    # ----- costruzione dei campi -----
    def _wheel_on_combo(self, event):
        self.sf.canvas.yview_scroll(int(-event.delta / 120), "units")
        return "break"

    def _line(self, label, build):
        ttk.Label(self.body, text=label).grid(row=self.row, column=0, sticky="nw", pady=5,
                                              padx=(0, 14))
        fr = ttk.Frame(self.body)
        fr.grid(row=self.row, column=1, sticky="ew", pady=5)
        build(fr)
        self.row += 1

    def _entry(self, label, key, width=44, hint="", default=""):
        var = self.v[key] = tk.StringVar(value=self.ini.get(key, default))

        def build(fr):
            ttk.Entry(fr, textvariable=var, width=width).pack(side="left")
            if hint:
                ttk.Label(fr, text=hint, style="Hint.TLabel").pack(side="left", padx=(8, 0))
        self._line(label, build)
        return var

    def _combo(self, label, key, values, width=42, hint="", default="", scan=None):
        var = self.v[key] = tk.StringVar(value=self.ini.get(key, default))
        holder = []

        def build(fr):
            cb = ttk.Combobox(fr, textvariable=var, values=values, width=width)
            bind_autocomplete(cb, var, lambda: values)
            cb.pack(side="left")
            cb.bind("<MouseWheel>", self._wheel_on_combo)
            holder.append(cb)
            if scan:
                scan_button(fr, var, "serial", entry=cb, on_code=lambda t: scan(None)).pack(
                    side="left", padx=(6, 0))
            if hint:
                ttk.Label(fr, text=hint, style="Hint.TLabel").pack(side="left", padx=(8, 0))
        self._line(label, build)
        return var, holder[0]

    def _text(self, label, key, height, hint="", suggest=""):
        value = self.ini.get(key, "")
        if isinstance(value, (list, tuple)):
            value = "\n".join(value)

        def build(fr):
            t = make_text(fr, height)
            t.pack(fill="x", expand=True)
            t.insert("1.0", value)
            self.t[key] = t
            if hint:
                ttk.Label(fr, text=hint, style="Hint.TLabel").pack(anchor="w")
            options = self._hist(suggest) if suggest else []
            if options:
                row = ttk.Frame(fr)
                row.pack(anchor="w", pady=(4, 0))
                ttk.Label(row, text="Aggiungi dai verbali precedenti:", style="Hint.TLabel").pack(
                    side="left")
                var = tk.StringVar()
                cb = ttk.Combobox(row, textvariable=var, values=options, width=34,
                                  state="readonly")
                cb.pack(side="left", padx=6)

                def add_line(event=None, t=t, var=var):
                    item = var.get().strip()
                    if not item:
                        return
                    if item.casefold() not in [x.strip().casefold()
                                               for x in t.get("1.0", "end").splitlines()]:
                        t.insert("end", ("\n" if t.get("1.0", "end").strip() else "") + item)
                    var.set("")
                cb.bind("<<ComboboxSelected>>", add_line)
                self._suggest_adders[key] = (var, add_line)
        self._line(label, build)

    def _motivi(self, options):
        selected = self.ini.get("motivi", [])

        def build(fr):
            for opt in options:
                var = tk.BooleanVar(value=opt in selected)
                self.checks[opt] = var
                ttk.Checkbutton(fr, text=opt, variable=var).pack(anchor="w")
        self._line("Motivazione", build)
        self._entry("Altro (motivo)", "altro", 50,
                    hint="se lo scrivi, la casella «Altro» viene spuntata")

    def _hist(self, campo, base=()):
        """Suggerimenti per un campo: i valori fissi, poi quelli gia' usati nei verbali precedenti."""
        vals, seen = [], set()
        for v in list(base) + self.hist["valori"].get(campo, []):
            v = (v or "").strip()
            if v and v.casefold() not in seen:
                seen.add(v.casefold())
                vals.append(v)
        return vals

    def _hcombo(self, label, key, campo, width=44, base=(), hint="", default=""):
        return self._combo(label, key, self._hist(campo, base), width, hint=hint,
                           default=default)

    @staticmethod
    def _on_pick(cb, handler):
        for ev in ("<<ComboboxSelected>>", "<FocusOut>", "<Return>"):
            cb.bind(ev, handler, add="+")

    def _user_combo(self, label, key, hint="scegli un utente in lista e compilo i suoi dati"):
        var, cb = self._combo(label, key, self.app.users_cache, 40, hint=hint)
        for ev in ("<<ComboboxSelected>>", "<FocusOut>", "<Return>"):
            cb.bind(ev, self._user_picked, add="+")
        self.user_key = key
        return var

    # ----- dati dell'utente scelto -----
    def _current_user(self):
        if not self.user_key:
            return None
        raw = re.sub(r"\s*\(matricola[^)]*\)\s*$", "", self.v[self.user_key].get())
        return self.app.users_by_key.get(key_name(raw)) if raw.strip() else None

    def _set_if_empty(self, key, value):
        if key in self.v and value and not self.v[key].get().strip():
            self.v[key].set(value)

    def _auto_set(self, key, value):
        """Scrivo il dato della persona nel campo se e' vuoto, o se c'ero io (per la persona di
        prima): cosi' cambiando persona cambiano anche i suoi dati. Quello che scrivi tu resta."""
        if key not in self.v:
            return
        value = value or ""
        cur = self.v[key].get().strip()
        if not cur or cur == self._auto_vals.get(key):
            self.v[key].set(value)
            self._auto_vals[key] = value

    def _user_picked(self, event=None):
        u = self._current_user()
        if not u:
            if self.user_key and not self.v[self.user_key].get().strip():
                self._owner_key = None
                self._owner_devs = []
                self._owner_lists(None)           # nome tolto: tornano gli elenchi di tutti
            return
        if self.tipo in ("trasferimento", "assegnazione") and self.v[self.user_key].get() != u["nome"]:
            self.v[self.user_key].set(u["nome"])      # nome scritto come in lista
        if self.tipo == "trasferimento":
            self._auto_set("contatti", u["telefono"])
        elif self.tipo == "presa_in_carico":
            wanted = "%s (matricola %s)" % (u["nome"], u["matricola"]) if u["matricola"] else u["nome"]
            if self.v["utente"].get() != wanted:
                self.v["utente"].set(wanted)
        elif self.tipo == "assegnazione":
            self._auto_set("matricola", u["matricola"])
            self._auto_set("titolo", self.hist["titoli"].get(u["nome"].casefold(), ""))
            self._dot_from_card(u)
            self._dot_hint()
        if self.tipo in ("trasferimento", "presa_in_carico"):
            self._user_devices(u)
        self._device_hint()

    def _dot_from_card(self, u):
        """Assegnazione: carico nell'elenco la dotazione scritta nella scheda della persona
        (notebook, smartphone, SIM, cuffie con seriale/IMEI). Non tocco quello che hai scritto tu."""
        auto = self._dot_card_rows
        if any(r not in auto for r in self.dotazione):
            return
        rows = VB.dotazione_da_scheda(u)
        self.dotazione = [dict(r) for r in rows]
        self._dot_card_rows = [dict(r) for r in rows]
        self._fill_dot()

    # ----- apparati gia' noti (dai verbali precedenti e dalla scheda utente) -----
    @staticmethod
    def _group(tip):
        tip = (tip or "").casefold()
        if any(w in tip for w in PC_WORDS):
            return "pc"
        if any(w in tip for w in PHONE_WORDS):
            return "phone"
        return ""

    def _device_hint(self, *_):
        """Se cambi la tipologia (PC o cellulare) metto il seriale/IMEI giusto della persona."""
        u = self._current_user()
        if not u or "tipologia" not in self.v:
            return
        group = self._group(self.v["tipologia"].get())
        if not group:
            return
        devs = VB.dispositivi_utente(self.hist, u["nome"], u)
        cur = self.v["seriale"].get().strip()
        cur_dev = next((d for d in devs if cur and VB._codice_key(d["seriale"]) ==
                        VB._codice_key(cur)), None)
        wanted = [d for d in devs if self._group(d["tipologia"]) == group]
        if cur and not (cur_dev and self._group(cur_dev["tipologia"]) != group):
            return                                   # il seriale scritto va bene: non lo tocco
        if cur_dev and self.v["marca_modello"].get().strip() == cur_dev["marca_modello"]:
            self.v["marca_modello"].set("")
        self.v["seriale"].set(wanted[0]["seriale"] if wanted else "")
        if wanted and not self.v["marca_modello"].get().strip():
            self.v["marca_modello"].set(wanted[0]["marca_modello"])

    def _device_label(self, d):
        return VB.etichetta_dispositivo(d)

    def _set_device_list(self, devs=None):
        if devs is None:
            devs = self.hist["dispositivi"][:80]
        self.dev_map = {}
        for d in devs:
            self.dev_map.setdefault(self._device_label(d), d)
        if hasattr(self, "dev_cb"):
            self.dev_cb.configure(values=list(self.dev_map))

    def _apply_device(self, dev, override=False):
        for k in ("tipologia", "marca_modello", "seriale"):
            val = dev.get(k, "")
            if k in self.v and (override or (val and not self.v[k].get().strip())):
                self.v[k].set(val)         # scegliendo dall'elenco sostituisco tutto, anche con il vuoto
        uk = self.user_key
        if uk and dev.get("utente") and not self.v[uk].get().strip():
            self.v[uk].set(dev["utente"])
            self._user_picked()

    def _set_serial_values(self, vals=None):
        """L'elenco dei seriali del campo: tutti, oppure solo quelli della persona scelta."""
        if self.serial_cb is None:
            return
        vals = list(vals) if vals is not None else self._serial_values()
        self._serial_all[:] = vals            # e' la stessa lista che usa il completamento automatico
        self.serial_cb.configure(values=vals)

    def _owner_lists(self, nome, devs=None):
        """Con una persona scelta mostro solo i SUOI apparati e seriali; senza, quelli di tutti.
        La spunta «anche gli altri» aggiunge in coda gli apparati delle altre persone."""
        allv = self.show_all.get()
        if nome and devs:
            self._set_device_list(devs + self.hist["dispositivi"][:80] if allv else devs)
            sers = [d["seriale"] for d in devs if d["seriale"]]
            if allv:
                have = {s.casefold() for s in sers}
                sers += [x for x in self._serial_values() if x.casefold() not in have]
            self._set_serial_values(sers)
            note = ("prima gli apparati di %s, poi quelli degli altri" if allv
                    else "solo gli apparati di %s") % nome
        elif nome:
            self._set_device_list()
            self._set_serial_values()
            note = "di %s non ho apparati: scegli o scrivi il seriale" % nome
        else:
            self._set_device_list()
            self._set_serial_values()
            note = ""
        if self.dev_hint_lbl is not None:
            self.dev_hint_lbl.configure(text=note or "compila tipologia, modello e seriale")

    def _toggle_all(self, *_):
        u = self._current_user()
        self._owner_lists(u["nome"] if u else None, self._owner_devs if u else None)

    def _user_devices(self, u):
        devs = VB.dispositivi_utente(self.hist, u["nome"], u)
        same = self._owner_key == key_name(u["nome"])
        prev = self._owner_devs
        self._owner_key, self._owner_devs = key_name(u["nome"]), devs
        self._owner_lists(u["nome"], devs)
        if "tipologia" not in self.v:
            return
        keys = ("tipologia", "marca_modello", "seriale")
        cur = tuple(self.v[k].get().strip() for k in keys)
        ser = VB._codice_key(cur[2])
        mine = (not any(cur) or cur == self._dev_auto
                or (not same and bool(ser) and any(VB._codice_key(d["seriale"]) == ser for d in prev)))
        if not mine:
            return                       # l'hai scritto o scelto tu: non lo tocco
        if devs:
            self._apply_device(devs[0], override=True)
            self._dev_auto = tuple(self.v[k].get().strip() for k in keys)
        elif any(cur):
            for k in keys:               # erano gli apparati della persona di prima
                self.v[k].set("")
            self._dev_auto = None

    def _device_chosen(self, event=None):
        dev = self.dev_map.get(self.dev_var.get())
        if dev:
            self._apply_device(dev, override=True)

    def _serial_picked(self, event=None):
        dev = VB.trova_dispositivo(self.hist, self.v["seriale"].get(), self.app.users_data)
        if dev:
            self._apply_device(dev)

    def _serial_values(self):
        vals, seen = [], set()
        for x in [d["seriale"] for d in self.hist["dispositivi"]] + \
                [u.get(k, "") for u in self.app.users_data for k in ("seriale_pc", "imei")] + \
                [d["seriale"] for u in self.app.users_data
                 for d in VB.dispositivi_utente({"dispositivi": []}, u["nome"], u)]:
            if x and x.casefold() not in seen:
                seen.add(x.casefold())
                vals.append(x)
        return vals[:300]

    def _device_row(self):
        self.dev_var = tk.StringVar()

        def build(fr):
            self.dev_cb = ttk.Combobox(fr, textvariable=self.dev_var, state="readonly", width=62)
            self.dev_cb.pack(side="left")
            self.dev_cb.bind("<<ComboboxSelected>>", self._device_chosen)
            self.dev_cb.bind("<MouseWheel>", self._wheel_on_combo)
            self.dev_hint_lbl = ttk.Label(fr, text="compila tipologia, modello e seriale",
                                          style="Hint.TLabel")
            self.dev_hint_lbl.pack(side="left", padx=(8, 0))
            ttk.Checkbutton(fr, text="mostra anche gli altri", variable=self.show_all,
                            command=self._toggle_all).pack(side="left", padx=(10, 0))
        self._line("Apparato già noto", build)
        self._set_device_list()

    def _dot_hint(self, *_):
        """Assegnazione: per un notebook o uno smartphone propongo il seriale/IMEI della persona."""
        u = self._current_user() if self.tipo == "assegnazione" else None
        if not u or not hasattr(self, "dv"):
            return
        cur = self.dv["codice"].get().strip()
        if cur and cur != self._dot_auto:
            return                                    # l'hai scritto tu: non lo tocco
        tip = VB.tipologia_da_dotazione(self.dv["tipo"].get())
        val = u["seriale_pc"] if tip == "Notebook" else u["imei"] if tip == "Smartphone" else ""
        self._dot_auto = val
        self.dv["codice"].set(val or "")

    def _tipologia(self, label="Tipologia *"):
        var, _ = self._combo(label, "tipologia", self._hist("tipologia", VB.TIPOLOGIE), 28)
        var.trace_add("write", self._device_hint)

    # ----- i quattro tipi di verbale -----
    def _form_trasferimento(self):
        self._device_row()
        self._tipologia()
        self._hcombo("Marca / Modello", "marca_modello", "marca_modello", 44)
        self._serial_all = self._serial_values()
        _, cb = self._combo("Numero di serie", "seriale", self._serial_all, 36, scan=self._serial_picked)
        self.serial_cb = cb
        self._on_pick(cb, self._serial_picked)
        luoghi = self.app.verbale_luoghi()
        self._hcombo("Provenienza (struttura)", "provenienza", "provenienza", 64, base=luoghi)
        self._hcombo("Destinazione (struttura)", "destinazione", "destinazione", 64, base=luoghi)
        self._user_combo("Assegnatario", "assegnatario")
        self._hcombo("Contatti", "contatti", "contatti", 36)
        self._motivi(VB.MOTIVI_TRASFERIMENTO)
        self._text("Note", "note", 3)
        self._entry("Data spedizione", "data_sped", 14, hint="se la lasci vuota uso la data del verbale")

    def _form_presa_in_carico(self):
        self._device_row()
        self._tipologia("Tipologia dispositivo *")
        self._hcombo("Marca e modello", "marca_modello", "marca_modello", 44)
        self._serial_all = self._serial_values()
        _, cb = self._combo("Numero seriale", "seriale", self._serial_all, 36, scan=self._serial_picked)
        self.serial_cb = cb
        self._on_pick(cb, self._serial_picked)
        self._user_combo("Utente assegnatario *", "utente")
        self._entry("N° INC", "inc", 24, hint="il ticket dell'incidente")
        self._entry("N° RIC", "ric", 24, hint="la richiesta")
        self._text("Elenco materiale consegnato", "materiale", 4,
                   hint="Una voce per riga (esempio: caricabatterie, cover).", suggest="materiale")
        self._motivi(VB.MOTIVI_PRESA_IN_CARICO)

    def _form_assegnazione(self):
        self._combo("Titolo", "titolo", TITOLI_ELENCO, 12,
                    hint="facoltativo: scrive «al dott.», «all'avv.», «alla dott.ssa»…")
        self._user_combo("Utente *", "utente")
        self._entry("Matricola", "matricola", 20)
        self._dotazione_box()
        self._text("Accessori inclusi", "accessori", 3,
                   hint="Una voce per riga (esempio: zaino porta notebook).", suggest="accessori")

    def _form_censimento(self):
        st = self.app.settings
        _, cb = self._hcombo("Oggetto *", "oggetto", "oggetto", 62,
                             hint="esempio: Dispositivi Microsoft Lumia 950")
        self._on_pick(cb, self._oggetto_picked)
        self._hcombo("Nel testo: «costituito da»", "descrizione", "descrizione", 62,
                     hint="esempio: dispositivi smartphone Microsoft Lumia")
        _, cb = self._hcombo("Sede (città)", "sede_citta", "sede_citta", 26,
                             base=self.app.sedi_cache, default=st.get("sede_citta", ""))
        self._on_pick(cb, self._sede_picked)
        self._hcombo("Indirizzo", "indirizzo", "indirizzo", 40, default=st.get("indirizzo", ""))
        self._hcombo("CAP", "cap", "cap", 10, default=st.get("cap", ""))
        # indirizzo e CAP messi da me (non scritti da te): se cambi citta' li sostituisco
        self._sede_auto = {k: self.v[k].get().strip() for k in ("indirizzo", "cap")
                           if self.v[k].get().strip() and not self.ini.get(k)}
        self._hcombo("Tecnico *", "tecnico", "tecnico", 36, base=[st.get("tecnico", "")],
                     default=st.get("tecnico", ""), hint="chi redige il verbale")
        self._hcombo("Indicazioni ricevute da", "collega", "collega", 40,
                     base=self.app.users_cache,
                     hint="facoltativo: se lo lasci vuoto la frase sulla dismissione non compare")
        var, _ = self._combo("Titolo colonna codice", "col_codice", VB.COLONNE_CODICE, 20,
                             default="IMEI")
        var.trace_add("write", lambda *_: self._fill_apparati())
        self._apparati_box()

    def _oggetto_picked(self, event=None):
        info = self.hist["oggetti"].get(self.v["oggetto"].get().strip().casefold())
        if info:
            self._set_if_empty("descrizione", info.get("descrizione", ""))

    def _sede_picked(self, event=None):
        citta = self.v["sede_citta"].get().strip().casefold()
        info = self.hist["sedi"].get(citta)
        st = self.app.settings
        if not info and citta and citta == st.get("sede_citta", "").strip().casefold():
            info = {"indirizzo": st.get("indirizzo", ""), "cap": st.get("cap", "")}
        for k in ("indirizzo", "cap"):
            if self.v[k].get().strip() and self.v[k].get().strip() == self._sede_auto.get(k):
                self.v[k].set("")                    # era quello della citta' di prima
                self._sede_auto.pop(k, None)
            new = (info or {}).get(k, "")
            if new and not self.v[k].get().strip():
                self.v[k].set(new)
                self._sede_auto[k] = new

    def _modello_picked(self, event=None):
        marca = self.hist["modelli"].get(self.v_modello.get().strip().casefold())
        cur = self.v_marca.get().strip()
        if marca and (not cur or cur == self._marca_auto):
            self.v_marca.set(marca)
            self._marca_auto = marca

    # ----- elenco degli apparecchi del censimento -----
    def _tree(self, parent, columns, height):
        tv = ttk.Treeview(parent, columns=[c[0] for c in columns], show="headings", height=height,
                          selectmode="extended")
        for c, title, width in columns:
            tv.heading(c, text=title)
            tv.column(c, width=width, anchor="w")
        return tv

    def _apparati_box(self):
        box = ttk.LabelFrame(self.body, text="Apparecchi censiti", padding=10)
        box.grid(row=self.row, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        self.row += 1
        box.columnconfigure(0, weight=1)
        self.tv_app = self._tree(box, (("n", "N.", 44), ("marca", "Marca", 170),
                                       ("modello", "Modello", 170), ("codice", "Codice", 280)), 7)
        self.tv_app.grid(row=0, column=0, sticky="ew")
        self.tv_app.bind("<Double-1>", lambda e: self._edit_apparato())
        self.tv_app.bind("<Delete>", lambda e: self._del_apparati())
        add = ttk.Frame(box)
        add.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        last = self.apparati[-1] if self.apparati else {}
        self._marca_auto = last.get("marca", "")
        self.v_marca = tk.StringVar(value=last.get("marca", ""))
        self.v_modello = tk.StringVar(value=last.get("modello", ""))
        self.v_codice = tk.StringVar()
        for label, var, width, campo in (("Marca", self.v_marca, 16, "ap_marca"),
                                         ("Modello", self.v_modello, 16, "ap_modello"),
                                         ("Codice", self.v_codice, 24, "")):
            ttk.Label(add, text=label).pack(side="left", padx=(0, 4))
            if campo:
                values = self._hist(campo)
                e = ttk.Combobox(add, textvariable=var, values=values, width=width)
                bind_autocomplete(e, var, lambda v=values: v)
                if campo == "ap_modello":
                    self._on_pick(e, self._modello_picked)
            else:
                e = ttk.Entry(add, textvariable=var, width=width)
            e.pack(side="left", padx=(0, 10) if var is not self.v_codice else (0, 4))
            e.bind("<Return>", lambda ev: self._add_apparato(), add="+")
            if var is self.v_codice:
                self.e_codice = e
                scan_button(add, var, multi=True, entry=e, on_code=lambda t: self._add_apparato(),
                            kind_fn=lambda: "imei" if self.v["col_codice"].get().strip().upper() == "IMEI"
                            else "serial").pack(side="left", padx=(0, 10))
        ttk.Button(add, text="Aggiungi", style="Primary.TButton",
                   command=self._add_apparato).pack(side="left")
        row2 = ttk.Frame(box)
        row2.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(row2, text="Rimuovi selezionati", command=self._del_apparati).pack(side="left")
        ttk.Button(row2, text="Incolla elenco…", command=self._paste_apparati).pack(
            side="left", padx=(8, 0))
        self.lbl_count = ttk.Label(row2, style="Hint.TLabel")
        self.lbl_count.pack(side="right")
        ttk.Label(box, style="Hint.TLabel", wraplength=800, justify="left",
                  text="Marca e modello restano scritti: per i prossimi apparecchi basta il codice "
                       "e Invio. Doppio clic su una riga per correggerla.").grid(
            row=3, column=0, sticky="w", pady=(6, 0))
        self.lbl_warn = ttk.Label(box, foreground=RED, wraplength=800, justify="left")
        self.lbl_warn.grid(row=4, column=0, sticky="w", pady=(4, 0))
        self._fill_apparati()

    def _fill_apparati(self):
        if not hasattr(self, "tv_app"):
            return
        self.tv_app.delete(*self.tv_app.get_children())
        for i, r in enumerate(self.apparati):
            self.tv_app.insert("", "end", iid=str(i), values=(
                i + 1, r.get("marca", ""), r.get("modello", ""), r.get("codice", "")))
        self.lbl_count.configure(text="Totale: %d" % len(self.apparati))
        col = self.v["col_codice"].get() or "IMEI"
        avvisi = VB.controlla_apparati(self.apparati, col)
        text = "\n".join(avvisi[:4]) + ("\n…e altri %d da controllare" % (len(avvisi) - 4)
                                        if len(avvisi) > 4 else "")
        self.lbl_warn.configure(text=text)

    def _add_apparato(self):
        marca, modello = self.v_marca.get().strip(), self.v_modello.get().strip()
        codice = self.v_codice.get().strip()
        if not (marca or modello or codice):
            return
        self.apparati.append({"marca": marca, "modello": modello, "codice": codice})
        self._marca_auto = marca                     # per il prossimo resta scritta questa
        self.v_codice.set("")
        self._fill_apparati()
        self.e_codice.focus_set()

    def _selected_indexes(self, tv):
        return sorted((int(i) for i in tv.selection()), reverse=True)

    def _del_apparati(self):
        for i in self._selected_indexes(self.tv_app):
            del self.apparati[i]
        self._fill_apparati()

    def _edit_apparato(self):
        idx = self._selected_indexes(self.tv_app)
        if len(idx) != 1:
            return
        r = self.apparati.pop(idx[0])
        self._marca_auto = ""
        self.v_marca.set(r.get("marca", ""))
        self.v_modello.set(r.get("modello", ""))
        self.v_codice.set(r.get("codice", ""))
        self._fill_apparati()
        self.e_codice.focus_set()

    def _paste_apparati(self):
        def ok(text):
            rows = VB.parse_apparati_text(text, self.v_marca.get().strip(),
                                          self.v_modello.get().strip())
            self.apparati.extend(rows)
            self._fill_apparati()
        PasteDialog(self, "Incolla elenco apparecchi",
                    "Incolla qui l'elenco, una riga per apparecchio. Puoi copiarlo da Excel "
                    "(colonne: marca, modello, codice) oppure scrivere solo i codici, uno per riga: "
                    "in quel caso uso la marca e il modello scritti nel modulo.", ok)

    # ----- dotazione dell'assegnazione -----
    def _dotazione_box(self):
        box = ttk.LabelFrame(self.body, text="Dotazione assegnata *", padding=10)
        box.grid(row=self.row, column=0, columnspan=2, sticky="ew", pady=(12, 8))
        self.row += 1
        box.columnconfigure(0, weight=1)
        self.tv_dot = self._tree(box, (("tipo", "Oggetto", 170), ("modello", "Modello", 190),
                                       ("codice", "S/N o IMEI", 180), ("iccid", "ICCID", 170),
                                       ("numero", "Numero", 110)), 5)
        self.tv_dot.grid(row=0, column=0, sticky="ew")
        self.tv_dot.bind("<Double-1>", lambda e: self._edit_dot())
        self.tv_dot.bind("<Delete>", lambda e: self._del_dot())
        self.dv = {k: tk.StringVar() for k in ("tipo", "modello", "codice", "iccid", "numero")}
        self.dv["tipo"].set(VB.TIPI_DOTAZIONE[0])
        self.dv["tipo"].trace_add("write", self._dot_hint)
        r1 = ttk.Frame(box)
        r1.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(r1, text="Oggetto").pack(side="left", padx=(0, 4))
        ttk.Combobox(r1, textvariable=self.dv["tipo"], values=VB.TIPI_DOTAZIONE, width=22).pack(
            side="left", padx=(0, 10))
        ttk.Label(r1, text="Modello").pack(side="left", padx=(0, 4))
        dot_models = self._hist("dot_modello")
        cb_mod = ttk.Combobox(r1, textvariable=self.dv["modello"], values=dot_models, width=22)
        bind_autocomplete(cb_mod, self.dv["modello"], lambda: dot_models)
        cb_mod.pack(side="left", padx=(0, 10))
        ttk.Label(r1, text="S/N o IMEI").pack(side="left", padx=(0, 4))
        e = ttk.Entry(r1, textvariable=self.dv["codice"], width=22)
        e.pack(side="left")
        e.bind("<Return>", lambda ev: self._add_dot())
        scan_button(r1, self.dv["codice"], entry=e, kind_fn=lambda: "imei" if VB.tipologia_da_dotazione(
            self.dv["tipo"].get()) == "Smartphone" else "serial").pack(side="left", padx=(6, 0))
        r2 = ttk.Frame(box)
        r2.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        ttk.Label(r2, text="Solo per la SIM:  ICCID").pack(side="left", padx=(0, 4))
        ttk.Entry(r2, textvariable=self.dv["iccid"], width=26).pack(side="left", padx=(0, 10))
        ttk.Label(r2, text="Numero").pack(side="left", padx=(0, 4))
        ttk.Entry(r2, textvariable=self.dv["numero"], width=18).pack(side="left", padx=(0, 10))
        ttk.Button(r2, text="Aggiungi", style="Primary.TButton", command=self._add_dot).pack(
            side="left")
        ttk.Button(r2, text="Rimuovi selezionati", command=self._del_dot).pack(
            side="left", padx=(8, 0))
        ttk.Label(box, style="Hint.TLabel", wraplength=800, justify="left",
                  text="Se non scrivi il codice resta una riga ______ da compilare a mano. "
                       "Doppio clic su una riga per correggerla.").grid(
            row=3, column=0, sticky="w", pady=(6, 0))
        self._fill_dot()

    def _fill_dot(self):
        if not hasattr(self, "tv_dot"):
            return
        self.tv_dot.delete(*self.tv_dot.get_children())
        for i, r in enumerate(self.dotazione):
            self.tv_dot.insert("", "end", iid=str(i), values=(
                r.get("tipo", ""), r.get("modello", ""), r.get("codice", ""),
                r.get("iccid", ""), r.get("numero", "")))

    def _add_dot(self):
        r = {k: v.get().strip() for k, v in self.dv.items()}
        if not r["tipo"]:
            return
        self.dotazione.append(r)
        for k in ("modello", "codice", "iccid", "numero"):
            self.dv[k].set("")
        self._fill_dot()

    def _del_dot(self):
        for i in self._selected_indexes(self.tv_dot):
            del self.dotazione[i]
        self._fill_dot()

    def _edit_dot(self):
        idx = self._selected_indexes(self.tv_dot)
        if len(idx) != 1:
            return
        r = self.dotazione.pop(idx[0])
        for k in self.dv:
            self.dv[k].set(r.get(k, ""))
        self._fill_dot()

    # ----- foto -----
    def _photo_box(self):
        obbl = VB.TIPI[self.tipo][2]
        box = ttk.LabelFrame(
            self.body, padding=10,
            text="Foto da allegare (%s)" % ("OBBLIGATORIA: almeno una" if obbl else "facoltativa"))
        box.grid(row=self.row, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        self.row += 1
        box.columnconfigure(0, weight=1)
        self.tv_ph = self._tree(box, (("nome", "File", 600),), 4)
        self.tv_ph.grid(row=0, column=0, sticky="ew")
        btns = ttk.Frame(box)
        btns.grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Button(btns, text="Aggiungi foto…", command=self._add_photos).pack(side="left")
        ttk.Button(btns, text="Incolla immagine (dagli appunti)", command=self._paste_photo).pack(
            side="left", padx=(8, 0))
        ttk.Button(btns, text="Rimuovi selezionate", command=self._del_photos).pack(
            side="left", padx=(8, 0))
        ttk.Label(box, style="Hint.TLabel", wraplength=800, justify="left",
                  text="Le foto vengono messe in fondo al verbale, in una pagina «Allegato "
                       "fotografico».").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self._fill_photos()

    def _fill_photos(self):
        self.tv_ph.delete(*self.tv_ph.get_children())
        for i, p in enumerate(self.photos):
            self.tv_ph.insert("", "end", iid=str(i), values=(os.path.basename(p),))

    def _add_photos(self):
        files = filedialog.askopenfilenames(
            parent=self, title="Scegli le foto da allegare",
            filetypes=[("Immagini", "*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tif *.tiff"),
                       ("Tutti i file", "*.*")])
        for f in files or ():
            if os.path.isfile(f) and f not in self.photos:
                self.photos.append(f)
        self._fill_photos()

    def _paste_photo(self):
        found = grab_clipboard_images(staging_dir())
        if not found:
            messagebox.showinfo(APP_NAME, "Negli appunti non c'è nessuna immagine.\nCopia una "
                                "foto o fai uno screenshot e riprova.", parent=self)
            return
        for f in found:
            if f not in self.photos:
                self.photos.append(f)
        self._fill_photos()

    def _del_photos(self):
        for i in self._selected_indexes(self.tv_ph):
            del self.photos[i]
        self._fill_photos()

    # ----- creazione -----
    def _flush_pending(self):
        """Se hai scritto un apparecchio o un oggetto e non hai premuto «Aggiungi», lo aggiungo io."""
        if self.tipo == "censimento" and self.v_codice.get().strip():
            self._add_apparato()
        if self.tipo == "assegnazione":
            codice = self.dv["codice"].get().strip()
            # il seriale proposto da me (dalla scheda utente) non e' un oggetto da aggiungere in silenzio
            if (self.dv["modello"].get().strip() or self.dv["iccid"].get().strip()
                    or (codice and codice != (self._dot_auto or ""))):
                self._add_dot()

    def collect(self):
        d = {k: v.get().strip() for k, v in self.v.items()}
        for k, t in self.t.items():
            raw = t.get("1.0", "end").strip()
            d[k] = [x.strip() for x in raw.splitlines() if x.strip()] if k in VERB_LIST_KEYS else raw
        if self.checks:
            d["motivi"] = [o for o, var in self.checks.items() if var.get()]
        if self.tipo == "censimento":
            d["apparati"] = [dict(r) for r in self.apparati]
        if self.tipo == "assegnazione":
            d["dotazione"] = [dict(r) for r in self.dotazione]
        return d

    def validate(self, d):
        dt = parse_date(d["data"])
        if not dt:
            return "La data non è valida.\nScrivila così: gg/mm/aaaa (esempio 15/11/2026)."
        d["data"] = fmt_date(dt)
        if d.get("data_sped"):
            ds = parse_date(d["data_sped"])
            if not ds:
                return "La data di spedizione non è valida.\nScrivila così: gg/mm/aaaa."
            d["data_sped"] = fmt_date(ds)
        if self.tipo == "trasferimento" and not d["tipologia"]:
            return "Scrivi la tipologia dell'apparato (esempio: Notebook)."
        if self.tipo == "presa_in_carico":
            if not d["tipologia"]:
                return "Scrivi la tipologia del dispositivo (esempio: Notebook)."
            if not d["utente"]:
                return "Scrivi l'utente assegnatario."
        if self.tipo == "assegnazione":
            if not d["utente"]:
                return "Scrivi il nome dell'utente a cui assegni il materiale."
            if not d["dotazione"]:
                return "Aggiungi almeno un oggetto alla dotazione assegnata."
        if self.tipo == "censimento":
            if not d["oggetto"]:
                return "Scrivi l'oggetto del verbale (esempio: Dispositivi Microsoft Lumia 950)."
            if not d["tecnico"]:
                return "Scrivi il nome del tecnico che redige il verbale."
            if not d["apparati"]:
                return "Aggiungi almeno un apparecchio all'elenco."
        if VB.TIPI[self.tipo][2] and not self.photos:
            return "Per il %s la foto è obbligatoria.\nAllega almeno una foto." % \
                   VB.TIPI[self.tipo][0].lower()
        return None

    def create(self):
        self._flush_pending()
        d = self.collect()
        err = self.validate(d)
        if err:
            messagebox.showwarning(APP_NAME, err, parent=self)
            return
        if self.tipo == "censimento":
            avvisi = VB.controlla_apparati(d["apparati"], d["col_codice"] or "IMEI")
            if avvisi:
                more = "\n…" if len(avvisi) > 8 else ""
                if not messagebox.askyesno(
                        APP_NAME, "Ci sono punti da controllare:\n\n%s%s\n\nCreo comunque il "
                                  "verbale?" % ("\n".join(avvisi[:8]), more), parent=self):
                    return
        self.firme_scelte = self.firme_panel.scelte()      # lette da make_verbale
        if self.on_create(self.tipo, d, list(self.photos), self):
            self.destroy()


class SerieVerbaliDialog(ModalDialog):
    """Verbali in serie: scegli piano o stanza e il tipo di verbale, e per ogni persona ne esce uno."""

    def __init__(self, app, on_create):
        super().__init__(app, "Verbali in serie")
        self.resizable(False, False)
        self.app, self.on_create = app, on_create
        self.eds = {b["nome"]: b for b in app.db.edifici()}
        self.v_ed = tk.StringVar(value=next(iter(self.eds), ""))
        self.v_piano, self.v_stanza = tk.StringVar(), tk.StringVar()
        self.v_tipo = tk.StringVar(value=VB.TIPI[SERIE_TIPI[0]][0])
        self.v_cosa = tk.StringVar(value=SERIE_COSA[0][1])
        self.v_motivo = tk.StringVar()
        self.v_data = tk.StringVar(value=fmt_date(date.today()))
        self.v_dir = tk.StringVar(value=app.settings.get("cartella_verbali", ""))
        self.v_firma = tk.BooleanVar(value=bool(saved_signature()))
        self.v_info = tk.StringVar(value="")
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Un verbale Word per ogni persona", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(body, style="Hint.TLabel", wraplength=520, justify="left",
                  text="Scegli dove siedono le persone e il tipo di verbale: i dati (nome, matricola, "
                       "seriali, dotazione) li prendo dalle loro schede. Chi non ha niente da "
                       "scrivere viene saltato e te lo dico."
                  ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 10))
        self.cb_ed = self._row(2, "Edificio", ttk.Combobox(body, textvariable=self.v_ed, state="readonly",
                                                           values=list(self.eds), width=34), body)
        self.cb_piano = self._row(3, "Piano", ttk.Combobox(body, textvariable=self.v_piano,
                                                           state="readonly", width=34), body)
        self.cb_stanza = self._row(4, "Stanza", ttk.Combobox(body, textvariable=self.v_stanza,
                                                             state="readonly", width=34), body)
        self.cb_tipo = self._row(5, "Verbale", ttk.Combobox(
            body, textvariable=self.v_tipo, state="readonly",
            values=[VB.TIPI[t][0] for t in SERIE_TIPI], width=34), body)
        self.cb_cosa = self._row(6, "Cosa consegni", ttk.Combobox(
            body, textvariable=self.v_cosa, state="readonly", values=[c[1] for c in SERIE_COSA],
            width=34), body)
        self.cb_motivo = self._row(7, "Motivo (facoltativo)", ttk.Combobox(
            body, textvariable=self.v_motivo, values=[""] + list(VB.MOTIVI_PRESA_IN_CARICO),
            width=34), body)
        self._row(8, "Data", ttk.Entry(body, textvariable=self.v_data, width=16), body)
        fr = ttk.Frame(body)
        ttk.Entry(fr, textvariable=self.v_dir, width=34).pack(side="left")
        ttk.Button(fr, text="Scegli…", command=self.pick_dir).pack(side="left", padx=(6, 0))
        self._row(9, "Salva nella cartella", fr, body)
        self.chk_firma = ttk.Checkbutton(body, text="Metti la mia firma (tecnico) in ogni verbale",
                                         variable=self.v_firma, command=self.update_info)
        self.chk_firma.grid(row=10, column=0, columnspan=3, sticky="w", pady=(8, 0))
        ttk.Label(body, style="Hint.TLabel", wraplength=520, justify="left",
                  text="La firma di chi riceve resta da fare (a mano sul foglio o col dito sul "
                       "telefono). Se non hai ancora salvato la tua firma, la scegli in «Dati per i "
                       "verbali».").grid(row=11, column=0, columnspan=3, sticky="w", pady=(2, 0))
        ttk.Label(body, wraplength=520, justify="left", textvariable=self.v_info).grid(
            row=12, column=0, columnspan=3, sticky="w", pady=(12, 0))
        btns = ttk.Frame(body)
        btns.grid(row=13, column=0, columnspan=3, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Crea i verbali", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))
        self.cb_ed.bind("<<ComboboxSelected>>", lambda e: self.refill_ed())
        self.cb_piano.bind("<<ComboboxSelected>>", lambda e: self.refill_piano())
        for cb in (self.cb_stanza, self.cb_tipo, self.cb_cosa, self.cb_motivo):
            cb.bind("<<ComboboxSelected>>", lambda e: self.update_info())
        self.v_data.trace_add("write", lambda *_: self.update_info())
        self.v_tipo.trace_add("write", lambda *_: self.update_info())
        self.refill_ed()

    def _row(self, row, label, widget, body):
        ttk.Label(body, text=label).grid(row=row, column=0, sticky="w", pady=4, padx=(0, 14))
        widget.grid(row=row, column=1, sticky="w", pady=4)
        return widget

    def pick_dir(self):
        d = filedialog.askdirectory(parent=self, title="Dove salvo i verbali Word?",
                                    initialdir=self.v_dir.get() or None)
        if d:
            self.v_dir.set(d)

    def refill_ed(self):
        b = self.eds.get(self.v_ed.get())
        self.floors = {}
        if b:
            for r in self.app.db.stanze_of(b["id"]):
                self.floors.setdefault(r["piano"], []).append(r)
        labels = [floor_label(n) for n in sorted(self.floors, reverse=True)]
        self.cb_piano.configure(values=labels)
        if self.v_piano.get() not in labels:
            self.v_piano.set(labels[0] if labels else "")
        self.refill_piano()

    def refill_piano(self):
        piano = next((n for n in self.floors if floor_label(n) == self.v_piano.get()), None)
        self.rooms = {"Tutto il piano": None}
        for r in sorted(self.floors.get(piano, []), key=lambda r: natural_key(r["codice"])):
            self.rooms["Stanza %s (%d persone)" % (r["codice"], r["n_persone"])] = r["id"]
        self.cb_stanza.configure(values=list(self.rooms))
        if self.v_stanza.get() not in self.rooms:
            self.v_stanza.set("Tutto il piano")
        self.update_info()

    def params(self):
        b = self.eds.get(self.v_ed.get())
        piano = next((n for n in self.floors if floor_label(n) == self.v_piano.get()), None)
        tipo = next((t for t in SERIE_TIPI if VB.TIPI[t][0] == self.v_tipo.get()), SERIE_TIPI[0])
        cosa = next((k for k, lab in SERIE_COSA if lab == self.v_cosa.get()), "pc")
        return {"eid": b["id"] if b else None, "piano": piano,
                "stanza_id": self.rooms.get(self.v_stanza.get()), "tipo": tipo, "cosa": cosa,
                "motivo": self.v_motivo.get().strip(), "data": self.v_data.get().strip(),
                "cartella": self.v_dir.get().strip(), "firma_tecnico": bool(self.v_firma.get())}

    def update_info(self):
        presa = self.params()["tipo"] == "presa_in_carico"
        self.cb_cosa.configure(state="readonly" if presa else "disabled")
        self.cb_motivo.configure(state="normal" if presa else "disabled")
        self.v_info.set(self.app.serie_preview(self.params()))

    def save(self):
        if self.on_create(self.params()):
            self.destroy()


class MergeDuplicatesDialog(ModalDialog):
    """Persone doppie: per ogni gruppo scegli quale tenere (di solito quella scritta «Cognome Nome»);
    le altre vengono unite a lei: interventi e postazione passano, i dati che mancano si completano."""

    def __init__(self, app, groups, on_merge):
        super().__init__(app, "Unisci persone doppie")
        self.on_merge = on_merge
        self.geometry("860x%d" % min(700, max(420, self.winfo_screenheight() - 140)))
        self.rows = []                                   # (gruppo, spunta, scelta, {etichetta: id})
        outer = ttk.Frame(self)
        outer.pack(fill="both", expand=True)
        bar = ttk.Frame(outer, padding=(16, 10))
        bar.pack(side="bottom", fill="x")
        ttk.Button(bar, text="Chiudi", command=self.destroy).pack(side="right")
        ttk.Button(bar, text="Unisci le persone spuntate", style="Save.TButton",
                   command=self.save).pack(side="right", padx=(0, 8))
        ttk.Button(bar, text="Spunta anche i probabili", command=self.tick_likely).pack(
            side="right", padx=(0, 8))
        self.v_count = tk.StringVar(value="")
        ttk.Label(bar, style="Hint.TLabel", textvariable=self.v_count).pack(side="left")
        sf = ScrollFrame(outer)
        sf.pack(fill="both", expand=True)
        body = sf.inner
        ttk.Label(body, style="Hint.TLabel", wraplength=800, justify="left",
                  text="Ho trovato queste persone che sembrano doppie. Tengo quella scritta «Cognome "
                       "Nome» (le liste e i verbali vanno in ordine di cognome) e le altre si "
                       "uniscono a lei: i suoi interventi e la postazione passano, e quello che le "
                       "manca (matricola, telefono, email, seriali, dotazione) lo prende dalle "
                       "altre. Spunta solo dove sei sicuro: i casi incerti non sono spuntati."
                  ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        for i, g in enumerate(groups, start=1):
            labels = {self.label(m): m["id"] for m in g["members"]}
            keep = next(k for k, v in labels.items() if v == g["keep"])
            on = tk.BooleanVar(value=bool(g["sure"]))
            choice = tk.StringVar(value=keep)
            ttk.Checkbutton(body, variable=on, command=self.update_count,
                            text="Unisci   (%s)" % "; ".join(g["why"])).grid(
                row=i * 2 - 1, column=0, columnspan=2, sticky="w", pady=(10, 0))
            ttk.Label(body, text="Tieni").grid(row=i * 2, column=0, sticky="w", padx=(24, 8))
            ttk.Combobox(body, textvariable=choice, values=list(labels), state="readonly",
                         width=92).grid(row=i * 2, column=1, sticky="w", pady=2)
            self.rows.append((g, on, choice, labels))
        self.update_count()

    def tick_likely(self):
        """Spunta anche i gruppi «probabili»: una sola persona ha la matricola (viene dalla lista
        INAIL) e le altre no, di solito sono copie scritte a mano."""
        for g, on, choice, labels in self.rows:
            if g.get("likely"):
                on.set(True)
        self.update_count()

    @staticmethod
    def label(m):
        bits = []
        if m.get("matricola"):
            bits.append("matr. %s" % m["matricola"])
        if m.get("n"):
            bits.append("%d interventi" % m["n"])
        if m.get("sede"):
            bits.append(m["sede"])
        return m["nome"] + ("   [%s]" % ", ".join(bits) if bits else "")

    def pairs(self):
        """[(id da togliere, id da tenere)] per i gruppi spuntati."""
        out = []
        for g, on, choice, labels in self.rows:
            if on.get():
                keep = labels.get(choice.get())
                out += [(m["id"], keep) for m in g["members"] if keep and m["id"] != keep]
        return out

    def update_count(self):
        n = len(self.pairs())
        self.v_count.set("%d persone da unire" % n if n else "Nessuna persona spuntata")

    def save(self):
        pairs = self.pairs()
        if not pairs:
            messagebox.showinfo(APP_NAME, "Spunta almeno un gruppo da unire.", parent=self)
            return
        if self.on_merge(pairs):
            self.destroy()


class Walk3DHost:
    """Ponte tra le visite 3D (corridoio3d.py) e i dati dell'app: stanze, persone, stampanti di un piano."""

    def __init__(self, app):
        self.app = app
        self.parent = app
        self._pimg = {}
        self._pcache = (None, 0.0, [])

    # --- dati
    def building(self, eid):
        return next((b for b in self.app.db.edifici() if b["id"] == eid), None)

    def building_name(self, eid):
        b = self.building(eid)
        return b["nome"] if b else ""

    def floors(self, eid):
        """I piani dell'edificio dal piu' basso (anche quelli fuori dall'intervallo ma con stanze o stampanti)."""
        b = self.building(eid)
        if not b:
            return []
        f = set(range(min(b["piano_min"], b["piano_max"]), max(b["piano_min"], b["piano_max"]) + 1))
        f.update(r["piano"] for r in self.app.db.stanze_of(eid))
        f.update(p["piano"] for p in self._printers(eid))
        return sorted(f)

    def _printers(self, eid):
        key, when, rows = self._pcache
        if key != eid or time.time() - when > 3.0:
            rows = self.app.db.printers_of(eid)
            self._pcache = (eid, time.time(), rows)
        return rows

    def floor_data(self, eid, piano):
        db = self.app.db
        rooms = []
        for r in db.stanze_of(eid, piano):
            ppl = db.persone_of(r["id"])
            miss = sum(1 for p in ppl for d in desk_devices(p) if not d["value"])
            rooms.append({"uid": r["uid"], "id": r["id"], "codice": r["codice"],
                          "scrivanie": r["scrivanie"] or 0, "miss": miss,
                          "persone": [{"nome": p["nome"], "ufficio": p.get("ufficio") or ""} for p in ppl]})
        prs = []
        for p in self._printers(eid):
            if p["piano"] != piano:
                continue
            prs.append({"uid": str(p["id"]), "nome": p["nome"], "modello": p["modello"],
                        "stanza": p.get("stanza_codice") or "", "red": p.get("man_richieste") or 0,
                        "low": bool(p.get("scorta_bassa")), "fuori": bool(p.get("fuori_servizio"))})
        return {"edificio": self.building_name(eid), "piano": piano, "label": floor_label(piano),
                "rooms": rooms, "printers": prs}

    def printer_image(self, modello):
        key = "%s_m" % printer_image_key(modello)
        if key not in self._pimg:
            try:
                from PIL import Image
                self._pimg[key] = Image.open(io.BytesIO(base64.b64decode(PRINTER_IMAGES[key]))).convert("RGBA")
            except Exception:
                self._pimg[key] = None
        return self._pimg[key]

    # --- la foto del personaggio: solo in questa cartella, mai nei backup, nei file per il telefono o su GitHub
    def avatar_path(self):
        return os.path.join(data_dir(), C3D.AVATAR_NAME)

    # --- azioni
    def open_room(self, room_id):
        if self.app.db.get_stanza(room_id):
            return RoomDialog(self.app, room_id)
        return None

    def open_walk(self, eid, piano):
        self.app.open_walk_3d(eid, piano)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("%s - Registro interventi" % APP_NAME)
        self.geometry("1240x%d" % min(900, max(680, self.winfo_screenheight() - 90)))
        self.minsize(1080, 740)
        self.configure(bg=LIGHT)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        self.db = Database(os.path.join(data_dir(), "interventi.db"))
        self.edit_id = None
        self.current_rows = []
        self.rows_by_id = {}
        self.users_data = []
        self.users_cache = []
        self.users_by_key = {}
        self.sedi_cache = []
        self.settings = load_settings()
        self.verb_rows = {}
        self.sort_col = "data"
        self.sort_rev = True
        self.buildings = []
        self.building_by_label = {}
        self.printer_imgs = {}
        self.toner_rows = []
        self.toner_rows_by_id = {}
        self.toner_printer_labels = {}
        self.manut_rows = []
        self.manut_rows_by_id = {}
        self.scorta_keys = {}
        self.thumb_cache = {}
        self.att_items = []
        self.att_removed = []
        self.users_visible = []
        self._filter_job = None
        self._status_job = None
        self._draw_job = None
        self._toner_job = None
        self._manut_job = None
        self._autofilled_for = None
        self._loading_form = False

        self._setup_style()
        self._build_menu()
        self._build_header()
        self._build_tabs()

        self.refresh_users()
        self.refresh_stats()
        self.refresh_buildings()
        self.apply_filters()
        self.refresh_verbali()
        self.clear_form()
        self.bind("<Control-s>", lambda e: self.save_intervento())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(700, self.check_model_names)

    # ---------------- nomi dei modelli ----------------
    def check_model_names(self):
        """Una volta sola: propone di correggere i vecchi nomi sbagliati (C3351i, C360i)."""
        if self.settings.get("modelli_rinomina_chiesto") == "1":
            return
        changed = False
        for wrong, right in MODELLI_DA_CORREGGERE:
            for name in self.known_model_names():
                if model_key(name) != model_key(wrong):
                    continue
                n_p, n_s = self.db.count_model(name)
                if not (n_p or n_s):
                    continue
                if messagebox.askyesno(
                        APP_NAME,
                        "Nel programma ci sono %d stampant%s (e %d voc%s della scorta toner) con il "
                        "modello \"%s\".\n\nSul fronte delle stampanti di Via delle Porte Nuove "
                        "c'è scritto \"%s\". Vuoi cambiare il nome in \"%s\"?\n\n"
                        "Se qualcuna è davvero un altro modello rispondi No: potrai cambiarlo "
                        "una stampante alla volta dalla sua scheda, oppure da File › Correggi il "
                        "nome di un modello stampante." % (
                            n_p, "e" if n_p == 1 else "i", n_s, "e" if n_s == 1 else "i",
                            name, right, right)):
                    self.db.rename_model(name, right)
                    changed = True
        self.settings["modelli_rinomina_chiesto"] = "1"
        save_settings(self.settings)
        if changed:
            self.after_printers_change()

    def known_model_names(self):
        names = list(self.db.models())
        for r in self.db.scorta_all():
            if model_key(r["modello"]) not in {model_key(x) for x in names}:
                names.append(r["modello"])
        return names

    def rename_model_dialog(self):
        names = self.known_model_names()
        if not names:
            messagebox.showinfo(APP_NAME, "Non ci sono ancora modelli di stampante nel programma.")
            return
        ModelRenameDialog(self, names, self.db.count_model, self.apply_model_rename)

    def apply_model_rename(self, old, new):
        n_p, n_s = self.db.count_model(old)
        if not messagebox.askyesno(
                APP_NAME, "Cambio \"%s\" in \"%s\" su %d stampant%s e %d voc%s della scorta "
                          "toner?\n\nPoi manda di nuovo i dati al telefono, così il cambio arriva "
                          "anche lì." % (old, new, n_p, "e" if n_p == 1 else "i", n_s,
                                         "e" if n_s == 1 else "i")):
            return False
        try:
            self.db.rename_model(old, new)
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a cambiare il nome:\n%s" % e)
            return False
        self.after_printers_change()
        messagebox.showinfo(APP_NAME, "Fatto: ora il modello è \"%s\"." % new)
        return True

    # ---------------- stile ----------------
    def _setup_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        bold = ("Segoe UI", 10, "bold")
        scale = self.winfo_fpixels("1i") / 96.0
        s.configure(".", background=LIGHT, foreground=NAVY, font=("Segoe UI", 10))
        s.configure("TFrame", background=LIGHT)
        s.configure("TLabel", background=LIGHT, foreground=NAVY)
        s.configure("Hint.TLabel", foreground=GREY)
        s.configure("Count.TLabel", font=("Segoe UI", 13, "bold"))
        s.configure("TLabelframe", background=LIGHT, bordercolor=BORDER)
        s.configure("TLabelframe.Label", background=LIGHT, foreground=NAVY, font=bold)
        s.configure("TNotebook", background=LIGHT, borderwidth=0)
        s.configure("TNotebook.Tab", padding=(20, 9), font=bold, background=BTN_BG, foreground=NAVY)
        s.map("TNotebook.Tab", background=[("selected", NAVY)], foreground=[("selected", WHITE)])
        s.configure("TButton", padding=(12, 6), background=BTN_BG, foreground=NAVY, borderwidth=0)
        s.map("TButton", background=[("active", BTN_BG_ACTIVE)])
        for name, bg, active in (("Primary", NAVY, NAVY_DARK), ("Save", GREEN, GREEN_DARK),
                                 ("Danger", RED, RED_DARK)):
            s.configure(name + ".TButton", background=bg, foreground=WHITE, font=bold)
            s.map(name + ".TButton", background=[("active", active), ("disabled", "#9aa8b6")],
                  foreground=[("disabled", "#eef2f6")])
        s.configure("TEntry", padding=5, fieldbackground=WHITE, bordercolor=BORDER)
        s.configure("TCombobox", padding=5, fieldbackground=WHITE, bordercolor=BORDER,
                    arrowcolor=NAVY)
        s.map("TCombobox", fieldbackground=[("readonly", WHITE)],
              foreground=[("readonly", TEXT)])
        s.configure("Treeview", rowheight=int(26 * scale), background=WHITE, fieldbackground=WHITE,
                    foreground=TEXT, bordercolor=BORDER)
        s.map("Treeview", background=[("selected", NAVY)], foreground=[("selected", WHITE)])
        s.configure("Treeview.Heading", background=NAVY, foreground=WHITE, font=bold,
                    padding=(6, 6), relief="flat")
        s.map("Treeview.Heading", background=[("active", NAVY_DARK)])

    def _build_menu(self):
        menubar = tk.Menu(self)
        m = tk.Menu(menubar, tearoff=0)
        m.add_command(label="Importa dati dal telefono…", command=self.import_from_phone)
        m.add_command(label="Esporta dati per il telefono…", command=self.export_for_phone)
        m.add_command(label="Importa elenco stanze e persone…", command=self.import_rooms_file)
        m.add_command(label="Importa persone da Excel (lista INAIL)…",
                      command=self.import_people_excel)
        m.add_command(label="Importa seriali dei PC da Excel…", command=self.import_pc_serials)
        m.add_command(label="Importa interventi da ServiceNow (Excel)…",
                      command=self.import_servicenow)
        m.add_command(label="Unisci persone doppie…", command=self.merge_duplicates_dialog)
        m.add_command(label="Unisci doppioni stampanti…", command=self.merge_printers_dialog)
        m.add_separator()
        m.add_command(label="Correggi il nome di un modello stampante…",
                      command=self.rename_model_dialog)
        m.add_command(label="Mappa della Toscana…", command=self.open_map)
        m.add_command(label="Passeggiata nel piano in 3D…", command=self.open_walk_3d)
        m.add_command(label="Edificio intero in 3D…", command=self.open_tower_3d)
        m.add_command(label="Verbali in serie (un Word per ogni persona)…", command=self.verbali_serie)
        m.add_command(label="Dati per i verbali Word…", command=self.verbali_settings)
        m.add_command(label="Apri la cartella dei modelli Word", command=self.open_templates_folder)
        m.add_separator()
        m.add_command(label="Salva una copia di backup (dati e foto)…", command=self.backup)
        m.add_command(label="Apri la cartella dei dati", command=self.open_data_folder)
        m.add_separator()
        m.add_command(label="Esci", command=self.on_close)
        menubar.add_cascade(label="File", menu=m)
        self.config(menu=menubar)

    def _build_header(self):
        header = tk.Frame(self, bg=WHITE)
        header.pack(fill="x")
        inner = tk.Frame(header, bg=WHITE)
        inner.pack(fill="x", padx=18, pady=10)

        try:
            self.logo_img = tk.PhotoImage(file=resource_path("logo_header.png"))
            tk.Label(inner, image=self.logo_img, bg=WHITE).pack(side="left")
        except tk.TclError:
            tk.Label(inner, text="INAIL", font=("Segoe UI", 26, "bold"), fg=NAVY,
                     bg=WHITE).pack(side="left")
        tk.Label(inner, text="Help Desk · Registro interventi", font=("Segoe UI", 17, "bold"),
                 fg=NAVY, bg=WHITE).pack(side="left", padx=(26, 0))

        self.stat_vars = {}
        self.stat_labels = {}
        stats = tk.Frame(inner, bg=WHITE)
        stats.pack(side="right")
        for key, label in (("aperti", "Da seguire"), ("oggi", "Oggi"),
                           ("mese", "Questo mese"), ("tot", "Totale")):
            box = tk.Frame(stats, bg=LIGHT, padx=14, pady=6, highlightthickness=1,
                           highlightbackground=BORDER)
            box.pack(side="left", padx=(8, 0))
            var = tk.StringVar(value="0")
            self.stat_vars[key] = var
            num = tk.Label(box, textvariable=var, font=("Segoe UI", 18, "bold"), fg=NAVY,
                           bg=LIGHT)
            num.pack()
            cap = tk.Label(box, text=label, font=("Segoe UI", 9), fg=GREY, bg=LIGHT)
            cap.pack()
            self.stat_labels[key] = num
            if key == "aperti":
                for w in (box, num, cap):
                    w.configure(cursor="hand2")
                    w.bind("<Button-1>", lambda e: self.show_open_interventi())

        bar = tk.Frame(self, height=5, bg=LIGHT)
        bar.pack(fill="x")
        for color in (GREEN, LIGHT, RED):
            tk.Frame(bar, bg=color, height=5).pack(side="left", fill="x", expand=True)

    def _build_tabs(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=14, pady=(10, 14))
        self.tab_new = ttk.Frame(self.nb, padding=18)
        self.tab_reg = ttk.Frame(self.nb, padding=14)
        self.tab_users = ttk.Frame(self.nb, padding=18)
        self.tab_print = ttk.Frame(self.nb, padding=10)
        self.tab_verb = ttk.Frame(self.nb, padding=14)
        self.nb.add(self.tab_new, text="Nuovo intervento")
        self.nb.add(self.tab_reg, text="Registro e filtri")
        self.nb.add(self.tab_users, text="Utenti")
        self.nb.add(self.tab_print, text="Stampanti")
        self.nb.add(self.tab_verb, text="Verbali")
        self._build_form()
        self._build_registry()
        self._build_users()
        self._build_printers()
        self._build_verbali()

    # ---------------- scheda: nuovo intervento ----------------
    def _build_form(self):
        f = self.tab_new
        f.columnconfigure(1, weight=1)
        self.form_title = ttk.Label(f, text="Nuovo intervento", font=("Segoe UI", 14, "bold"))
        self.form_title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.var_dt = tk.StringVar()
        self.var_user = tk.StringVar()
        self.var_matricola = tk.StringVar()
        self.var_stato = tk.StringVar(value="Risolto")
        self.var_sede = tk.StringVar()
        self.var_tel = tk.StringVar()
        self.var_seriale = tk.StringVar()
        self.var_imei = tk.StringVar()
        self.var_ticket = tk.StringVar()

        ttk.Label(f, text="Data e ora").grid(row=1, column=0, sticky="nw", pady=6, padx=(0, 14))
        row = ttk.Frame(f)
        row.grid(row=1, column=1, sticky="w", pady=4)
        ttk.Entry(row, textvariable=self.var_dt, width=20).pack(side="left")
        ttk.Button(row, text="Adesso", command=self.set_now).pack(side="left", padx=(8, 0))
        ttk.Label(row, text="formato gg/mm/aaaa oo:mm", style="Hint.TLabel").pack(
            side="left", padx=(10, 0))

        ttk.Label(f, text="Utente *").grid(row=2, column=0, sticky="nw", pady=6, padx=(0, 14))
        uf = ttk.Frame(f)
        uf.grid(row=2, column=1, sticky="ew", pady=4)
        self.cb_user = ttk.Combobox(uf, textvariable=self.var_user, width=30)
        self.cb_user.grid(row=0, column=0, sticky="w")
        bind_autocomplete(self.cb_user, self.var_user, lambda: self.users_cache)
        self.cb_user.bind("<Return>", self.autocomplete_user)
        ttk.Label(uf, text="Matricola").grid(row=0, column=1, padx=(20, 6))
        self.ent_matricola = ttk.Entry(uf, textvariable=self.var_matricola, width=12)
        self.ent_matricola.grid(row=0, column=2, sticky="w")
        self.ent_matricola.bind("<Return>", self.autocomplete_matricola)
        self.ent_matricola.bind("<FocusOut>", self.autocomplete_matricola)
        ttk.Label(uf, text="Sede").grid(row=0, column=3, padx=(20, 6))
        self.cb_sede = ttk.Combobox(uf, textvariable=self.var_sede, width=18)
        self.cb_sede.grid(row=0, column=4, sticky="w")
        bind_autocomplete(self.cb_sede, self.var_sede, lambda: self.sedi_cache)
        self.lbl_hint = ttk.Label(uf, text="", style="Hint.TLabel", wraplength=760, justify="left")
        self.lbl_hint.grid(row=1, column=0, columnspan=5, sticky="w", pady=(3, 0))
        self.var_user.trace_add("write", self.update_user_hint)

        ttk.Label(f, text="Postazione\ndell'utente").grid(row=3, column=0, sticky="nw", pady=6,
                                                         padx=(0, 14))
        pf = ttk.Frame(f)
        pf.grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(pf, text="Telefono").grid(row=0, column=0, sticky="w")
        ttk.Entry(pf, textvariable=self.var_tel, width=18).grid(row=0, column=1, padx=(6, 18))
        ttk.Label(pf, text="Seriale PC").grid(row=0, column=2, sticky="w")
        e_ser = ttk.Entry(pf, textvariable=self.var_seriale, width=20)
        e_ser.grid(row=0, column=3, padx=(6, 4))
        scan_button(pf, self.var_seriale, "serial", entry=e_ser).grid(row=0, column=4, padx=(0, 14))
        ttk.Label(pf, text="IMEI cellulare").grid(row=0, column=5, sticky="w")
        e_imei = ttk.Entry(pf, textvariable=self.var_imei, width=20)
        e_imei.grid(row=0, column=6, padx=(6, 4))
        scan_button(pf, self.var_imei, "imei", entry=e_imei).grid(row=0, column=7, padx=(0, 14))
        ttk.Button(pf, text="Copia dati utente", command=self.copy_user_data).grid(row=0, column=8)
        ttk.Label(pf, style="Hint.TLabel",
                  text="Si compilano da soli se l'utente è già in lista; quello che scrivi qui "
                       "viene salvato con l'utente. «Fotocamera…» legge il codice a barre o il QR "
                       "dell'etichetta.").grid(row=1, column=0, columnspan=9, sticky="w", pady=(3, 0))

        ttk.Label(f, text="N. ticket\n(facoltativo)").grid(row=4, column=0, sticky="nw", pady=6,
                                                           padx=(0, 14))
        tk4 = ttk.Frame(f)
        tk4.grid(row=4, column=1, sticky="w", pady=4)
        ttk.Entry(tk4, textvariable=self.var_ticket, width=24).pack(side="left")
        ttk.Label(tk4, text="Stato").pack(side="left", padx=(24, 6))
        ttk.Combobox(tk4, textvariable=self.var_stato, values=STATI_INTERVENTO, width=12,
                     state="readonly").pack(side="left")
        ttk.Label(tk4, style="Hint.TLabel",
                  text="Aperto o In attesa = lo ritrovi in «Da seguire» (in alto a destra)."
                  ).pack(side="left", padx=(12, 0))

        ttk.Label(f, text="Per cosa mi hanno\nchiamato *").grid(row=5, column=0, sticky="nw",
                                                               pady=6, padx=(0, 14))
        self.t_richiesta = make_text(f, 3)
        self.t_richiesta.grid(row=5, column=1, sticky="ew", pady=4)

        ttk.Label(f, text="Come ho risolto").grid(row=6, column=0, sticky="nw", pady=6,
                                                  padx=(0, 14))
        self.t_soluzione = make_text(f, 5)
        self.t_soluzione.grid(row=6, column=1, sticky="ew", pady=4)

        ttk.Label(f, text="Note").grid(row=7, column=0, sticky="nw", pady=6, padx=(0, 14))
        self.t_note = make_text(f, 2)
        self.t_note.grid(row=7, column=1, sticky="ew", pady=4)

        ttk.Label(f, text="Foto allegate").grid(row=8, column=0, sticky="nw", pady=6, padx=(0, 14))
        af = ttk.Frame(f)
        af.grid(row=8, column=1, sticky="ew", pady=4)
        ab = ttk.Frame(af)
        ab.pack(fill="x")
        ttk.Button(ab, text="Aggiungi foto…", command=self.add_photos).pack(side="left")
        ttk.Button(ab, text="Incolla immagine", command=self.paste_photo).pack(
            side="left", padx=(8, 0))
        ttk.Button(ab, text="Apri cartella foto", command=self.open_att_folder).pack(
            side="left", padx=(8, 0))
        ttk.Label(ab, style="Hint.TLabel",
                  text="Restano salvate con l'intervento: dalla cartella le trascini nel ticket."
                  ).pack(side="left", padx=(12, 0))
        self.att_strip = ttk.Frame(af)
        self.att_strip.pack(fill="x", pady=(6, 0))

        btns = ttk.Frame(f)
        btns.grid(row=9, column=1, sticky="w", pady=(14, 0))
        self.btn_save = ttk.Button(btns, text="Salva intervento", style="Save.TButton",
                                   command=self.save_intervento)
        self.btn_save.pack(side="left")
        self.btn_clear = ttk.Button(btns, text="Svuota", command=self.clear_form)
        self.btn_clear.pack(side="left", padx=(8, 0))
        self.lbl_status = ttk.Label(f, text="", foreground=GREEN, font=("Segoe UI", 10, "bold"))
        self.lbl_status.grid(row=10, column=1, sticky="w", pady=(10, 0))
        ttk.Label(f, text="* campi obbligatori. Scorciatoia: Ctrl+S per salvare.",
                  style="Hint.TLabel").grid(row=11, column=1, sticky="w", pady=(4, 0))

    def set_now(self):
        self.var_dt.set(datetime.now().strftime("%d/%m/%Y %H:%M"))

    def set_status(self, text, color=GREEN):
        self.lbl_status.configure(text=text, foreground=color)
        if self._status_job:
            self.after_cancel(self._status_job)
        self._status_job = self.after(6000, lambda: self.lbl_status.configure(text=""))

    def _set_station(self, sede="", tel="", seriale="", imei="", matricola=""):
        self.var_sede.set(sede)
        self.var_tel.set(tel)
        self.var_seriale.set(seriale)
        self.var_imei.set(imei)
        self.var_matricola.set(matricola)

    def update_user_hint(self, *_):
        typed = self.var_user.get().strip()
        key = key_name(typed)
        match = self.users_by_key.get(key) if typed else None
        if not self._loading_form:
            if match and self._autofilled_for != key:
                self._autofilled_for = key
                self._set_station(match["sede"], match["telefono"], match["seriale_pc"],
                                  match["imei"], match["matricola"])
            elif not match and self._autofilled_for is not None:
                self._autofilled_for = None
                self._set_station()
        if not typed:
            self.lbl_hint.configure(text="", foreground=GREY)
        elif match:
            self.lbl_hint.configure(text="Utente già in lista.", foreground=GREEN)
        else:
            terms = typed.casefold().split()
            similar = [u for u in self.users_cache if matches_all(terms, u)]
            if similar:
                self.lbl_hint.configure(
                    text="Simili in lista: %s. Se è la stessa persona scegli il nome dal menu; "
                         "altrimenti verrà aggiunto come nuovo utente." % ", ".join(similar[:5]),
                    foreground=NAVY)
            else:
                self.lbl_hint.configure(
                    text="Nuovo utente: verrà aggiunto alla lista quando salvi.", foreground=NAVY)

    def autocomplete_user(self, event=None):
        terms = self.var_user.get().casefold().split()
        if not terms:
            return
        found = [u for u in self.users_cache if matches_all(terms, u)]
        if len(found) == 1:
            self.var_user.set(found[0])
            self.cb_user.icursor("end")

    def autocomplete_matricola(self, event=None):
        """Scrivi la matricola e (se il nome è vuoto) trovo io l'utente."""
        m = self.var_matricola.get().strip()
        if not m or self.var_user.get().strip():
            return
        found = self.db.find_user_by_matricola(m)
        if found:
            self.var_user.set(found["nome"])
            self.cb_user.icursor("end")

    def copy_user_data(self):
        lines = []
        known = self.db.find_user_like(self.var_user.get()) if self.var_user.get().strip() else None
        for label, text in (("Utente", self.var_user.get()), ("Matricola", self.var_matricola.get()),
                            ("Sede", self.var_sede.get()),
                            ("Telefono", self.var_tel.get()),
                            ("Email", (known or {}).get("email") or ""),
                            ("Seriale PC", self.var_seriale.get()),
                            ("IMEI cellulare", self.var_imei.get())):
            if text.strip():
                lines.append("%s: %s" % (label, text.strip()))
        if not lines:
            messagebox.showinfo(APP_NAME, "Non c'è ancora nessun dato da copiare.")
            return
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self.set_status("Dati dell'utente copiati: ora puoi incollarli nel ticket.")

    # --- foto allegate ---
    def thumb(self, path):
        if path not in self.thumb_cache:
            b64 = make_thumb_b64(path)
            img = None
            if b64:
                try:
                    img = tk.PhotoImage(data=b64)
                except tk.TclError:
                    img = None
            self.thumb_cache[path] = img
        return self.thumb_cache[path]

    def attachment_items(self, iid):
        return [{"id": a["id"], "nome": a["nome"], "new": False,
                 "path": os.path.join(attachments_root(), *a["file"].split("/"))}
                for a in self.db.allegati_of(iid)]

    def render_strip(self, parent, items, removable):
        for child in parent.winfo_children():
            child.destroy()
        for i, it in enumerate(items):
            cell = ttk.Frame(parent)
            cell.grid(row=i // 8, column=i % 8, padx=(0, 10), pady=(0, 6), sticky="n")
            img = self.thumb(it["path"])
            if img:
                lbl = tk.Label(cell, image=img, bg=LIGHT, cursor="hand2")
                lbl.image = img
            else:
                lbl = tk.Label(cell, text="foto", width=9, height=4, bg=WHITE, fg=GREY,
                               relief="solid", cursor="hand2")
            lbl.pack()
            lbl.bind("<Button-1>", lambda e, path=it["path"]: self.open_file(path))
            ttk.Label(cell, text=one_line(it["nome"], 14), style="Hint.TLabel").pack()
            if removable:
                ttk.Button(cell, text="Togli", command=lambda idx=i: self.remove_photo(idx)).pack(
                    pady=(2, 0))

    def update_att_strip(self):
        self.render_strip(self.att_strip, self.att_items, True)

    def open_file(self, path):
        if not open_path(path):
            messagebox.showinfo(APP_NAME, "Non riesco ad aprire il file:\n%s" % path)

    def _queue_photo(self, path, staged=False):
        if any(it["path"] == path for it in self.att_items):
            return
        self.att_items.append({"id": None, "nome": os.path.basename(path), "path": path,
                               "new": True, "staged": staged})

    def add_photos(self):
        paths = filedialog.askopenfilenames(
            title="Scegli le foto da allegare",
            filetypes=[("Immagini", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.tif *.tiff *.heic"),
                       ("Tutti i file", "*.*")])
        for path in paths:
            self._queue_photo(path)
        self.update_att_strip()

    def paste_photo(self):
        paths = grab_clipboard_images(staging_dir())
        if not paths:
            messagebox.showinfo(APP_NAME, "Negli appunti non c'è nessuna immagine.\n"
                                "Fai uno screenshot (ad esempio Win+Maiusc+S) e premi di nuovo "
                                "\"Incolla immagine\".")
            return
        for path in paths:
            self._queue_photo(path, staged=os.path.dirname(path) == staging_dir())
        self.update_att_strip()

    def remove_photo(self, idx):
        it = self.att_items.pop(idx)
        if it["id"] is not None:
            self.att_removed.append(it)
        elif it.get("staged"):
            try:
                os.remove(it["path"])
            except OSError:
                pass
        self.update_att_strip()

    def discard_staged(self):
        for it in self.att_items:
            if it.get("staged"):
                try:
                    os.remove(it["path"])
                except OSError:
                    pass

    def commit_attachments(self, iid):
        """Salva le foto scelte nella cartella dell'intervento; toglie quelle rimosse."""
        for it in self.att_removed:
            self.db.delete_allegato(it["id"])
            try:
                os.remove(it["path"])
            except OSError:
                pass
        added = 0
        dest = attachment_dir(iid)
        for it in self.att_items:
            if not it["new"]:
                continue
            try:
                name = store_image(it["path"], dest)
            except OSError as e:
                messagebox.showerror(APP_NAME, "Non sono riuscito a salvare la foto %s:\n%s"
                                     % (it["nome"], e))
                continue
            self.db.add_allegato(iid, name, "%s/%s" % (iid, name))
            added += 1
            if it.get("staged"):
                try:
                    os.remove(it["path"])
                except OSError:
                    pass
        return added

    def open_att_folder(self):
        if self.edit_id and any(not it["new"] for it in self.att_items):
            self.open_folder(attachment_dir(self.edit_id))
        else:
            messagebox.showinfo(APP_NAME, "La cartella con le foto esiste dopo il salvataggio. "
                                "Salva l'intervento, poi riaprilo dal registro (doppio clic) "
                                "oppure usa \"Apri cartella foto\" nel registro.")

    def open_folder(self, folder):
        if not open_path(folder):
            messagebox.showinfo(APP_NAME, "Le foto sono in:\n%s" % folder)

    def open_row_folder(self):
        r = self.selected_row()
        if not r:
            messagebox.showinfo(APP_NAME, "Seleziona prima un intervento dall'elenco.")
        elif not r["n_foto"]:
            messagebox.showinfo(APP_NAME, "Questo intervento non ha foto allegate.")
        else:
            self.open_folder(attachment_dir(r["id"]))

    def clear_form(self):
        self._loading_form = True
        self.edit_id = None
        self.discard_staged()
        self.att_items = []
        self.att_removed = []
        self._autofilled_for = None
        self.set_now()
        self.var_user.set("")
        self._set_station()
        self.var_ticket.set("")
        self.var_stato.set("Risolto")
        for w in (self.t_richiesta, self.t_soluzione, self.t_note):
            w.delete("1.0", "end")
        self._loading_form = False
        self.form_title.configure(text="Nuovo intervento")
        self.btn_save.configure(text="Salva intervento")
        self.btn_clear.configure(text="Svuota")
        self.update_user_hint()
        self.update_att_strip()
        self.cb_user.focus_set()

    def save_intervento(self):
        if self.nb.select() != str(self.tab_new):
            return
        dt = parse_datetime(self.var_dt.get())
        utente = normalize_name(self.var_user.get())
        sede = normalize_name(self.var_sede.get())
        tel = self.var_tel.get().strip()
        seriale = self.var_seriale.get().strip()
        imei = self.var_imei.get().strip()
        matricola = self.var_matricola.get().strip()
        stato = self.var_stato.get() if self.var_stato.get() in STATI_INTERVENTO else "Risolto"
        richiesta = self.t_richiesta.get("1.0", "end").strip()
        soluzione = self.t_soluzione.get("1.0", "end").strip()
        note = self.t_note.get("1.0", "end").strip()
        ticket = self.var_ticket.get().strip()

        if not dt:
            messagebox.showwarning(APP_NAME, "La data non è valida.\nScrivila così: "
                                   "gg/mm/aaaa oo:mm (esempio 13/11/2026 14:30).")
            return
        if not utente:
            messagebox.showwarning(APP_NAME, "Scrivi il nome dell'utente che ti ha chiamato.")
            self.cb_user.focus_set()
            return
        if not richiesta:
            messagebox.showwarning(APP_NAME, "Scrivi per cosa ti hanno chiamato.")
            self.t_richiesta.focus_set()
            return

        data_ora = dt.strftime("%Y-%m-%d %H:%M")
        try:
            if self.edit_id:
                created = self.db.update_intervento(self.edit_id, data_ora, utente, richiesta,
                                                    soluzione, note, ticket, sede, tel, seriale,
                                                    imei, matricola, stato)
                msg = "Modifiche salvate."
                saved_id = self.edit_id
            else:
                saved_id, created = self.db.add_intervento(data_ora, utente, richiesta,
                                                           soluzione, note, ticket, sede, tel,
                                                           seriale, imei, matricola, stato)
                msg = "Intervento salvato."
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return
        n_new = self.commit_attachments(saved_id)
        if n_new:
            msg += " Foto allegate: %d." % n_new
        if created:
            msg += " Nuovo utente aggiunto alla lista: %s." % utente
        if stato != "Risolto":
            msg += " Resta da seguire (%s)." % stato
        self.clear_form()
        self.refresh_users()
        self.refresh_stats()
        self.apply_filters()
        self.set_status(msg)

    # ---------------- scheda: registro e filtri ----------------
    def _build_registry(self):
        t = self.tab_reg
        t.columnconfigure(0, weight=1)
        t.rowconfigure(2, weight=3)
        t.rowconfigure(3, weight=1)

        self.var_f_user = tk.StringVar(value="(Tutti)")
        self.var_f_sede = tk.StringVar(value="(Tutte)")
        self.var_f_dal = tk.StringVar()
        self.var_f_al = tk.StringVar()
        self.var_f_ticket_mode = tk.StringVar(value="Tutti")
        self.var_f_ticket = tk.StringVar()
        self.var_f_text = tk.StringVar()
        self.var_f_stato = tk.StringVar(value=FILTRO_STATO[0])

        box = ttk.LabelFrame(t, text="Filtri", padding=10)
        box.grid(row=0, column=0, sticky="ew")

        def select_all(combo):
            combo.bind("<FocusIn>", lambda e: self.after_idle(lambda: combo.select_range(0, "end")))

        r1 = ttk.Frame(box)
        r1.pack(fill="x")
        ttk.Label(r1, text="Utente").pack(side="left")
        self.cb_filter_user = ttk.Combobox(r1, textvariable=self.var_f_user, width=22)
        self.cb_filter_user.pack(side="left", padx=(6, 14))
        bind_autocomplete(self.cb_filter_user, self.var_f_user,
                          lambda: ["(Tutti)"] + self.users_cache)
        select_all(self.cb_filter_user)
        ttk.Label(r1, text="Sede").pack(side="left")
        self.cb_filter_sede = ttk.Combobox(r1, textvariable=self.var_f_sede, width=14)
        self.cb_filter_sede.pack(side="left", padx=(6, 14))
        bind_autocomplete(self.cb_filter_sede, self.var_f_sede,
                          lambda: ["(Tutte)"] + self.sedi_cache)
        select_all(self.cb_filter_sede)
        ttk.Label(r1, text="Ticket").pack(side="left")
        ttk.Combobox(r1, textvariable=self.var_f_ticket_mode, width=12, state="readonly",
                     values=["Tutti", "Con ticket", "Senza ticket"]).pack(side="left",
                                                                         padx=(6, 14))
        ttk.Label(r1, text="N. ticket").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_f_ticket, width=10).pack(side="left", padx=(6, 14))
        ttk.Label(r1, text="Cerca nel testo").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_f_text, width=22).pack(side="left", padx=(6, 0),
                                                                  fill="x", expand=True)

        r2 = ttk.Frame(box)
        r2.pack(fill="x", pady=(10, 0))
        ttk.Label(r2, text="Dal").pack(side="left")
        self.ent_dal = ttk.Entry(r2, textvariable=self.var_f_dal, width=11)
        self.ent_dal.pack(side="left", padx=(6, 12))
        ttk.Label(r2, text="Al").pack(side="left")
        self.ent_al = ttk.Entry(r2, textvariable=self.var_f_al, width=11)
        self.ent_al.pack(side="left", padx=(6, 18))
        ttk.Label(r2, text="Periodo veloce:", style="Hint.TLabel").pack(side="left")
        for label, kind in (("Oggi", "oggi"), ("7 giorni", "7g"), ("Questo mese", "mese"),
                            ("Quest'anno", "anno"), ("Tutto", "tutto")):
            ttk.Button(r2, text=label, command=lambda k=kind: self.set_period(k)).pack(
                side="left", padx=(6, 0))
        ttk.Button(r2, text="Azzera filtri", style="Primary.TButton",
                   command=self.reset_filters).pack(side="right")

        r3 = ttk.Frame(box)
        r3.pack(fill="x", pady=(10, 0))
        ttk.Label(r3, text="Stato").pack(side="left")
        ttk.Combobox(r3, textvariable=self.var_f_stato, width=32, state="readonly",
                     values=FILTRO_STATO).pack(side="left", padx=(6, 14))
        ttk.Button(r3, text="Mostra solo quelli da richiamare",
                   command=lambda: self.var_f_stato.set(FILTRO_STATO[1])).pack(side="left")
        ttk.Button(r3, text="Mappa della Toscana…", command=self.open_map).pack(side="right")
        ttk.Label(r3, style="Hint.TLabel",
                  text="Aperto / In attesa = interventi ancora da seguire o da richiamare."
                  ).pack(side="left", padx=(12, 0))

        for v in (self.var_f_user, self.var_f_sede, self.var_f_dal, self.var_f_al,
                  self.var_f_ticket_mode, self.var_f_ticket, self.var_f_text, self.var_f_stato):
            v.trace_add("write", self.schedule_filter)

        cnt = ttk.Frame(t)
        cnt.grid(row=1, column=0, sticky="ew", pady=(10, 6))
        self.lbl_count = ttk.Label(cnt, text="", style="Count.TLabel")
        self.lbl_count.pack(side="left")
        self.lbl_filter_desc = ttk.Label(cnt, text="", style="Hint.TLabel")
        self.lbl_filter_desc.pack(side="left", padx=(14, 0))

        tf = ttk.Frame(t)
        tf.grid(row=2, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)
        self.columns = ("data", "utente", "matricola", "sede", "ticket", "stato", "foto",
                        "richiesta", "soluzione")
        self.col_titles = {"data": "Data e ora", "utente": "Utente", "matricola": "Matricola",
                           "sede": "Sede", "ticket": "Ticket", "stato": "Stato", "foto": "Foto",
                           "richiesta": "Per cosa mi hanno chiamato",
                           "soluzione": "Come ho risolto"}
        self.tree = ttk.Treeview(tf, columns=self.columns, show="headings", selectmode="browse")
        widths = {"data": 120, "utente": 150, "matricola": 75, "sede": 85, "ticket": 70,
                  "stato": 75, "foto": 45, "richiesta": 280, "soluzione": 280}
        for c in self.columns:
            self.tree.heading(c, text=self.col_titles[c], command=lambda c=c: self.sort_by(c))
            self.tree.column(c, width=widths[c], anchor="center" if c == "foto" else "w",
                             stretch=(c in ("richiesta", "soluzione")))
        for stato, color in STATO_COLORI.items():
            self.tree.tag_configure(stato, foreground=color)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<<TreeviewSelect>>", self.show_detail)
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())
        self.tree.bind("<Delete>", lambda e: self.delete_selected())

        df = ttk.Frame(t)
        df.grid(row=3, column=0, sticky="nsew", pady=(8, 0))
        df.columnconfigure(0, weight=1)
        df.rowconfigure(0, weight=1)
        self.detail = tk.Text(df, height=8, wrap="word", font=("Segoe UI", 10), bg=WHITE, fg=TEXT,
                              relief="flat", highlightthickness=1, highlightbackground=BORDER,
                              padx=8, pady=6, state="disabled")
        self.detail.tag_configure("h", font=("Segoe UI", 10, "bold"), foreground=NAVY)
        for stato, color in STATO_COLORI.items():
            self.detail.tag_configure("st_" + stato, font=("Segoe UI", 10, "bold"),
                                      foreground=color)
        dsb = ttk.Scrollbar(df, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.grid(row=0, column=0, sticky="nsew")
        dsb.grid(row=0, column=1, sticky="ns")
        self.det_strip = ttk.Frame(df)
        self.det_strip.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        bf = ttk.Frame(t)
        bf.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Modifica", command=self.edit_selected).pack(side="left")
        ttk.Button(bf, text="Duplica", command=self.duplicate_selected).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Segna come risolto", style="Save.TButton",
                   command=self.mark_resolved).pack(side="left", padx=(8, 0))
        ttk.Button(bf, text="Elimina", style="Danger.TButton",
                   command=self.delete_selected).pack(side="left", padx=(8, 0))
        ttk.Button(bf, text="Apri cartella foto", command=self.open_row_folder).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Esporta in PDF", style="Primary.TButton",
                   command=lambda: self.export("pdf")).pack(side="right")
        ttk.Button(bf, text="Esporta in Excel", style="Primary.TButton",
                   command=lambda: self.export("xlsx")).pack(side="right", padx=(0, 8))

    def set_period(self, kind):
        dal, al = period_dates(kind)
        self.var_f_dal.set(fmt_date(dal) if dal else "")
        self.var_f_al.set(fmt_date(al) if al else "")

    def reset_filters(self):
        self.var_f_user.set("(Tutti)")
        self.var_f_sede.set("(Tutte)")
        self.var_f_dal.set("")
        self.var_f_al.set("")
        self.var_f_ticket_mode.set("Tutti")
        self.var_f_ticket.set("")
        self.var_f_text.set("")
        self.var_f_stato.set(FILTRO_STATO[0])

    def show_open_interventi(self):
        self.reset_filters()
        self.var_f_stato.set(FILTRO_STATO[1])
        self.nb.select(self.tab_reg)

    # --- visite 3D (passeggiata nel piano, edificio intero)
    def _walk_host(self):
        if C3D is None:
            messagebox.showinfo(
                APP_NAME, "Le visite 3D non sono disponibili: nel programma manca il file "
                          "corridoio3d.py. Caricalo su GitHub insieme a helpdesk.py e aspetta che "
                          "il programma si ricrei.")
            return None
        if self.__dict__.get("_w3d") is None:
            self._w3d = Walk3DHost(self)
        return self._w3d

    def _open_3d_window(self, attr, factory):
        old = self.__dict__.get(attr)
        if old is not None:
            try:
                if old.winfo_exists():
                    old.close()
            except Exception:
                pass
        win = factory()
        setattr(self, attr, win)
        return win

    def open_walk_3d(self, eid=None, piano=None, mode=None):
        """Passeggiata nel piano (di default il piano scelto in «Piani e stanze»)."""
        host = self._walk_host()
        if host is None:
            return None
        b = self.current_building()
        if eid is None:
            if not b:
                messagebox.showinfo(APP_NAME, "Crea prima un edificio (scheda Edificio).")
                return None
            eid = b["id"]
        if piano is None:
            piano = self.rooms_floor()
            if piano is None:
                piano = 0
        if not self.db.stanze_of(eid, piano):
            messagebox.showinfo(
                APP_NAME, "Su questo piano non ci sono ancora stanze: creale in «Piani e stanze» "
                          "(o importa l'elenco con File > Importa elenco stanze e persone).")
            return None
        return self._open_3d_window("_walk_win", lambda: C3D.WalkWindow(host, eid, piano, mode))

    def open_tower_3d(self):
        """L'edificio scelto, intero, con i suoi piani impilati."""
        host = self._walk_host()
        if host is None:
            return None
        b = self.current_building()
        if not b:
            messagebox.showinfo(APP_NAME, "Crea prima un edificio (scheda Edificio).")
            return None
        return self._open_3d_window("_tower_win", lambda: C3D.TowerWindow(host, b["id"]))

    def open_map(self):
        MapDialog(self)

    def show_sede_in_registry(self, sede, periodo="tutto", solo_aperti=False):
        """Dalla mappa al Registro: gli interventi di una sede (nel periodo, o solo quelli da seguire)."""
        self.reset_filters()
        self.var_f_sede.set(sede)
        if solo_aperti:
            self.var_f_stato.set(FILTRO_STATO[1])
        elif periodo == "30g":
            self.var_f_dal.set(fmt_date(date.today() - timedelta(days=29)))
            self.var_f_al.set(fmt_date(date.today()))
        elif periodo in ("anno", "mese"):
            self.set_period(periodo)
        self.nb.select(self.tab_reg)

    def schedule_filter(self, *_):
        if self._filter_job:
            self.after_cancel(self._filter_job)
        self._filter_job = self.after(250, self.apply_filters)

    def get_filters(self):
        user = self.var_f_user.get().strip()
        if user.casefold() in ("(tutti)", "tutti"):
            user = ""
        sede = self.var_f_sede.get().strip()
        if sede.casefold() in ("(tutte)", "tutte"):
            sede = ""
        dal_txt = self.var_f_dal.get().strip()
        al_txt = self.var_f_al.get().strip()
        dal = parse_date(dal_txt) if dal_txt else None
        al = parse_date(al_txt) if al_txt else None
        self.ent_dal.configure(foreground=RED if dal_txt and not dal else TEXT)
        self.ent_al.configure(foreground=RED if al_txt and not al else TEXT)
        mode = {"Con ticket": "con", "Senza ticket": "senza"}.get(
            self.var_f_ticket_mode.get(), "tutti")
        st = self.var_f_stato.get()
        stato = "aperti" if st == FILTRO_STATO[1] else (st if st in STATI_INTERVENTO else "")
        return {"utente": user, "sede": sede, "dal": dal, "al": al, "ticket_mode": mode,
                "ticket": self.var_f_ticket.get().strip(), "testo": self.var_f_text.get().strip(),
                "stato": stato}

    @staticmethod
    def describe_filters(f):
        parts = []
        if f["utente"]:
            parts.append("utente \"%s\"" % f["utente"])
        if f["sede"]:
            parts.append("sede \"%s\"" % f["sede"])
        if f["dal"] or f["al"]:
            parts.append("periodo %s - %s" % (fmt_date(f["dal"]) if f["dal"] else "inizio",
                                              fmt_date(f["al"]) if f["al"] else "oggi"))
        if f["ticket_mode"] == "con":
            parts.append("solo con ticket")
        elif f["ticket_mode"] == "senza":
            parts.append("solo senza ticket")
        if f["ticket"]:
            parts.append("ticket \"%s\"" % f["ticket"])
        if f["testo"]:
            parts.append("testo \"%s\"" % f["testo"])
        if f.get("stato") == "aperti":
            parts.append("solo da richiamare (aperti e in attesa)")
        elif f.get("stato"):
            parts.append("stato %s" % f["stato"])
        return "; ".join(parts) if parts else "nessun filtro (tutti gli interventi)"

    def apply_filters(self):
        self._filter_job = None
        f = self.get_filters()
        self.current_rows = self.db.search(**f)
        self.render_rows()
        total = self.db.stats()[2]
        users = len({r["utente"] for r in self.current_rows})
        self.lbl_count.configure(text="Interventi trovati: %d  (su %d totali)"
                                      % (len(self.current_rows), total))
        self.lbl_filter_desc.configure(text="Utenti diversi: %d" % users)

    def render_rows(self):
        keys = {
            "data": lambda r: r["data_ora"],
            "utente": lambda r: r["utente"].casefold(),
            "matricola": lambda r: r["matricola"].casefold(),
            "sede": lambda r: r["sede"].casefold(),
            "ticket": lambda r: r["ticket"].casefold(),
            "stato": lambda r: STATI_INTERVENTO.index(r["stato"]) if r["stato"] in
                                STATI_INTERVENTO else 9,
            "foto": lambda r: r["n_foto"],
            "richiesta": lambda r: r["richiesta"].casefold(),
            "soluzione": lambda r: r["soluzione"].casefold(),
        }
        self.current_rows.sort(key=keys[self.sort_col], reverse=self.sort_rev)
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        self.rows_by_id = {}
        for r in self.current_rows:
            self.rows_by_id[str(r["id"])] = r
            self.tree.insert("", "end", iid=str(r["id"]), tags=(r["stato"],), values=(
                fmt_dt(r["data_ora"]), r["utente"], r["matricola"], r["sede"], r["ticket"],
                r["stato"], r["n_foto"] or "", one_line(r["richiesta"]),
                one_line(r["soluzione"])))
        for c in self.columns:
            arrow = (" ▼" if self.sort_rev else " ▲") if c == self.sort_col else ""
            self.tree.heading(c, text=self.col_titles[c] + arrow)
        if selected and selected[0] in self.rows_by_id:
            self.tree.selection_set(selected[0])
        else:
            self.show_detail()

    def sort_by(self, col):
        if self.sort_col == col:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_col, self.sort_rev = col, (col == "data")
        self.render_rows()

    def selected_row(self):
        sel = self.tree.selection()
        return self.rows_by_id.get(sel[0]) if sel else None

    def show_detail(self, *_):
        r = self.selected_row()
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        if r:
            self.detail.insert("end", "Data: ", "h")
            self.detail.insert("end", fmt_dt(r["data_ora"]) + "     ")
            self.detail.insert("end", "Utente: ", "h")
            self.detail.insert("end", r["utente"] + "     ")
            if r["matricola"]:
                self.detail.insert("end", "Matricola: ", "h")
                self.detail.insert("end", r["matricola"] + "     ")
            self.detail.insert("end", "Sede: ", "h")
            self.detail.insert("end", (r["sede"] or "-") + "     ")
            self.detail.insert("end", "Ticket: ", "h")
            self.detail.insert("end", (r["ticket"] or "nessuno") + "     ")
            self.detail.insert("end", "Stato: ", "h")
            self.detail.insert("end", r["stato"] + "\n", "st_" + r["stato"])
            station = [(l, r[k]) for l, k in (("Telefono", "telefono"),
                                              ("Seriale PC", "seriale_pc"), ("IMEI", "imei"))
                       if r[k]]
            if station:
                for label, value in station:
                    self.detail.insert("end", label + ": ", "h")
                    self.detail.insert("end", value + "     ")
                self.detail.insert("end", "\n")
            self.detail.insert("end", "\nPer cosa mi hanno chiamato\n", "h")
            self.detail.insert("end", (r["richiesta"] or "-") + "\n\n")
            self.detail.insert("end", "Come ho risolto\n", "h")
            self.detail.insert("end", (r["soluzione"] or "-") + "\n")
            if r["note"]:
                self.detail.insert("end", "\nNote\n", "h")
                self.detail.insert("end", r["note"] + "\n")
        else:
            self.detail.insert("end", "Seleziona un intervento per vedere tutti i dettagli. "
                                      "Doppio clic per modificarlo.", "")
        self.detail.configure(state="disabled")
        items = self.attachment_items(r["id"]) if (r and r["n_foto"]) else []
        self.render_strip(self.det_strip, items, False)

    def edit_selected(self):
        r = self.selected_row()
        if not r:
            messagebox.showinfo(APP_NAME, "Seleziona prima un intervento dall'elenco.")
            return
        self.discard_staged()
        self._loading_form = True
        self.edit_id = r["id"]
        self.var_dt.set(fmt_dt(r["data_ora"]))
        self.var_user.set(r["utente"])
        self._set_station(r["sede"], r["telefono"], r["seriale_pc"], r["imei"], r["matricola"])
        self._autofilled_for = key_name(r["utente"])
        self.var_ticket.set(r["ticket"])
        self.var_stato.set(r["stato"] if r["stato"] in STATI_INTERVENTO else "Risolto")
        for w, key in ((self.t_richiesta, "richiesta"), (self.t_soluzione, "soluzione"),
                       (self.t_note, "note")):
            w.delete("1.0", "end")
            w.insert("1.0", r[key])
        self._loading_form = False
        self.att_items = self.attachment_items(r["id"])
        self.att_removed = []
        self.update_att_strip()
        self.form_title.configure(text="Modifica intervento del %s" % fmt_dt(r["data_ora"]))
        self.btn_save.configure(text="Salva modifiche")
        self.btn_clear.configure(text="Annulla modifica")
        self.update_user_hint()
        self.nb.select(self.tab_new)

    def duplicate_selected(self):
        """Copia l'intervento selezionato in un intervento NUOVO: stessi dati, data e ora di
        adesso, nessuna foto (restano all'originale). L'originale non viene toccato."""
        r = self.selected_row()
        if not r:
            messagebox.showinfo(APP_NAME, "Seleziona prima un intervento dall'elenco.")
            return
        self.discard_staged()
        self._loading_form = True
        self.edit_id = None
        self.set_now()
        self.var_user.set(r["utente"])
        self._set_station(r["sede"], r["telefono"], r["seriale_pc"], r["imei"], r["matricola"])
        self._autofilled_for = key_name(r["utente"])
        self.var_ticket.set(r["ticket"])
        self.var_stato.set(r["stato"] if r["stato"] in STATI_INTERVENTO else "Risolto")
        for w, key in ((self.t_richiesta, "richiesta"), (self.t_soluzione, "soluzione"),
                       (self.t_note, "note")):
            w.delete("1.0", "end")
            w.insert("1.0", r[key])
        self._loading_form = False
        self.att_items = []
        self.att_removed = []
        self.update_att_strip()
        self.form_title.configure(text="Nuovo intervento (copia di quello del %s)"
                                  % fmt_dt(r["data_ora"]))
        self.btn_save.configure(text="Salva la copia")
        self.btn_clear.configure(text="Annulla copia")
        self.update_user_hint()
        self.nb.select(self.tab_new)
        self.set_status("Copia pronta: data e ora sono quelle di adesso e le foto non vengono "
                        "copiate. Cambia ciò che serve e salva.")

    def mark_resolved(self):
        r = self.selected_row()
        if not r:
            messagebox.showinfo(APP_NAME, "Seleziona prima un intervento dall'elenco.")
            return
        if r["stato"] == "Risolto":
            messagebox.showinfo(APP_NAME, "Questo intervento è già segnato come risolto.")
            return
        self.db.set_stato(r["id"], "Risolto")
        self.refresh_stats()
        self.apply_filters()

    def delete_selected(self):
        r = self.selected_row()
        if not r:
            return
        if not messagebox.askyesno(APP_NAME, "Eliminare questo intervento?\n\n%s - %s\n%s" % (
                fmt_dt(r["data_ora"]), r["utente"], one_line(r["richiesta"], 80)), icon="warning"):
            return
        self.db.delete_intervento(r["id"])
        shutil.rmtree(attachment_dir(r["id"]), ignore_errors=True)
        self.refresh_stats()
        self.apply_filters()

    def ask_save_path(self, kind, base_name):
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        if kind == "xlsx":
            return filedialog.asksaveasfilename(
                title="Esporta in Excel", defaultextension=".xlsx",
                initialfile="%s_%s.xlsx" % (base_name, stamp), filetypes=[("File Excel", "*.xlsx")])
        return filedialog.asksaveasfilename(
            title="Esporta in PDF", defaultextension=".pdf",
            initialfile="%s_%s.pdf" % (base_name, stamp), filetypes=[("File PDF", "*.pdf")])

    def finish_export(self, path):
        if messagebox.askyesno(APP_NAME, "File creato:\n%s\n\nVuoi aprirlo ora?" % path):
            try:
                os.startfile(path)  # solo Windows
            except Exception:
                pass

    def export(self, kind):
        if not self.current_rows:
            messagebox.showinfo(APP_NAME, "Non ci sono interventi da esportare con i filtri "
                                          "attuali.")
            return
        path = self.ask_save_path(kind, "interventi")
        if not path:
            return
        desc = self.describe_filters(self.get_filters())
        try:
            if kind == "xlsx":
                export_excel(path, self.current_rows, desc)
            else:
                export_pdf(path, self.current_rows, desc)
        except Exception as e:
            messagebox.showerror(APP_NAME, "Esportazione non riuscita (il file è aperto in un "
                                           "altro programma?):\n%s" % e)
            return
        self.finish_export(path)

    # ---------------- scheda: utenti ----------------
    def _build_users(self):
        u = self.tab_users
        u.columnconfigure(0, weight=1)
        u.rowconfigure(2, weight=1)
        self.var_u_search = tk.StringVar()
        self.var_u_sede = tk.StringVar(value="(Tutte)")

        top = ttk.Frame(u)
        top.grid(row=0, column=0, sticky="ew")
        ttk.Button(top, text="Nuovo utente…", style="Save.TButton",
                   command=self.new_user).pack(side="left")
        ttk.Label(top, text="Gli utenti si aggiungono da soli anche quando salvi un intervento "
                            "con un nome nuovo.", style="Hint.TLabel").pack(side="left",
                                                                           padx=(14, 0))
        sf = ttk.Frame(u)
        sf.grid(row=1, column=0, sticky="ew", pady=(12, 6))
        ttk.Label(sf, text="Cerca").pack(side="left")
        e_usearch = ttk.Entry(sf, textvariable=self.var_u_search, width=30)
        e_usearch.pack(side="left", padx=(8, 0))
        scan_button(sf, self.var_u_search, "any", entry=e_usearch,
                    title="Leggi un codice per cercare l'utente").pack(side="left", padx=(6, 0))
        self.var_u_search.trace_add("write", lambda *_: self.fill_users_tree())
        ttk.Label(sf, text="Sede").pack(side="left", padx=(18, 0))
        self.cb_u_sede = ttk.Combobox(sf, textvariable=self.var_u_sede, width=16,
                                      state="readonly")
        self.cb_u_sede.pack(side="left", padx=(8, 0))
        self.var_u_sede.trace_add("write", lambda *_: self.fill_users_tree())
        ttk.Button(sf, text="Esporta dati utenti…", style="Primary.TButton",
                   command=self.export_users).pack(side="right")

        mid = ttk.Frame(u)
        mid.grid(row=2, column=0, sticky="nsew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)
        cols = ("nome", "ufficio", "matricola", "sede", "telefono", "email", "seriale", "imei",
                "dot", "n")
        self.users_tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("nome", "Utente", 200, True), ("ufficio", "Ufficio", 190, False),
                ("matricola", "Matricola", 90, False),
                ("sede", "Sede", 100, False),
                ("telefono", "Telefono", 110, False), ("email", "Email", 210, False),
                ("seriale", "Seriale PC", 140, False),
                ("imei", "IMEI cellulare", 140, False), ("dot", "Dotazione", 190, False),
                ("n", "Interventi", 70, False)):
            self.users_tree.heading(c, text=title)
            self.users_tree.column(c, width=width, anchor="center" if c == "n" else "w",
                                   stretch=stretch)
        usb = ttk.Scrollbar(mid, orient="vertical", command=self.users_tree.yview)
        self.users_tree.configure(yscrollcommand=usb.set)
        self.users_tree.grid(row=0, column=0, sticky="nsew")
        usb.grid(row=0, column=1, sticky="ns")
        self.users_tree.bind("<Double-1>", lambda e: self.edit_user())

        bf = ttk.Frame(u)
        bf.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Modifica scheda…", style="Primary.TButton",
                   command=self.edit_user).pack(side="left")
        ttk.Button(bf, text="Vedi i suoi interventi",
                   command=self.show_user_interventions).pack(side="left", padx=(8, 0))
        ttk.Button(bf, text="Scrivi mail", command=self.mail_user).pack(side="left", padx=(8, 0))
        ttk.Button(bf, text="Unisci doppie…", command=self.merge_duplicates_dialog).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Elimina", style="Danger.TButton",
                   command=self.delete_user).pack(side="left", padx=(8, 0))
        if VB is not None:
            vb = ttk.Menubutton(bf, text="Crea verbale per questo utente…")
            vmenu = tk.Menu(vb, tearoff=0)
            for tipo in VB.TIPI_ORDINE:
                vmenu.add_command(label=VB.TIPI[tipo][0] + "…",
                                  command=lambda t=tipo: self.new_verbale_for_user(t))
            vb["menu"] = vmenu
            vb.pack(side="left", padx=(24, 0))
            ttk.Button(bf, text="Suoi verbali", command=self.show_user_verbali).pack(
                side="left", padx=(8, 0))

    def refresh_users(self):
        self.users_data = self.db.users_with_counts()
        self.users_cache = [u["nome"] for u in self.users_data]
        self.users_by_key = {key_name(u["nome"]): u for u in self.users_data}
        self.sedi_cache = self.db.sedi()
        self.cb_user.configure(values=self.users_cache)
        self.cb_sede.configure(values=self.sedi_cache)
        self.cb_filter_user.configure(values=["(Tutti)"] + self.users_cache)
        self.cb_filter_sede.configure(values=["(Tutte)"] + self.sedi_cache)
        self.cb_u_sede.configure(values=["(Tutte)"] + self.sedi_cache)
        self.fill_users_tree()
        self.update_user_hint()
        self.refresh_rooms()

    def fill_users_tree(self):
        terms = self.var_u_search.get().casefold().split()
        sede = self.var_u_sede.get().strip()
        sede = "" if sede in ("", "(Tutte)") else sede.casefold()
        self.users_tree.delete(*self.users_tree.get_children())
        self.users_visible = []
        for u in self.users_data:
            if sede and sede != u["sede"].casefold():
                continue
            dot = dotazione_text(u["dotazione"])
            text = " ".join((u["nome"], u["ufficio"], u["matricola"], u["sede"], u["telefono"],
                             u["email"], u["seriale_pc"], u["imei"], dot))
            if terms and not matches_all(terms, text):
                continue
            self.users_visible.append(u)
            self.users_tree.insert("", "end", iid=str(u["id"]), values=(
                u["nome"], u["ufficio"], u["matricola"], u["sede"], u["telefono"], u["email"],
                u["seriale_pc"], u["imei"], dot, u["n"]))

    def describe_user_scope(self):
        parts = []
        if self.var_u_sede.get() not in ("(Tutte)", ""):
            parts.append("sede \"%s\"" % self.var_u_sede.get())
        if self.var_u_search.get().strip():
            parts.append("ricerca \"%s\"" % self.var_u_search.get().strip())
        return "; ".join(parts) if parts else "tutti gli utenti"

    def export_users(self):
        if not self.users_visible:
            messagebox.showinfo(APP_NAME, "Non ci sono utenti da esportare con i filtri attuali.")
            return
        UserExportDialog(self, len(self.users_visible), self.describe_user_scope(),
                         self.do_export_users)

    def do_export_users(self, kind, keys):
        path = self.ask_save_path(kind, "utenti")
        if not path:
            return False
        try:
            if kind == "xlsx":
                export_users_excel(path, self.users_visible, keys, self.describe_user_scope())
            else:
                export_users_pdf(path, self.users_visible, keys, self.describe_user_scope())
        except Exception as e:
            messagebox.showerror(APP_NAME, "Esportazione non riuscita (il file è aperto in un "
                                           "altro programma?):\n%s" % e)
            return False
        self.finish_export(path)
        return True

    def selected_user(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima un utente dall'elenco.")
            return None
        return next((u for u in self.users_data if str(u["id"]) == sel[0]), None)

    def new_user(self):
        UserDialog(self, None, self.sedi_cache, lambda data: self.save_user_card(None, data),
                   uffici=self.db.uffici())

    def edit_user(self):
        user = self.selected_user()
        if user:
            posto = self.db.posto_of_user(user["id"])
            seat = ("%s · Postazione %d  (%s · %s)" % (posto["codice"], posto["numero"],
                                                        posto["edificio"], floor_label(posto["piano"]))
                    if posto else None)
            UserDialog(self, user, self.sedi_cache, lambda data: self.save_user_card(user, data),
                       seat=seat, on_room=(lambda: self.open_room_anywhere(posto["stanza_id"]))
                       if posto else None, uffici=self.db.uffici())

    def save_user_card(self, user, data):
        nome = normalize_name(data["nome"])
        if not nome:
            messagebox.showwarning(APP_NAME, "Scrivi il nome e cognome dell'utente.")
            return False
        fields = (normalize_name(data["sede"]), data["telefono"].strip(),
                  data["seriale_pc"].strip(), data["imei"].strip(), data["matricola"].strip())
        ufficio = " ".join(str(data.get("ufficio") or "").split())
        dotazione = clean_dotazione(data.get("dotazione"))
        email = clean_email(data.get("email"))
        if email and not EMAIL_RE.match(email):
            messagebox.showwarning(APP_NAME, "L'email non sembra valida (manca la @ o il "
                                             "dominio): %s" % email)
            return False
        try:
            if user is None:
                uid, created = self.db.add_user(nome, *fields)
                if not created:
                    messagebox.showinfo(APP_NAME, "%s è già nella lista." % nome)
                    return False
                self.db.set_user_extra(uid, ufficio, dotazione)
                self.db.set_user_email(uid, email)
                self.set_status("Utente aggiunto alla lista: %s." % nome)
            else:
                other = self.db.find_user(nome)
                if other and other["id"] != user["id"]:
                    if not messagebox.askyesno(
                            APP_NAME, "\"%s\" esiste già.\nUnire i due utenti? I %d interventi di "
                                      "\"%s\" passeranno a \"%s\" (restano i dati già salvati di "
                                      "\"%s\"; quelli che gli mancano li prendo da \"%s\")."
                                      % (other["nome"], user["n"], user["nome"], other["nome"],
                                         other["nome"], user["nome"])):
                        return False
                    self.db.merge_users(user["id"], other["id"])
                else:
                    self.db.update_user(user["id"], nome, *fields)
                    self.db.set_user_extra(user["id"], ufficio, dotazione)
                    self.db.set_user_email(user["id"], email)
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return False
        self.refresh_users()
        self.apply_filters()
        return True

    def mail_user(self):
        user = self.selected_user()
        if not user:
            return
        if not user.get("email"):
            messagebox.showinfo(APP_NAME, "\"%s\" non ha un'email: scrivila nella sua scheda."
                                % user["nome"])
            return
        if not mail_to(user["email"]):
            messagebox.showwarning(APP_NAME, "L'email di \"%s\" non sembra valida: %s"
                                   % (user["nome"], user["email"]))

    def show_user_interventions(self):
        user = self.selected_user()
        if not user:
            return
        self.reset_filters()
        self.var_f_user.set(user["nome"])
        self.nb.select(self.tab_reg)

    def merge_duplicates_dialog(self):
        groups = find_duplicate_users(self.db.users_with_counts())
        if not groups:
            messagebox.showinfo(APP_NAME, "Non ho trovato persone doppie.")
            return
        MergeDuplicatesDialog(self, groups, self.do_merge_duplicates)

    def do_merge_duplicates(self, pairs):
        if not messagebox.askyesno(
                APP_NAME, "Unire %d persone?\n\nGli interventi e la postazione di quelle che "
                          "scompaiono passano a quella che resta. Non si può annullare." % len(pairs)):
            return False
        done, interventi, diversi, errors = 0, 0, 0, []
        for src, tgt in pairs:
            try:
                r = self.db.merge_users(src, tgt)
            except (ValueError, sqlite3.Error) as e:
                errors.append(str(e))
                continue
            done += 1
            interventi += r["interventi"]
            diversi += 1 if r["diversi"] else 0
        self.refresh_users()
        self.apply_filters()
        lines = ["Persone unite: %d (interventi spostati: %d)." % (done, interventi)]
        if diversi:
            lines.append("In %d casi le due persone avevano un dato diverso (per esempio due "
                         "telefoni): ho tenuto quello della persona che resta." % diversi)
        if errors:
            lines.append("Non riuscite: %d (%s)" % (len(errors), errors[0]))
        lines.append("\nPer il telefono: File › Esporta dati per il telefono, poi caricalo; il "
                     "telefono toglie da solo le persone unite.")
        messagebox.showinfo(APP_NAME, "\n".join(lines))
        return done > 0

    def delete_user(self):
        user = self.selected_user()
        if not user:
            return
        if user["n"]:
            messagebox.showinfo(APP_NAME, "\"%s\" ha %d interventi, quindi non si può eliminare.\n"
                                "Puoi cambiargli il nome con quello di un altro utente per "
                                "unirli." % (user["nome"], user["n"]))
            return
        if messagebox.askyesno(APP_NAME, "Eliminare l'utente \"%s\"?" % user["nome"]):
            self.db.delete_user(user["id"])
            self.refresh_users()

    # ---------------- scheda: stampanti ----------------
    def _build_printers(self):
        p = self.tab_print
        p.columnconfigure(0, weight=1)
        p.rowconfigure(0, weight=1)
        self.nb_print = ttk.Notebook(p)
        self.nb_print.grid(row=0, column=0, sticky="nsew")
        self.tab_building = ttk.Frame(self.nb_print, padding=12)
        self.tab_toner = ttk.Frame(self.nb_print, padding=12)
        self.tab_bolle = ttk.Frame(self.nb_print, padding=12)
        self.tab_manut = ttk.Frame(self.nb_print, padding=12)
        self.tab_scorta = ttk.Frame(self.nb_print, padding=12)
        self.tab_rooms = ttk.Frame(self.nb_print, padding=12)
        self.nb_print.add(self.tab_building, text="Edificio e stampanti")
        self.nb_print.add(self.tab_toner, text="Cambi toner")
        self.nb_print.add(self.tab_bolle, text="Bolle toner")
        self.nb_print.add(self.tab_manut, text="Interventi tecnico")
        self.nb_print.add(self.tab_scorta, text="Scorta toner")
        self.nb_print.add(self.tab_rooms, text="Piani e stanze")
        self._build_building_view()
        self._build_toner_log()
        self._build_bolle()
        self._build_manut_log()
        self._build_scorta()
        self._build_rooms_view()

    def get_printer_image(self, model, size):
        key = "%s_%s" % (printer_image_key(model), size)
        if key not in self.printer_imgs:
            try:
                self.printer_imgs[key] = tk.PhotoImage(data=PRINTER_IMAGES[key])
            except Exception:
                self.printer_imgs[key] = None
        return self.printer_imgs[key]

    def models_list(self):
        models = list(MODELLI_STAMPANTI)
        for m in self.db.models():
            if m not in models:
                models.append(m)
        return models

    def _build_building_view(self):
        t = self.tab_building
        t.columnconfigure(0, weight=1)
        t.rowconfigure(1, weight=1)
        bar = ttk.Frame(t)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(bar, text="Edificio").pack(side="left")
        self.var_building = tk.StringVar()
        self.cb_building = ttk.Combobox(bar, textvariable=self.var_building, state="readonly",
                                        width=44)
        self.cb_building.pack(side="left", padx=(8, 12))
        self.cb_building.bind("<<ComboboxSelected>>", lambda e: self.on_building_selected())
        ttk.Button(bar, text="Nuovo edificio…", style="Save.TButton",
                   command=self.new_building).pack(side="left")
        ttk.Button(bar, text="Modifica edificio…", command=self.edit_building).pack(
            side="left", padx=(8, 0))
        ttk.Button(bar, text="Elimina edificio", style="Danger.TButton",
                   command=self.delete_building).pack(side="left", padx=(8, 0))
        ttk.Button(bar, text="Unisci doppioni stampanti…", command=self.merge_printers_dialog).pack(
            side="left", padx=(8, 0))
        ttk.Button(bar, text="Edificio in 3D", command=self.open_tower_3d).pack(
            side="left", padx=(8, 0))

        cf = ttk.Frame(t)
        cf.grid(row=1, column=0, sticky="nsew")
        cf.columnconfigure(0, weight=1)
        cf.rowconfigure(0, weight=1)
        self.canvas = tk.Canvas(cf, bg=LIGHT, highlightthickness=1, highlightbackground=BORDER)
        sb = ttk.Scrollbar(cf, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=sb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.canvas.bind("<Configure>", lambda e: self.schedule_draw())
        self.canvas.bind("<MouseWheel>",
                         lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))
        ttk.Label(t, text="Clicca su una stampante per registrare il cambio toner o vaschetta, "
                          "vedere lo storico e gli interventi del tecnico. Con «+ Aggiungi "
                          "stampante» la metti sul piano giusto. Clicca sul nome del piano "
                          "(a sinistra) per aprire la pianta e le stanze.",
                  style="Hint.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))
        leg = ttk.Frame(t)
        leg.grid(row=3, column=0, sticky="w", pady=(4, 0))
        for fill, txt, fg, spiega in (
                (RED, "1", WHITE, "ticket del tecnico Konica ancora aperto (il numero dice quanti)"),
                (BADGE_GIALLO, "!", BADGE_GIALLO_TESTO, "scorta toner in esaurimento (sotto la soglia)")):
            tk.Label(leg, text=txt, bg=fill, fg=fg, width=2, font=("Segoe UI", 9, "bold")).pack(
                side="left")
            ttk.Label(leg, text=spiega, style="Hint.TLabel").pack(side="left", padx=(4, 18))
        ttk.Label(leg, text="Sotto ogni stampante: da quanti giorni è stato cambiato ogni colore "
                            "(gg = giorni, mai = nessun cambio registrato).",
                  style="Hint.TLabel").pack(side="left")
        ttk.Label(t, text="Barretta sotto ogni colore: STIMA di quanto toner resta (dai giorni "
                          "dall'ultimo cambio e da quanto sono durati i cambi precedenti), non il "
                          "livello vero letto dalla stampante. Tratteggiata = servono più cambi per "
                          "stimare; contorno rosso = quasi finito.",
                  style="Hint.TLabel", wraplength=900, justify="left").grid(
            row=4, column=0, sticky="w", pady=(2, 0))

    def current_building(self):
        return self.building_by_label.get(self.var_building.get())

    def refresh_buildings(self, select_id=None):
        keep = select_id
        if keep is None:
            cur = self.current_building()
            keep = cur["id"] if cur else None
        self.buildings = self.db.edifici()
        self.building_by_label = {}
        for b in self.buildings:
            label = b["nome"] + (" (%s)" % b["sede"] if b["sede"] else "")
            self.building_by_label[label] = b
        labels = list(self.building_by_label)
        self.cb_building.configure(values=labels)
        target = next((l for l, b in self.building_by_label.items() if b["id"] == keep),
                      labels[0] if labels else "")
        self.var_building.set(target)
        self.refresh_toner_filter_lists()
        self.schedule_draw()
        self.apply_toner_filters()
        self.apply_manut_filters()
        self.refresh_scorta()
        self.refresh_rooms()

    def on_building_selected(self):
        self.schedule_draw()
        self.refresh_rooms()

    def schedule_draw(self):
        if self._draw_job:
            self.after_cancel(self._draw_job)
        self._draw_job = self.after(120, self.draw_building)

    def draw_building(self):
        self._draw_job = None
        c = self.canvas
        c.delete("all")
        b = self.current_building()
        width = max(c.winfo_width(), 820)
        if not b:
            c.create_text(width / 2, 120, fill=GREY, font=("Segoe UI", 13),
                          text="Nessun edificio. Premi «Nuovo edificio…» per crearne uno.")
            c.configure(scrollregion=(0, 0, width, 260))
            return
        by_floor = {}
        for p in self.db.printers_of(b["id"]):
            by_floor.setdefault(p["piano"], []).append(p)

        x0, x1 = 30, width - 40
        label_w, floor_h, roof_h = 130, 178, 74
        y = 18
        # tetto con il nome dell'edificio
        c.create_polygon(x0 - 12, y + roof_h, x0 + 70, y, x1 - 70, y, x1 + 12, y + roof_h,
                         fill=NAVY, outline=NAVY)
        title = b["nome"] + (" · %s" % b["sede"] if b["sede"] else "")
        c.create_text((x0 + x1) / 2, y + roof_h / 2 + 6, text=title, fill=WHITE,
                      font=("Segoe UI", 14, "bold"))
        y += roof_h

        floors = list(range(b["piano_max"], b["piano_min"] - 1, -1))
        for idx, fl in enumerate(floors):
            top = y
            # prima le stampanti: cosi' so quanto deve essere alto il piano (con le scritte sotto)
            px = x0 + label_w + 26
            lowest = top + floor_h
            for p in by_floor.get(fl, []):
                slot, p_bottom = self._draw_printer(c, p, px, top + 10)
                px += slot + 28
                lowest = max(lowest, p_bottom + 12) if p_bottom else lowest
            bottom = max(top + floor_h, lowest)
            # poi lo sfondo del piano, messo sotto le stampanti già disegnate
            bg = [c.create_rectangle(x0, top, x1, bottom, fill="#efe7cf" if idx % 2 == 0 else "#e7ddc0",
                                     outline="#7d7355", width=2)]
            wx = x0 + label_w + 20
            while wx + 44 < x1 - 10:
                bg.append(c.create_rectangle(wx, top + 14, wx + 44, bottom - 14, fill="#dbe7f0",
                                             outline="#c5d3e0"))
                wx += 96
            ftag = "fl%d" % fl
            bg.append(c.create_rectangle(x0, top, x0 + label_w, bottom, fill=NAVY, outline=NAVY,
                                         tags=(ftag,)))
            for item in reversed(bg):
                c.tag_lower(item)
            c.create_text(x0 + label_w / 2, (top + bottom) / 2, text=floor_label(fl), fill=WHITE,
                          font=("Segoe UI", 12, "bold"), width=label_w - 14, tags=(ftag,))
            c.tag_bind(ftag, "<Button-1>", lambda e, fl=fl: self.floor_menu(fl, e))
            c.tag_bind(ftag, "<Enter>", lambda e: c.configure(cursor="hand2"))
            c.tag_bind(ftag, "<Leave>", lambda e: c.configure(cursor=""))

            tag = "add%d" % fl
            bx1, bx0 = x1 - 14, x1 - 184
            cy = (top + bottom) / 2
            c.create_rectangle(bx0, cy - 17, bx1, cy + 17, fill=GREEN, outline=GREEN, tags=(tag,))
            c.create_text((bx0 + bx1) / 2, cy, text="+ Aggiungi stampante", fill=WHITE,
                          font=("Segoe UI", 9, "bold"), tags=(tag,))
            c.tag_bind(tag, "<Button-1>", lambda e, fl=fl: self.add_printer(fl))
            c.tag_bind(tag, "<Enter>", lambda e: c.configure(cursor="hand2"))
            c.tag_bind(tag, "<Leave>", lambda e: c.configure(cursor=""))
            y = bottom
        c.create_rectangle(x0 - 34, y, x1 + 34, y + 12, fill=GREEN, outline=GREEN)
        c.configure(scrollregion=(0, 0, width, y + 30))

    @staticmethod
    def _draw_toner_bar(c, cx, y, pct, hexcol, tag, w=64, h=5):
        """Barretta della STIMA del toner: piena = tanto, contorno rosso = quasi finito,
        tratteggiata = non ci sono abbastanza cambi per stimare. Restituisce la y del fondo."""
        x0 = cx - w / 2
        basso = pct is not None and pct <= TONER_BASSO
        c.create_rectangle(x0, y, x0 + w, y + h, fill="#e1e6ec", outline=RED if basso else "#e1e6ec",
                           tags=(tag,))
        if pct is None:
            c.create_line(x0 + 2, y + h / 2, x0 + w - 2, y + h / 2, fill="#9aa5b1", dash=(3, 3),
                          tags=(tag,))
        elif pct > 0:
            c.create_rectangle(x0, y, x0 + w * pct / 100.0, y + h, fill=hexcol, outline="",
                               tags=(tag,))
        return y + h

    def _draw_printer(self, c, p, px, py):
        """Disegna una stampante con bollini e scritte sotto; restituisce (larghezza occupata, y del fondo)."""
        tag = "pr%d" % p["id"]
        img = self.get_printer_image(p["modello"], "s")
        iw, ih = (img.width(), img.height()) if img else (70, 96)
        slot = max(iw, 118)
        cx = px + slot / 2
        ix = px + (slot - iw) / 2
        if img:
            c.create_image(ix, py, image=img, anchor="nw", tags=(tag,))
        else:
            c.create_rectangle(ix, py, ix + iw, py + ih, fill=WHITE, outline=NAVY, tags=(tag,))
            c.create_text(cx, py + ih / 2, text="Stampante", fill=NAVY, tags=(tag,))
        fuori = bool(p.get("fuori_servizio"))
        if fuori:
            # croce rossa sopra la foto: la stampante e' rotta
            m = 6
            c.create_line(ix + m, py + m, ix + iw - m, py + ih - m, fill=RED, width=5,
                          capstyle="round", tags=(tag,))
            c.create_line(ix + iw - m, py + m, ix + m, py + ih - m, fill=RED, width=5,
                          capstyle="round", tags=(tag,))

        # bollini: rosso = ticket Konica ancora aperto (col numero), giallo = scorta del modello bassa
        red = p.get("man_richieste") or 0
        low = p.get("scorta_bassa") or []
        badges = []
        if red:
            badges.append((RED, str(red), WHITE))
        if low:
            badges.append((BADGE_GIALLO, "!", BADGE_GIALLO_TESTO))
        r = 11
        bcx, bcy = ix + iw - 2, py + r
        for fill, txt, fg in badges:
            c.create_oval(bcx - r, bcy - r, bcx + r, bcy + r, fill=fill, outline=WHITE, width=2,
                          tags=(tag,))
            c.create_text(bcx, bcy, text=txt, fill=fg, font=("Segoe UI", 9, "bold"), tags=(tag,))
            bcx -= 2 * r + 4

        lines = [(p["nome"] or p["modello"] or "Stampante", ("Segoe UI", 9, "bold"), NAVY)]
        if p["nome"] and p["modello"]:
            lines.append((p["modello"], ("Segoe UI", 8), GREY))
        if (p.get("seriale") or "").strip():
            lines.append(("S/N %s" % p["seriale"].strip(), ("Segoe UI", 8), GREY))
        if fuori:
            lines.append(("FUORI SERVIZIO", ("Segoe UI", 9, "bold"), RED))
        if red:
            lines.append(("tecnico richiesto (%d)" % red, ("Segoe UI", 8, "bold"), RED))
        if low:
            lines.append(("scorta bassa: %s" % ", ".join("%s (%d)" % (col, q) for col, q in low),
                          ("Segoe UI", 8, "bold"), BADGE_GIALLO_TESTO))
        lines.append(("Cambiato:", ("Segoe UI", 8, "bold"), GREY, None))
        ultimi = p.get("ultimi") or {}
        stima = p.get("stima") or {}
        for colore, hexcol, _t in COLORI_TONER:
            lines.append(("%s: %s" % (colore, ago_short(days_ago(ultimi.get(colore)))),
                          ("Segoe UI", 8), TONER_TESTO.get(colore, TEXT),
                          (stima.get(colore, {}).get("pct"), hexcol)))
        y = py + ih + 4
        for text, font, color, *extra in [l if len(l) == 4 else (*l, None) for l in lines]:
            item = c.create_text(cx, y, text=text, anchor="n", width=slot, font=font, fill=color,
                                 justify="center", tags=(tag,))
            box = c.bbox(item)
            y = box[3] + 2 if isinstance(box, tuple) and len(box) == 4 else y + 16
            if extra and extra[0] is not None:
                y = self._draw_toner_bar(c, cx, y, extra[0][0], extra[0][1], tag) + 2
        c.tag_bind(tag, "<Button-1>", lambda e, pid=p["id"]: self.open_printer(pid))
        c.tag_bind(tag, "<Enter>", lambda e: c.configure(cursor="hand2"))
        c.tag_bind(tag, "<Leave>", lambda e: c.configure(cursor=""))
        return slot, y

    # ---------------- scheda: piani e stanze ----------------
    def _build_rooms_view(self):
        t = self.tab_rooms
        t.columnconfigure(0, weight=1)
        t.rowconfigure(1, weight=1)
        self.plan_view = tk.StringVar(value="foto")
        self.plan_mark = tk.BooleanVar(value=False)
        self.var_rfloor = tk.StringVar()
        self._plan_job = None
        self._plan_img = None
        self.rooms_rows, self.rooms_people_map = [], {}
        self.plan_items = {"mode": "schema", "tiles": [], "markers": []}
        bar = ttk.Frame(t)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(bar, text="Edificio").pack(side="left")
        self.cb_rbuilding = ttk.Combobox(bar, textvariable=self.var_building, state="readonly",
                                         width=36)
        self.cb_rbuilding.pack(side="left", padx=(8, 12))
        self.cb_rbuilding.bind("<<ComboboxSelected>>", lambda e: self.on_building_selected())
        ttk.Label(bar, text="Piano").pack(side="left")
        self.cb_rfloor = ttk.Combobox(bar, textvariable=self.var_rfloor, state="readonly", width=14)
        self.cb_rfloor.pack(side="left", padx=(8, 12))
        self.cb_rfloor.bind("<<ComboboxSelected>>", lambda e: self.refresh_rooms())
        ttk.Button(bar, text="+ Nuova stanza…", style="Save.TButton",
                   command=self.new_room).pack(side="left")
        ttk.Button(bar, text="Incolla più stanze…", command=self.bulk_rooms).pack(
            side="left", padx=(8, 0))
        ttk.Button(bar, text="Passeggiata nel piano 3D", command=self.open_walk_3d).pack(
            side="left", padx=(8, 0))
        ttk.Button(bar, text="Edificio in 3D", command=self.open_tower_3d).pack(
            side="left", padx=(8, 0))
        self.lbl_rfloor = ttk.Label(bar, style="Hint.TLabel")
        self.lbl_rfloor.pack(side="left", padx=(16, 0))

        body = ttk.Frame(t)
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=2, uniform="r")
        body.columnconfigure(1, weight=3, uniform="r")
        body.rowconfigure(0, weight=1)
        left = ttk.Frame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        left.rowconfigure(4, weight=1)
        ttk.Label(left, text="Stanze del piano", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, sticky="w")
        self.tree_rooms = ttk.Treeview(left, columns=("stanza", "persone", "occupate", "come"),
                                       show="headings", height=8, selectmode="browse")
        for col, text, w in (("stanza", "Stanza", 80), ("persone", "Persone", 60),
                             ("occupate", "Scrivanie", 90), ("come", "Come arrivarci", 220)):
            self.tree_rooms.heading(col, text=text)
            self.tree_rooms.column(col, width=w, anchor="w")
        self.tree_rooms.grid(row=1, column=0, sticky="nsew")
        self.tree_rooms.bind("<Double-1>", lambda e: self.open_selected_room())
        self.tree_rooms.bind("<<TreeviewSelect>>", lambda e: self.schedule_plan())
        rb = ttk.Frame(left)
        rb.grid(row=2, column=0, sticky="w", pady=(6, 10))
        ttk.Button(rb, text="Apri la stanza…", command=self.open_selected_room).pack(side="left")
        ttk.Button(rb, text="Togli il segno dalla pianta", command=self.unmark_selected).pack(
            side="left", padx=(8, 0))
        ttk.Button(rb, text="Elimina la stanza", style="Danger.TButton",
                   command=self.delete_selected_room).pack(side="left", padx=(8, 0))
        ttk.Label(left, text="Persone del piano", font=("Segoe UI", 10, "bold")).grid(
            row=3, column=0, sticky="w")
        self.tree_rpeople = ttk.Treeview(
            left, columns=("persona", "stanza", "posto", "tel", "pc", "imei"), show="headings",
            height=8, selectmode="browse")
        for col, text, w in (("persona", "Persona", 130), ("stanza", "Stanza", 55),
                             ("posto", "Posto", 45), ("tel", "Telefono", 85),
                             ("pc", "Seriale PC", 100), ("imei", "IMEI", 120)):
            self.tree_rpeople.heading(col, text=text)
            self.tree_rpeople.column(col, width=w, anchor="w")
        self.tree_rpeople.grid(row=4, column=0, sticky="nsew")
        self.tree_rpeople.bind("<Double-1>", lambda e: self.open_person_room())

        right = ttk.Frame(body)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        ctl = ttk.Frame(right)
        ctl.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Radiobutton(ctl, text="Foto della pianta", variable=self.plan_view, value="foto",
                        command=self.schedule_plan).pack(side="left")
        ttk.Radiobutton(ctl, text="Schema", variable=self.plan_view, value="schema",
                        command=self.schedule_plan).pack(side="left", padx=(8, 0))
        ttk.Radiobutton(ctl, text="Scrivanie", variable=self.plan_view, value="scrivanie",
                        command=self.schedule_plan).pack(side="left", padx=(8, 14))
        ttk.Button(ctl, text="Carica la foto della pianta…", command=self.load_plan_photo).pack(
            side="left")
        ttk.Button(ctl, text="Togli la foto", command=self.remove_plan_photo).pack(
            side="left", padx=(8, 14))
        ttk.Checkbutton(ctl, text="Segna le stanze sulla foto", variable=self.plan_mark,
                        command=self.schedule_plan).pack(side="left")
        cf = ttk.Frame(right)
        cf.grid(row=1, column=0, sticky="nsew")
        cf.columnconfigure(0, weight=1)
        cf.rowconfigure(0, weight=1)
        self.pcanvas = tk.Canvas(cf, bg=WHITE, highlightthickness=1, highlightbackground=BORDER)
        psb = ttk.Scrollbar(cf, orient="vertical", command=self.pcanvas.yview)
        self.pcanvas.configure(yscrollcommand=psb.set)
        self.pcanvas.grid(row=0, column=0, sticky="nsew")
        psb.grid(row=0, column=1, sticky="ns")
        self.pcanvas.bind("<Configure>", lambda e: self.schedule_plan())
        self.pcanvas.bind("<Button-1>", self.on_plan_click)
        self.pcanvas.bind("<MouseWheel>",
                          lambda e: self.pcanvas.yview_scroll(int(-e.delta / 120), "units"))
        ttk.Label(t, style="Hint.TLabel", wraplength=1100, justify="left",
                  text="Per segnare le stanze sulla foto: scegli una stanza nell'elenco, attiva "
                       "«Segna le stanze sulla foto» e clicca il punto dove si trova. La scritta "
                       "sulla pianta è «codice · numero di persone» (rosso se le persone sono più "
                       "delle scrivanie). Le stampanti del piano sono sotto la pianta, con i "
                       "bollini.").grid(row=2, column=0, sticky="w", pady=(6, 0))

    def rooms_floor(self):
        b = self.current_building()
        if not b:
            return None
        return next((n for n in range(b["piano_max"], b["piano_min"] - 1, -1)
                     if floor_label(n) == self.var_rfloor.get()), None)

    def refresh_rooms(self, keep_floor=None):
        if not hasattr(self, "tree_rooms"):
            return
        b = self.current_building()
        self.tree_rooms.delete(*self.tree_rooms.get_children())
        self.tree_rpeople.delete(*self.tree_rpeople.get_children())
        self.rooms_rows, self.rooms_people_map = [], {}
        self.cb_rbuilding.configure(values=list(getattr(self, "building_by_label", {})))
        if not b:
            self.cb_rfloor.configure(values=[])
            self.var_rfloor.set("")
            self.lbl_rfloor.configure(text="")
            self.schedule_plan()
            return
        floors = list(range(b["piano_max"], b["piano_min"] - 1, -1))
        self.cb_rfloor.configure(values=[floor_label(n) for n in floors])
        want = keep_floor if keep_floor is not None else self.rooms_floor()
        if want is None or want not in floors:
            want = 0 if 0 in floors else floors[0]
        self.var_rfloor.set(floor_label(want))
        self.rooms_rows = self.db.stanze_of(b["id"], want)
        for r in self.rooms_rows:
            occ, _ = room_occupancy(r)
            self.tree_rooms.insert("", "end", iid=str(r["id"]),
                                   values=(r["codice"], r["n_persone"], occ or "-",
                                           " ".join((r["come"] or "").split())[:80]))
        people = self.db.persone_piano(b["id"], want)
        for p in people:
            self.tree_rpeople.insert("", "end", iid=str(p["utente_id"]),
                                     values=(p["nome"], p["codice"], p["numero"], p["telefono"],
                                             p["seriale_pc"], p["imei"]))
            self.rooms_people_map[str(p["utente_id"])] = p["stanza_id"]
        n_pr = len([x for x in self.db.printers_of(b["id"]) if x["piano"] == want])
        self.lbl_rfloor.configure(text="%d stanze · %d persone · %d stampanti"
                                       % (len(self.rooms_rows), len(people), n_pr))
        self.schedule_plan()

    def selected_room_id(self):
        sel = self.tree_rooms.selection()
        return int(sel[0]) if sel else None

    def need_room(self):
        rid = self.selected_room_id()
        if rid is None:
            messagebox.showinfo(APP_NAME, "Seleziona prima una stanza dall'elenco.")
        return rid

    def new_room(self, floor=None):
        b = self.current_building()
        if not b:
            messagebox.showinfo(APP_NAME, "Crea prima un edificio.")
            return
        fl = floor if floor is not None else self.rooms_floor()
        if fl is None:
            return
        RoomNewDialog(self, "%s · %s" % (b["nome"], floor_label(fl)),
                      lambda cod, scr: self.save_new_room(b["id"], fl, cod, scr))

    def save_new_room(self, building_id, floor, codice, scrivanie):
        txt = (scrivanie or "").strip()
        if txt and not txt.isdigit():
            messagebox.showwarning(APP_NAME, "Le scrivanie totali devono essere un numero (oppure "
                                             "lascia vuoto).")
            return False
        try:
            self.db.add_stanza(building_id, floor, codice, int(txt or 0))
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return False
        self.refresh_rooms(keep_floor=floor)
        self.set_status("Stanza %s creata." % clean_room_code(codice))
        return True

    def bulk_rooms(self):
        b = self.current_building()
        fl = self.rooms_floor()
        if not b or fl is None:
            messagebox.showinfo(APP_NAME, "Crea prima un edificio.")
            return
        PasteDialog(self, "Incolla più stanze",
                    "Una stanza per riga. Per indicare anche le postazioni scrivi per esempio "
                    "T12; 3  oppure  T12 - 3  (o due colonne copiate da Excel). Senza numero "
                    "crei solo la stanza; più codici sulla stessa riga si separano con la virgola. "
                    "Le stanze già presenti non si duplicano: se scrivi un numero diverso di "
                    "postazioni lo aggiorno.\n%s · %s" % (b["nome"], floor_label(fl)),
                    lambda text: self.bulk_rooms_ok(b["id"], fl, text))

    def bulk_rooms_ok(self, building_id, floor, text):
        r = self.db.add_stanze_bulk(building_id, floor, text)
        msg = "Stanze create: %d." % len(r["nuove"])
        con = [c for c in r["nuove"] if c in r["scrivanie"]]
        if con:
            msg += "\nCon le postazioni indicate: %d." % len(con)
        if r["aggiornate"]:
            msg += "\nPostazioni aggiornate in: %s." % ", ".join(r["aggiornate"])
        if r["esistenti"]:
            msg += "\nGià presenti (non duplicate): %s." % ", ".join(r["esistenti"])
        if r["scartate"]:
            msg += "\nScartate perché più lunghe di %d caratteri: %s." % (
                ROOM_CODE_MAX, ", ".join(r["scartate"]))
        self.refresh_rooms(keep_floor=floor)
        messagebox.showinfo(APP_NAME, msg)
        return r

    def open_room(self, room_id):
        if self.db.get_stanza(room_id):
            RoomDialog(self, room_id)

    def on_desk_click(self, kind, room_id, utente_id):
        """Clic su una scrivania della vista «Scrivanie»: chi c'è seduto, il suo PC e il S/N.
        Su una scrivania libera (o sulla scritta della stanza) si apre la stanza."""
        if kind == "occ" and utente_id is not None:
            DeskDialog(self, self, room_id, utente_id, open_room=lambda: self.open_room(room_id))
        else:
            self.open_room(room_id)

    def edit_person_by_id(self, utente_id, parent=None, after=None):
        """La scheda completa di una persona (come in Utenti), aperta da una scrivania."""
        user = next((u for u in self.users_data if u["id"] == utente_id), None)
        if not user:
            return
        posto = self.db.posto_of_user(utente_id)
        seat = ("%s · Postazione %d" % (posto["codice"], posto["numero"])) if posto else None

        def on_save(data):
            if not self.save_user_card(user, data):
                return False
            if after:
                after()
            return True
        UserDialog(parent or self, user, self.sedi_cache, on_save, seat=seat,
                   uffici=self.db.uffici())

    def open_selected_room(self):
        rid = self.need_room()
        if rid is not None:
            self.open_room(rid)

    def open_person_room(self):
        sel = self.tree_rpeople.selection()
        if sel and sel[0] in self.rooms_people_map:
            self.open_room(self.rooms_people_map[sel[0]])

    def open_room_anywhere(self, room_id):
        """Dalla scheda utente: va alla stanza giusta (edificio e piano) e la apre."""
        room = self.db.get_stanza(room_id)
        if not room:
            return
        self.refresh_buildings(select_id=room["edificio_id"])
        self.refresh_rooms(keep_floor=room["piano"])
        self.nb.select(self.tab_print)
        self.nb_print.select(self.tab_rooms)
        RoomDialog(self, room_id)

    def delete_selected_room(self):
        rid = self.need_room()
        if rid is None:
            return
        room = self.db.get_stanza(rid)
        msg = "Eliminare la stanza %s?" % room["codice"]
        if room["n_persone"]:
            msg += "\nLe %d persone assegnate restano tra gli utenti ma senza stanza." % room[
                "n_persone"]
        msg += "\nLe foto della stanza vengono tolte."
        if messagebox.askyesno(APP_NAME, msg, icon="warning"):
            self.db.delete_stanza(rid)
            self.refresh_rooms()

    def floor_menu(self, floor, event):
        """Click sul nome del piano nella vista Edificio: apri la pianta o aggiungi una stanza."""
        m = tk.Menu(self, tearoff=0)
        m.add_command(label="Apri pianta e stanze di questo piano",
                      command=lambda: self.open_floor_plan(floor))
        m.add_command(label="Nuova stanza su questo piano…", command=lambda: self.new_room(floor))
        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()

    def open_floor_plan(self, floor):
        self.refresh_rooms(keep_floor=floor)
        self.nb_print.select(self.tab_rooms)

    # --- la pianta (foto con i segni, oppure schema) ---
    def schedule_plan(self):
        if getattr(self, "_plan_job", None):
            self.after_cancel(self._plan_job)
        self._plan_job = self.after(80, self.draw_plan)

    def plan_photo_path(self, b, floor):
        pi = self.db.pianta_of(b["id"], floor)
        if pi and pi["file"]:
            path = os.path.join(piante_root(), *pi["file"].split("/"))
            if os.path.isfile(path):
                return path
        return ""

    def draw_plan(self):
        self._plan_job = None
        c = self.pcanvas
        c.delete("all")
        self._plan_img = None
        b, fl = self.current_building(), self.rooms_floor()
        width = max(c.winfo_width(), 520)
        if not b or fl is None:
            c.create_text(width / 2, 100, fill=GREY, font=("Segoe UI", 12),
                          text="Nessun edificio.")
            c.configure(scrollregion=(0, 0, width, 200))
            self.plan_items = {"mode": "schema", "tiles": [], "markers": []}
            return
        rooms = self.rooms_rows
        sel = self.selected_room_id()
        marking = bool(self.plan_mark.get())
        info = None
        if self.plan_view.get() == "foto":
            path = self.plan_photo_path(b, fl)
            info = plan_image_b64(path, width - 24) if path else None
        y = 10
        if self.plan_view.get() == "scrivanie":
            people = {r["id"]: self.db.persone_of(r["id"]) for r in rooms}
            lay = desk_layout(rooms, people, width - 4, y0=4)
            self.plan_items = {"mode": "scrivanie", "layout": lay, "tiles": [], "markers": []}
            if lay["rows"]:
                draw_desk_rows(c, lay)
            else:
                c.create_text(12, 14, anchor="nw", fill=GREY, font=("Segoe UI", 10),
                              text="In questo piano non ci sono ancora stanze: "
                                   "premi «+ Nuova stanza…» o «Incolla più stanze…».")
            y = 4 + lay["height"] + 16
            if lay["rows"]:
                c.create_text(12, y, anchor="nw", fill=GREY, font=("Segoe UI", 9), width=width - 30,
                              text="Clicca una scrivania per vedere il numero di serie del PC (S/N), "
                                   "l'IMEI e la dotazione. Una scrivania tratteggiata è libera: "
                                   "cliccala per aprire la stanza e aggiungere una persona.")
                y += 36
        elif info:
            b64, w, h = info
            try:
                self._plan_img = tk.PhotoImage(data=b64)
            except Exception:
                self._plan_img = None
            ox = oy = 10
            c.create_image(ox, oy, image=self._plan_img, anchor="nw")
            markers = plan_markers(rooms, w, h)
            for m in markers:
                x, yy = ox + m["x"], oy + m["y"]
                chosen = marking and m["id"] == sel
                edge = RED if m["over"] else NAVY
                c.create_rectangle(x - m["w"] / 2, yy - m["h"] / 2, x + m["w"] / 2,
                                   yy + m["h"] / 2, fill=NAVY if chosen else WHITE, outline=edge,
                                   width=2)
                c.create_text(x, yy, text=m["label"], fill=WHITE if chosen else edge,
                              font=("Segoe UI", 9, "bold"))
            self.plan_items = {"mode": "foto", "ox": ox, "oy": oy, "w": w, "h": h,
                               "markers": markers, "tiles": []}
            y = oy + h + 14
        else:
            tiles, th = grid_tiles(rooms, width - 24)
            for t in tiles:
                chosen = marking and t["id"] == sel
                edge = RED if t["over"] else NAVY
                c.create_rectangle(t["x0"], t["y0"], t["x1"], t["y1"],
                                   fill="#fde8ea" if t["over"] else "#f7fafd", outline=edge,
                                   width=3 if chosen else 2)
                c.create_text((t["x0"] + t["x1"]) / 2, t["y0"] + 24, text=t["label"], fill=NAVY,
                              font=("Segoe UI", 11, "bold"))
                if t["occ"]:
                    c.create_text((t["x0"] + t["x1"]) / 2, t["y0"] + 46, text=t["occ"],
                                  fill=RED if t["over"] else GREY, font=("Segoe UI", 8, "bold"))
            self.plan_items = {"mode": "schema", "tiles": tiles, "markers": []}
            y = th + 6
            if self.plan_view.get() == "foto":
                c.create_text(12, y, anchor="nw", fill=GREY, font=("Segoe UI", 9), width=width - 30,
                              text="Per questo piano non c'è ancora la foto della pianta: "
                                   "premi «Carica la foto della pianta…». Intanto vedi lo schema.")
                y += 40
        c.create_text(12, y, text="Stampanti del piano", anchor="nw", fill=NAVY,
                      font=("Segoe UI", 11, "bold"))
        y += 26
        printers = [p for p in self.db.printers_of(b["id"]) if p["piano"] == fl]
        bottom, px = y, 12
        if not printers:
            c.create_text(12, y, text="Nessuna stampante su questo piano.", anchor="nw", fill=GREY,
                          font=("Segoe UI", 9))
            bottom = y + 20
        for p in printers:
            slot, pb = self._draw_printer(c, p, px, y)
            px += slot + 20
            bottom = max(bottom, pb)
        c.configure(scrollregion=(0, 0, width, bottom + 20))

    def on_plan_click(self, event):
        c = self.pcanvas
        x, y = c.canvasx(event.x), c.canvasy(event.y)
        it = self.plan_items
        if it["mode"] == "scrivanie":
            hit = desk_hit(it["layout"], x, y)
            if hit:
                self.on_desk_click(*hit)
            return
        if it["mode"] == "foto":
            fx, fy = x - it["ox"], y - it["oy"]
            hit = plan_hit(it["markers"], fx, fy)
            if self.plan_mark.get():
                if hit:
                    self.tree_rooms.selection_set(str(hit))
                    self.schedule_plan()
                    return
                frac = plan_click_fraction(fx, fy, it["w"], it["h"])
                if frac is None:
                    return
                rid = self.selected_room_id()
                if rid is None:
                    self.set_status("Prima scegli una stanza nell'elenco a sinistra.", RED)
                    return
                self.place_marker(rid, frac[0], frac[1])
            elif hit:
                self.open_room(hit)
        else:
            hit = tile_hit(it["tiles"], x, y)
            if hit:
                self.open_room(hit)

    def place_marker(self, room_id, px, py):
        """Segna la stanza sulla pianta; poi la scelta passa alla prossima stanza non ancora segnata."""
        room = self.db.get_stanza(room_id)
        self.db.set_stanza_pos(room_id, px, py)
        self.refresh_rooms(keep_floor=room["piano"])
        nxt = next((r for r in self.rooms_rows if r["px"] is None and r["id"] != room_id), None)
        if nxt:
            self.tree_rooms.selection_set(str(nxt["id"]))
        self.set_status("%s segnata sulla pianta." % room["codice"])

    def unmark_selected(self):
        rid = self.need_room()
        if rid is not None:
            self.db.set_stanza_pos(rid, None, None)
            self.refresh_rooms()

    def start_marking(self, room_id):
        """Dalla scheda della stanza: si va alla pianta già pronti a segnare quella stanza."""
        room = self.db.get_stanza(room_id)
        if not room:
            return
        self.refresh_buildings(select_id=room["edificio_id"])
        self.plan_view.set("foto")
        self.plan_mark.set(True)
        self.refresh_rooms(keep_floor=room["piano"])
        self.tree_rooms.selection_set(str(room_id))
        self.nb.select(self.tab_print)
        self.nb_print.select(self.tab_rooms)
        self.schedule_plan()

    def load_plan_photo(self):
        b, fl = self.current_building(), self.rooms_floor()
        if not b or fl is None:
            messagebox.showinfo(APP_NAME, "Crea prima un edificio.")
            return
        path = filedialog.askopenfilename(
            title="Scegli la foto della pianta di %s" % floor_label(fl),
            filetypes=[("Immagini", "*.jpg *.jpeg *.png *.bmp *.webp"), ("Tutti i file", "*.*")])
        if not path:
            return
        try:
            save_plan_photo(self.db, b["id"], fl, path)
        except ValueError as e:
            messagebox.showwarning(APP_NAME, str(e))
            return
        self.plan_view.set("foto")
        self.refresh_rooms(keep_floor=fl)
        self.set_status("Pianta di %s salvata." % floor_label(fl))

    def remove_plan_photo(self):
        b, fl = self.current_building(), self.rooms_floor()
        if not b or fl is None or not self.db.pianta_of(b["id"], fl):
            messagebox.showinfo(APP_NAME, "Per questo piano non c'è nessuna foto della pianta.")
            return
        if messagebox.askyesno(APP_NAME, "Togliere la foto della pianta di %s?\nI segni delle "
                                         "stanze restano salvati." % floor_label(fl)):
            self.db.delete_pianta(b["id"], fl)
            self.refresh_rooms(keep_floor=fl)

    def open_printer(self, printer_id):
        PrinterDialog(self, printer_id)

    def new_building(self):
        BuildingDialog(self, None, self.sedi_cache, lambda data: self.save_building(None, data))

    def edit_building(self):
        b = self.current_building()
        if not b:
            messagebox.showinfo(APP_NAME, "Non c'è nessun edificio da modificare.")
            return
        BuildingDialog(self, b, self.sedi_cache, lambda data: self.save_building(b, data))

    def save_building(self, building, data):
        nome = " ".join(data["nome"].split())
        sede = normalize_name(data["sede"])
        pmin, pmax = data["piano_min"], data["piano_max"]
        if not nome:
            messagebox.showwarning(APP_NAME, "Scrivi il nome dell'edificio.")
            return False
        if pmin > pmax:
            messagebox.showwarning(APP_NAME, "Il piano più basso non può essere sopra il piano "
                                             "più alto.")
            return False
        try:
            if building is None:
                bid = self.db.add_edificio(nome, sede, pmin, pmax)
            else:
                outside = [p for p in self.db.printers_of(building["id"])
                           if not (pmin <= p["piano"] <= pmax)]
                if outside:
                    messagebox.showwarning(
                        APP_NAME, "Ci sono %d stampanti sui piani che vuoi togliere.\n"
                                  "Spostale prima su un altro piano (Modifica stampante)."
                                  % len(outside))
                    return False
                outside_rooms = [r for r in self.db.stanze_of(building["id"])
                                 if not (pmin <= r["piano"] <= pmax)]
                if outside_rooms:
                    messagebox.showwarning(
                        APP_NAME, "Ci sono %d stanze sui piani che vuoi togliere.\n"
                                  "Eliminale prima da Piani e stanze." % len(outside_rooms))
                    return False
                bid = building["id"]
                self.db.update_edificio(bid, nome, sede, pmin, pmax)
        except sqlite3.IntegrityError:
            messagebox.showwarning(APP_NAME, "Esiste già un edificio con questo nome.")
            return False
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return False
        self.refresh_users()
        self.refresh_buildings(select_id=bid)
        return True

    def delete_building(self):
        b = self.current_building()
        if not b:
            return
        n_s, n_t = self.db.edificio_stats(b["id"])
        msg = "Eliminare l'edificio \"%s\"?" % b["nome"]
        if n_s:
            msg += "\nVerranno eliminate anche %d stampanti e i %d cambi toner registrati." % (
                n_s, n_t)
        if messagebox.askyesno(APP_NAME, msg, icon="warning"):
            self.db.delete_edificio(b["id"])
            self.refresh_buildings(select_id=-1)

    def add_printer(self, floor):
        b = self.current_building()
        if not b:
            messagebox.showinfo(APP_NAME, "Crea prima un edificio.")
            return
        PrinterEditDialog(self, None, b, floor, self.models_list(),
                          lambda data: self.save_printer(None, b["id"], data),
                          rooms_fn=lambda piano: self.db.stanze_of(b["id"], piano))

    def save_printer(self, printer_id, building_id, data):
        nome = " ".join(data["nome"].split())
        modello = " ".join(data["modello"].split())
        if not nome and not modello:
            messagebox.showwarning(APP_NAME, "Scrivi almeno il nome o il modello della stampante.")
            return False
        seriale = clean_serial(data.get("seriale"))
        other = self.db.printer_with_serial(seriale, printer_id)
        if other and not messagebox.askyesno(
                APP_NAME, "Il numero di serie %s e' gia' della stampante \"%s\" (%s, %s).\n\n"
                          "Un numero di serie dovrebbe essere diverso per ogni stampante: "
                          "potrebbe essere un errore di lettura. Vuoi salvare lo stesso?" % (
                              seriale, other["nome"] or other["modello"] or "senza nome",
                              other["edificio"], floor_label(other["piano"]))):
            return False
        args = (data["piano"], nome, modello, data["ip"].strip(), data["note"].strip())
        # senza la voce nei dati, lo stato "fuori servizio" resta com'era
        fuori = bool(data["fuori_servizio"]) if "fuori_servizio" in data else None
        stanza = data["stanza_uid"] if "stanza_uid" in data else None   # senza la voce resta com'era
        try:
            if printer_id is None:
                self.db.add_stampante(building_id, *args, seriale=seriale,
                                      fuori_servizio=bool(fuori), stanza_uid=stanza or "")
            else:
                self.db.update_stampante(printer_id, *args, seriale=seriale, fuori_servizio=fuori,
                                         stanza_uid=stanza)
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return False
        self.after_printers_change()
        return True

    def after_printers_change(self):
        self.refresh_toner_filter_lists()
        self.schedule_draw()
        self.apply_toner_filters()
        self.apply_manut_filters()
        self.refresh_scorta()

    def after_toner_change(self):
        self.schedule_draw()
        self.apply_toner_filters()
        self.refresh_scorta()

    def after_manut_change(self):
        self.schedule_draw()
        self.apply_manut_filters()

    # ---------------- scheda: cambi toner ----------------
    def _build_toner_log(self):
        t = self.tab_toner
        t.columnconfigure(0, weight=1)
        t.rowconfigure(2, weight=1)
        self.var_t_ed = tk.StringVar(value="(Tutti)")
        self.var_t_pr = tk.StringVar(value="(Tutte)")
        self.var_t_col = tk.StringVar(value="(Tutti)")
        self.var_t_dal = tk.StringVar()
        self.var_t_al = tk.StringVar()

        box = ttk.LabelFrame(t, text="Filtri", padding=10)
        box.grid(row=0, column=0, sticky="ew")
        r1 = ttk.Frame(box)
        r1.pack(fill="x")
        ttk.Label(r1, text="Edificio").pack(side="left")
        self.cb_t_ed = ttk.Combobox(r1, textvariable=self.var_t_ed, state="readonly", width=30)
        self.cb_t_ed.pack(side="left", padx=(6, 14))
        ttk.Label(r1, text="Stampante").pack(side="left")
        self.cb_t_pr = ttk.Combobox(r1, textvariable=self.var_t_pr, state="readonly", width=44)
        self.cb_t_pr.pack(side="left", padx=(6, 14))
        ttk.Label(r1, text="Colore").pack(side="left")
        ttk.Combobox(r1, textvariable=self.var_t_col, state="readonly", width=12,
                     values=["(Tutti)"] + TONER_NOMI).pack(side="left", padx=(6, 0))

        r2 = ttk.Frame(box)
        r2.pack(fill="x", pady=(10, 0))
        ttk.Label(r2, text="Dal").pack(side="left")
        self.ent_t_dal = ttk.Entry(r2, textvariable=self.var_t_dal, width=11)
        self.ent_t_dal.pack(side="left", padx=(6, 12))
        ttk.Label(r2, text="Al").pack(side="left")
        self.ent_t_al = ttk.Entry(r2, textvariable=self.var_t_al, width=11)
        self.ent_t_al.pack(side="left", padx=(6, 18))
        ttk.Label(r2, text="Periodo veloce:", style="Hint.TLabel").pack(side="left")
        for label, kind in (("Oggi", "oggi"), ("7 giorni", "7g"), ("Questo mese", "mese"),
                            ("Quest'anno", "anno"), ("Tutto", "tutto")):
            ttk.Button(r2, text=label, command=lambda k=kind: self.set_toner_period(k)).pack(
                side="left", padx=(6, 0))
        ttk.Button(r2, text="Azzera filtri", style="Primary.TButton",
                   command=self.reset_toner_filters).pack(side="right")
        for v in (self.var_t_ed, self.var_t_pr, self.var_t_col, self.var_t_dal, self.var_t_al):
            v.trace_add("write", self.schedule_toner_filter)

        cnt = ttk.Frame(t)
        cnt.grid(row=1, column=0, sticky="ew", pady=(10, 6))
        self.lbl_t_count = ttk.Label(cnt, text="", style="Count.TLabel")
        self.lbl_t_count.pack(side="left")
        self.lbl_t_colors = ttk.Label(cnt, text="", style="Hint.TLabel")
        self.lbl_t_colors.pack(side="left", padx=(14, 0))

        tf = ttk.Frame(t)
        tf.grid(row=2, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)
        cols = ("data", "sede", "edificio", "piano", "stampante", "modello", "colore", "note")
        self.toner_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("data", "Data e ora", 130, False), ("sede", "Sede", 90, False),
                ("edificio", "Edificio", 170, False), ("piano", "Piano", 100, False),
                ("stampante", "Stampante", 170, False), ("modello", "Modello", 110, False),
                ("colore", "Colore", 80, False), ("note", "Note", 240, True)):
            self.toner_tree.heading(c, text=title)
            self.toner_tree.column(c, width=width, anchor="w", stretch=stretch)
        for nome in TONER_NOMI:
            self.toner_tree.tag_configure(nome, foreground=TONER_TESTO[nome])
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.toner_tree.yview)
        self.toner_tree.configure(yscrollcommand=sb.set)
        self.toner_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.toner_tree.bind("<Delete>", lambda e: self.delete_toner_selected())

        bf = ttk.Frame(t)
        bf.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Elimina il cambio selezionato", style="Danger.TButton",
                   command=self.delete_toner_selected).pack(side="left")
        ttk.Button(bf, text="Esporta in PDF", style="Primary.TButton",
                   command=lambda: self.export_toner("pdf")).pack(side="right")
        ttk.Button(bf, text="Esporta in Excel", style="Primary.TButton",
                   command=lambda: self.export_toner("xlsx")).pack(side="right", padx=(0, 8))

    def refresh_toner_filter_lists(self):
        eds = ["(Tutti)"] + [b["nome"] for b in self.buildings]
        self.cb_t_ed.configure(values=eds)
        labels = ["(Tutte)"]
        self.toner_printer_labels = {}
        for p in self.db.all_printers():
            label = "%s · %s · %s" % (p["edificio"], floor_label(p["piano"]),
                                      p["nome"] or p["modello"] or "Stampante")
            if label in self.toner_printer_labels:
                label += " #%d" % p["id"]
            self.toner_printer_labels[label] = p["id"]
            labels.append(label)
        self.cb_t_pr.configure(values=labels)
        if self.var_t_ed.get() not in eds:
            self.var_t_ed.set("(Tutti)")
        if self.var_t_pr.get() not in labels:
            self.var_t_pr.set("(Tutte)")
        self.refresh_manut_filter_lists(eds, labels)

    def set_toner_period(self, kind):
        dal, al = period_dates(kind)
        self.var_t_dal.set(fmt_date(dal) if dal else "")
        self.var_t_al.set(fmt_date(al) if al else "")

    def reset_toner_filters(self):
        self.var_t_ed.set("(Tutti)")
        self.var_t_pr.set("(Tutte)")
        self.var_t_col.set("(Tutti)")
        self.var_t_dal.set("")
        self.var_t_al.set("")

    def schedule_toner_filter(self, *_):
        if self._toner_job:
            self.after_cancel(self._toner_job)
        self._toner_job = self.after(250, self.apply_toner_filters)

    def get_toner_filters(self):
        ed_name = self.var_t_ed.get()
        edificio_id = next((b["id"] for b in self.buildings if b["nome"] == ed_name), None)
        stampante_id = self.toner_printer_labels.get(self.var_t_pr.get())
        colore = self.var_t_col.get()
        dal_txt, al_txt = self.var_t_dal.get().strip(), self.var_t_al.get().strip()
        dal = parse_date(dal_txt) if dal_txt else None
        al = parse_date(al_txt) if al_txt else None
        self.ent_t_dal.configure(foreground=RED if dal_txt and not dal else TEXT)
        self.ent_t_al.configure(foreground=RED if al_txt and not al else TEXT)
        return {"edificio_id": edificio_id, "stampante_id": stampante_id,
                "colore": colore if colore in TONER_NOMI else "", "dal": dal, "al": al}

    def describe_toner_filters(self):
        parts = []
        if self.var_t_ed.get() != "(Tutti)":
            parts.append("edificio \"%s\"" % self.var_t_ed.get())
        if self.var_t_pr.get() != "(Tutte)":
            parts.append("stampante \"%s\"" % self.var_t_pr.get())
        if self.var_t_col.get() in TONER_NOMI:
            parts.append("colore %s" % self.var_t_col.get())
        dal = parse_date(self.var_t_dal.get()) if self.var_t_dal.get().strip() else None
        al = parse_date(self.var_t_al.get()) if self.var_t_al.get().strip() else None
        if dal or al:
            parts.append("periodo %s - %s" % (fmt_date(dal) if dal else "inizio",
                                              fmt_date(al) if al else "oggi"))
        return "; ".join(parts) if parts else "nessun filtro (tutti i cambi toner)"

    def apply_toner_filters(self):
        self._toner_job = None
        self.toner_rows = self.db.search_toner(**self.get_toner_filters())
        self.toner_tree.delete(*self.toner_tree.get_children())
        self.toner_rows_by_id = {}
        for r in self.toner_rows:
            self.toner_rows_by_id[str(r["id"])] = r
            self.toner_tree.insert("", "end", iid=str(r["id"]), tags=(r["colore"],), values=(
                fmt_dt(r["data_ora"]), r["sede"], r["edificio"], floor_label(r["piano"]),
                r["stampante"], r["modello"], r["colore"], r["note"]))
        counts = Counter(r["colore"] for r in self.toner_rows)
        self.lbl_t_count.configure(text="Cambi toner trovati: %d" % len(self.toner_rows))
        self.lbl_t_colors.configure(text="   ".join(
            "%s %d" % (n, counts[n]) for n in TONER_NOMI if counts[n]))

    def delete_toner_selected(self):
        sel = self.toner_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima un cambio toner dall'elenco.")
            return
        r = self.toner_rows_by_id.get(sel[0])
        if r and messagebox.askyesno(APP_NAME, "Eliminare questo cambio toner?\n\n%s - %s - %s" % (
                fmt_dt(r["data_ora"]), r["stampante"] or r["modello"], r["colore"])):
            self.db.delete_toner(r["id"])
            self.after_toner_change()

    def export_toner(self, kind):
        if not self.toner_rows:
            messagebox.showinfo(APP_NAME, "Non ci sono cambi toner da esportare con i filtri "
                                          "attuali.")
            return
        path = self.ask_save_path(kind, "cambi_toner")
        if not path:
            return
        desc = self.describe_toner_filters()
        try:
            if kind == "xlsx":
                export_toner_excel(path, self.toner_rows, desc)
            else:
                export_toner_pdf(path, self.toner_rows, desc)
        except Exception as e:
            messagebox.showerror(APP_NAME, "Esportazione non riuscita (il file è aperto in un "
                                           "altro programma?):\n%s" % e)
            return
        self.finish_export(path)

    # ---------------- scheda: interventi del tecnico Konica ----------------
    def _build_manut_log(self):
        t = self.tab_manut
        t.columnconfigure(0, weight=1)
        t.rowconfigure(2, weight=1)
        self.var_m_ed = tk.StringVar(value="(Tutti)")
        self.var_m_pr = tk.StringVar(value="(Tutte)")
        self.var_m_stato = tk.StringVar(value="(Tutti)")
        self.var_m_dal = tk.StringVar()
        self.var_m_al = tk.StringVar()
        self.var_m_text = tk.StringVar()

        box = ttk.LabelFrame(t, text="Filtri", padding=10)
        box.grid(row=0, column=0, sticky="ew")
        r1 = ttk.Frame(box)
        r1.pack(fill="x")
        ttk.Label(r1, text="Edificio").pack(side="left")
        self.cb_m_ed = ttk.Combobox(r1, textvariable=self.var_m_ed, state="readonly", width=26)
        self.cb_m_ed.pack(side="left", padx=(6, 14))
        ttk.Label(r1, text="Stampante").pack(side="left")
        self.cb_m_pr = ttk.Combobox(r1, textvariable=self.var_m_pr, state="readonly", width=40)
        self.cb_m_pr.pack(side="left", padx=(6, 14))
        ttk.Label(r1, text="Stato").pack(side="left")
        ttk.Combobox(r1, textvariable=self.var_m_stato, state="readonly", width=11,
                     values=["(Tutti)"] + STATI_TECNICO).pack(side="left", padx=(6, 14))
        ttk.Label(r1, text="Cerca").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_m_text, width=18).pack(side="left", padx=(6, 0),
                                                                  fill="x", expand=True)

        r2 = ttk.Frame(box)
        r2.pack(fill="x", pady=(10, 0))
        ttk.Label(r2, text="Dal").pack(side="left")
        self.ent_m_dal = ttk.Entry(r2, textvariable=self.var_m_dal, width=11)
        self.ent_m_dal.pack(side="left", padx=(6, 12))
        ttk.Label(r2, text="Al").pack(side="left")
        self.ent_m_al = ttk.Entry(r2, textvariable=self.var_m_al, width=11)
        self.ent_m_al.pack(side="left", padx=(6, 18))
        ttk.Label(r2, text="Periodo veloce:", style="Hint.TLabel").pack(side="left")
        for label, kind in (("Oggi", "oggi"), ("7 giorni", "7g"), ("Questo mese", "mese"),
                            ("Quest'anno", "anno"), ("Tutto", "tutto")):
            ttk.Button(r2, text=label, command=lambda k=kind: self.set_manut_period(k)).pack(
                side="left", padx=(6, 0))
        ttk.Button(r2, text="Azzera filtri", style="Primary.TButton",
                   command=self.reset_manut_filters).pack(side="right")
        for v in (self.var_m_ed, self.var_m_pr, self.var_m_stato, self.var_m_dal, self.var_m_al,
                  self.var_m_text):
            v.trace_add("write", self.schedule_manut_filter)

        cnt = ttk.Frame(t)
        cnt.grid(row=1, column=0, sticky="ew", pady=(10, 6))
        self.lbl_m_count = ttk.Label(cnt, text="", style="Count.TLabel")
        self.lbl_m_count.pack(side="left")
        self.lbl_m_hint = ttk.Label(cnt, text="", style="Hint.TLabel")
        self.lbl_m_hint.pack(side="left", padx=(14, 0))

        tf = ttk.Frame(t)
        tf.grid(row=2, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)
        cols = ("data", "edificio", "piano", "stampante", "ticket", "stato", "data_tec", "tecnico",
                "problema", "intervento")
        self.manut_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("data", "Richiesta del", 120, False), ("edificio", "Edificio", 140, False),
                ("piano", "Piano", 85, False), ("stampante", "Stampante", 130, False),
                ("ticket", "Ticket", 70, False), ("stato", "Stato", 75, False),
                ("data_tec", "Tecnico venuto", 95, False), ("tecnico", "Tecnico", 100, False),
                ("problema", "Problema segnalato", 180, True),
                ("intervento", "Intervento svolto", 180, True)):
            self.manut_tree.heading(c, text=title)
            self.manut_tree.column(c, width=width, anchor="w", stretch=stretch)
        self.manut_tree.tag_configure("Richiesto", foreground=RED)
        self.manut_tree.tag_configure("Eseguito", foreground=GREEN_DARK)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.manut_tree.yview)
        self.manut_tree.configure(yscrollcommand=sb.set)
        self.manut_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.manut_tree.bind("<Double-1>", lambda e: self.edit_manutenzione())
        self.manut_tree.bind("<Delete>", lambda e: self.delete_manut_selected())

        bf = ttk.Frame(t)
        bf.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Nuovo intervento del tecnico…", style="Save.TButton",
                   command=self.new_manutenzione).pack(side="left")
        ttk.Button(bf, text="Modifica…", command=self.edit_manutenzione).pack(side="left",
                                                                            padx=(8, 0))
        ttk.Button(bf, text="Segna come eseguito", command=self.mark_manut_done).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Elimina", style="Danger.TButton",
                   command=self.delete_manut_selected).pack(side="left", padx=(8, 0))
        ttk.Button(bf, text="Esporta in PDF", style="Primary.TButton",
                   command=lambda: self.export_manut("pdf")).pack(side="right")
        ttk.Button(bf, text="Esporta in Excel", style="Primary.TButton",
                   command=lambda: self.export_manut("xlsx")).pack(side="right", padx=(0, 8))

    def refresh_manut_filter_lists(self, eds, labels):
        self.cb_m_ed.configure(values=eds)
        self.cb_m_pr.configure(values=labels)
        if self.var_m_ed.get() not in eds:
            self.var_m_ed.set("(Tutti)")
        if self.var_m_pr.get() not in labels:
            self.var_m_pr.set("(Tutte)")

    def set_manut_period(self, kind):
        dal, al = period_dates(kind)
        self.var_m_dal.set(fmt_date(dal) if dal else "")
        self.var_m_al.set(fmt_date(al) if al else "")

    def reset_manut_filters(self):
        self.var_m_ed.set("(Tutti)")
        self.var_m_pr.set("(Tutte)")
        self.var_m_stato.set("(Tutti)")
        self.var_m_dal.set("")
        self.var_m_al.set("")
        self.var_m_text.set("")

    def schedule_manut_filter(self, *_):
        if self._manut_job:
            self.after_cancel(self._manut_job)
        self._manut_job = self.after(250, self.apply_manut_filters)

    def get_manut_filters(self):
        ed_name = self.var_m_ed.get()
        edificio_id = next((b["id"] for b in self.buildings if b["nome"] == ed_name), None)
        stampante_id = self.toner_printer_labels.get(self.var_m_pr.get())
        st = self.var_m_stato.get()
        dal_txt, al_txt = self.var_m_dal.get().strip(), self.var_m_al.get().strip()
        dal = parse_date(dal_txt) if dal_txt else None
        al = parse_date(al_txt) if al_txt else None
        self.ent_m_dal.configure(foreground=RED if dal_txt and not dal else TEXT)
        self.ent_m_al.configure(foreground=RED if al_txt and not al else TEXT)
        return {"edificio_id": edificio_id, "stampante_id": stampante_id,
                "stato": st if st in STATI_TECNICO else "", "dal": dal, "al": al,
                "testo": self.var_m_text.get().strip()}

    def describe_manut_filters(self):
        parts = []
        if self.var_m_ed.get() != "(Tutti)":
            parts.append("edificio \"%s\"" % self.var_m_ed.get())
        if self.var_m_pr.get() != "(Tutte)":
            parts.append("stampante \"%s\"" % self.var_m_pr.get())
        if self.var_m_stato.get() in STATI_TECNICO:
            parts.append("stato %s" % self.var_m_stato.get())
        dal = parse_date(self.var_m_dal.get()) if self.var_m_dal.get().strip() else None
        al = parse_date(self.var_m_al.get()) if self.var_m_al.get().strip() else None
        if dal or al:
            parts.append("periodo %s - %s" % (fmt_date(dal) if dal else "inizio",
                                              fmt_date(al) if al else "oggi"))
        if self.var_m_text.get().strip():
            parts.append("testo \"%s\"" % self.var_m_text.get().strip())
        return "; ".join(parts) if parts else "nessun filtro (tutti gli interventi del tecnico)"

    def apply_manut_filters(self):
        self._manut_job = None
        self.manut_rows = self.db.manutenzioni_search(**self.get_manut_filters())
        self.manut_tree.delete(*self.manut_tree.get_children())
        self.manut_rows_by_id = {}
        for m in self.manut_rows:
            self.manut_rows_by_id[str(m["id"])] = m
            self.manut_tree.insert("", "end", iid=str(m["id"]), tags=(m["stato"],), values=(
                fmt_dt(m["data_ora"]), m["edificio"], floor_label(m["piano"]),
                m["stampante"] or m["modello"], m["ticket"], m["stato"], tec_date(m),
                m["tecnico"], one_line(m["problema"], 70), one_line(m["intervento"], 70)))
        richieste = sum(1 for m in self.manut_rows if m["stato"] == "Richiesto")
        self.lbl_m_count.configure(text="Interventi trovati: %d" % len(self.manut_rows))
        self.lbl_m_hint.configure(
            text=("Da eseguire: %d" % richieste) if richieste else "")

    def selected_manut(self):
        sel = self.manut_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima un intervento dall'elenco.")
            return None
        return self.manut_rows_by_id.get(sel[0])

    def open_manut_editor(self, printer_id=None, manut=None, after=None, lock=False, parent=None):
        current = dict(manut) if manut else {"stampante_id": printer_id}

        def on_save(data):
            ok = self.save_manutenzione(manut["id"] if manut else None, data)
            if ok and after:
                after()
            return ok

        ManutenzioneDialog(parent or self, self.toner_printer_labels, current, on_save,
                           lock_printer=lock)

    def new_manutenzione(self):
        pid = self.toner_printer_labels.get(self.var_m_pr.get())
        self.open_manut_editor(pid)

    def edit_manutenzione(self):
        m = self.selected_manut()
        if m:
            self.open_manut_editor(m["stampante_id"], m)

    def save_manutenzione(self, mid, data):
        args = (data["stampante_id"], data["data_ora"], data["data_tecnico"],
                data["ticket"].strip(), data["tecnico"].strip(), data["problema"],
                data["intervento"], data["stato"])
        try:
            if mid is None:
                self.db.add_manutenzione(*args)
            else:
                self.db.update_manutenzione(mid, *args)
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return False
        self.after_manut_change()
        return True

    def set_manut_done(self, m):
        data_t = m["data_tecnico"] or date.today().isoformat()
        self.db.set_stato_manutenzione(m["id"], "Eseguito", data_t)
        self.after_manut_change()

    def mark_manut_done(self):
        m = self.selected_manut()
        if not m:
            return
        if m["stato"] == "Eseguito":
            messagebox.showinfo(APP_NAME, "Questo intervento è già segnato come eseguito.")
            return
        self.set_manut_done(m)

    def delete_manut_selected(self):
        m = self.selected_manut()
        if m and messagebox.askyesno(
                APP_NAME, "Eliminare questo intervento del tecnico?\n\n%s - %s\n%s" % (
                    fmt_dt(m["data_ora"]), m["stampante"] or m["modello"],
                    one_line(m["problema"], 70))):
            self.db.delete_manutenzione(m["id"])
            self.after_manut_change()

    def export_manut(self, kind):
        if not self.manut_rows:
            messagebox.showinfo(APP_NAME, "Non ci sono interventi del tecnico da esportare con "
                                          "i filtri attuali.")
            return
        path = self.ask_save_path(kind, "interventi_tecnico")
        if not path:
            return
        desc = self.describe_manut_filters()
        try:
            if kind == "xlsx":
                export_manut_excel(path, self.manut_rows, desc)
            else:
                export_manut_pdf(path, self.manut_rows, desc)
        except Exception as e:
            messagebox.showerror(APP_NAME, "Esportazione non riuscita (il file è aperto in un "
                                           "altro programma?):\n%s" % e)
            return
        self.finish_export(path)

    # ---------------- scheda: scorta toner ----------------
    def _build_scorta(self):
        t = self.tab_scorta
        t.columnconfigure(0, weight=1)
        t.rowconfigure(2, weight=1)
        self.lbl_scorta_alert = ttk.Label(t, text="", font=("Segoe UI", 12, "bold"),
                                          foreground=GREEN, wraplength=1050, justify="left")
        self.lbl_scorta_alert.grid(row=0, column=0, sticky="w")
        ttk.Label(t, style="Hint.TLabel", wraplength=1050, justify="left",
                  text="Scrivi quanti toner (e vaschette) hai in magazzino per ogni modello: a ogni "
                       "cambio registrato ne tolgo uno da solo e ti avviso quando stanno per finire. "
                       "Quelli che non imposti non vengono contati.").grid(
            row=1, column=0, sticky="w", pady=(2, 8))
        tf = ttk.Frame(t)
        tf.grid(row=2, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)
        cols = ("modello", "colore", "quantita", "soglia", "stato")
        self.scorta_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("modello", "Modello stampante", 220, True), ("colore", "Toner / vaschetta", 140, False),
                ("quantita", "In magazzino", 110, False), ("soglia", "Avviso a", 90, False),
                ("stato", "Stato", 140, False)):
            self.scorta_tree.heading(c, text=title)
            self.scorta_tree.column(c, width=width, anchor="center" if c in (
                "quantita", "soglia") else "w", stretch=stretch)
        self.scorta_tree.tag_configure("Finito", foreground=RED)
        self.scorta_tree.tag_configure("In esaurimento", foreground="#b36b00")
        self.scorta_tree.tag_configure("OK", foreground=GREEN_DARK)
        self.scorta_tree.tag_configure("Non impostata", foreground=GREY)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.scorta_tree.yview)
        self.scorta_tree.configure(yscrollcommand=sb.set)
        self.scorta_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.scorta_tree.bind("<Double-1>", lambda e: self.scorta_edit())
        bf = ttk.Frame(t)
        bf.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Ho ricevuto toner…", style="Save.TButton",
                   command=self.scorta_receive).pack(side="left")
        ttk.Button(bf, text="Modifica quantità e avviso…", command=self.scorta_edit).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Non contare più questo toner", command=self.scorta_untrack).pack(
            side="left", padx=(8, 0))

    # ---------------- scheda: bolle dei toner ----------------
    BOLLE_FILTRI = (("Da montare", "da_montare"), ("Montati", "montato"), ("Tutte", None))

    def _build_bolle(self):
        t = self.tab_bolle
        t.columnconfigure(0, weight=1)
        t.rowconfigure(4, weight=1)
        self.lbl_bolle = ttk.Label(t, text="", font=("Segoe UI", 12, "bold"), foreground=GREEN,
                                   wraplength=1050, justify="left")
        self.lbl_bolle.grid(row=0, column=0, sticky="w")
        ttk.Label(t, style="Hint.TLabel", wraplength=1050, justify="left",
                  text="Le bolle dei toner le registri con il telefono: quando il toner arriva (entra "
                       "nell'inventario del piano e nella scorta) e quando lo monti (inquadrando il seriale "
                       "della stampante). Qui le vedi dopo File › Importa dati dal telefono, con la foto della "
                       "bolla.").grid(row=1, column=0, sticky="w", pady=(2, 8))
        ff = ttk.LabelFrame(t, text="Toner nei piani (arrivati, ancora da montare)", padding=8)
        ff.grid(row=2, column=0, sticky="ew")
        ff.columnconfigure(0, weight=1)
        cols = ("edificio", "piano") + tuple(TONER_NOMI) + ("totale",)
        self.bolle_floor_tree = ttk.Treeview(ff, columns=cols, show="headings", height=5,
                                             selectmode="browse")
        for c in cols:
            self.bolle_floor_tree.heading(c, text=("Edificio" if c == "edificio" else "Piano" if c == "piano"
                                                   else "Totale" if c == "totale" else c))
            self.bolle_floor_tree.column(c, width=200 if c == "edificio" else 120 if c == "piano" else 90,
                                         anchor="w" if c in ("edificio", "piano") else "center",
                                         stretch=c == "edificio")
        self.bolle_floor_tree.grid(row=0, column=0, sticky="ew")
        bar = ttk.Frame(t)
        bar.grid(row=3, column=0, sticky="ew", pady=(10, 4))
        ttk.Label(bar, text="Mostra:").pack(side="left")
        self.var_bolle_stato = tk.StringVar(value="Da montare")
        cb = ttk.Combobox(bar, textvariable=self.var_bolle_stato, state="readonly", width=14,
                          values=[n for n, _ in self.BOLLE_FILTRI])
        cb.pack(side="left", padx=(6, 0))
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_bolle())
        ttk.Button(bar, text="Vedi la foto della bolla", command=self.bolla_photo).pack(
            side="left", padx=(16, 0))
        tf = ttk.Frame(t)
        tf.grid(row=4, column=0, sticky="nsew")
        tf.columnconfigure(0, weight=1)
        tf.rowconfigure(0, weight=1)
        cols = ("stato", "colore", "edificio", "piano", "bolla", "arrivo", "montaggio", "stampante")
        self.bolle_tree = ttk.Treeview(tf, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("stato", "Stato", 100, False), ("colore", "Toner", 100, False),
                ("edificio", "Edificio", 170, True), ("piano", "Piano", 100, False),
                ("bolla", "Bolla", 190, True), ("arrivo", "Arrivato il", 130, False),
                ("montaggio", "Montato il", 130, False), ("stampante", "Montato su", 200, True)):
            self.bolle_tree.heading(c, text=title)
            self.bolle_tree.column(c, width=width, anchor="w", stretch=stretch)
        self.bolle_tree.tag_configure("da_montare", foreground="#b36b00")
        self.bolle_tree.tag_configure("montato", foreground=GREEN_DARK)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.bolle_tree.yview)
        self.bolle_tree.configure(yscrollcommand=sb.set)
        self.bolle_tree.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")
        self.bolle_tree.bind("<Double-1>", lambda e: self.bolla_photo())
        self.bolle_by_iid = {}

    def refresh_bolle(self):
        inv = self.db.bolle_inventario()
        self.bolle_floor_tree.delete(*self.bolle_floor_tree.get_children())
        for n, r in enumerate(inv):
            self.bolle_floor_tree.insert("", "end", iid="f%d" % n, values=(
                r["edificio"], floor_label(r["piano"])) + tuple(r["colori"][c] or "" for c in TONER_NOMI)
                + (r["totale"],))
        tot = sum(r["totale"] for r in inv)
        self.lbl_bolle.configure(
            foreground="#b36b00" if tot else GREEN,
            text=("Toner da montare nei piani: %d" % tot) if tot else "Nessun toner da montare nei piani.")
        self.nb_print.tab(self.tab_bolle, text="Bolle toner" + (" (%d)" % tot if tot else ""))
        stato = dict(self.BOLLE_FILTRI).get(self.var_bolle_stato.get())
        keep = self.bolle_tree.selection()
        keep_id = self.bolle_by_iid.get(keep[0]) if keep else None
        self.bolle_tree.delete(*self.bolle_tree.get_children())
        self.bolle_by_iid = {}
        for n, b in enumerate(self.db.bolle_all(stato)):
            iid = "b%d" % n
            self.bolle_by_iid[iid] = b["id"]
            bolla = b["numero"] or b["codice"] or "(senza numero)"
            if b["numero"] and b["codice"]:
                bolla += " · " + b["codice"]
            self.bolle_tree.insert("", "end", iid=iid, tags=(b["stato"],), values=(
                "Montato" if b["stato"] == "montato" else "Da montare", b["colore"], b["edificio"],
                floor_label(b["piano"]), bolla, fmt_dt(b["data_arrivo"]),
                fmt_dt(b["data_montaggio"]) if b["data_montaggio"] else "",
                b["montata_nome"] or ("" if b["stato"] != "montato" else "(stampante non più presente)")))
            if keep_id == b["id"]:
                self.bolle_tree.selection_set(iid)

    def bolla_photo(self):
        sel = self.bolle_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima una bolla dall'elenco.")
            return
        row = self.db.conn.execute("SELECT foto FROM bolle WHERE id = ?",
                                   (self.bolle_by_iid.get(sel[0]),)).fetchone()
        path = os.path.join(bolle_foto_root(), row["foto"]) if row and row["foto"] else ""
        if not path or not os.path.isfile(path):
            messagebox.showinfo(APP_NAME, "Per questa bolla non c'è la foto: la scatti sul telefono "
                                          "quando registri l'arrivo.")
            return
        self.open_file(path)

    def scorta_models(self):
        models = list(self.models_list())
        for r in self.db.scorta_all():
            if r["modello"].casefold() not in {m.casefold() for m in models}:
                models.append(r["modello"])
        return models

    def refresh_scorta(self):
        by_key = {(r["modello"].casefold(), r["colore"]): r for r in self.db.scorta_all()}
        keep = self.scorta_tree.selection()
        keep_key = self.scorta_keys.get(keep[0]) if keep else None
        self.scorta_tree.delete(*self.scorta_tree.get_children())
        self.scorta_keys = {}
        n = 0
        for modello in self.scorta_models():
            for colore in TONER_NOMI:
                row = by_key.get((modello.casefold(), colore))
                iid = "s%d" % n
                n += 1
                self.scorta_keys[iid] = (modello, colore)
                if row:
                    stato = scorta_status(row["quantita"], row["soglia"])
                    values = (modello, colore, row["quantita"], row["soglia"], stato)
                else:
                    stato = "Non impostata"
                    values = (modello, colore, "-", "", stato)
                self.scorta_tree.insert("", "end", iid=iid, tags=(stato,), values=values)
                if keep_key and keep_key == (modello, colore):
                    self.scorta_tree.selection_set(iid)
        low = self.db.scorta_bassa()
        if low:
            self.lbl_scorta_alert.configure(
                foreground=RED, text="Attenzione, da ordinare: " + "; ".join(
                    "%s %s (ne restano %d)" % (r["colore"], r["modello"], r["quantita"])
                    for r in low))
        else:
            tracked = len(by_key)
            self.lbl_scorta_alert.configure(
                foreground=GREEN, text="Scorta toner: tutto a posto." if tracked
                else "Scorta toner: non hai ancora impostato nulla.")
        self.nb_print.tab(self.tab_scorta, text="Scorta toner" + (" ⚠" if low else ""))
        self.nb.tab(self.tab_print, text="Stampanti" + (" ⚠" if low else ""))
        if hasattr(self, "bolle_tree"):
            self.refresh_bolle()

    def _scorta_selected(self):
        sel = self.scorta_tree.selection()
        return self.scorta_keys.get(sel[0]) if sel else None

    def scorta_receive(self):
        key = self._scorta_selected()
        ScortaDialog(self, self.scorta_models(), "add", key[0] if key else "",
                     key[1] if key else "", 1, SOGLIA_SCORTA, self.save_scorta)

    def scorta_edit(self):
        key = self._scorta_selected()
        if not key:
            messagebox.showinfo(APP_NAME, "Seleziona prima una riga dall'elenco.")
            return
        row = self.db.scorta_get(*key)
        ScortaDialog(self, self.scorta_models(), "set", key[0], key[1],
                     row["quantita"] if row else 0, row["soglia"] if row else SOGLIA_SCORTA,
                     self.save_scorta)

    def scorta_untrack(self):
        key = self._scorta_selected()
        if not key:
            messagebox.showinfo(APP_NAME, "Seleziona prima una riga dall'elenco.")
            return
        if not self.db.scorta_get(*key):
            messagebox.showinfo(APP_NAME, "Questo toner non è contato: non c'è niente da togliere.")
            return
        if messagebox.askyesno(APP_NAME, "Smettere di contare \"%s - %s\"?\nNon ti avviserò più "
                                         "quando sta per finire." % (key[1], key[0])):
            self.db.delete_scorta(*key)
            self.refresh_scorta()

    def save_scorta(self, mode, data):
        modello = " ".join(data["modello"].split())
        if not modello:
            messagebox.showwarning(APP_NAME, "Scegli o scrivi il modello della stampante.")
            return False
        if data["colore"] not in TONER_NOMI:
            messagebox.showwarning(APP_NAME, "Scegli il toner o la vaschetta.")
            return False
        try:
            if mode == "add":
                self.db.add_scorta(modello, data["colore"], data["quantita"])
            else:
                self.db.set_scorta(modello, data["colore"], data["quantita"], data["soglia"])
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return False
        self.refresh_scorta()
        return True

    # ---------------- scheda: verbali Word ----------------
    def _build_verbali(self):
        v = self.tab_verb
        v.columnconfigure(0, weight=1)
        v.rowconfigure(2, weight=1)
        self.verb_rows = {}
        self.var_v_tipo = tk.StringVar(value="(Tutti)")
        self.var_v_search = tk.StringVar()
        if VB is None:
            ttk.Label(v, foreground=RED, wraplength=900, justify="left",
                      text="Manca il file verbali.py (e verbali_modelli.py) accanto a helpdesk.py.\n"
                           "Caricali su GitHub insieme a helpdesk.py e ricostruisci il programma."
                      ).grid(row=0, column=0, sticky="w")
            return
        top = ttk.LabelFrame(v, text="Crea un nuovo verbale Word", padding=10)
        top.grid(row=0, column=0, sticky="ew")
        for tipo in VB.TIPI_ORDINE:
            ttk.Button(top, text=VB.TIPI[tipo][0] + "…", style="Primary.TButton",
                       command=lambda t=tipo: self.new_verbale(t)).pack(side="left", padx=(0, 8))
        ttk.Button(top, text="Verbali in serie…", command=self.verbali_serie).pack(
            side="left", padx=(16, 8))
        mb = ttk.Menubutton(top, text="Modelli Word")
        menu = tk.Menu(mb, tearoff=0)
        menu.add_command(label="Apri la cartella dei modelli", command=self.open_templates_folder)
        menu.add_command(label="Ripristina i modelli originali…", command=self.restore_verbali_models)
        mb["menu"] = menu
        mb.pack(side="right")
        ttk.Button(top, text="Dati per i verbali…", command=self.verbali_settings).pack(
            side="right", padx=(0, 8))

        sf = ttk.Frame(v)
        sf.grid(row=1, column=0, sticky="ew", pady=(12, 6))
        ttk.Label(sf, text="Verbali creati").pack(side="left")
        ttk.Label(sf, text="Tipo").pack(side="left", padx=(18, 0))
        cb = ttk.Combobox(sf, textvariable=self.var_v_tipo, width=26, state="readonly",
                          values=["(Tutti)"] + [VB.TIPI[t][0] for t in VB.TIPI_ORDINE])
        cb.pack(side="left", padx=(8, 0))
        ttk.Label(sf, text="Cerca").pack(side="left", padx=(18, 0))
        ttk.Entry(sf, textvariable=self.var_v_search, width=30).pack(side="left", padx=(8, 0))
        self.var_v_tipo.trace_add("write", lambda *_: self.refresh_verbali())
        self.var_v_search.trace_add("write", lambda *_: self.refresh_verbali())
        self.lbl_verb_count = ttk.Label(sf, style="Hint.TLabel")
        self.lbl_verb_count.pack(side="right")

        mid = ttk.Frame(v)
        mid.grid(row=2, column=0, sticky="nsew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)
        cols = ("data", "tipo", "soggetto", "dettaglio", "foto", "firme", "stato", "file")
        self.verb_tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        self.verb_tree.tag_configure("da_fare", foreground=RED)
        for c, title, width, stretch in (
                ("data", "Data", 80, False), ("tipo", "Verbale", 170, False),
                ("soggetto", "Utente / oggetto", 190, False), ("dettaglio", "Dettaglio", 260, True),
                ("foto", "Foto", 45, False), ("firme", "Firme", 190, False),
                ("stato", "Word", 100, False), ("file", "File Word", 190, False)):
            self.verb_tree.heading(c, text=title)
            self.verb_tree.column(c, width=width, anchor="center" if c == "foto" else "w",
                                  stretch=stretch)
        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.verb_tree.yview)
        self.verb_tree.configure(yscrollcommand=vsb.set)
        self.verb_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.verb_tree.bind("<Double-1>", lambda e: self.open_verbale_file())

        bf = ttk.Frame(v)
        bf.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Crea il Word…", style="Save.TButton",
                   command=self.create_word_for_selected).pack(side="left", padx=(0, 8))
        ttk.Button(bf, text="Apri il Word", style="Primary.TButton",
                   command=self.open_verbale_file).pack(side="left")
        ttk.Button(bf, text="Apri la cartella", command=self.open_verbale_folder).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Modifica e ricrea…", command=self.edit_verbale).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Firme…", command=self.edit_firme).pack(side="left", padx=(8, 0))
        ttk.Button(bf, text="Togli dall'elenco", style="Danger.TButton",
                   command=self.delete_verbale).pack(side="left", padx=(8, 0))
        ttk.Label(v, style="Hint.TLabel", wraplength=1000, justify="left",
                  text="I verbali fatti sul telefono (con foto e firme) arrivano con File > Importa "
                       "dati dal telefono: qui compaiono con «Word da creare»; selezionali e "
                       "premi «Crea il Word…». Il file Word resta nella cartella che scegli. "
                       "«Modifica e ricrea» corregge un dato e rifà il Word. «Togli dall'elenco» "
                       "non cancella il file Word (se il verbale è anche sul telefono, toglilo "
                       "anche lì, altrimenti torna con la prossima importazione)."
                  ).grid(row=4, column=0, sticky="w", pady=(8, 0))

    def verbali_storico(self):
        """Quello che hai scritto nei verbali gia' fatti (serve a suggerire e compilare da solo)."""
        rows = []
        if VB is not None:
            for r in self.db.search_verbali():
                try:
                    dati = json.loads(r["dati"] or "{}")
                except ValueError:
                    continue
                rows.append((r["modificato"] or "", r["tipo"], dati))
        rows.sort(key=lambda x: x[0], reverse=True)
        return VB.storico([(t, d) for _, t, d in rows]) if VB is not None else {}

    def verbale_luoghi(self):
        """Suggerimenti per provenienza/destinazione: strutture e sedi note."""
        found = {}
        for name in [b["nome"] for b in self.buildings] + list(self.sedi_cache):
            if name and name.strip():
                found.setdefault(name.strip().casefold(), name.strip())
        return sorted(found.values(), key=str.casefold)

    def refresh_verbali(self):
        if VB is None or not hasattr(self, "verb_tree"):
            return
        label = self.var_v_tipo.get()
        tipo = next((t for t in VB.TIPI_ORDINE if VB.TIPI[t][0] == label), "")
        rows = self.db.search_verbali(tipo, self.var_v_search.get())
        self.verb_rows = {r["id"]: r for r in rows}
        self.verb_tree.delete(*self.verb_tree.get_children())
        da_fare = 0
        for r in rows:
            d = parse_iso_date(r["data"])
            nome_tipo = VB.TIPI[r["tipo"]][0] if r["tipo"] in VB.TIPI else r["tipo"]
            stato = verbale_stato_word(r)
            da_fare += stato != "Creato"
            self.verb_tree.insert("", "end", iid=str(r["id"]), tags=(
                () if stato == "Creato" else ("da_fare",)), values=(
                fmt_date(d) if d else r["data"], nome_tipo, r["soggetto"], r["dettaglio"],
                r["n_foto"] or "", verbale_firme_text(r), stato, os.path.basename(r["file"])))
        self.lbl_verb_count.configure(
            text="%d verbali" % len(rows) + (" · %d da creare/rifare" % da_fare if da_fare else ""))

    def selected_verbale(self):
        sel = self.verb_tree.selection() if hasattr(self, "verb_tree") else ()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima un verbale dall'elenco.")
            return None
        return self.verb_rows.get(int(sel[0]))

    def verbali_ready(self):
        if VB is None:
            messagebox.showinfo(APP_NAME, "Mancano i file verbali.py e verbali_modelli.py: "
                                "caricali su GitHub insieme a helpdesk.py.")
            return False
        return True

    def new_verbale(self, tipo, ini=None):
        if not self.verbali_ready():
            return
        VerbaleDialog(self, tipo, ini, self.make_verbale)

    def new_verbale_for_user(self, tipo):
        """Dalla scheda Utenti: apre il verbale con i dati dell'utente già scritti."""
        user = self.selected_user()
        if not user or not self.verbali_ready():
            return
        nome = user["nome"]
        if tipo == "trasferimento":
            ini = {"assegnatario": nome, "contatti": user["telefono"]}
        elif tipo == "presa_in_carico":
            ini = {"utente": "%s (matricola %s)" % (nome, user["matricola"])
                   if user["matricola"] else nome}
        elif tipo == "assegnazione":
            ini = {"utente": nome, "matricola": user["matricola"]}
        else:
            ini = {}
        self.new_verbale(tipo, ini)

    # ----- verbali in serie: un Word per ogni persona di una stanza o di un piano -----
    def verbali_serie(self):
        if not self.verbali_ready():
            return
        if not self.db.edifici():
            messagebox.showinfo(APP_NAME, "Prima inserisci l'edificio e le stanze con le persone "
                                "(scheda Stampanti › Piani e stanze).")
            return
        SerieVerbaliDialog(self, self.create_verbali_serie)

    def serie_persone(self, eid, piano, stanza_id=None):
        """Le persone sedute nella stanza (o in tutto il piano), con i dati delle loro schede."""
        out = []
        for r in self.db.stanze_of(eid, piano):
            if stanza_id is None or r["id"] == stanza_id:
                for p in self.db.persone_of(r["id"]):
                    out.append(dict(p, stanza=r["codice"]))
        return out

    def serie_plan(self, params):
        if params.get("eid") is None or params.get("piano") is None:
            return {"verbali": [], "saltati": []}
        titoli = self.verbali_storico().get("titoli", {}) if VB is not None else {}
        return plan_verbali_serie(
            self.serie_persone(params["eid"], params["piano"], params.get("stanza_id")),
            params["tipo"], params["data"], {"cosa": params.get("cosa"), "motivo": params.get("motivo")},
            titoli)

    def serie_preview(self, params):
        dt = parse_date(params["data"])
        plan = self.serie_plan(dict(params, data=fmt_date(dt) if dt else params["data"]))
        n, k = len(plan["verbali"]), len(plan["saltati"])
        if not n and not k:
            return "In questa stanza o piano non siedono persone."
        text = "Verranno creati %d verbali." % n
        if k:
            text += " Saltate %d persone (non hanno niente da scrivere)." % k
        if not dt:
            text += "\nLa data non è valida: scrivila così gg/mm/aaaa."
        return text

    def create_verbali_serie(self, params):
        dt = parse_date(params["data"])
        if not dt:
            messagebox.showwarning(APP_NAME, "La data non è valida.\nScrivila così: gg/mm/aaaa "
                                             "(esempio 15/11/2026).")
            return False
        folder = params["cartella"].strip()
        if not folder:
            messagebox.showwarning(APP_NAME, "Scegli la cartella dove salvare i verbali.")
            return False
        firma = None
        if params.get("firma_tecnico"):
            firma = saved_signature()
            if not firma:
                messagebox.showwarning(APP_NAME, "Non c'è ancora un'immagine della tua firma: "
                                       "scegline una in «Dati per i verbali», oppure togli la spunta.")
                return False
        tipo = params["tipo"]
        plan = self.serie_plan(dict(params, data=fmt_date(dt)))
        if not plan["verbali"]:
            messagebox.showinfo(APP_NAME, "Non c'è nessun verbale da creare: " + (
                "nessuna persona in questa stanza o piano." if not plan["saltati"] else
                "nessuna persona ha i dati necessari."))
            return False
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError as e:
            messagebox.showerror(APP_NAME, "Non riesco a usare la cartella:\n%s" % e)
            return False
        done, error = [], None
        for item in plan["verbali"]:
            d = item["d"]
            uid = VB.nuovo_codice()
            foto_dir = os.path.join(verbali_foto_root(), uid)
            stem = os.path.splitext(VB.nome_file_proposto(tipo, dict(d, utente=item["disp"])))[0]
            path = unique_path(folder, stem, ".docx")
            try:
                firme_names, firme = {}, {}
                if firma:
                    firme_names, _ = VB.applica_scelte_firme(foto_dir, {}, set(),
                                                             {"tecnico": "salvata"}, firma)
                    firme = {k: os.path.join(foto_dir, v) for k, v in firme_names.items()}
                VB.crea_verbale(tipo, d, path, templates_dir(), photos=[], imp=self.settings,
                                firme=firme, avvisi=[], cartaceo=set())
            except VB.VerbaleError as e:
                self._drop_files([], foto_dir)
                error = "%s: %s" % (item["disp"], e)
                break
            except Exception as e:
                self._drop_files([], foto_dir)
                error = "%s: %s" % (item["disp"], e)
                break
            dati = dict(d)
            dati["_foto"], dati["_firme"], dati["_cartaceo"] = [], firme_names, {}
            soggetto, dettaglio = riassunto_verbale(tipo, d)
            now = now_utc_iso()
            try:
                self.db.add_verbale(uid, tipo, dt.isoformat(), soggetto, dettaglio, dati, path, 0,
                                    modificato=now, word_il=now)
            except sqlite3.Error as e:
                error = "%s: il Word è stato creato ma non è nell'elenco (%s)" % (item["disp"], e)
                break
            done.append(item["disp"])
        if done:
            self.settings["cartella_verbali"] = folder
            save_settings(self.settings)
            self.refresh_verbali()
        lines = ["Verbali creati: %d (cartella: %s)." % (len(done), folder)]
        if plan["saltati"]:
            lines.append("\nSaltati (%d):" % len(plan["saltati"]))
            lines += ["   %s: %s" % (n, why) for n, why in plan["saltati"][:15]]
            if len(plan["saltati"]) > 15:
                lines.append("   … e altri %d" % (len(plan["saltati"]) - 15))
        if error:
            lines.append("\nMi sono fermato qui: %s" % error)
            lines.append("Quelli già creati restano nell'elenco dei verbali.")
        messagebox.showwarning(APP_NAME, "\n".join(lines)) if error else \
            messagebox.showinfo(APP_NAME, "\n".join(lines))
        if done and not error and messagebox.askyesno(APP_NAME, "Apro la cartella dei verbali?"):
            open_path(folder)
        return bool(done)

    def show_user_verbali(self):
        user = self.selected_user()
        if not user or not self.verbali_ready():
            return
        self.var_v_tipo.set("(Tutti)")
        self.var_v_search.set(user["nome"])
        self.nb.select(self.tab_verb)

    def create_word_for_selected(self):
        """Fa il Word di un verbale già nell'elenco (per esempio arrivato dal telefono)."""
        r = self.selected_verbale()
        if not r or not self.verbali_ready():
            return
        try:
            dati = json.loads(r["dati"] or "{}")
        except ValueError:
            dati = {}
        foto_dir = os.path.join(verbali_foto_root(), r["uid"])
        photos = [os.path.join(foto_dir, n) for n in dati.get("_foto", [])
                  if os.path.isfile(os.path.join(foto_dir, n))]
        d = {k: v for k, v in dati.items() if not k.startswith("_")}
        self.make_verbale(r["tipo"], d, photos, self, r)

    def make_verbale(self, tipo, d, photos, parent, edit_row=None):
        """Chiamata dal modulo: sceglie dove salvare, crea il Word e lo registra nell'elenco."""
        if edit_row and edit_row["file"]:
            init_dir, init_name = os.path.split(edit_row["file"])
        else:
            init_dir = self.settings.get("cartella_verbali", "")
            init_name = VB.nome_file_proposto(tipo, d)
        kw = {"initialdir": init_dir} if init_dir and os.path.isdir(init_dir) else {}
        path = filedialog.asksaveasfilename(
            parent=parent, title="Salva il verbale Word", defaultextension=".docx",
            initialfile=init_name, filetypes=[("Documento Word", "*.docx")], **kw)
        if not path:
            return False
        uid = edit_row["uid"] if edit_row else VB.nuovo_codice()
        foto_dir = os.path.join(verbali_foto_root(), uid)
        stamp = datetime.now().strftime("%Y%m%d%H%M%S")
        old = {}
        if edit_row:
            try:
                old = json.loads(edit_row["dati"] or "{}")
            except ValueError:
                old = {}
        firme_names = {k: v for k, v in (old.get("_firme") or {}).items()
                       if os.path.isfile(os.path.join(foto_dir, v))}
        cart = {k for k, v in (old.get("_cartaceo") or {}).items()
                if v and k in VB.FIRME_CHIAVI and k not in firme_names}
        scelte = getattr(parent, "firme_scelte", None)    # scelte fatte nel riquadro «Firme»
        if isinstance(scelte, dict) and scelte:
            try:
                firme_names, cart = VB.applica_scelte_firme(
                    foto_dir, firme_names, cart, scelte, saved_signature())
            except VB.VerbaleError as e:
                messagebox.showerror(APP_NAME, str(e), parent=parent)
                return False
        firme = {k: os.path.join(foto_dir, v) for k, v in firme_names.items()}
        avvisi = []
        ready, created = [], []
        try:
            os.makedirs(foto_dir, exist_ok=True)
            for i, src in enumerate(photos, 1):
                if os.path.normcase(os.path.dirname(os.path.abspath(src))) == \
                        os.path.normcase(os.path.abspath(foto_dir)):
                    ready.append(src)           # già preparata in una creazione precedente
                else:
                    target = VB.prepara_foto(src, foto_dir, "foto_%s_%02d" % (stamp, i))
                    ready.append(target)
                    created.append(target)
            VB.crea_verbale(tipo, d, path, templates_dir(), photos=ready, imp=self.settings,
                            firme=firme, avvisi=avvisi, cartaceo=cart)
        except VB.VerbaleError as e:
            self._drop_files(created, foto_dir)
            messagebox.showerror(APP_NAME, str(e), parent=parent)
            return False
        except Exception as e:
            self._drop_files(created, foto_dir)
            messagebox.showerror(APP_NAME, "Non sono riuscito a creare il verbale:\n%s" % e,
                                 parent=parent)
            return False
        # foto non più usate (tolte in "Modifica e ricrea")
        keep = {os.path.normcase(os.path.abspath(p)) for p in list(ready) + list(firme.values())}
        try:
            for name in os.listdir(foto_dir):
                full = os.path.abspath(os.path.join(foto_dir, name))
                if os.path.normcase(full) not in keep:
                    os.remove(full)
        except OSError:
            pass
        dati = dict(d)
        dati["_foto"] = [os.path.basename(p) for p in ready]
        dati["_firme"] = firme_names
        dati["_cartaceo"] = {k: True for k in sorted(cart)}
        soggetto, dettaglio = riassunto_verbale(tipo, d)
        data_iso = (parse_date(d["data"]) or date.today()).isoformat()
        now = now_utc_iso()
        try:
            if edit_row:
                self.db.update_verbale(edit_row["id"], data_iso, soggetto, dettaglio, dati, path,
                                       len(ready), modificato=now, word_il=now)
            else:
                self.db.add_verbale(uid, tipo, data_iso, soggetto, dettaglio, dati, path,
                                    len(ready), modificato=now, word_il=now)
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Il Word è stato creato ma non sono riuscito a "
                                 "metterlo nell'elenco:\n%s\n\nFile: %s" % (e, path), parent=parent)
        # ricordo cartella e (per i censimenti) tecnico e sede per la volta dopo
        self.settings["cartella_verbali"] = os.path.dirname(path)
        if tipo == "censimento":
            for k in ("tecnico", "sede_citta", "indirizzo", "cap"):
                if d.get(k):
                    self.settings[k] = d[k]
        save_settings(self.settings)
        self.refresh_verbali()
        if avvisi:
            messagebox.showwarning(APP_NAME, "\n\n".join(avvisi), parent=parent)
        if messagebox.askyesno(APP_NAME, "Verbale creato:\n%s\n\nVuoi aprirlo ora?" % path,
                               parent=parent):
            open_path(path)
        return True

    def _drop_files(self, files, folder):
        for f in files:
            try:
                os.remove(f)
            except OSError:
                pass
        try:
            os.rmdir(folder)        # solo se è vuota
        except OSError:
            pass

    def open_verbale_file(self):
        r = self.selected_verbale()
        if not r:
            return
        if not r["file"] or not os.path.isfile(r["file"]):
            messagebox.showinfo(APP_NAME, "Il file Word non si trova più qui:\n%s\n\nSe lo hai "
                                "spostato o cancellato puoi rifarlo con «Modifica e ricrea…»."
                                % r["file"])
            return
        if not open_path(r["file"]):
            messagebox.showinfo(APP_NAME, "Il file è qui:\n%s" % r["file"])

    def open_verbale_folder(self):
        r = self.selected_verbale()
        if not r:
            return
        folder = os.path.dirname(r["file"])
        if r["file"] and os.path.isfile(r["file"]):
            try:
                subprocess.Popen(["explorer", "/select,", os.path.normpath(r["file"])])
                return
            except Exception:
                pass
        if folder and os.path.isdir(folder):
            if open_path(folder):
                return
        messagebox.showinfo(APP_NAME, "La cartella non si trova più:\n%s" % folder)

    def edit_verbale(self):
        r = self.selected_verbale()
        if not r or not self.verbali_ready():
            return
        try:
            ini = json.loads(r["dati"] or "{}")
        except ValueError:
            ini = {}
        foto_dir = os.path.join(verbali_foto_root(), r["uid"])
        ini["foto"] = [os.path.join(foto_dir, n) for n in ini.pop("_foto", [])]
        firme_names = ini.pop("_firme", None) or {}
        ini["firme_esistenti"] = {k: os.path.join(foto_dir, v) for k, v in firme_names.items()
                                  if os.path.isfile(os.path.join(foto_dir, v))}
        ini["firme_cartaceo"] = sorted(k for k, v in (ini.pop("_cartaceo", None) or {}).items()
                                       if v and k not in ini["firme_esistenti"])
        if r["tipo"] not in VB.TIPI:
            messagebox.showinfo(APP_NAME, "Tipo di verbale non riconosciuto.")
            return
        VerbaleDialog(self, r["tipo"], ini,
                      lambda tipo, d, photos, parent: self.make_verbale(tipo, d, photos, parent, r))

    def edit_firme(self):
        """Cambia le firme di un verbale già nell'elenco (firma cartacea, togli la firma, la mia
        firma salvata per il tecnico) senza riaprire tutto il modulo."""
        r = self.selected_verbale()
        if not r or not self.verbali_ready():
            return
        if r["tipo"] not in VB.TIPI:
            messagebox.showinfo(APP_NAME, "Tipo di verbale non riconosciuto.")
            return
        try:
            dati = json.loads(r["dati"] or "{}")
        except ValueError:
            dati = {}
        folder = os.path.join(verbali_foto_root(), r["uid"])
        esistenti = {k: os.path.join(folder, v) for k, v in (dati.get("_firme") or {}).items()
                     if os.path.isfile(os.path.join(folder, v))}
        cart = sorted(k for k, v in (dati.get("_cartaceo") or {}).items()
                      if v and k not in esistenti)
        FirmeDialog(self, r["tipo"], esistenti, cart, lambda scelte: self.save_firme(r, scelte))

    def save_firme(self, r, scelte):
        """Salva le scelte sulle firme del verbale r. True se riuscito."""
        try:
            dati = json.loads(r["dati"] or "{}")
        except ValueError:
            dati = {}
        folder = os.path.join(verbali_foto_root(), r["uid"])
        names = {k: v for k, v in (dati.get("_firme") or {}).items()
                 if os.path.isfile(os.path.join(folder, v))}
        cart = {k for k, v in (dati.get("_cartaceo") or {}).items()
                if v and k in VB.FIRME_CHIAVI and k not in names}
        try:
            names, cart = VB.applica_scelte_firme(folder, names, cart, scelte, saved_signature())
        except VB.VerbaleError as e:
            messagebox.showerror(APP_NAME, str(e))
            return False
        for fname in os.listdir(folder) if os.path.isdir(folder) else ():
            if fname.startswith("firma_") and fname not in names.values():
                try:
                    os.remove(os.path.join(folder, fname))
                except OSError:
                    pass
        dati["_firme"] = names
        dati["_cartaceo"] = {k: True for k in sorted(cart)}
        try:
            self.db.update_verbale(r["id"], r["data"], r["soggetto"], r["dettaglio"], dati,
                                   r["file"], r["n_foto"], modificato=now_utc_iso())
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare le firme:\n%s" % e)
            return False
        self.refresh_verbali()
        return True

    def delete_verbale(self):
        r = self.selected_verbale()
        if not r:
            return
        if not messagebox.askyesno(
                APP_NAME, "Togliere questo verbale dall'elenco?\n\n%s - %s\n\nIl file Word NON "
                          "viene cancellato: resta dove l'hai salvato. Vengono eliminate dal "
                          "programma le foto allegate e i dati per «Modifica e ricrea»."
                          % (r["soggetto"], os.path.basename(r["file"])), icon="warning"):
            return
        self.db.delete_verbale(r["id"])
        shutil.rmtree(os.path.join(verbali_foto_root(), r["uid"]), ignore_errors=True)
        self.refresh_verbali()

    def verbali_settings(self):
        if not self.verbali_ready():
            return

        def save(values):
            self.settings.update(values)
            if not save_settings(self.settings):
                messagebox.showwarning(APP_NAME, "Non sono riuscito a salvare le impostazioni.")
        VerbaliSettingsDialog(self, self.settings, save)

    def open_templates_folder(self):
        if not self.verbali_ready():
            return
        try:
            VB.ensure_templates(templates_dir())
        except Exception as e:
            messagebox.showerror(APP_NAME, "Non riesco a preparare i modelli:\n%s" % e)
            return
        if not open_path(templates_dir()):
            messagebox.showinfo(APP_NAME, "I modelli sono in:\n%s" % templates_dir())

    def restore_verbali_models(self):
        if not self.verbali_ready():
            return
        if not messagebox.askyesno(
                APP_NAME, "Rimettere i quattro modelli Word originali?\n\nI modelli che hai "
                          "modificato non vanno persi: ne resta una copia con scritto «(prima del "
                          "ripristino)» nella stessa cartella."):
            return
        try:
            kept = VB.restore_templates(templates_dir())
        except Exception as e:
            messagebox.showerror(APP_NAME, "Ripristino non riuscito:\n%s" % e)
            return
        messagebox.showinfo(APP_NAME, "Modelli ripristinati.\n\nCopie dei precedenti: %d.\nCartella:"
                            "\n%s" % (len(kept), templates_dir()))

    # ---------------- varie ----------------
    def refresh_stats(self):
        oggi, mese, tot = self.db.stats()
        aperti = self.db.count_aperti()
        self.stat_vars["aperti"].set(str(aperti))
        self.stat_labels["aperti"].configure(fg=RED if aperti else NAVY)
        self.stat_vars["oggi"].set(str(oggi))
        self.stat_vars["mese"].set(str(mese))
        self.stat_vars["tot"].set(str(tot))

    def backup(self):
        path = filedialog.asksaveasfilename(
            title="Salva una copia di backup (dati e foto)", defaultextension=".zip",
            initialfile="backup_helpdesk_%s.zip" % datetime.now().strftime("%Y%m%d"),
            filetypes=[("Archivio zip", "*.zip")])
        if not path:
            return
        tmp_db = os.path.join(data_dir(), "_backup_tmp.db")
        try:
            self.db.backup_to(tmp_db)
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
                z.write(tmp_db, "interventi.db")
                root = attachments_root()
                for folder, dirs, files in os.walk(root):
                    dirs[:] = [d for d in dirs if d != "_temp"]
                    for name in files:
                        full = os.path.join(folder, name)
                        z.write(full, os.path.join("allegati", os.path.relpath(full, root)))
        except Exception as e:
            messagebox.showerror(APP_NAME, "Backup non riuscito:\n%s" % e)
            return
        finally:
            try:
                os.remove(tmp_db)
            except OSError:
                pass
        messagebox.showinfo(APP_NAME, "Copia salvata in:\n%s\n\nContiene database e foto. Per "
                            "ripristinarla estrai lo zip nella cartella dei dati "
                            "(File > Apri la cartella dei dati)." % path)

    def export_for_phone(self):
        foto = messagebox.askyesno(
            APP_NAME, "Vuoi mandare al telefono anche le FOTO degli interventi?\n\n"
            "Sì = ci sono tutti i dati con le foto (il file è più pesante).\n"
            "No = utenti e interventi senza foto (file leggero).")
        path = filedialog.asksaveasfilename(
            title="Esporta dati per il telefono", defaultextension=".json",
            initialfile="dati_per_telefono_%s.json" % datetime.now().strftime("%Y%m%d_%H%M"),
            filetypes=[("File dati", "*.json")])
        if not path:
            return
        try:
            data, n_foto = build_phone_export(self.db, with_photos=foto, settings=self.settings)
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False)
        except (OSError, sqlite3.Error) as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a creare il file:\n%s" % e)
            return
        messagebox.showinfo(APP_NAME, "File creato con %d utenti, %d interventi, %d foto, "
                            "%d stampanti, %d cambi toner, %d interventi del tecnico e %d "
                            "verbali:\n%s\n\n"
                            "Mandalo al telefono (AirDrop, email o iCloud) e nell'app tocca "
                            "Scambio > Ricevi dal PC."
                            % (len(data["utenti"]), len(data["interventi"]), n_foto,
                               len(data["stampanti"]), len(data["toner"]),
                               len(data["manutenzioni"]), len(data["verbali"]), path))

    def import_from_phone(self):
        path = filedialog.askopenfilename(
            title="Importa dati dal telefono",
            filetypes=[("File dati", "*.json"), ("Tutti i file", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            res = import_phone_data(self.db, data)
        except (OSError, ValueError) as e:
            messagebox.showerror(APP_NAME, "Non riesco a leggere il file:\n%s" % e)
            return
        self.refresh_users()
        self.refresh_stats()
        self.apply_filters()
        self.refresh_verbali()
        if hasattr(self, "refresh_rooms"):
            self.refresh_rooms()
        lines = ["Interventi nuovi: %d" % res["interventi_nuovi"],
                 "Interventi aggiornati: %d" % res["interventi_aggiornati"],
                 "Già presenti (saltati): %d" % res["interventi_gia_presenti"],
                 "Foto importate: %d" % res["foto"],
                 "Utenti nuovi: %d, aggiornati: %d" % (res["utenti_nuovi"], res["utenti_aggiornati"])]
        if res["utenti_uniti"]:
            lines.append("Persone già unite sul PC (il telefono ne aveva ancora una copia): %d"
                         % res["utenti_uniti"])
        if res["toner_nuovi"] or res["toner_gia_presenti"] or res["toner_saltati"]:
            lines.append("Cambi toner nuovi: %d (già presenti: %d)"
                         % (res["toner_nuovi"], res["toner_gia_presenti"]))
        if res["toner_saltati"]:
            lines.append("Cambi toner ignorati (stampante non trovata): %d" % res["toner_saltati"])
        if res["bolle_nuove"] or res["bolle_aggiornate"] or res["bolle_gia_presenti"] \
                or res["bolle_saltate"]:
            lines.append("Bolle dei toner: nuove %d, aggiornate %d, già presenti %d (foto: %d)"
                         % (res["bolle_nuove"], res["bolle_aggiornate"], res["bolle_gia_presenti"],
                            res["bolle_foto"]))
        if res["bolle_saltate"]:
            lines.append("Bolle ignorate (edificio non trovato o dati non validi): %d"
                         % res["bolle_saltate"])
        if res["manut_nuove"] or res["manut_aggiornate"] or res["manut_gia_presenti"] \
                or res["manut_saltate"]:
            lines.append("Interventi del tecnico: nuovi %d, aggiornati %d, già presenti %d"
                         % (res["manut_nuove"], res["manut_aggiornate"],
                            res["manut_gia_presenti"]))
        if res["manut_saltate"]:
            lines.append("Interventi del tecnico ignorati (stampante non trovata): %d"
                         % res["manut_saltate"])
        if res["verbali_nuovi"] or res["verbali_aggiornati"] or res["verbali_gia_presenti"] \
                or res["verbali_saltati"]:
            lines.append("Verbali: nuovi %d, aggiornati %d, già presenti %d (foto: %d)"
                         % (res["verbali_nuovi"], res["verbali_aggiornati"],
                            res["verbali_gia_presenti"], res["verbali_foto"]))
            if res["verbali_nuovi"] or res["verbali_aggiornati"]:
                lines.append("   -> scheda «Verbali»: «Crea il Word…» per quelli da creare.")
        if res["verbali_saltati"]:
            lines.append("Verbali ignorati (non validi o programma senza verbali.py): %d"
                         % res["verbali_saltati"])
        if res["stanze_nuove"] or res["stanze_aggiornate"] or res["stanze_gia_presenti"] \
                or res["postazioni"] or res["piante"] or res["stanze_foto"]:
            lines.append("Stanze: nuove %d, aggiornate %d, già presenti %d"
                         % (res["stanze_nuove"], res["stanze_aggiornate"],
                            res["stanze_gia_presenti"]))
            lines.append("Persone nelle stanze aggiornate: %d, piante: %d, foto delle stanze: %d"
                         % (res["postazioni"], res["piante"], res["stanze_foto"]))
        if res["stanze_saltate"] or res["postazioni_saltate"]:
            lines.append("Stanze o persone ignorate (edificio, stanza o utente non trovati): %d"
                         % (res["stanze_saltate"] + res["postazioni_saltate"]))
        if res["saltati"]:
            lines.append("Righe non valide ignorate: %d" % res["saltati"])
        if res["scorta_bassa"]:
            lines.append("\nATTENZIONE, scorta toner bassa: " + "; ".join(res["scorta_bassa"]))
        if res["manut_nuove"] or res["manut_aggiornate"]:
            self.after_manut_change()
        if res["toner_nuovi"]:
            self.after_toner_change()
        if res["bolle_nuove"] or res["bolle_aggiornate"]:
            self.refresh_scorta()
        messagebox.showinfo(APP_NAME, "Importazione completata.\n\n" + "\n".join(lines))

    def merge_printers_dialog(self):
        b = self.current_building()
        if not b:
            messagebox.showinfo(APP_NAME, "Crea prima un edificio (scheda Stampanti).")
            return
        props = self.db.printer_merge_proposals(b["id"])
        if not props:
            messagebox.showinfo(APP_NAME, "Non ci sono stampanti doppie da unire in «%s».\n\n"
                                          "Cerco le stampanti senza numero di serie che hanno, allo stesso "
                                          "piano e con lo stesso modello, una stampante col numero di serie."
                                % b["nome"])
            return
        MergePrintersDialog(self, b, props, self.do_merge_printers)

    def do_merge_printers(self, pairs):
        n_t = n_m = 0
        for old_id, keep_id in pairs:
            r = self.db.merge_printers(old_id, keep_id)
            n_t += r["toner"]
            n_m += r["manutenzioni"]
        self.after_printers_change()
        return n_t, n_m

    def import_rooms_file(self):
        b = self.current_building()
        if not b:
            messagebox.showinfo(APP_NAME, "Crea prima un edificio (scheda Stampanti), poi importa "
                                          "l'elenco.")
            return
        path = filedialog.askopenfilename(
            title="Importa elenco stanze e persone",
            filetypes=[("Elenco stanze e persone", "*.json"), ("Tutti i file", "*.*")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict) or data.get("tipo") != ELENCO_TIPO:
                raise ValueError("Questo non è un elenco di stanze e persone dell'app Help Desk.")
        except (OSError, ValueError) as e:
            messagebox.showerror(APP_NAME, "Non riesco a leggere il file:\n%s" % e)
            return
        n_s, n_p = len(data.get("stanze") or []), len(data.get("stampanti") or [])
        if not messagebox.askyesno(
                APP_NAME, "Nel file ci sono %d stanze e %d stampanti.\n\nLe metto nell'edificio "
                          "«%s». Quello che c'è già non viene duplicato né sovrascritto.\n\n"
                          "Vuoi continuare?" % (n_s, n_p, b["nome"])):
            return
        try:
            res = import_rooms_list(self.db, data, b["id"])
        except (ValueError, sqlite3.Error) as e:
            messagebox.showerror(APP_NAME, "Importazione non riuscita:\n%s" % e)
            return
        self.refresh_users()
        self.after_printers_change()
        lines = ["Stanze: nuove %d, già presenti %d (scrivanie aggiornate: %d)"
                 % (res["stanze_nuove"], res["stanze_esistenti"], res["stanze_aggiornate"]),
                 "Persone: nuove %d, già in elenco %d; messe in stanza %d, già lì %d"
                 % (res["persone_nuove"], res["persone_esistenti"], res["persone_assegnate"],
                    res["persone_gia_in_stanza"]),
                 "Uffici scritti: %d" % res["uffici_scritti"],
                 "Matricole scritte: %d; dotazioni (per esempio il monitor) scritte: %d"
                 % (res["matricole_scritte"], res["dotazioni_scritte"]),
                 "Persone senza stanza aggiunte: %d" % res["persone_senza_stanza"],
                 "Stampanti: nuove %d, aggiornate %d, già presenti %d, vecchie senza seriale "
                 "completate (non duplicate) %d"
                 % (res["stampanti_nuove"], res["stampanti_aggiornate"],
                    res["stampanti_gia_presenti"], res["stampanti_adottate"])]
        if res["stampanti_da_unire"]:
            lines.append("\nATTENZIONE: %d stampanti nuove hanno una vecchia simile (stesso piano e "
                         "modello) e non ho indovinato quale sia. Usa File › «Unisci doppioni "
                         "stampanti…»: sposta lo storico sulla nuova e toglie la vecchia."
                         % res["stampanti_da_unire"])
        if res["persone_altrove"]:
            lines.append("\nPersone lasciate dove sono (siedono già in un'altra stanza):")
            lines += ["   " + t for t in res["persone_altrove"][:12]]
            if len(res["persone_altrove"]) > 12:
                lines.append("   … e altre %d" % (len(res["persone_altrove"]) - 12))
        if res["scartati"]:
            lines.append("\nNon importati:")
            lines += ["   " + t for t in res["scartati"][:12]]
        messagebox.showinfo(APP_NAME, "Importazione completata.\n\n" + "\n".join(lines))

    def import_people_excel(self):
        paths = filedialog.askopenfilenames(
            title="Importa persone da Excel (lista INAIL)",
            filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")])
        if not paths:
            return
        files, errors = [], []
        for p in paths:
            name = os.path.basename(p)
            try:
                title, people = read_people_xlsx(p)
            except ValueError as e:
                errors.append("%s: %s" % (name, e))
                continue
            if not people:
                errors.append("%s: non ci sono persone" % name)
                continue
            files.append({"name": name, "title": title, "persone": people,
                          "sede": guess_people_sede(title, name)})
        if errors:
            messagebox.showwarning(APP_NAME, "Questi file non li posso usare:\n\n" +
                                   "\n".join(errors))
        if not files:
            return
        PeopleImportDialog(self, files, self.db.sedi(), self.preview_people_import,
                           self.apply_people_import)

    def preview_people_import(self, lists, update_sede):
        s = plan_people_import(self.db.users_basic(), lists, update_sede)["stats"]
        text = ["Nei file ci sono %d persone: %d nuove da aggiungere, %d già nell'elenco."
                % (s["persone"], s["nuove"], s["presenti"])]
        if s["presenti"]:
            text.append("A chi c'è già scrivo: matricola a %d, cellulare a %d, email a %d, "
                        "sede a %d; ordine per cognome a %d."
                        % (s["matricole"], s["telefoni"], s["email"],
                           s["sedi_scritte"] + s["sedi_cambiate"], s["ordinate"]))
        if s["sede_diversa"]:
            text.append("%d hanno già una sede diversa: la lascio com'è (spunta la casella qui "
                        "sopra per cambiarla)." % s["sede_diversa"])
        if s["conflitti"]:
            text.append("%d da controllare a mano (nome uguale, matricola diversa)."
                        % len(s["conflitti"]))
        return " ".join(text)

    def apply_people_import(self, lists, update_sede):
        try:
            plan = plan_people_import(self.db.users_basic(), lists, update_sede)
            apply_people_plan(self.db, plan)
        except (ValueError, sqlite3.Error) as e:
            messagebox.showerror(APP_NAME, "Importazione non riuscita:\n%s" % e)
            return False
        s = plan["stats"]
        self.refresh_users()
        lines = ["Persone nei file: %d" % s["persone"],
                 "Persone nuove aggiunte: %d" % s["nuove"],
                 "Già nell'elenco: %d (aggiornate %d, senza niente da aggiungere %d)"
                 % (s["presenti"], s["aggiornate"], s["invariate"]),
                 "Matricole scritte: %d; cellulari scritti: %d; uffici scritti: %d; "
                 "email scritte: %d" % (s["matricole"], s["telefoni"], s["uffici"], s["email"]),
                 "Sedi scritte: %d; sedi cambiate: %d" % (s["sedi_scritte"], s["sedi_cambiate"]),
                 "Messe in ordine per cognome: %d" % s["ordinate"]]
        if s["telefono_diverso"]:
            lines.append("Cellulare diverso da quello che avevi già (non toccato): %d"
                         % s["telefono_diverso"])
        if s["sede_diversa"]:
            lines.append("Sede diversa lasciata com'era: %d" % s["sede_diversa"])
        if s["senza_telefono"]:
            lines.append("Senza cellulare nel file: %d" % s["senza_telefono"])
        if s["doppi"]:
            lines.append("Persone presenti in più file (contate una volta): %d" % s["doppi"])
        if s["conflitti"]:
            lines.append("\nDa controllare a mano:")
            lines += ["   " + t for t in s["conflitti"][:12]]
            if len(s["conflitti"]) > 12:
                lines.append("   … e altre %d" % (len(s["conflitti"]) - 12))
        lines.append("\nPer portare tutto sul telefono: File › Esporta dati per il telefono.")
        messagebox.showinfo(APP_NAME, "Importazione completata.\n\n" + "\n".join(lines))
        return True

    def import_pc_serials(self):
        paths = filedialog.askopenfilenames(
            title="Importa seriali dei PC da Excel",
            filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")])
        if not paths:
            return
        parts, errors = [], []
        for p in paths:
            try:
                parts.append(read_pc_xlsx(p))
            except ValueError as e:
                errors.append("%s: %s" % (os.path.basename(p), e))
        if errors:
            messagebox.showwarning(APP_NAME, "Questi file non li posso usare:\n\n" +
                                   "\n".join(errors))
        if not parts:
            return
        users = self.db.users_basic()
        plan = plan_pc_serials(users, merge_pc_entries(parts))
        s = plan["stats"]
        if not messagebox.askyesno(
                APP_NAME, "Nei file ci sono %d persone.\n\nScrivo il seriale del PC a %d persone "
                          "dell'app.\nGià scritto uguale: %d. Già scritto diverso (non lo "
                          "tocco): %d.\nNel file ma non nell'app (le ignoro, non le aggiungo): "
                          "%d.\nSenza PC nel file: %d.%s\n\nVuoi continuare?"
                % (s["nel_file"], s["scritti"], s["gia_uguali"], len(s["diversi"]),
                   s["non_in_app"], s["senza_pc"] + len(s["nome_non_seriale"]),
                   ("\nA %d seriali già scritti tolgo il codice dell'ente davanti (60016-)."
                    % s["prefisso_tolto"]) if s["prefisso_tolto"] else "")):
            return
        try:
            apply_pc_plan(self.db, plan)
        except (ValueError, sqlite3.Error) as e:
            messagebox.showerror(APP_NAME, "Importazione non riuscita:\n%s" % e)
            return
        self.refresh_users()
        lines = ["Seriali scritti: %d" % s["scritti"],
                 "Già uguali: %d; già diversi (non toccati): %d"
                 % (s["gia_uguali"], len(s["diversi"]))]
        if s["prefisso_tolto"]:
            lines.append("Seriali scritti con 60016- davanti, sistemati: %d" % s["prefisso_tolto"])
        lines += [
                  "Persone nel file ma non nell'app (ignorate): %d" % s["non_in_app"]]

        def more(title, items, limit=15):
            if items:
                lines.append("\n%s (%d):" % (title, len(items)))
                lines.extend("   " + t for t in items[:limit])
                if len(items) > limit:
                    lines.append("   … e altre %d" % (len(items) - limit))
        more("Il nome del PC non contiene un seriale (lasciato vuoto)", s["nome_non_seriale"])
        more("Seriale diverso da quello già scritto (non toccato)", s["diversi"])
        more("Più PC: ho scritto il primo", s["piu_pc"])
        more("Da controllare a mano", s["ambigui"])
        # chi nell'app resta senza seriale del PC (gli sportelli non sono persone)
        mancano = sorted(((u.get("ordine") or u["nome"]) for u in self.db.users_basic()
                          if not u.get("seriale_pc") and not u["nome"].casefold().startswith(
                              "sportello")), key=str.casefold)
        more("Nell'app senza seriale del PC: li scrivi tu a mano", mancano, 40)
        lines.append("\nPer portarli sul telefono: File › Esporta dati per il telefono.")
        messagebox.showinfo(APP_NAME, "Importazione completata.\n\n" + "\n".join(lines))

    def import_servicenow(self):
        if SN is None:
            messagebox.showerror(APP_NAME, "Manca il file servicenow.py accanto al programma.")
            return
        paths = filedialog.askopenfilenames(
            title="Importa interventi da ServiceNow",
            filetypes=[("Excel", "*.xlsx"), ("Tutti i file", "*.*")])
        if not paths:
            return
        rows, errors = [], []
        for p in paths:
            try:
                rows += SN.read_rows(p)
            except ValueError as e:
                errors.append("%s: %s" % (os.path.basename(p), e))
        if errors:
            messagebox.showwarning(APP_NAME, "Questi file non li posso usare:\n\n" +
                                   "\n".join(errors))
        if not rows:
            return
        try:
            tickets = SN.group_tickets(rows)
            plan = SN.plan_import(tickets, self.db.users_basic(), self.db.interventi_per_import(),
                                  self.db.sedi(), SN.technician_name(rows))
        except (ValueError, sqlite3.Error) as e:
            messagebox.showerror(APP_NAME, "Non riesco a preparare l'importazione:\n%s" % e)
            return
        s = plan["stats"]
        if not s["nuovi"] and not (plan["soluzioni"] or plan["riassegna"]):
            messagebox.showinfo(
                APP_NAME, "Non c'è niente da importare.\n\nTicket nel file: %d (da %d righe).\n"
                          "Già nell'app: %d. Annullati (saltati): %d."
                % (s["ticket"], s["righe_file"], s["gia_presenti"], s["annullati"]))
            return
        sedi = (" Sedi nuove da creare: %s." % ", ".join(plan["sedi_nuove"])
                if plan["sedi_nuove"] else "")
        if not messagebox.askyesno(
                APP_NAME,
                "Nel file ci sono %d righe, cioè %d ticket (uno per ticket, senza doppioni).\n\n"
                "Interventi nuovi da aggiungere: %d.\nGià nell'app: %d. Annullati (saltati): %d.\n\n"
                "Persone nuove da creare: %d.%s\nUtente preso da chi ha aperto il ticket: %d; "
                "dedotto da matricola, seriale o testo: %d; non riconosciuto (va su «%s»): %d.\n\n"
                "Vuoi continuare?"
                % (s["righe_file"], s["ticket"], s["nuovi"], s["gia_presenti"], s["annullati"],
                   len([u for u in plan["utenti_nuovi"] if u["nome"] != SN.DA_ASSEGNARE]), sedi,
                   s["da_richiedente"] + s["da_chiamante"], s["dedotti"], SN.DA_ASSEGNARE,
                   s["senza_utente"])):
            return
        try:
            SN.apply_plan(self.db, plan)
        except (ValueError, sqlite3.Error) as e:
            messagebox.showerror(APP_NAME, "Importazione non riuscita:\n%s" % e)
            return
        self.refresh_users()
        self.refresh_stats()
        self.apply_filters()
        report = ["Interventi aggiunti: %d" % s["nuovi"],
                  "Già nell'app (non toccati): %d" % s["gia_presenti"],
                  "Annullati (saltati): %d" % s["annullati"],
                  "Persone nuove create: %d" % len([u for u in plan["utenti_nuovi"]
                                                     if u["nome"] != SN.DA_ASSEGNARE]),
                  "Sedi nuove create: %s" % (", ".join(plan["sedi_nuove"]) or "nessuna"),
                  "Sede scritta a persone che non l'avevano: %d" % s["sede_scritta"],
                  "Persone con sede diversa da quella del ticket (lasciata com'era): %d"
                  % s["sede_diversa"],
                  "Utente da chi ha aperto il ticket: %d" % (s["da_richiedente"] + s["da_chiamante"]),
                  "Utente dedotto: %d (%s)" % (s["dedotti"], ", ".join(
                      "%s %d" % (k, v) for k, v in sorted(s["per_come"].items())) or "-"),
                  "Utente non riconosciuto (su «%s»): %d" % (SN.DA_ASSEGNARE, s["senza_utente"])]
        if s["soluzioni_aggiunte"]:
            report.append("Soluzioni aggiunte a interventi già presenti: %d" % s["soluzioni_aggiunte"])
        if s["riassegnati"]:
            report.append("Interventi con utente corretto: %d" % s["riassegnati"])
        extra = []
        if plan["utenti_nuovi"]:
            extra.append("\nPersone create (controlla cognome e nome):")
            extra += ["   %s%s" % (u["nome"], " - %s" % u["sede"] if u["sede"] else "")
                      for u in plan["utenti_nuovi"]]
        if plan["controllare"]:
            extra.append("\nDa controllare a mano:")
            extra += ["   " + t for t in plan["controllare"]]
        path = os.path.join(data_dir(), "rapporto_importazione_servicenow.txt")
        saved = ""
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(report + extra) + "\n")
            saved = "\n\nIl rapporto completo è in:\n%s" % path
        except OSError:
            pass
        short = report + (["\nDa controllare a mano: %d (nel rapporto)" % len(plan["controllare"])]
                          if plan["controllare"] else [])
        if len(plan["utenti_nuovi"]):
            short.append("Persone nuove nel rapporto: %d" % len(plan["utenti_nuovi"]))
        short.append("\nPer portare tutto sul telefono: File › Esporta dati per il telefono.")
        messagebox.showinfo(APP_NAME, "Importazione completata.\n\n" + "\n".join(short) + saved)

    def open_data_folder(self):
        try:
            os.startfile(data_dir())  # solo Windows
        except Exception:
            messagebox.showinfo(APP_NAME, "I dati sono in:\n%s" % data_dir())

    def on_close(self):
        shutil.rmtree(staging_dir(), ignore_errors=True)
        try:
            self.db.close()
        finally:
            self.destroy()


def main():
    try:  # schermo nitido su Windows con scala 125%/150%
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()


if __name__ == "__main__":
    main()
