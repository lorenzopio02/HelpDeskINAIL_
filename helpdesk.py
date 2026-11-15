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
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import uuid
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
    import verbali as VB          # verbali Word (trasferimento, consegna, assegnazione, censimento)
except ImportError:
    VB = None

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

SEDI_TOSCANA = ["Arezzo", "Carrara", "DR Toscana", "Firenze", "Grosseto", "Livorno", "Lucca",
                "Massa", "Pisa", "Pistoia", "Prato", "Siena"]

# (nome, colore a schermo, colore leggibile su carta bianca)
COLORI_TONER = [("Nero", "#1c1c1c", "#1c1c1c"), ("Azzurro", "#1fa3dc", "#0d7fb0"),
                ("Magenta", "#d6247f", "#c0176c"), ("Giallo", "#f2c500", "#a67c00"),
                ("Vaschetta", "#8a96a3", "#4f5d6b")]
TONER_NOMI = [c[0] for c in COLORI_TONER]
TONER_TESTO = {c[0]: c[2] for c in COLORI_TONER}

MODELLI_STAMPANTI = ["bizhub C360i", "bizhub C3351i"]

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
    ("sede", "Sede", True, 30, 16),
    ("telefono", "Telefono", True, 35, 18), ("seriale_pc", "Seriale PC", True, 45, 22),
    ("imei", "IMEI cellulare", True, 45, 20), ("n", "N. interventi", False, 22, 14),
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
        return "c3351i"
    return "c360i"

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


def staging_dir():
    """Cartella temporanea per le immagini incollate prima di salvare l'intervento."""
    return os.path.join(attachments_root(), "_temp")


# ---------------------------------------------------------------------------
# Verbali Word: impostazioni e cartelle
# ---------------------------------------------------------------------------
SETTINGS_DEFAULT = {"tecnico": "", "sede_citta": "Firenze", "indirizzo": "Via Delle Porte Nuove, 61",
                    "cap": "50144", "cartella_verbali": ""}


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


def templates_dir():
    """Cartella dei modelli Word dei verbali (si possono modificare a mano)."""
    return os.path.join(data_dir(), "Modelli verbali")


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
    try:
        firme = json.loads(r.get("dati") or "{}").get("_firme") or {}
    except ValueError:
        firme = {}
    nomi = [label for k, label in (("utente", "utente"), ("tecnico", "tecnico")) if firme.get(k)]
    return " + ".join(nomi) if nomi else ""


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
    note        TEXT NOT NULL DEFAULT ''
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
        for col in ("uid", "modificato"):
            if col not in self._columns("interventi"):
                self.conn.execute(
                    "ALTER TABLE interventi ADD COLUMN %s TEXT NOT NULL DEFAULT ''" % col)
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_int_uid ON interventi(uid)")
        for table in ("edifici", "stampanti", "toner"):
            if "uid" not in self._columns(table):
                self.conn.execute("ALTER TABLE %s ADD COLUMN uid TEXT NOT NULL DEFAULT ''" % table)
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
        rows = self.conn.execute("SELECT nome FROM utenti ORDER BY nome COLLATE NOCASE")
        return [r["nome"] for r in rows]

    def users_with_counts(self):
        rows = self.conn.execute(
            "SELECT u.id, u.nome, u.matricola, u.sede, u.telefono, u.seriale_pc, u.imei, "
            "u.modificato, "
            "COUNT(i.id) AS n, "
            "MAX(i.data_ora) AS ultimo "
            "FROM utenti u LEFT JOIN interventi i ON i.utente_id = u.id "
            "GROUP BY u.id ORDER BY u.nome COLLATE NOCASE"
        )
        return [dict(r) for r in rows]

    def find_user(self, name):
        row = self.conn.execute(
            "SELECT id, nome, matricola, sede, telefono, seriale_pc, imei, modificato "
            "FROM utenti WHERE nome = ?",
            (normalize_name(name),)
        ).fetchone()
        return dict(row) if row else None

    def find_user_by_matricola(self, matricola):
        m = (matricola or "").strip()
        if not m:
            return None
        row = self.conn.execute(
            "SELECT id, nome, matricola, sede, telefono, seriale_pc, imei, modificato "
            "FROM utenti WHERE matricola = ? COLLATE NOCASE ORDER BY nome LIMIT 1", (m,)
        ).fetchone()
        return dict(row) if row else None

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
        return cur.lastrowid, True

    def add_user(self, name, sede="", telefono="", seriale_pc="", imei="", matricola="",
                 modificato=None):
        with self.conn:
            return self._get_or_create_user(name, sede, telefono, seriale_pc, imei, matricola,
                                            update=False, modificato=modificato)

    def update_user(self, user_id, name, sede="", telefono="", seriale_pc="", imei="",
                    matricola="", modificato=None):
        with self.conn:
            self.conn.execute(
                "UPDATE utenti SET nome=?, sede=?, telefono=?, seriale_pc=?, imei=?, matricola=?, "
                "modificato=? WHERE id=?", (normalize_name(name), sede, telefono, seriale_pc, imei,
                                            matricola, modificato or now_utc_iso(), user_id))

    def merge_users(self, source_id, target_id):
        with self.conn:
            self.conn.execute(
                "UPDATE interventi SET utente_id = ? WHERE utente_id = ?", (target_id, source_id)
            )
            self.conn.execute("DELETE FROM utenti WHERE id = ?", (source_id,))

    def delete_user(self, user_id):
        n = self.conn.execute(
            "SELECT COUNT(*) FROM interventi WHERE utente_id = ?", (user_id,)
        ).fetchone()[0]
        if n:
            return False
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
        with self.conn:
            self.conn.execute("DELETE FROM stampanti WHERE edificio_id=?", (eid,))
            self.conn.execute("DELETE FROM edifici WHERE id=?", (eid,))

    # --- stampanti ---
    def printers_of(self, edificio_id):
        rows = self.conn.execute(
            "SELECT s.id, s.edificio_id, s.piano, s.nome, s.modello, s.ip, s.note, "
            "(SELECT t.colore || '|' || t.data_ora FROM toner t WHERE t.stampante_id = s.id "
            " ORDER BY t.data_ora DESC, t.id DESC LIMIT 1) AS ultimo, "
            "(SELECT COUNT(*) FROM manutenzioni m WHERE m.stampante_id = s.id "
            " AND m.stato = 'Richiesto') AS man_richieste "
            "FROM stampanti s WHERE s.edificio_id = ? ORDER BY s.piano DESC, s.id",
            (edificio_id,))
        return [dict(r) for r in rows]

    def all_printers(self):
        rows = self.conn.execute(
            "SELECT s.id, s.piano, s.nome, s.modello, e.nome AS edificio FROM stampanti s "
            "JOIN edifici e ON e.id = s.edificio_id "
            "ORDER BY e.nome COLLATE NOCASE, s.piano DESC, s.id")
        return [dict(r) for r in rows]

    def get_stampante(self, sid):
        row = self.conn.execute(
            "SELECT s.id, s.edificio_id, s.piano, s.nome, s.modello, s.ip, s.note, "
            "e.nome AS edificio, e.sede FROM stampanti s JOIN edifici e ON e.id = s.edificio_id "
            "WHERE s.id = ?", (sid,)).fetchone()
        return dict(row) if row else None

    def models(self):
        rows = self.conn.execute(
            "SELECT DISTINCT modello FROM stampanti WHERE modello <> '' ORDER BY modello")
        return [r[0] for r in rows]

    def add_stampante(self, edificio_id, piano, nome, modello, ip, note):
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO stampanti(edificio_id, piano, nome, modello, ip, note) "
                "VALUES (?,?,?,?,?,?)", (edificio_id, piano, nome, modello, ip, note))
        return cur.lastrowid

    def update_stampante(self, sid, piano, nome, modello, ip, note):
        with self.conn:
            self.conn.execute(
                "UPDATE stampanti SET piano=?, nome=?, modello=?, ip=?, note=? WHERE id=?",
                (piano, nome, modello, ip, note, sid))

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
    if not isinstance(data, dict) or data.get("tipo") not in ("helpdesk-telefono", "helpdesk-pc"):
        raise ValueError("Questo non è un file dati dell'app Help Desk.")
    res = {"utenti_nuovi": 0, "utenti_aggiornati": 0, "interventi_nuovi": 0,
           "interventi_aggiornati": 0, "interventi_gia_presenti": 0, "foto": 0, "saltati": 0,
           "toner_nuovi": 0, "toner_gia_presenti": 0, "toner_saltati": 0, "scorta_bassa": [],
           "manut_nuove": 0, "manut_aggiornate": 0, "manut_gia_presenti": 0, "manut_saltate": 0,
           "verbali_nuovi": 0, "verbali_aggiornati": 0, "verbali_gia_presenti": 0,
           "verbali_saltati": 0, "verbali_foto": 0}

    for u in data.get("utenti") or []:
        nome = normalize_name(_clean(u.get("nome")))
        if not nome:
            continue
        cur = db.find_user(nome)
        # un file di una versione vecchia dell'app non ha la matricola: in quel caso si tiene
        # quella già salvata
        matricola = _clean(u.get("matricola")) if "matricola" in u else (
            cur["matricola"] if cur else "")
        fields = (normalize_name(_clean(u.get("sede"))), _clean(u.get("telefono")),
                  _clean(u.get("seriale_pc")), _clean(u.get("imei")), matricola)
        mod = _clean(u.get("modificato"))
        if not cur:
            db.add_user(nome, *fields, modificato=mod or now_utc_iso())
            res["utenti_nuovi"] += 1
        elif mod and mod > (cur.get("modificato") or "") and \
                fields != (cur["sede"], cur["telefono"], cur["seriale_pc"], cur["imei"],
                           cur["matricola"]):
            db.update_user(cur["id"], cur["nome"], *fields, modificato=mod)
            res["utenti_aggiornati"] += 1

    if data.get("tipo") != "helpdesk-telefono":
        return res
    import_phone_toner(db, data.get("toner") or [], res)
    import_phone_manutenzioni(db, data.get("manutenzioni") or [], res)
    import_phone_verbali(db, data.get("verbali") or [], res)
    for it in data.get("interventi") or []:
        uid = _clean(it.get("uid"))
        utente = normalize_name(_clean(it.get("utente")))
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
        for name in os.listdir(folder):          # tolgo i file che sul telefono non ci sono piu'
            if name not in keep:
                try:
                    os.remove(os.path.join(folder, name))
                except OSError:
                    pass
        dati["_foto"] = foto_names
        dati["_firme"] = firme
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


def photo_data_url(path):
    """Foto -> testo (data URL) per il file del telefono, ridotta per non appesantirlo."""
    try:
        from PIL import Image, ImageOps
        with Image.open(path) as im:
            im = ImageOps.exif_transpose(im)
            im.thumbnail((PHONE_PHOTO_PX, PHONE_PHOTO_PX))
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


def build_phone_export(db, with_photos=False, settings=None):
    """Dati del PC per il telefono: utenti e interventi (con le foto, se richiesto).
    Gli interventi nati sul PC ricevono un identificativo (uid), così quando il telefono li
    rimanda indietro il programma li riconosce e non li duplica."""
    now = now_utc_iso()
    db.ensure_sync_ids()
    users = [{k: u[k] for k in ("nome", "matricola", "sede", "telefono", "seriale_pc", "imei",
                                "modificato")}
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
    stampanti = [{k: r[k] for k in ("uid", "edificio_uid", "piano", "nome", "modello", "ip", "note")}
                 for r in db.conn.execute(
                     "SELECT s.uid, e.uid AS edificio_uid, s.piano, s.nome, s.modello, s.ip, "
                     "s.note FROM stampanti s JOIN edifici e ON e.id = s.edificio_id "
                     "ORDER BY s.id").fetchall()]
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
    return ({"tipo": "helpdesk-pc", "versione": 5, "creato": now, "utenti": users,
             "interventi": items, "edifici": edifici, "stampanti": stampanti,
             "toner": toner, "manutenzioni": manutenzioni, "scorta": db.scorta_all(),
             "scorta_aggiornata": now, "verbali": verbali,
             "impostazioni_verbali": {k: _clean((settings or {}).get(k))
                                      for k in ("tecnico", "sede_citta", "indirizzo", "cap")}},
            n_foto + n_foto_v)


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
        item = {"uid": r["uid"], "tipo": r["tipo"], "dati": dati, "foto": [], "firme": {},
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
    """Scheda di un utente: nome, sede e dati della postazione."""

    FIELDS = (("Nome e cognome *", "nome"), ("Matricola aziendale", "matricola"),
              ("Sede", "sede"), ("Telefono", "telefono"),
              ("Seriale del PC", "seriale_pc"), ("IMEI cellulare aziendale", "imei"))

    def __init__(self, parent, user, sedi, on_save):
        super().__init__(parent, "Nuovo utente" if user is None else "Scheda utente")
        self.resizable(False, False)
        self.on_save = on_save
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        user = user or {}
        self.vars = {k: tk.StringVar(value=user.get(k, "")) for _, k in self.FIELDS}
        for i, (label, key) in enumerate(self.FIELDS):
            ttk.Label(body, text=label).grid(row=i, column=0, sticky="w", pady=5, padx=(0, 14))
            if key == "sede":
                w = ttk.Combobox(body, textvariable=self.vars[key], values=sedi, width=34)
                bind_autocomplete(w, self.vars[key], lambda: sedi)
            else:
                w = ttk.Entry(body, textvariable=self.vars[key], width=36)
            w.grid(row=i, column=1, sticky="ew", pady=5)
            if i == 0:
                w.focus_set()
        ttk.Label(body, style="Hint.TLabel", wraplength=420, justify="left",
                  text="Se cambi il nome con quello di un altro utente già in lista, i due "
                       "utenti vengono uniti.").grid(row=len(self.FIELDS), column=0,
                                                     columnspan=2, sticky="w", pady=(8, 0))
        btns = ttk.Frame(body)
        btns.grid(row=len(self.FIELDS) + 1, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))
        self.bind("<Return>", lambda e: self.save())

    def save(self):
        data = {k: v.get() for k, v in self.vars.items()}
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


class PrinterEditDialog(ModalDialog):
    def __init__(self, parent, printer, building, floor, models, on_save):
        super().__init__(parent, "Nuova stampante" if printer is None else "Modifica stampante")
        self.resizable(False, False)
        self.on_save = on_save
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

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="Piano").grid(row=0, column=0, sticky="w", pady=5, padx=(0, 14))
        ttk.Combobox(body, textvariable=self.v_piano, state="readonly", width=24,
                     values=[floor_label(n) for n in self.floors]).grid(
            row=0, column=1, sticky="w", pady=5)
        ttk.Label(body, text="Nome o posizione").grid(row=1, column=0, sticky="w", pady=5,
                                                     padx=(0, 14))
        e = ttk.Entry(body, textvariable=self.v_nome, width=40)
        e.grid(row=1, column=1, sticky="ew", pady=5)
        e.focus_set()
        ttk.Label(body, text="Modello").grid(row=2, column=0, sticky="w", pady=5, padx=(0, 14))
        ttk.Combobox(body, textvariable=self.v_modello, values=models, width=38).grid(
            row=2, column=1, sticky="ew", pady=5)
        ttk.Label(body, text="Indirizzo IP").grid(row=3, column=0, sticky="w", pady=5,
                                                  padx=(0, 14))
        ttk.Entry(body, textvariable=self.v_ip, width=24).grid(row=3, column=1, sticky="w",
                                                               pady=5)
        ttk.Label(body, text="Note").grid(row=4, column=0, sticky="w", pady=5, padx=(0, 14))
        ttk.Entry(body, textvariable=self.v_note, width=40).grid(row=4, column=1, sticky="ew",
                                                                 pady=5)
        ttk.Label(body, style="Hint.TLabel", wraplength=400, justify="left",
                  text="Nome o posizione, ad esempio \"corridoio ovest\" o \"stanza 12\". "
                       "Il modello sceglie la foto nella grafica.").grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(6, 0))
        btns = ttk.Frame(body)
        btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))

    def save(self):
        data = {"piano": self.labels.get(self.v_piano.get(), 0), "nome": self.v_nome.get(),
                "modello": self.v_modello.get(), "ip": self.v_ip.get(), "note": self.v_note.get()}
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
                 "Dove: %s - %s" % (p["edificio"], floor_label(p["piano"]))]
        if p["ip"]:
            lines.append("IP: %s" % p["ip"])
        if p["note"]:
            lines.append("Note: %s" % p["note"])
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

        PrinterEditDialog(self, self.printer, building, None, self.app.models_list(), on_save)

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
        btns = ttk.Frame(body)
        btns.grid(row=len(self.FIELDS) + 1, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Annulla", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Salva", style="Save.TButton", command=self.save).pack(
            side="right", padx=(0, 8))

    def save(self):
        self.on_save({k: v.get().strip() for k, v in self.vars.items()})
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

    def _combo(self, label, key, values, width=42, hint="", default=""):
        var = self.v[key] = tk.StringVar(value=self.ini.get(key, default))
        holder = []

        def build(fr):
            cb = ttk.Combobox(fr, textvariable=var, values=values, width=width)
            bind_autocomplete(cb, var, lambda: values)
            cb.pack(side="left")
            cb.bind("<MouseWheel>", self._wheel_on_combo)
            holder.append(cb)
            if hint:
                ttk.Label(fr, text=hint, style="Hint.TLabel").pack(side="left", padx=(8, 0))
        self._line(label, build)
        return var, holder[0]

    def _text(self, label, key, height, hint=""):
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

    def _user_picked(self, event=None):
        u = self._current_user()
        if not u:
            return
        if self.tipo in ("trasferimento", "assegnazione") and self.v[self.user_key].get() != u["nome"]:
            self.v[self.user_key].set(u["nome"])      # nome scritto come in lista
        if self.tipo == "trasferimento":
            self._set_if_empty("contatti", u["telefono"])
        elif self.tipo == "presa_in_carico":
            wanted = "%s (matricola %s)" % (u["nome"], u["matricola"]) if u["matricola"] else u["nome"]
            if self.v["utente"].get() != wanted:
                self.v["utente"].set(wanted)
        elif self.tipo == "assegnazione":
            self._set_if_empty("matricola", u["matricola"])
        self._device_hint()

    def _device_hint(self, *_):
        """Se l'apparato è un PC o un cellulare, propongo il seriale/IMEI scritto nella scheda utente."""
        u = self._current_user()
        if not u or "tipologia" not in self.v:
            return
        tip = self.v["tipologia"].get().casefold()
        if any(w in tip for w in PC_WORDS):
            self._set_if_empty("seriale", u["seriale_pc"])
        elif any(w in tip for w in PHONE_WORDS):
            self._set_if_empty("seriale", u["imei"])

    def _tipologia(self, label="Tipologia *"):
        var, _ = self._combo(label, "tipologia", VB.TIPOLOGIE, 28)
        var.trace_add("write", self._device_hint)

    # ----- i quattro tipi di verbale -----
    def _form_trasferimento(self):
        self._tipologia()
        self._entry("Marca / Modello", "marca_modello", 44)
        self._entry("Numero di serie", "seriale", 36)
        luoghi = self.app.verbale_luoghi()
        self._combo("Provenienza (struttura)", "provenienza", luoghi, 64)
        self._combo("Destinazione (struttura)", "destinazione", luoghi, 64)
        self._user_combo("Assegnatario", "assegnatario")
        self._entry("Contatti", "contatti", 36)
        self._motivi(VB.MOTIVI_TRASFERIMENTO)
        self._text("Note", "note", 3)
        self._entry("Data spedizione", "data_sped", 14, hint="se la lasci vuota uso la data del verbale")

    def _form_presa_in_carico(self):
        self._tipologia("Tipologia dispositivo *")
        self._entry("Marca e modello", "marca_modello", 44)
        self._entry("Numero seriale", "seriale", 36)
        self._user_combo("Utente assegnatario *", "utente")
        self._entry("N° INC", "inc", 24, hint="il ticket dell'incidente")
        self._entry("N° RIC", "ric", 24, hint="la richiesta")
        self._text("Elenco materiale consegnato", "materiale", 4,
                   hint="Una voce per riga (esempio: caricabatterie, cover).")
        self._motivi(VB.MOTIVI_PRESA_IN_CARICO)

    def _form_assegnazione(self):
        self._combo("Titolo", "titolo", TITOLI_ELENCO, 12,
                    hint="facoltativo: scrive «al dott.», «all'avv.», «alla dott.ssa»…")
        self._user_combo("Utente *", "utente")
        self._entry("Matricola", "matricola", 20)
        self._dotazione_box()
        self._text("Accessori inclusi", "accessori", 3,
                   hint="Una voce per riga (esempio: zaino porta notebook).")

    def _form_censimento(self):
        st = self.app.settings
        self._entry("Oggetto *", "oggetto", 62, hint="esempio: Dispositivi Microsoft Lumia 950")
        self._entry("Nel testo: «costituito da»", "descrizione", 62,
                    hint="esempio: dispositivi smartphone Microsoft Lumia")
        self._combo("Sede (città)", "sede_citta", self.app.sedi_cache, 26,
                    default=st.get("sede_citta", ""))
        self._entry("Indirizzo", "indirizzo", 40, default=st.get("indirizzo", ""))
        self._entry("CAP", "cap", 10, default=st.get("cap", ""))
        self._entry("Tecnico *", "tecnico", 36, default=st.get("tecnico", ""),
                    hint="chi redige il verbale")
        self._entry("Indicazioni ricevute da", "collega", 40,
                    hint="facoltativo: se lo lasci vuoto la frase sulla dismissione non compare")
        var, _ = self._combo("Titolo colonna codice", "col_codice", VB.COLONNE_CODICE, 20,
                             default="IMEI")
        var.trace_add("write", lambda *_: self._fill_apparati())
        self._apparati_box()

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
        self.v_marca = tk.StringVar(value=last.get("marca", ""))
        self.v_modello = tk.StringVar(value=last.get("modello", ""))
        self.v_codice = tk.StringVar()
        for label, var, width in (("Marca", self.v_marca, 16), ("Modello", self.v_modello, 16),
                                  ("Codice", self.v_codice, 24)):
            ttk.Label(add, text=label).pack(side="left", padx=(0, 4))
            e = ttk.Entry(add, textvariable=var, width=width)
            e.pack(side="left", padx=(0, 10))
            e.bind("<Return>", lambda ev: self._add_apparato())
            if var is self.v_codice:
                self.e_codice = e
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
        r1 = ttk.Frame(box)
        r1.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Label(r1, text="Oggetto").pack(side="left", padx=(0, 4))
        ttk.Combobox(r1, textvariable=self.dv["tipo"], values=VB.TIPI_DOTAZIONE, width=22).pack(
            side="left", padx=(0, 10))
        ttk.Label(r1, text="Modello").pack(side="left", padx=(0, 4))
        ttk.Entry(r1, textvariable=self.dv["modello"], width=24).pack(side="left", padx=(0, 10))
        ttk.Label(r1, text="S/N o IMEI").pack(side="left", padx=(0, 4))
        e = ttk.Entry(r1, textvariable=self.dv["codice"], width=22)
        e.pack(side="left")
        e.bind("<Return>", lambda ev: self._add_dot())
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
        if self.tipo == "assegnazione" and (self.dv["modello"].get().strip()
                                            or self.dv["codice"].get().strip()
                                            or self.dv["iccid"].get().strip()):
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
        if self.on_create(self.tipo, d, list(self.photos), self):
            self.destroy()


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
        m.add_separator()
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
        ttk.Entry(pf, textvariable=self.var_seriale, width=22).grid(row=0, column=3, padx=(6, 18))
        ttk.Label(pf, text="IMEI cellulare").grid(row=0, column=4, sticky="w")
        ttk.Entry(pf, textvariable=self.var_imei, width=22).grid(row=0, column=5, padx=(6, 18))
        ttk.Button(pf, text="Copia dati utente", command=self.copy_user_data).grid(row=0, column=6)
        ttk.Label(pf, style="Hint.TLabel",
                  text="Si compilano da soli se l'utente è già in lista; quello che scrivi qui "
                       "viene salvato con l'utente.").grid(row=1, column=0, columnspan=7,
                                                          sticky="w", pady=(3, 0))

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
        for label, var in (("Utente", self.var_user), ("Matricola", self.var_matricola),
                           ("Sede", self.var_sede),
                           ("Telefono", self.var_tel), ("Seriale PC", self.var_seriale),
                           ("IMEI cellulare", self.var_imei)):
            if var.get().strip():
                lines.append("%s: %s" % (label, var.get().strip()))
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
        ttk.Entry(sf, textvariable=self.var_u_search, width=30).pack(side="left", padx=(8, 0))
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
        cols = ("nome", "matricola", "sede", "telefono", "seriale", "imei", "n")
        self.users_tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("nome", "Utente", 230, True), ("matricola", "Matricola", 100, False),
                ("sede", "Sede", 110, False),
                ("telefono", "Telefono", 130, False), ("seriale", "Seriale PC", 170, False),
                ("imei", "IMEI cellulare", 170, False), ("n", "Interventi", 80, False)):
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

    def fill_users_tree(self):
        terms = self.var_u_search.get().casefold().split()
        sede = self.var_u_sede.get().strip()
        sede = "" if sede in ("", "(Tutte)") else sede.casefold()
        self.users_tree.delete(*self.users_tree.get_children())
        self.users_visible = []
        for u in self.users_data:
            if sede and sede != u["sede"].casefold():
                continue
            text = " ".join((u["nome"], u["matricola"], u["sede"], u["telefono"],
                             u["seriale_pc"], u["imei"]))
            if terms and not matches_all(terms, text):
                continue
            self.users_visible.append(u)
            self.users_tree.insert("", "end", iid=str(u["id"]), values=(
                u["nome"], u["matricola"], u["sede"], u["telefono"], u["seriale_pc"], u["imei"],
                u["n"]))

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
        UserDialog(self, None, self.sedi_cache, lambda data: self.save_user_card(None, data))

    def edit_user(self):
        user = self.selected_user()
        if user:
            UserDialog(self, user, self.sedi_cache, lambda data: self.save_user_card(user, data))

    def save_user_card(self, user, data):
        nome = normalize_name(data["nome"])
        if not nome:
            messagebox.showwarning(APP_NAME, "Scrivi il nome e cognome dell'utente.")
            return False
        fields = (normalize_name(data["sede"]), data["telefono"].strip(),
                  data["seriale_pc"].strip(), data["imei"].strip(), data["matricola"].strip())
        try:
            if user is None:
                _, created = self.db.add_user(nome, *fields)
                if not created:
                    messagebox.showinfo(APP_NAME, "%s è già nella lista." % nome)
                    return False
                self.set_status("Utente aggiunto alla lista: %s." % nome)
            else:
                other = self.db.find_user(nome)
                if other and other["id"] != user["id"]:
                    if not messagebox.askyesno(
                            APP_NAME, "\"%s\" esiste già.\nUnire i due utenti? I %d interventi di "
                                      "\"%s\" passeranno a \"%s\" (restano i dati già salvati di "
                                      "\"%s\")." % (other["nome"], user["n"], user["nome"],
                                                    other["nome"], other["nome"])):
                        return False
                    self.db.merge_users(user["id"], other["id"])
                else:
                    self.db.update_user(user["id"], nome, *fields)
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return False
        self.refresh_users()
        self.apply_filters()
        return True

    def show_user_interventions(self):
        user = self.selected_user()
        if not user:
            return
        self.reset_filters()
        self.var_f_user.set(user["nome"])
        self.nb.select(self.tab_reg)

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
        self.tab_manut = ttk.Frame(self.nb_print, padding=12)
        self.tab_scorta = ttk.Frame(self.nb_print, padding=12)
        self.nb_print.add(self.tab_building, text="Edificio e stampanti")
        self.nb_print.add(self.tab_toner, text="Cambi toner")
        self.nb_print.add(self.tab_manut, text="Interventi tecnico")
        self.nb_print.add(self.tab_scorta, text="Scorta toner")
        self._build_building_view()
        self._build_toner_log()
        self._build_manut_log()
        self._build_scorta()

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
        self.cb_building.bind("<<ComboboxSelected>>", lambda e: self.schedule_draw())
        ttk.Button(bar, text="Nuovo edificio…", style="Save.TButton",
                   command=self.new_building).pack(side="left")
        ttk.Button(bar, text="Modifica edificio…", command=self.edit_building).pack(
            side="left", padx=(8, 0))
        ttk.Button(bar, text="Elimina edificio", style="Danger.TButton",
                   command=self.delete_building).pack(side="left", padx=(8, 0))

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
                          "stampante» la metti sul piano giusto.",
                  style="Hint.TLabel").grid(row=2, column=0, sticky="w", pady=(6, 0))

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
            top, bottom = y, y + floor_h
            c.create_rectangle(x0, top, x1, bottom, fill="#efe7cf" if idx % 2 == 0 else "#e7ddc0",
                               outline="#7d7355", width=2)
            # finestre appena accennate dietro alle stampanti
            wx = x0 + label_w + 20
            while wx + 44 < x1 - 10:
                c.create_rectangle(wx, top + 14, wx + 44, top + floor_h - 14, fill="#dbe7f0",
                                   outline="#c5d3e0")
                wx += 96
            c.create_rectangle(x0, top, x0 + label_w, bottom, fill=NAVY, outline=NAVY)
            c.create_text(x0 + label_w / 2, top + floor_h / 2, text=floor_label(fl), fill=WHITE,
                          font=("Segoe UI", 12, "bold"), width=label_w - 14)

            px = x0 + label_w + 26
            for p in by_floor.get(fl, []):
                px += self._draw_printer(c, p, px, top + 10) + 28

            tag = "add%d" % fl
            bx1, bx0 = x1 - 14, x1 - 184
            cy = top + floor_h / 2
            c.create_rectangle(bx0, cy - 17, bx1, cy + 17, fill=GREEN, outline=GREEN, tags=(tag,))
            c.create_text((bx0 + bx1) / 2, cy, text="+ Aggiungi stampante", fill=WHITE,
                          font=("Segoe UI", 9, "bold"), tags=(tag,))
            c.tag_bind(tag, "<Button-1>", lambda e, fl=fl: self.add_printer(fl))
            c.tag_bind(tag, "<Enter>", lambda e: c.configure(cursor="hand2"))
            c.tag_bind(tag, "<Leave>", lambda e: c.configure(cursor=""))
            y = bottom
        c.create_rectangle(x0 - 34, y, x1 + 34, y + 12, fill=GREEN, outline=GREEN)
        c.configure(scrollregion=(0, 0, width, y + 30))

    def _draw_printer(self, c, p, px, py):
        """Disegna una stampante con le scritte sotto; restituisce la larghezza occupata."""
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

        lines = [(p["nome"] or p["modello"] or "Stampante", ("Segoe UI", 9, "bold"), NAVY)]
        if p["nome"] and p["modello"]:
            lines.append((p["modello"], ("Segoe UI", 8), GREY))
        if p["ultimo"]:
            colore, when = p["ultimo"].split("|", 1)
            lines.append(("ultimo %s: %s" % (
                "cambio" if colore == "Vaschetta" else "toner", "%s %s" % (
                    colore, fmt_dt(when)[:5])), ("Segoe UI", 8), TONER_TESTO.get(colore, TEXT)))
        if p.get("man_richieste"):
            lines.append(("tecnico richiesto (%d)" % p["man_richieste"], ("Segoe UI", 8, "bold"),
                          RED))
        y = py + ih + 4
        for text, font, color in lines:
            item = c.create_text(cx, y, text=text, anchor="n", width=slot, font=font, fill=color,
                                 justify="center", tags=(tag,))
            box = c.bbox(item)
            y = box[3] + 2 if isinstance(box, tuple) and len(box) == 4 else y + 16
        c.tag_bind(tag, "<Button-1>", lambda e, pid=p["id"]: self.open_printer(pid))
        c.tag_bind(tag, "<Enter>", lambda e: c.configure(cursor="hand2"))
        c.tag_bind(tag, "<Leave>", lambda e: c.configure(cursor=""))
        return slot

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
                          lambda data: self.save_printer(None, b["id"], data))

    def save_printer(self, printer_id, building_id, data):
        nome = " ".join(data["nome"].split())
        modello = " ".join(data["modello"].split())
        if not nome and not modello:
            messagebox.showwarning(APP_NAME, "Scrivi almeno il nome o il modello della stampante.")
            return False
        args = (data["piano"], nome, modello, data["ip"].strip(), data["note"].strip())
        try:
            if printer_id is None:
                self.db.add_stampante(building_id, *args)
            else:
                self.db.update_stampante(printer_id, *args)
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
                ("foto", "Foto", 45, False), ("firme", "Firme", 110, False),
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
                            firme=firme, avvisi=avvisi)
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
        ini.pop("_firme", None)
        if r["tipo"] not in VB.TIPI:
            messagebox.showinfo(APP_NAME, "Tipo di verbale non riconosciuto.")
            return
        VerbaleDialog(self, r["tipo"], ini,
                      lambda tipo, d, photos, parent: self.make_verbale(tipo, d, photos, parent, r))

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
        lines = ["Interventi nuovi: %d" % res["interventi_nuovi"],
                 "Interventi aggiornati: %d" % res["interventi_aggiornati"],
                 "Già presenti (saltati): %d" % res["interventi_gia_presenti"],
                 "Foto importate: %d" % res["foto"],
                 "Utenti nuovi: %d, aggiornati: %d" % (res["utenti_nuovi"], res["utenti_aggiornati"])]
        if res["toner_nuovi"] or res["toner_gia_presenti"] or res["toner_saltati"]:
            lines.append("Cambi toner nuovi: %d (già presenti: %d)"
                         % (res["toner_nuovi"], res["toner_gia_presenti"]))
        if res["toner_saltati"]:
            lines.append("Cambi toner ignorati (stampante non trovata): %d" % res["toner_saltati"])
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
        if res["saltati"]:
            lines.append("Righe non valide ignorate: %d" % res["saltati"])
        if res["scorta_bassa"]:
            lines.append("\nATTENZIONE, scorta toner bassa: " + "; ".join(res["scorta_bassa"]))
        if res["manut_nuove"] or res["manut_aggiornate"]:
            self.after_manut_change()
        if res["toner_nuovi"]:
            self.after_toner_change()
        messagebox.showinfo(APP_NAME, "Importazione completata.\n\n" + "\n".join(lines))

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
