# -*- coding: utf-8 -*-
"""
Help Desk INAIL - Registro interventi
Registra gli interventi di help desk (utente, richiesta, soluzione, note, ticket),
con filtri, conteggio, elenco utenti che si costruisce da solo ed esportazione
in Excel/PDF.
"""
import os
import re
import sqlite3
import sys
from collections import Counter
from datetime import date, datetime, timedelta

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from printer_images import IMAGES as PRINTER_IMAGES
except ImportError:
    PRINTER_IMAGES = {}

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

SEDI_TOSCANA = ["Arezzo", "Carrara", "DR Toscana", "Firenze", "Grosseto", "Livorno", "Lucca", "Massa",
                "Pisa", "Pistoia", "Prato", "Siena"]

# (nome, colore a schermo, colore leggibile su carta bianca)
COLORI_TONER = [("Nero", "#1c1c1c", "#1c1c1c"), ("Azzurro", "#1fa3dc", "#0d7fb0"),
                ("Magenta", "#d6247f", "#c0176c"), ("Giallo", "#f2c500", "#a67c00")]
TONER_NOMI = [c[0] for c in COLORI_TONER]
TONER_TESTO = {c[0]: c[2] for c in COLORI_TONER}

MODELLI_STAMPANTI = ["bizhub C360i", "bizhub C3351i"]


def floor_label(n):
    if n == 0:
        return "Piano terra"
    if n == -1:
        return "Seminterrato"
    if n < -1:
        return "Interrato %d" % (-n)
    return "%d° piano" % n


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
"""

SELECT_INTERVENTI = (
    "SELECT i.id, i.data_ora, u.nome AS utente, i.sede, i.richiesta, i.soluzione, "
    "i.note, i.ticket, u.telefono, u.seriale_pc, u.imei "
    "FROM interventi i JOIN utenti u ON u.id = i.utente_id"
)

SELECT_TONER = (
    "SELECT t.id, t.data_ora, t.colore, t.note, s.id AS stampante_id, s.nome AS stampante, "
    "s.modello, s.piano, e.id AS edificio_id, e.nome AS edificio, e.sede "
    "FROM toner t JOIN stampanti s ON s.id = t.stampante_id "
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
        for col in ("sede", "telefono", "seriale_pc", "imei"):
            if col not in self._columns("utenti"):
                self.conn.execute(
                    "ALTER TABLE utenti ADD COLUMN %s TEXT NOT NULL DEFAULT ''" % col)
        if "sede" not in self._columns("interventi"):
            self.conn.execute("ALTER TABLE interventi ADD COLUMN sede TEXT NOT NULL DEFAULT ''")
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
            "SELECT u.id, u.nome, u.sede, u.telefono, u.seriale_pc, u.imei, COUNT(i.id) AS n "
            "FROM utenti u LEFT JOIN interventi i ON i.utente_id = u.id "
            "GROUP BY u.id ORDER BY u.nome COLLATE NOCASE"
        )
        return [dict(r) for r in rows]

    def find_user(self, name):
        row = self.conn.execute(
            "SELECT id, nome, sede, telefono, seriale_pc, imei FROM utenti WHERE nome = ?",
            (normalize_name(name),)
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
                            update=True):
        name = normalize_name(name)
        row = self.conn.execute("SELECT id FROM utenti WHERE nome = ?", (name,)).fetchone()
        if row:
            if update:
                self.conn.execute(
                    "UPDATE utenti SET sede=?, telefono=?, seriale_pc=?, imei=? WHERE id=?",
                    (sede, telefono, seriale_pc, imei, row["id"]))
            return row["id"], False
        cur = self.conn.execute(
            "INSERT INTO utenti(nome, sede, telefono, seriale_pc, imei) VALUES (?,?,?,?,?)",
            (name, sede, telefono, seriale_pc, imei))
        return cur.lastrowid, True

    def add_user(self, name, sede="", telefono="", seriale_pc="", imei=""):
        with self.conn:
            return self._get_or_create_user(name, sede, telefono, seriale_pc, imei, update=False)

    def update_user(self, user_id, name, sede="", telefono="", seriale_pc="", imei=""):
        with self.conn:
            self.conn.execute(
                "UPDATE utenti SET nome=?, sede=?, telefono=?, seriale_pc=?, imei=? WHERE id=?",
                (normalize_name(name), sede, telefono, seriale_pc, imei, user_id))

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
                       sede="", telefono="", seriale_pc="", imei=""):
        with self.conn:
            uid, created = self._get_or_create_user(utente, sede, telefono, seriale_pc, imei)
            cur = self.conn.execute(
                "INSERT INTO interventi(data_ora, utente_id, sede, richiesta, soluzione, note, "
                "ticket) VALUES (?,?,?,?,?,?,?)",
                (data_ora, uid, sede, richiesta, soluzione, note, ticket),
            )
        return cur.lastrowid, created

    def update_intervento(self, iid, data_ora, utente, richiesta, soluzione, note, ticket,
                          sede="", telefono="", seriale_pc="", imei=""):
        with self.conn:
            uid, created = self._get_or_create_user(utente, sede, telefono, seriale_pc, imei)
            self.conn.execute(
                "UPDATE interventi SET data_ora=?, utente_id=?, sede=?, richiesta=?, soluzione=?, "
                "note=?, ticket=? WHERE id=?",
                (data_ora, uid, sede, richiesta, soluzione, note, ticket, iid),
            )
        return created

    def delete_intervento(self, iid):
        with self.conn:
            self.conn.execute("DELETE FROM interventi WHERE id = ?", (iid,))

    def search(self, utente="", sede="", dal=None, al=None, ticket_mode="tutti", ticket="",
               testo=""):
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
        sql += " ORDER BY i.data_ora DESC, i.id DESC"
        rows = [dict(r) for r in self.conn.execute(sql, params)]

        terms = utente.casefold().split()
        if terms:
            rows = [r for r in rows if matches_all(terms, r["utente"])]
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
            " ORDER BY t.data_ora DESC, t.id DESC LIMIT 1) AS ultimo "
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

    # --- cambi toner ---
    def add_toner(self, stampante_id, colori, data_ora, note=""):
        with self.conn:
            for colore in colori:
                self.conn.execute(
                    "INSERT INTO toner(stampante_id, colore, data_ora, note) VALUES (?,?,?,?)",
                    (stampante_id, colore, data_ora, note))

    def delete_toner(self, tid):
        with self.conn:
            self.conn.execute("DELETE FROM toner WHERE id=?", (tid,))

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
        ["Data e ora", "Utente", "Sede", "N. ticket", "Per cosa mi hanno chiamato",
         "Come ho risolto", "Note"],
        [[fmt_dt(r["data_ora"]), r["utente"], r["sede"], r["ticket"], r["richiesta"],
          r["soluzione"], r["note"]] for r in rows],
        [17, 24, 14, 14, 50, 50, 35], text_cols=(4,))
    per_user = [[nome, cnt] for nome, cnt in Counter(r["utente"] for r in rows).most_common()]
    per_sede = [[(sede or "(non indicata)"), cnt]
                for sede, cnt in Counter(r["sede"] for r in rows).most_common()]
    _excel_summary(wb, "Riepilogo interventi", filtri_desc, "Totale interventi", len(rows),
                   [(["Utente", "Interventi"], per_user), (["Sede", "Interventi"], per_sede)])
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


def export_pdf(path, rows, filtri_desc):
    build_pdf(
        path, "Registro interventi Help Desk", filtri_desc, "Interventi",
        ["Data e ora", "Utente", "Sede", "Ticket", "Per cosa mi hanno chiamato",
         "Come ho risolto", "Note"],
        [[fmt_dt(r["data_ora"]), r["utente"], r["sede"], r["ticket"], r["richiesta"],
          r["soluzione"], r["note"]] for r in rows],
        [22, 30, 22, 18, 62, 66, 47], "Nessun intervento per i filtri selezionati.")


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

    FIELDS = (("Nome e cognome *", "nome"), ("Sede", "sede"), ("Telefono", "telefono"),
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

        form = ttk.LabelFrame(body, text="Registra un cambio toner", padding=10)
        form.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        r1 = ttk.Frame(form)
        r1.pack(fill="x")
        ttk.Label(r1, text="Colori cambiati:").pack(side="left")
        for nome, _, testo in COLORI_TONER:
            tk.Checkbutton(r1, text=nome, variable=self.colors[nome], bg=LIGHT, fg=testo,
                           activebackground=LIGHT, activeforeground=testo, selectcolor=WHITE,
                           font=("Segoe UI", 11, "bold")).pack(side="left", padx=(14, 0))
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
        ttk.Button(r3, text="Registra cambio toner", style="Save.TButton",
                   command=self.register).pack(side="left")
        self.lbl_status = ttk.Label(r3, text="", foreground=GREEN, font=("Segoe UI", 10, "bold"))
        self.lbl_status.pack(side="left", padx=12)

        hist = ttk.Frame(body)
        hist.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        hist.columnconfigure(0, weight=1)
        hist.rowconfigure(1, weight=1)
        ttk.Label(hist, text="Storico cambi toner di questa stampante",
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
        self.app.db.add_toner(self.printer_id, colori, dt.strftime("%Y-%m-%d %H:%M"),
                              self.var_note.get().strip())
        for v in self.colors.values():
            v.set(False)
        self.var_note.set("")
        self.lbl_status.configure(text="Registrato: %s - %s" % (", ".join(colori),
                                                               dt.strftime("%d/%m/%Y %H:%M")))
        self.set_now()
        self.refresh()
        self.app.after_toner_change()

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
        if messagebox.askyesno(APP_NAME, "Eliminare questa stampante?%s" % extra,
                               icon="warning", parent=self):
            self.app.db.delete_stampante(self.printer_id)
            self.app.after_printers_change()
            self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("%s - Registro interventi" % APP_NAME)
        self.geometry("1240x860")
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
        self.sort_col = "data"
        self.sort_rev = True
        self.buildings = []
        self.building_by_label = {}
        self.printer_imgs = {}
        self.toner_rows = []
        self.toner_rows_by_id = {}
        self.toner_printer_labels = {}
        self._filter_job = None
        self._status_job = None
        self._draw_job = None
        self._toner_job = None
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
        m.add_command(label="Salva una copia di backup…", command=self.backup)
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
        stats = tk.Frame(inner, bg=WHITE)
        stats.pack(side="right")
        for key, label in (("oggi", "Oggi"), ("mese", "Questo mese"), ("tot", "Totale")):
            box = tk.Frame(stats, bg=LIGHT, padx=16, pady=6, highlightthickness=1,
                           highlightbackground=BORDER)
            box.pack(side="left", padx=(8, 0))
            var = tk.StringVar(value="0")
            self.stat_vars[key] = var
            tk.Label(box, textvariable=var, font=("Segoe UI", 18, "bold"), fg=NAVY,
                     bg=LIGHT).pack()
            tk.Label(box, text=label, font=("Segoe UI", 9), fg=GREY, bg=LIGHT).pack()

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
        self.nb.add(self.tab_new, text="Nuovo intervento")
        self.nb.add(self.tab_reg, text="Registro e filtri")
        self.nb.add(self.tab_users, text="Utenti")
        self.nb.add(self.tab_print, text="Stampanti")
        self._build_form()
        self._build_registry()
        self._build_users()
        self._build_printers()

    # ---------------- scheda: nuovo intervento ----------------
    def _build_form(self):
        f = self.tab_new
        f.columnconfigure(1, weight=1)
        self.form_title = ttk.Label(f, text="Nuovo intervento", font=("Segoe UI", 14, "bold"))
        self.form_title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))

        self.var_dt = tk.StringVar()
        self.var_user = tk.StringVar()
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
        self.cb_user = ttk.Combobox(uf, textvariable=self.var_user, width=34)
        self.cb_user.grid(row=0, column=0, sticky="w")
        bind_autocomplete(self.cb_user, self.var_user, lambda: self.users_cache)
        self.cb_user.bind("<Return>", self.autocomplete_user)
        ttk.Label(uf, text="Sede").grid(row=0, column=1, padx=(20, 6))
        self.cb_sede = ttk.Combobox(uf, textvariable=self.var_sede, width=20)
        self.cb_sede.grid(row=0, column=2, sticky="w")
        bind_autocomplete(self.cb_sede, self.var_sede, lambda: self.sedi_cache)
        self.lbl_hint = ttk.Label(uf, text="", style="Hint.TLabel", wraplength=760, justify="left")
        self.lbl_hint.grid(row=1, column=0, columnspan=3, sticky="w", pady=(3, 0))
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
        ttk.Entry(f, textvariable=self.var_ticket, width=24).grid(row=4, column=1, sticky="w",
                                                                  pady=4)

        ttk.Label(f, text="Per cosa mi hanno\nchiamato *").grid(row=5, column=0, sticky="nw",
                                                               pady=6, padx=(0, 14))
        self.t_richiesta = make_text(f, 4)
        self.t_richiesta.grid(row=5, column=1, sticky="ew", pady=4)

        ttk.Label(f, text="Come ho risolto").grid(row=6, column=0, sticky="nw", pady=6,
                                                  padx=(0, 14))
        self.t_soluzione = make_text(f, 6)
        self.t_soluzione.grid(row=6, column=1, sticky="ew", pady=4)

        ttk.Label(f, text="Note").grid(row=7, column=0, sticky="nw", pady=6, padx=(0, 14))
        self.t_note = make_text(f, 3)
        self.t_note.grid(row=7, column=1, sticky="ew", pady=4)

        btns = ttk.Frame(f)
        btns.grid(row=8, column=1, sticky="w", pady=(14, 0))
        self.btn_save = ttk.Button(btns, text="Salva intervento", style="Save.TButton",
                                   command=self.save_intervento)
        self.btn_save.pack(side="left")
        self.btn_clear = ttk.Button(btns, text="Svuota", command=self.clear_form)
        self.btn_clear.pack(side="left", padx=(8, 0))
        self.lbl_status = ttk.Label(f, text="", foreground=GREEN, font=("Segoe UI", 10, "bold"))
        self.lbl_status.grid(row=9, column=1, sticky="w", pady=(10, 0))
        ttk.Label(f, text="* campi obbligatori. Scorciatoia: Ctrl+S per salvare.",
                  style="Hint.TLabel").grid(row=10, column=1, sticky="w", pady=(4, 0))

    def set_now(self):
        self.var_dt.set(datetime.now().strftime("%d/%m/%Y %H:%M"))

    def set_status(self, text, color=GREEN):
        self.lbl_status.configure(text=text, foreground=color)
        if self._status_job:
            self.after_cancel(self._status_job)
        self._status_job = self.after(6000, lambda: self.lbl_status.configure(text=""))

    def _set_station(self, sede="", tel="", seriale="", imei=""):
        self.var_sede.set(sede)
        self.var_tel.set(tel)
        self.var_seriale.set(seriale)
        self.var_imei.set(imei)

    def update_user_hint(self, *_):
        typed = self.var_user.get().strip()
        key = key_name(typed)
        match = self.users_by_key.get(key) if typed else None
        if not self._loading_form:
            if match and self._autofilled_for != key:
                self._autofilled_for = key
                self._set_station(match["sede"], match["telefono"], match["seriale_pc"],
                                  match["imei"])
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

    def copy_user_data(self):
        lines = []
        for label, var in (("Utente", self.var_user), ("Sede", self.var_sede),
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

    def clear_form(self):
        self._loading_form = True
        self.edit_id = None
        self._autofilled_for = None
        self.set_now()
        self.var_user.set("")
        self._set_station()
        self.var_ticket.set("")
        for w in (self.t_richiesta, self.t_soluzione, self.t_note):
            w.delete("1.0", "end")
        self._loading_form = False
        self.form_title.configure(text="Nuovo intervento")
        self.btn_save.configure(text="Salva intervento")
        self.btn_clear.configure(text="Svuota")
        self.update_user_hint()
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
                                                    imei)
                msg = "Modifiche salvate."
            else:
                _, created = self.db.add_intervento(data_ora, utente, richiesta, soluzione, note,
                                                    ticket, sede, tel, seriale, imei)
                msg = "Intervento salvato."
        except sqlite3.Error as e:
            messagebox.showerror(APP_NAME, "Non sono riuscito a salvare:\n%s" % e)
            return
        if created:
            msg += " Nuovo utente aggiunto alla lista: %s." % utente
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

        for v in (self.var_f_user, self.var_f_sede, self.var_f_dal, self.var_f_al,
                  self.var_f_ticket_mode, self.var_f_ticket, self.var_f_text):
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
        self.columns = ("data", "utente", "sede", "ticket", "richiesta", "soluzione")
        self.col_titles = {"data": "Data e ora", "utente": "Utente", "sede": "Sede",
                           "ticket": "Ticket", "richiesta": "Per cosa mi hanno chiamato",
                           "soluzione": "Come ho risolto"}
        self.tree = ttk.Treeview(tf, columns=self.columns, show="headings", selectmode="browse")
        widths = {"data": 125, "utente": 150, "sede": 95, "ticket": 80, "richiesta": 320,
                  "soluzione": 320}
        for c in self.columns:
            self.tree.heading(c, text=self.col_titles[c], command=lambda c=c: self.sort_by(c))
            self.tree.column(c, width=widths[c], anchor="w",
                             stretch=(c in ("richiesta", "soluzione")))
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
        dsb = ttk.Scrollbar(df, orient="vertical", command=self.detail.yview)
        self.detail.configure(yscrollcommand=dsb.set)
        self.detail.grid(row=0, column=0, sticky="nsew")
        dsb.grid(row=0, column=1, sticky="ns")

        bf = ttk.Frame(t)
        bf.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Modifica", command=self.edit_selected).pack(side="left")
        ttk.Button(bf, text="Elimina", style="Danger.TButton",
                   command=self.delete_selected).pack(side="left", padx=(8, 0))
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
        return {"utente": user, "sede": sede, "dal": dal, "al": al, "ticket_mode": mode,
                "ticket": self.var_f_ticket.get().strip(), "testo": self.var_f_text.get().strip()}

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
            "sede": lambda r: r["sede"].casefold(),
            "ticket": lambda r: r["ticket"].casefold(),
            "richiesta": lambda r: r["richiesta"].casefold(),
            "soluzione": lambda r: r["soluzione"].casefold(),
        }
        self.current_rows.sort(key=keys[self.sort_col], reverse=self.sort_rev)
        selected = self.tree.selection()
        self.tree.delete(*self.tree.get_children())
        self.rows_by_id = {}
        for r in self.current_rows:
            self.rows_by_id[str(r["id"])] = r
            self.tree.insert("", "end", iid=str(r["id"]), values=(
                fmt_dt(r["data_ora"]), r["utente"], r["sede"], r["ticket"],
                one_line(r["richiesta"]), one_line(r["soluzione"])))
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
            self.detail.insert("end", "Sede: ", "h")
            self.detail.insert("end", (r["sede"] or "-") + "     ")
            self.detail.insert("end", "Ticket: ", "h")
            self.detail.insert("end", (r["ticket"] or "nessuno") + "\n")
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

    def edit_selected(self):
        r = self.selected_row()
        if not r:
            messagebox.showinfo(APP_NAME, "Seleziona prima un intervento dall'elenco.")
            return
        self._loading_form = True
        self.edit_id = r["id"]
        self.var_dt.set(fmt_dt(r["data_ora"]))
        self.var_user.set(r["utente"])
        self._set_station(r["sede"], r["telefono"], r["seriale_pc"], r["imei"])
        self._autofilled_for = key_name(r["utente"])
        self.var_ticket.set(r["ticket"])
        for w, key in ((self.t_richiesta, "richiesta"), (self.t_soluzione, "soluzione"),
                       (self.t_note, "note")):
            w.delete("1.0", "end")
            w.insert("1.0", r[key])
        self._loading_form = False
        self.form_title.configure(text="Modifica intervento del %s" % fmt_dt(r["data_ora"]))
        self.btn_save.configure(text="Salva modifiche")
        self.btn_clear.configure(text="Annulla modifica")
        self.update_user_hint()
        self.nb.select(self.tab_new)

    def delete_selected(self):
        r = self.selected_row()
        if not r:
            return
        if not messagebox.askyesno(APP_NAME, "Eliminare questo intervento?\n\n%s - %s\n%s" % (
                fmt_dt(r["data_ora"]), r["utente"], one_line(r["richiesta"], 80)), icon="warning"):
            return
        self.db.delete_intervento(r["id"])
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

        mid = ttk.Frame(u)
        mid.grid(row=2, column=0, sticky="nsew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(0, weight=1)
        cols = ("nome", "sede", "telefono", "seriale", "imei", "n")
        self.users_tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        for c, title, width, stretch in (
                ("nome", "Utente", 230, True), ("sede", "Sede", 110, False),
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

    def refresh_users(self):
        self.users_data = self.db.users_with_counts()
        self.users_cache = [u["nome"] for u in self.users_data]
        self.users_by_key = {key_name(u["nome"]): u for u in self.users_data}
        self.sedi_cache = self.db.sedi()
        self.cb_user.configure(values=self.users_cache)
        self.cb_sede.configure(values=self.sedi_cache)
        self.cb_filter_user.configure(values=["(Tutti)"] + self.users_cache)
        self.cb_filter_sede.configure(values=["(Tutte)"] + self.sedi_cache)
        self.fill_users_tree()
        self.update_user_hint()

    def fill_users_tree(self):
        terms = self.var_u_search.get().casefold().split()
        self.users_tree.delete(*self.users_tree.get_children())
        for u in self.users_data:
            text = " ".join((u["nome"], u["sede"], u["telefono"], u["seriale_pc"], u["imei"]))
            if terms and not matches_all(terms, text):
                continue
            self.users_tree.insert("", "end", iid=str(u["id"]), values=(
                u["nome"], u["sede"], u["telefono"], u["seriale_pc"], u["imei"], u["n"]))

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
                  data["seriale_pc"].strip(), data["imei"].strip())
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
        self.nb_print.add(self.tab_building, text="Edificio e stampanti")
        self.nb_print.add(self.tab_toner, text="Cambi toner")
        self._build_building_view()
        self._build_toner_log()

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
        ttk.Label(t, text="Clicca su una stampante per registrare il cambio toner e vedere lo "
                          "storico. Con «+ Aggiungi stampante» la metti sul piano giusto.",
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
            lines.append(("ultimo toner: %s %s" % (colore, fmt_dt(when)[:5]), ("Segoe UI", 8),
                          TONER_TESTO.get(colore, TEXT)))
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

    def after_toner_change(self):
        self.schedule_draw()
        self.apply_toner_filters()

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

    # ---------------- varie ----------------
    def refresh_stats(self):
        oggi, mese, tot = self.db.stats()
        self.stat_vars["oggi"].set(str(oggi))
        self.stat_vars["mese"].set(str(mese))
        self.stat_vars["tot"].set(str(tot))

    def backup(self):
        path = filedialog.asksaveasfilename(
            title="Salva una copia di backup", defaultextension=".db",
            initialfile="backup_interventi_%s.db" % datetime.now().strftime("%Y%m%d"),
            filetypes=[("Database", "*.db")])
        if not path:
            return
        try:
            self.db.backup_to(path)
        except Exception as e:
            messagebox.showerror(APP_NAME, "Backup non riuscito:\n%s" % e)
            return
        messagebox.showinfo(APP_NAME, "Copia salvata in:\n%s" % path)

    def open_data_folder(self):
        try:
            os.startfile(data_dir())  # solo Windows
        except Exception:
            messagebox.showinfo(APP_NAME, "I dati sono in:\n%s" % data_dir())

    def on_close(self):
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
