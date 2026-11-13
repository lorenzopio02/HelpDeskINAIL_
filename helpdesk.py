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
from tkinter import filedialog, messagebox, simpledialog, ttk

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
"""

SELECT_INTERVENTI = (
    "SELECT i.id, i.data_ora, u.nome AS utente, i.richiesta, i.soluzione, "
    "i.note, i.ticket FROM interventi i JOIN utenti u ON u.id = i.utente_id"
)


class Database:
    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        self.conn.commit()

    def close(self):
        self.conn.close()

    # --- utenti ---
    def users(self):
        rows = self.conn.execute("SELECT nome FROM utenti ORDER BY nome COLLATE NOCASE")
        return [r["nome"] for r in rows]

    def users_with_counts(self):
        rows = self.conn.execute(
            "SELECT u.id, u.nome, COUNT(i.id) AS n FROM utenti u "
            "LEFT JOIN interventi i ON i.utente_id = u.id "
            "GROUP BY u.id ORDER BY u.nome COLLATE NOCASE"
        )
        return [dict(r) for r in rows]

    def find_user(self, name):
        row = self.conn.execute(
            "SELECT id, nome FROM utenti WHERE nome = ?", (normalize_name(name),)
        ).fetchone()
        return dict(row) if row else None

    def _get_or_create_user(self, name):
        name = normalize_name(name)
        row = self.conn.execute("SELECT id FROM utenti WHERE nome = ?", (name,)).fetchone()
        if row:
            return row["id"], False
        cur = self.conn.execute("INSERT INTO utenti(nome) VALUES (?)", (name,))
        return cur.lastrowid, True

    def add_user(self, name):
        with self.conn:
            return self._get_or_create_user(name)

    def rename_user(self, user_id, new_name):
        with self.conn:
            self.conn.execute(
                "UPDATE utenti SET nome = ? WHERE id = ?", (normalize_name(new_name), user_id)
            )

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
    def add_intervento(self, data_ora, utente, richiesta, soluzione, note, ticket):
        with self.conn:
            uid, created = self._get_or_create_user(utente)
            cur = self.conn.execute(
                "INSERT INTO interventi(data_ora, utente_id, richiesta, soluzione, note, ticket) "
                "VALUES (?,?,?,?,?,?)",
                (data_ora, uid, richiesta, soluzione, note, ticket),
            )
        return cur.lastrowid, created

    def update_intervento(self, iid, data_ora, utente, richiesta, soluzione, note, ticket):
        with self.conn:
            uid, created = self._get_or_create_user(utente)
            self.conn.execute(
                "UPDATE interventi SET data_ora=?, utente_id=?, richiesta=?, soluzione=?, "
                "note=?, ticket=? WHERE id=?",
                (data_ora, uid, richiesta, soluzione, note, ticket, iid),
            )
        return created

    def delete_intervento(self, iid):
        with self.conn:
            self.conn.execute("DELETE FROM interventi WHERE id = ?", (iid,))

    def search(self, utente="", dal=None, al=None, ticket_mode="tutti", ticket="", testo=""):
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

    def backup_to(self, dest_path):
        dest = sqlite3.connect(dest_path)
        try:
            self.conn.backup(dest)
        finally:
            dest.close()


# --------------------------------------------------------------------------
# Esportazione Excel / PDF
# --------------------------------------------------------------------------
def export_excel(path, rows, filtri_desc):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "Interventi"
    headers = ["Data e ora", "Utente", "N. ticket", "Per cosa mi hanno chiamato",
               "Come ho risolto", "Note"]
    ws.append(headers)
    for r in rows:
        ws.append([fmt_dt(r["data_ora"]), r["utente"], r["ticket"],
                   r["richiesta"], r["soluzione"], r["note"]])

    head_fill = PatternFill("solid", fgColor=NAVY.lstrip("#"))
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for width, col in zip([17, 24, 14, 50, 50, 35], "ABCDEF"):
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.column == 3:
                cell.number_format = "@"
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions

    # Foglio riepilogo con il conteggio per utente
    rs = wb.create_sheet("Riepilogo")
    rs["A1"] = "Riepilogo interventi"
    rs["A1"].font = Font(bold=True, size=14, color=NAVY.lstrip("#"))
    rs["A2"] = "Generato il " + datetime.now().strftime("%d/%m/%Y %H:%M")
    rs["A3"] = "Filtri: " + filtri_desc
    rs["A4"] = "Totale interventi"
    rs["B4"] = len(rows)
    rs["A4"].font = rs["B4"].font = Font(bold=True)
    rs["A6"], rs["B6"] = "Utente", "Interventi"
    for cell in (rs["A6"], rs["B6"]):
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head_fill
    n = 7
    for nome, cnt in Counter(r["utente"] for r in rows).most_common():
        rs.cell(row=n, column=1, value=nome)
        rs.cell(row=n, column=2, value=cnt)
        n += 1
    rs.column_dimensions["A"].width = 34
    rs.column_dimensions["B"].width = 14
    wb.save(path)


def export_pdf(path, rows, filtri_desc):
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
    st_info = ParagraphStyle("i", parent=base["Normal"], fontSize=9, textColor=colors.HexColor(GREY),
                             leading=12)
    st_head = ParagraphStyle("h", parent=base["Normal"], fontName="Helvetica-Bold", fontSize=8.5,
                             textColor=colors.white, leading=10)
    st_cell = ParagraphStyle("c", parent=base["Normal"], fontSize=8, leading=10)

    def para(text, style=st_cell):
        text = (text or "").replace("’", "'").replace("‘", "'")
        return Paragraph(escape(text).replace("\n", "<br/>"), style)

    doc = SimpleDocTemplate(path, pagesize=landscape(A4), leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=12 * mm, bottomMargin=14 * mm,
                            title="Registro interventi Help Desk")
    width = landscape(A4)[0] - 30 * mm
    story = []

    # intestazione con logo a sinistra
    left = ""
    try:
        iw, ih = ImageReader(resource_path("logo.png")).getSize()
        lw = 48 * mm
        left = Image(resource_path("logo.png"), width=lw, height=lw * ih / iw)
    except Exception:
        left = Paragraph("<b>INAIL</b>", st_title)
    right = [
        Paragraph("Registro interventi Help Desk", st_title),
        Paragraph("Interventi: <b>%d</b> &nbsp;|&nbsp; Filtri: %s &nbsp;|&nbsp; Generato il %s"
                  % (len(rows), escape(filtri_desc), datetime.now().strftime("%d/%m/%Y %H:%M")),
                  st_info),
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
        story.append(Paragraph("Nessun intervento per i filtri selezionati.", st_info))
    else:
        data = [[para(h, st_head) for h in
                 ("Data e ora", "Utente", "Ticket", "Per cosa mi hanno chiamato",
                  "Come ho risolto", "Note")]]
        for r in rows:
            data.append([para(fmt_dt(r["data_ora"])), para(r["utente"]), para(r["ticket"]),
                         para(r["richiesta"]), para(r["soluzione"]), para(r["note"])])
        cols = [24 * mm, 32 * mm, 20 * mm, 70 * mm, 75 * mm, width - 221 * mm]
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
        canvas.drawString(15 * mm, 8 * mm, "Help Desk INAIL - Registro interventi")
        canvas.drawRightString(landscape(A4)[0] - 15 * mm, 8 * mm, "Pagina %d" % doc_.page)
        canvas.restoreState()

    doc.build(story, onFirstPage=footer, onLaterPages=footer)


# --------------------------------------------------------------------------
# Interfaccia
# --------------------------------------------------------------------------
NAV_KEYS = {"Up", "Down", "Left", "Right", "Return", "Escape", "Tab", "Home", "End", "Prior",
            "Next", "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Caps_Lock"}


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


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("%s - Registro interventi" % APP_NAME)
        self.geometry("1200x820")
        self.minsize(1040, 700)
        self.configure(bg=LIGHT)
        try:
            self.iconbitmap(resource_path("icon.ico"))
        except Exception:
            pass

        self.db = Database(os.path.join(data_dir(), "interventi.db"))
        self.edit_id = None
        self.current_rows = []
        self.rows_by_id = {}
        self.users_cache = []
        self.sort_col = "data"
        self.sort_rev = True
        self._filter_job = None
        self._status_job = None

        self._setup_style()
        self._build_menu()
        self._build_header()
        self._build_tabs()

        self.refresh_users()
        self.refresh_stats()
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
        self.nb.add(self.tab_new, text="Nuovo intervento")
        self.nb.add(self.tab_reg, text="Registro e filtri")
        self.nb.add(self.tab_users, text="Utenti")
        self._build_form()
        self._build_registry()
        self._build_users()

    # ---------------- scheda: nuovo intervento ----------------
    def _build_form(self):
        f = self.tab_new
        f.columnconfigure(1, weight=1)
        self.form_title = ttk.Label(f, text="Nuovo intervento", font=("Segoe UI", 14, "bold"))
        self.form_title.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        self.var_dt = tk.StringVar()
        self.var_user = tk.StringVar()
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
        uf.columnconfigure(0, weight=1)
        self.cb_user = ttk.Combobox(uf, textvariable=self.var_user, width=40)
        self.cb_user.grid(row=0, column=0, sticky="w")
        bind_autocomplete(self.cb_user, self.var_user, lambda: self.users_cache)
        self.cb_user.bind("<Return>", self.autocomplete_user)
        self.lbl_hint = ttk.Label(uf, text="", style="Hint.TLabel", wraplength=700, justify="left")
        self.lbl_hint.grid(row=1, column=0, sticky="w", pady=(3, 0))
        self.var_user.trace_add("write", self.update_user_hint)

        ttk.Label(f, text="N. ticket\n(facoltativo)").grid(row=3, column=0, sticky="nw", pady=6,
                                                           padx=(0, 14))
        ttk.Entry(f, textvariable=self.var_ticket, width=24).grid(row=3, column=1, sticky="w",
                                                                  pady=4)

        ttk.Label(f, text="Per cosa mi hanno\nchiamato *").grid(row=4, column=0, sticky="nw",
                                                               pady=6, padx=(0, 14))
        self.t_richiesta = make_text(f, 4)
        self.t_richiesta.grid(row=4, column=1, sticky="ew", pady=4)

        ttk.Label(f, text="Come ho risolto").grid(row=5, column=0, sticky="nw", pady=6,
                                                  padx=(0, 14))
        self.t_soluzione = make_text(f, 7)
        self.t_soluzione.grid(row=5, column=1, sticky="ew", pady=4)

        ttk.Label(f, text="Note").grid(row=6, column=0, sticky="nw", pady=6, padx=(0, 14))
        self.t_note = make_text(f, 3)
        self.t_note.grid(row=6, column=1, sticky="ew", pady=4)

        btns = ttk.Frame(f)
        btns.grid(row=7, column=1, sticky="w", pady=(14, 0))
        self.btn_save = ttk.Button(btns, text="Salva intervento", style="Save.TButton",
                                   command=self.save_intervento)
        self.btn_save.pack(side="left")
        self.btn_clear = ttk.Button(btns, text="Svuota", command=self.clear_form)
        self.btn_clear.pack(side="left", padx=(8, 0))
        self.lbl_status = ttk.Label(f, text="", foreground=GREEN, font=("Segoe UI", 10, "bold"))
        self.lbl_status.grid(row=8, column=1, sticky="w", pady=(10, 0))
        ttk.Label(f, text="* campi obbligatori. Scorciatoia: Ctrl+S per salvare.",
                  style="Hint.TLabel").grid(row=9, column=1, sticky="w", pady=(4, 0))

    def set_now(self):
        self.var_dt.set(datetime.now().strftime("%d/%m/%Y %H:%M"))

    def set_status(self, text, color=GREEN):
        self.lbl_status.configure(text=text, foreground=color)
        if self._status_job:
            self.after_cancel(self._status_job)
        self._status_job = self.after(6000, lambda: self.lbl_status.configure(text=""))

    def update_user_hint(self, *_):
        typed = self.var_user.get().strip()
        if not typed:
            self.lbl_hint.configure(text="", foreground=GREY)
            return
        norm = normalize_name(typed).casefold()
        if any(u.casefold() == norm for u in self.users_cache):
            self.lbl_hint.configure(text="Utente già in lista.", foreground=GREEN)
            return
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

    def clear_form(self):
        self.edit_id = None
        self.set_now()
        self.var_user.set("")
        self.var_ticket.set("")
        for w in (self.t_richiesta, self.t_soluzione, self.t_note):
            w.delete("1.0", "end")
        self.form_title.configure(text="Nuovo intervento")
        self.btn_save.configure(text="Salva intervento")
        self.btn_clear.configure(text="Svuota")
        self.cb_user.focus_set()

    def save_intervento(self):
        if self.nb.select() != str(self.tab_new):
            return
        dt = parse_datetime(self.var_dt.get())
        utente = normalize_name(self.var_user.get())
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
                                                    soluzione, note, ticket)
                msg = "Modifiche salvate."
            else:
                _, created = self.db.add_intervento(data_ora, utente, richiesta, soluzione,
                                                    note, ticket)
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
        self.var_f_dal = tk.StringVar()
        self.var_f_al = tk.StringVar()
        self.var_f_ticket_mode = tk.StringVar(value="Tutti")
        self.var_f_ticket = tk.StringVar()
        self.var_f_text = tk.StringVar()

        box = ttk.LabelFrame(t, text="Filtri", padding=10)
        box.grid(row=0, column=0, sticky="ew")

        r1 = ttk.Frame(box)
        r1.pack(fill="x")
        ttk.Label(r1, text="Utente").pack(side="left")
        self.cb_filter_user = ttk.Combobox(r1, textvariable=self.var_f_user, width=26)
        self.cb_filter_user.pack(side="left", padx=(6, 18))
        bind_autocomplete(self.cb_filter_user, self.var_f_user,
                          lambda: ["(Tutti)"] + self.users_cache)
        self.cb_filter_user.bind("<FocusIn>", lambda e: self.after_idle(
            lambda: self.cb_filter_user.select_range(0, "end")))
        ttk.Label(r1, text="Dal").pack(side="left")
        self.ent_dal = ttk.Entry(r1, textvariable=self.var_f_dal, width=11)
        self.ent_dal.pack(side="left", padx=(6, 12))
        ttk.Label(r1, text="Al").pack(side="left")
        self.ent_al = ttk.Entry(r1, textvariable=self.var_f_al, width=11)
        self.ent_al.pack(side="left", padx=(6, 18))
        ttk.Label(r1, text="Ticket").pack(side="left")
        ttk.Combobox(r1, textvariable=self.var_f_ticket_mode, width=12, state="readonly",
                     values=["Tutti", "Con ticket", "Senza ticket"]).pack(side="left", padx=(6, 18))
        ttk.Label(r1, text="N. ticket").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_f_ticket, width=12).pack(side="left", padx=(6, 0))

        r2 = ttk.Frame(box)
        r2.pack(fill="x", pady=(10, 0))
        ttk.Label(r2, text="Cerca nel testo").pack(side="left")
        ttk.Entry(r2, textvariable=self.var_f_text, width=34).pack(side="left", padx=(6, 18))
        ttk.Label(r2, text="Periodo veloce:", style="Hint.TLabel").pack(side="left")
        for label, kind in (("Oggi", "oggi"), ("7 giorni", "7g"), ("Questo mese", "mese"),
                            ("Quest'anno", "anno"), ("Tutto", "tutto")):
            ttk.Button(r2, text=label, command=lambda k=kind: self.set_period(k)).pack(
                side="left", padx=(6, 0))
        ttk.Button(r2, text="Azzera filtri", style="Primary.TButton",
                   command=self.reset_filters).pack(side="right")

        for v in (self.var_f_user, self.var_f_dal, self.var_f_al, self.var_f_ticket_mode,
                  self.var_f_ticket, self.var_f_text):
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
        self.columns = ("data", "utente", "ticket", "richiesta", "soluzione")
        self.col_titles = {"data": "Data e ora", "utente": "Utente", "ticket": "Ticket",
                           "richiesta": "Per cosa mi hanno chiamato", "soluzione": "Come ho risolto"}
        self.tree = ttk.Treeview(tf, columns=self.columns, show="headings", selectmode="browse")
        widths = {"data": 130, "utente": 170, "ticket": 90, "richiesta": 340, "soluzione": 340}
        for c in self.columns:
            self.tree.heading(c, text=self.col_titles[c], command=lambda c=c: self.sort_by(c))
            self.tree.column(c, width=widths[c], anchor="w", stretch=(c in ("richiesta", "soluzione")))
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
        self.detail = tk.Text(df, height=7, wrap="word", font=("Segoe UI", 10), bg=WHITE, fg=TEXT,
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
        today = date.today()
        if kind == "oggi":
            dal, al = today, today
        elif kind == "7g":
            dal, al = today - timedelta(days=6), today
        elif kind == "mese":
            dal, al = today.replace(day=1), today
        elif kind == "anno":
            dal, al = today.replace(month=1, day=1), today
        else:
            dal, al = None, None
        self.var_f_dal.set(fmt_date(dal) if dal else "")
        self.var_f_al.set(fmt_date(al) if al else "")

    def reset_filters(self):
        self.var_f_user.set("(Tutti)")
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
        dal_txt = self.var_f_dal.get().strip()
        al_txt = self.var_f_al.get().strip()
        dal = parse_date(dal_txt) if dal_txt else None
        al = parse_date(al_txt) if al_txt else None
        self.ent_dal.configure(foreground=RED if dal_txt and not dal else TEXT)
        self.ent_al.configure(foreground=RED if al_txt and not al else TEXT)
        mode = {"Con ticket": "con", "Senza ticket": "senza"}.get(
            self.var_f_ticket_mode.get(), "tutti")
        return {"utente": user, "dal": dal, "al": al, "ticket_mode": mode,
                "ticket": self.var_f_ticket.get().strip(), "testo": self.var_f_text.get().strip()}

    @staticmethod
    def describe_filters(f):
        parts = []
        if f["utente"]:
            parts.append("utente \"%s\"" % f["utente"])
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
                fmt_dt(r["data_ora"]), r["utente"], r["ticket"],
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
            self.detail.insert("end", "Ticket: ", "h")
            self.detail.insert("end", (r["ticket"] or "nessuno") + "\n\n")
            self.detail.insert("end", "Per cosa mi hanno chiamato\n", "h")
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
        self.edit_id = r["id"]
        self.var_dt.set(fmt_dt(r["data_ora"]))
        self.var_user.set(r["utente"])
        self.var_ticket.set(r["ticket"])
        for w, key in ((self.t_richiesta, "richiesta"), (self.t_soluzione, "soluzione"),
                       (self.t_note, "note")):
            w.delete("1.0", "end")
            w.insert("1.0", r[key])
        self.form_title.configure(text="Modifica intervento del %s" % fmt_dt(r["data_ora"]))
        self.btn_save.configure(text="Salva modifiche")
        self.btn_clear.configure(text="Annulla modifica")
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

    def export(self, kind):
        if not self.current_rows:
            messagebox.showinfo(APP_NAME, "Non ci sono interventi da esportare con i filtri "
                                          "attuali.")
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M")
        if kind == "xlsx":
            path = filedialog.asksaveasfilename(
                title="Esporta in Excel", defaultextension=".xlsx",
                initialfile="interventi_%s.xlsx" % stamp,
                filetypes=[("File Excel", "*.xlsx")])
        else:
            path = filedialog.asksaveasfilename(
                title="Esporta in PDF", defaultextension=".pdf",
                initialfile="interventi_%s.pdf" % stamp, filetypes=[("File PDF", "*.pdf")])
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
        if messagebox.askyesno(APP_NAME, "File creato:\n%s\n\nVuoi aprirlo ora?" % path):
            try:
                os.startfile(path)  # solo Windows
            except Exception:
                pass

    # ---------------- scheda: utenti ----------------
    def _build_users(self):
        u = self.tab_users
        u.columnconfigure(0, weight=1)
        u.rowconfigure(2, weight=1)
        self.var_new_user = tk.StringVar()
        self.var_u_search = tk.StringVar()

        top = ttk.Frame(u)
        top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="Aggiungi utente").pack(side="left")
        e = ttk.Entry(top, textvariable=self.var_new_user, width=34)
        e.pack(side="left", padx=(8, 8))
        e.bind("<Return>", lambda ev: self.add_user_manual())
        ttk.Button(top, text="Aggiungi", style="Save.TButton",
                   command=self.add_user_manual).pack(side="left")
        ttk.Label(u, text="Gli utenti si aggiungono da soli anche quando salvi un intervento "
                          "con un nome nuovo.", style="Hint.TLabel").grid(
            row=1, column=0, sticky="w", pady=(4, 10))

        mid = ttk.Frame(u)
        mid.grid(row=2, column=0, sticky="nsew")
        mid.columnconfigure(0, weight=1)
        mid.rowconfigure(1, weight=1)
        sf = ttk.Frame(mid)
        sf.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(sf, text="Cerca").pack(side="left")
        ttk.Entry(sf, textvariable=self.var_u_search, width=30).pack(side="left", padx=(8, 0))
        self.var_u_search.trace_add("write", lambda *_: self.fill_users_tree())

        self.users_tree = ttk.Treeview(mid, columns=("nome", "n"), show="headings",
                                       selectmode="browse")
        self.users_tree.heading("nome", text="Utente")
        self.users_tree.heading("n", text="Interventi")
        self.users_tree.column("nome", width=420, anchor="w")
        self.users_tree.column("n", width=110, anchor="center", stretch=False)
        usb = ttk.Scrollbar(mid, orient="vertical", command=self.users_tree.yview)
        self.users_tree.configure(yscrollcommand=usb.set)
        self.users_tree.grid(row=1, column=0, sticky="nsew")
        usb.grid(row=1, column=1, sticky="ns")
        self.users_tree.bind("<Double-1>", lambda e: self.show_user_interventions())

        bf = ttk.Frame(u)
        bf.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(bf, text="Vedi i suoi interventi", style="Primary.TButton",
                   command=self.show_user_interventions).pack(side="left")
        ttk.Button(bf, text="Rinomina / Unisci", command=self.rename_user).pack(
            side="left", padx=(8, 0))
        ttk.Button(bf, text="Elimina", style="Danger.TButton",
                   command=self.delete_user).pack(side="left", padx=(8, 0))

    def refresh_users(self):
        self.users_cache = self.db.users()
        self.cb_user.configure(values=self.users_cache)
        self.cb_filter_user.configure(values=["(Tutti)"] + self.users_cache)
        self.fill_users_tree()
        self.update_user_hint()

    def fill_users_tree(self):
        terms = self.var_u_search.get().casefold().split()
        self.users_tree.delete(*self.users_tree.get_children())
        for row in self.db.users_with_counts():
            if terms and not matches_all(terms, row["nome"]):
                continue
            self.users_tree.insert("", "end", iid=str(row["id"]), values=(row["nome"], row["n"]))

    def selected_user(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Seleziona prima un utente dall'elenco.")
            return None
        nome, n = self.users_tree.item(sel[0], "values")
        return {"id": int(sel[0]), "nome": nome, "n": int(n)}

    def add_user_manual(self):
        name = normalize_name(self.var_new_user.get())
        if not name:
            return
        _, created = self.db.add_user(name)
        self.var_new_user.set("")
        self.refresh_users()
        if created:
            self.set_status("Utente aggiunto alla lista: %s." % name)
        else:
            messagebox.showinfo(APP_NAME, "%s è già nella lista." % name)

    def show_user_interventions(self):
        user = self.selected_user()
        if not user:
            return
        self.reset_filters()
        self.var_f_user.set(user["nome"])
        self.nb.select(self.tab_reg)

    def rename_user(self):
        user = self.selected_user()
        if not user:
            return
        new = simpledialog.askstring(
            APP_NAME, "Nuovo nome per \"%s\".\nSe scrivi il nome di un altro utente già in lista, "
                      "i due utenti vengono uniti." % user["nome"],
            initialvalue=user["nome"], parent=self)
        new = normalize_name(new)
        if not new or new == user["nome"]:
            return
        other = self.db.find_user(new)
        if other and other["id"] != user["id"]:
            if not messagebox.askyesno(
                    APP_NAME, "\"%s\" esiste già.\nUnire i due utenti? I %d interventi di \"%s\" "
                              "passeranno a \"%s\"." % (other["nome"], user["n"], user["nome"],
                                                        other["nome"])):
                return
            self.db.merge_users(user["id"], other["id"])
        else:
            self.db.rename_user(user["id"], new)
        self.refresh_users()
        self.apply_filters()

    def delete_user(self):
        user = self.selected_user()
        if not user:
            return
        if user["n"]:
            messagebox.showinfo(APP_NAME, "\"%s\" ha %d interventi, quindi non si può eliminare.\n"
                                "Puoi rinominarlo o unirlo a un altro utente." % (user["nome"],
                                                                                 user["n"]))
            return
        if messagebox.askyesno(APP_NAME, "Eliminare l'utente \"%s\"?" % user["nome"]):
            self.db.delete_user(user["id"])
            self.refresh_users()

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
