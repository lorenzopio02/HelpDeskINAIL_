# -*- coding: utf-8 -*-
"""
Importazione dei ticket di ServiceNow (un file Excel esportato dall'elenco) come interventi.

Una riga del file puo' essere un ticket o un «passaggio» di un ticket tra gruppi: le righe con
lo stesso numero di ticket (RIC..., INC...) diventano un solo intervento. Questo file non usa
tkinter e non tocca il database: prepara un piano (plan_import) che poi si applica.
"""
import json
import re
import unicodedata
import uuid
from datetime import datetime, timezone

TICKET_RE = re.compile(r"^[A-Z]{2,6}\d{5,9}$")
TICKET_IN_TEXT_RE = re.compile(r"[A-Z]{2,6}\d{5,9}")
MATRICOLA_RE = re.compile(r"(?<![A-Za-z0-9])x\s?f\s?[-.]?\s?(\d{4,6})(?![0-9])", re.I)
SN_RE = re.compile(r"(?<![A-Za-z0-9])(?:SN|S/N|seriale|serial number)\s*[:=]?\s*([A-Z0-9]{5,15})"
                   r"(?![A-Za-z0-9])", re.I)
ENTE_SERIAL_RE = re.compile(r"(?<![A-Za-z0-9])\d{5}\s*-\s*([A-Z0-9]{5,12})(?![A-Za-z0-9])", re.I)
IMEI_RE = re.compile(r"IMEI\s*[:=]?\s*(\d{14,16})", re.I)
TITOLO_COGNOME_RE = re.compile(
    r"(?i:\b(?:dott\.?\s?ssa|dott\.?|dottoressa|dottore|sig\.?\s?ra|sig\.?|signor[ae]?|ing\.?|"
    r"avv\.?|prof\.?\s?ssa|prof\.?|dr\.?|d\.ssa|geom\.?|rag\.?|collega|ispettore|ispettrice))"
    r"\s+(?-i:([A-ZÀ-Ý][a-zà-ÿ']{3,}))")

SEGNO = "Importato da ServiceNow"
DEDOTTO = "utente dedotto"
DA_ASSEGNARE = "Da assegnare (ServiceNow)"

# nome della colonna (minuscolo) -> campo
INTESTAZIONI = {
    "ticket": ("task", "ticket", "numero ticket", "numero", "number"),
    "aperto": ("aperto", "opened", "data apertura", "data di apertura"),
    "apertoda": ("aperto da", "opened by"),
    "chiamante": ("chiamante", "richiesto per", "richiesto da", "utente interessato", "caller",
                  "requested for", "beneficiario"),
    "stato": ("stato", "state", "status"),
    "breve": ("breve descrizione", "short description"),
    "descr": ("descrizione", "description"),
    "sede": ("sede", "location", "ubicazione"),
    "ci": ("ci interessato", "elemento di configurazione", "configuration item"),
    "soluzione": ("note di chiusura", "note di risoluzione", "close notes", "resolution notes",
                  "soluzione", "note di risoluzione/chiusura"),
    "actor": ("actor",),
}
# nomi di battesimo italiani comuni: servono solo a capire se un nome scritto nel testo è
# «Nome Cognome» o «Cognome Nome» (di una persona che non è ancora nell'app)
NOMI_BATTESIMO = set("""
adele adriana adriano agnese agostino alba alberto alessandra alessandro alessia alessio alfredo alice
alida alina alvaro amalia amedeo andrea angela angelica angelo anita anna annalisa annamaria anselmo
antonella antonello antonia antonino antonio arianna armando arturo augusto aurora barbara beatrice
benedetta benedetto benito bernardo bianca brunella bruno carla carlo carmela carmelo carmine carolina
caterina cecilia cinzia claudia claudio clelia clemente concetta consuelo corrado cosimo cristian cristiana
cristina cristoforo damiano daniela daniele danilo dario davide debora deborah delia denise diana diego
dina dino domenica domenico donatella donato edoardo elena eleonora elia elisa elisabetta emanuela
emanuele emilia emilio emma enrica enrico enzo erica ernesto eugenio eva fabiana fabio fabrizia fabrizio
fausto federica federico felice ferdinando fernanda filippo fiorella flavia flavio francesca francesco
franco gabriella gabriele gaetano gaia gemma generoso gennaro giacomo gian gianfranco gianluca gianni
gianpaolo gina gino gioia giordano giorgia giorgio giovanna giovanni giulia giuliana giuliano giulio
giuseppe giuseppina gloria graziella grazia gregorio guglielmo guido gustavo ida ilaria irene iris
isabella italo ivan ivana ivano jessica katia laura lavinia leonardo leone letizia liana lidia lina
lisa livia loredana lorena lorenza lorenzo luana luca luciana luciano lucia lucio ludovica luigi luigia
luisa luisella manuel manuela marcella marcello marco margherita maria mariangela mariano marilena
marina mario marisa marta martina marzia massimiliano massimo matilde matteo maura maurizia maurizio
michela michele milena mirella miriam monica nadia nadine natalia natalina nathalie nicola nicoletta
nicole nino nives noemi nunzia olga omar oriana orlando oscar ottavia paola paolo pasquale patrizia
patrizio paulo pia piera piero pietro pina pio rachele raffaele raffaella raimondo renata renato
riccardo rita roberta roberto rocco romina romeo romolo rosa rosalba rosanna rosaria rosario rosaura
rosella rossana rossella ruggero sabrina salvatore samanta samantha samuele sandra sandro sara
saverio serena sergio silvana silvano silvia silvio simona simone sofia solange sonia stefania
stefano sveva tania teresa tiziana tiziano tommaso ugo umberto valentina valentino valeria valerio
vanessa vera veronica vincenzo vincenza viola virginia vita vito vittoria vittorio viviana walter
wanda ylenia yuri zaira iole fanny cristiano gabriel
""".split())

PARTICELLE = {"de", "di", "del", "della", "dei", "degli", "delle", "dello", "da", "dal", "dalla",
              "la", "lo", "le", "li", "van", "von", "mc", "san", "santa", "d", "dell", "dall"}


# ------------------------------------------------------------------------------ piccoli attrezzi
def _now_iso():
    n = datetime.now(timezone.utc)
    return n.strftime("%Y-%m-%dT%H:%M:%S.") + "%03dZ" % (n.microsecond // 1000)


def _s(v):
    return "" if v is None else str(v).strip()


def fold(text):
    """Minuscolo, senza accenti e senza apostrofi: «Paternò» e «PATERNO'» sono uguali."""
    t = unicodedata.normalize("NFKD", str(text or ""))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return t.casefold().replace("'", "").replace("’", "").replace("`", "")


def tokens(text):
    return re.findall(r"[a-z0-9]+", fold(text))


def words_key(name):
    return " ".join(sorted(tokens(name)))


def _accent_end(word):
    """«Paterno'» -> «Paternò» (l'apostrofo finale di ServiceNow è un accento)."""
    m = re.match(r"^(.*[A-Za-zÀ-ÿ])'$", word)
    if not m:
        return word
    end = {"a": "à", "e": "è", "i": "ì", "o": "ò", "u": "ù"}.get(m.group(1)[-1].lower())
    return m.group(1)[:-1] + end if end else word


def cognome_nome(name, incerto=False):
    """«NOME COGNOME» (come scrive ServiceNow) -> «Cognome Nome». L'ultima parola è il cognome
    (con le particelle davanti: «Della Polla»); chi non ha almeno due parole resta com'è.
    Con incerto=True (un nome scritto nel testo di un ticket) se la prima parola non è un nome di
    battesimo e l'ultima sì, il nome è già «Cognome Nome» e non si gira."""
    words = [_accent_end(w) for w in " ".join(str(name or "").split()).title().split()]
    if len(words) < 2:
        return " ".join(words)
    if incerto and fold(words[0]) not in NOMI_BATTESIMO and fold(words[-1]) in NOMI_BATTESIMO:
        return " ".join(words)
    i = len(words) - 1
    while i - 1 >= 1 and fold(words[i - 1]) in PARTICELLE:
        i -= 1
    return " ".join(words[i:] + words[:i])


def parse_when(v):
    """La data di apertura come «2026-11-17 12:04» (vuoto se non si capisce)."""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    t = _s(v)
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
                "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(t, fmt).strftime("%Y-%m-%d %H:%M")
        except ValueError:
            pass
    return ""


def map_stato(text):
    """Stato dell'app per uno stato di ServiceNow; None = ticket da non importare (annullato)."""
    t = fold(text)
    if t.startswith(("annull", "cancel", "rifiut")):
        return None
    if t.startswith(("chius", "risol", "closed", "resolved", "complet")):
        return "Risolto"
    if "attesa" in t or t.startswith(("hold", "pending", "sospes")):
        return "In attesa"
    return "Aperto"


def map_sede(text, known=()):
    """La sede dell'app per la sede di ServiceNow: «Firenze- Porte Nuove» -> «Firenze»,
    «Direzione Regionale Toscana» -> «DR Toscana»; le altre (Lucca, Ravenna...) restano col loro
    nome, e se non ci sono ancora nell'app si creano."""
    t = " ".join(str(text or "").split())
    if not t:
        return ""
    low = fold(t)
    m = re.match(r"direzione regionale\s+(?:per\s+(?:la|il|l)\s+)?(.+)$", low)
    if m:
        rest = t[len(t) - len(m.group(1)):] if t.casefold().endswith(m.group(1)) else m.group(1)
        return "DR " + _accent_end(" ".join(rest.title().split()))
    if "porte nuove" in low or "firenze" in low:
        return "Firenze"
    for k in known:
        if fold(k) == low:
            return k
    return " ".join(_accent_end(w) for w in t.split())


# ------------------------------------------------------------------------------ testo del ticket
_SALUTO_INIZIO = re.compile(
    r"^\s*(?:buon\s?giorno|buon\s?pomeriggio|buona\s?sera|buonasera|salve|ciao|gentili|gentile)"
    r"\b[\s,.!:;-]*", re.I)
_CHIUSURA = (r"(?:saluti\s+e\s+grazie|grazie\s+e\s+saluti|grazie(?:\s+mille)?(?:\s+e\s+(?:buona\s+giornata|"
             r"buon\s+lavoro|buona\s+serata|cordiali\s+saluti|distinti\s+saluti|saluti))?|molte\s+grazie|"
             r"cordiali\s+saluti|distinti\s+saluti|saluti|buona\s+giornata|buon\s+lavoro|a\s+presto)")
_CHIUSURA_FINE = re.compile(r"(?:^|(?<=[.!?:;,])\s+|\n+)" + _CHIUSURA + r"[\s.!,]*$", re.I)
_CHIUSURA_RIGA = re.compile(r"^" + _CHIUSURA + r"[\s.!,]*$", re.I)


_RIGA_TICKET = re.compile(r"^(?:n\.?\s?b\.?\s*)?(?:il\s+)?ticket\b.{0,40}\b(?:in\s+gestione|gestito|risolto|aperto)\b", re.I)


def telefono_tecnico(tickets):
    """Il numero che il tecnico scrive in fondo ai ticket («potete contattarmi al numero: ...»): il più frequente."""
    cnt = {}
    for t in tickets:
        for m in re.finditer(r"(?i)contattarmi[^\n]{0,60}?(\+?\d[\d ]{7,14}\d)", t.get("descr") or ""):
            d = re.sub(r"\D", "", m.group(1))
            cnt[d] = cnt.get(d, 0) + 1
    return max(cnt, key=cnt.get) if cnt else ""


def _intestazione_tecnico(firma):
    """«Ticket in gestione dal tecnico del presidio di Firenze, Mario Rossi.» e simili."""
    parole = [re.escape(w) for w in str(firma or "").split()]
    nome = r"\s+".join(parole) if parole else r"(?!)"
    return re.compile(
        r"^\W*(?:il\s+)?ticket\s+(?:è\s+)?(?:in\s+gestione|gestito|risolto)\b[^\n]{0,140}?"
        r"(?:" + nome + r"|presidio(?:\s+territoriale)?(?:\s+(?:di|a)\s+firenze)?|\b(?:di|a)\s+firenze)"
        r"(?:\s*[:,]?\s*" + nome + r")?[\s.:,;-]*", re.I)


def clean_text(text, firma="", tel=""):
    """Il testo del ticket pulito: senza l'intestazione «Ticket in gestione dal tecnico...», senza
    saluti di apertura e di chiusura né la firma del tecnico (firma = il suo nome), con la
    maiuscola iniziale e il punto finale. L'ortografia non si tocca: il testo resta quello scritto
    in ServiceNow."""
    t = _s(text).replace("\r", "\n").replace("\xa0", " ")
    lines = [" ".join(l.split()) for l in t.split("\n")]
    t = "\n".join(l for l in lines if l)
    m = _intestazione_tecnico(firma).match(t)
    if m:
        rest = t[m.end():].lstrip(" ")
        first, _, after = rest.partition("\n")
        if after and re.fullmatch(r"(?:[A-ZÀ-Ý][a-zà-ÿ']+\s?){1,4}\.?", first.strip()):
            rest = after                           # il nome del tecnico, scritto sulla riga dopo
        t = rest.lstrip()
    lines = t.split("\n")
    fk = words_key(firma) if firma else ""
    ftok = set(tokens(firma))
    keep = []
    for l in lines:
        d = re.sub(r"\D", "", l)
        if (_RIGA_TICKET.match(l) and (ftok and ftok <= set(tokens(l)) or re.search(r"(?i)tecnico|presidio", l))
                and len(l) < 200):
            continue                               # «N.B ticket aperto dal tecnico ... – Mario Rossi»
        if tel and tel in d and len(d) <= len(tel) + 2 and (
                re.search(r"(?i)contatt|numero|cell|tel|chiam", l) or words_key(re.sub(r"[\d+.]", "", l)) in ("", fk)):
            continue                               # «potete contattarmi al numero: ...» del tecnico
        keep.append(l)
    lines = keep

    def firma_o_telefono(l, prev):
        if fk and words_key(l) == fk:
            return True
        return bool(re.fullmatch(r"[+\d][\d\s./-]{6,}\.?", l)) and prev
    while lines:                                   # in fondo: saluti, nome del tecnico, suo telefono
        last = lines[-1]
        before = lines[-2] if len(lines) > 1 else ""
        prev_sig = bool(_CHIUSURA_RIGA.match(before)) or (fk and words_key(before) == fk)
        if _CHIUSURA_RIGA.match(last) or firma_o_telefono(last, prev_sig):
            lines.pop()
        else:
            break
    t = "\n".join(lines)
    for _ in range(2):
        t = _SALUTO_INIZIO.sub("", t, count=1)
    for _ in range(3):
        new = _CHIUSURA_FINE.sub("", t).rstrip()
        if new == t:
            break
        t = new
    t = re.sub(r"\s+([,.;:!?])", r"\1", t).strip(" \n,;-")
    t = re.sub(r"\n{2,}", "\n", t)
    for i, c in enumerate(t):
        if c.isalpha():
            t = t[:i] + c.upper() + t[i + 1:]
            break
    if t and t[-1].isalnum() or t.endswith(")"):
        t += "."
    return t[:4000]


# ------------------------------------------------------------------------------ nomi scritti nel testo
_TITOLI = (r"(?:dott\.?\s?ss?a\.?|dottoressa|dottore|dott\.?|ingegnere|ing\.?|avvocato|avv\.?|"
           r"professoressa|professore|prof\.?\s?ssa|prof\.?|ispettore|ispettrice|signora|signore|"
           r"signor|sig\.?\s?ra|sig\.?)")
_NOME = r"((?:[A-ZÀ-Ý][a-zà-ÿ']+\s+){1,3}[A-ZÀ-Ý][a-zà-ÿ']+)"
NOME_CON_MATRICOLA_RE = re.compile(_NOME + r"\s*[,(:]?\s*(?i:con\s+)?(?i:matricola)\s*[:.]?\s*(?i:x\s?f\s?(\d{4,6}))")
NOME_CON_TITOLO_RE = re.compile(r"(?i:\b" + _TITOLI + r")[\s:]*" + _NOME)
_NO_NOME = {"buongiorno", "buonasera", "buon", "pomeriggio", "salve", "si", "il", "lo", "la", "le", "l",
            "dott", "dottor", "dottore", "dottoressa", "ssa", "sa", "ing", "ingegnere", "avv", "avvocato",
            "prof", "professore", "professoressa", "ispettore", "ispettrice", "utente", "collega", "per",
            "del", "della", "sig", "signor", "signora", "dr", "firenze", "inail", "microsoft", "teams",
            "outlook", "windows", "cartella", "clinica", "direzione", "regionale", "porte", "nuove",
            "piano", "sede", "ufficio", "centro", "medico", "legale", "stampante", "konica", "minolta",
            "dell", "lenovo", "hp", "word", "excel", "pdf", "vpn", "wifi", "carcli", "toscana"}


def names_in_text(text):
    """I nomi di persona scritti nel testo, con la matricola se c'è: [(nome, matricola)], prima quelli
    con matricola. Solo «Dott.ssa Nome Cognome» o «Nome Cognome con matricola XF...»."""
    found, seen = [], set()

    def add(raw, mat):
        w = raw.split()
        while w and fold(w[0]) in _NO_NOME:
            w = w[1:]
        if len(w) < 2 or any(len(x.replace("'", "")) < 2 or fold(x) in _NO_NOME for x in w):
            return
        k = " ".join(sorted(fold(x) for x in w))
        if k in seen:
            return
        seen.add(k)
        found.append((" ".join(w), ("XF" + mat) if mat else ""))
    for m in NOME_CON_MATRICOLA_RE.finditer(text):
        add(m.group(1), m.group(2))
    for m in NOME_CON_TITOLO_RE.finditer(text):
        add(m.group(1), "")
    return found


# ------------------------------------------------------------------------------ lettura del file
def read_rows(path):
    """Legge un Excel di ServiceNow e restituisce le righe con i campi che servono. Riconosce le
    colonne dal titolo. Solleva ValueError se non sembra un elenco di ticket."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, data_only=True)
    except Exception as e:
        raise ValueError("non riesco ad aprire il file Excel (%s)" % e)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for hi, head in enumerate(rows[:10]):
            names = [fold(_s(v)) if v is not None else "" for v in head]
            col = {}
            for field, aliases in INTESTAZIONI.items():
                for a in aliases:                    # la prima colonna di quel nome
                    if a in names:
                        col[field] = names.index(a)
                        break
            if "ticket" not in col or "descr" not in col and "breve" not in col:
                continue
            out = []
            for r in rows[hi + 1:]:
                def get(k):
                    i = col.get(k)
                    return r[i] if i is not None and i < len(r) else None
                num = _s(get("ticket")).upper().replace(" ", "")
                if not TICKET_RE.match(num):
                    continue
                out.append({"ticket": num, "aperto": get("aperto"), "apertoda": _s(get("apertoda")),
                            "chiamante": _s(get("chiamante")), "stato": _s(get("stato")),
                            "breve": _s(get("breve")), "descr": _s(get("descr")),
                            "sede": _s(get("sede")), "ci": _s(get("ci")),
                            "soluzione": _s(get("soluzione")), "actor": _s(get("actor"))})
            if out:
                return out
    raise ValueError("non sembra un elenco di ticket di ServiceNow: servono almeno le colonne "
                     "del numero del ticket e della descrizione")


def technician_name(rows):
    """Il nome di chi lavora i ticket (la colonna Actor, la più frequente), se c'è."""
    names = [r["actor"] for r in rows if r.get("actor")]
    return max(set(names), key=names.count) if names else ""


def group_tickets(rows):
    """Una riga per ticket: le righe con lo stesso numero si uniscono. Vale la data di apertura
    più vecchia, lo stato dell'ultima riga e, per gli altri dati, il primo valore scritto."""
    out = {}
    for r in rows:
        t = out.get(r["ticket"])
        if t is None:
            out[r["ticket"]] = dict(r, aperto=parse_when(r["aperto"]), righe=1)
            continue
        t["righe"] += 1
        w = parse_when(r["aperto"])
        if w and (not t["aperto"] or w < t["aperto"]):
            t["aperto"] = w
        if r["stato"]:
            t["stato"] = r["stato"]
        for k in ("apertoda", "chiamante", "breve", "descr", "sede", "ci", "soluzione"):
            if not t[k] and r[k]:
                t[k] = r[k]
    return list(out.values())


# ------------------------------------------------------------------------------ il piano
class _Index:
    """Le persone dell'app, per ritrovarle dal nome, dalla matricola, dal seriale, dall'IMEI."""

    def __init__(self, users):
        self.by_words, self.by_mat, self.by_serial, self.by_imei, self.by_cognome = {}, {}, {}, {}, {}
        self.all = []
        for u in users:
            for name in {u["nome"], u.get("ordine") or u["nome"]}:
                lst = self.by_words.setdefault(words_key(name), [])
                if u not in lst:
                    lst.append(u)
            m = MATRICOLA_RE.search(u.get("matricola") or "") or re.match(
                r"(?i)^\s*(?:xf)?\s*(\d{4,6})\s*$", u.get("matricola") or "")
            if m:
                self.by_mat.setdefault(m.group(1), []).append(u)
            for s in self._serials(u):
                self.by_serial.setdefault(s, []).append(u)
            im = re.sub(r"\D", "", u.get("imei") or "")
            if len(im) >= 14:
                self.by_imei.setdefault(im, []).append(u)
            first = ((u.get("ordine") or u["nome"] or "").split() or [""])[0]
            cogn = tokens(first)
            if cogn and len(cogn[0]) >= 4:
                self.by_cognome.setdefault(cogn[0], []).append(u)
            self.all.append((u, set(tokens(u["nome"])), set(tokens(u.get("ordine") or ""))))

    @staticmethod
    def _serials(u):
        found = set()
        vals = [u.get("seriale_pc") or ""]
        try:
            for d in json.loads(u.get("dotazione") or "[]"):
                vals.append(str(d.get("seriale") or ""))
        except (ValueError, AttributeError, TypeError):
            pass
        for v in vals:
            v = " ".join(v.split())
            m = ENTE_SERIAL_RE.search(v)
            v = (m.group(1) if m else v).replace(" ", "").upper()
            if len(v) >= 5:
                found.add(v)
        return found

    def by_name(self, name):
        return self.by_words.get(words_key(name), [])

    def by_name_loose(self, name):
        """Chi ha le stesse parole del nome, oppure una parte (almeno due) di quelle di un nome più
        lungo: «Federica Anna Paternò» per «Paternò Federica». Solo se la persona è una sola."""
        want = set(tokens(name))
        if len(want) < 2:
            return []
        hits = []
        for u, a, b in self.all:
            for have in (a, b):
                if len(have) >= 2 and len(have & want) >= 2 and (have <= want or want <= have):
                    hits.append(u)
                    break
        return _uniq(hits)


def _uniq(users):
    seen, out = set(), []
    for u in users:
        if u["id"] not in seen:
            seen.add(u["id"])
            out.append(u)
    return out


def _find_by_clues(ix, t, escludi):
    """Chi è la persona dai dati scritti nel ticket. Restituisce (persona, come, da_controllare)."""
    text = " ".join([t["breve"], t["descr"], t["ci"]])
    esc = set(escludi)

    def ok(lst):
        return [u for u in _uniq(lst) if u["id"] not in esc]

    mats = []
    for m in MATRICOLA_RE.finditer(text):
        mats += ok(ix.by_mat.get(m.group(1), []))
    mats = _uniq(mats)
    sers = []
    for m in list(SN_RE.finditer(text)) + list(ENTE_SERIAL_RE.finditer(text)):
        sers += ok(ix.by_serial.get(m.group(1).upper(), []))
    sers = _uniq(sers)
    imeis = []
    for m in IMEI_RE.finditer(text):
        imeis += ok(ix.by_imei.get(m.group(1), []))
    imeis = _uniq(imeis)
    if mats:
        dubbio = len(mats) > 1 or (sers and sers[0]["id"] != mats[0]["id"])
        return mats[0], "matricola", bool(dubbio)
    if sers:
        return sers[0], "seriale", len(sers) > 1
    if imeis:
        return imeis[0], "IMEI", len(imeis) > 1
    tk = tokens(t["descr"] + " " + t["breve"])
    for size in (4, 3, 2):                       # il nome intero scritto nel testo
        for i in range(len(tk) - size + 1):
            hits = ok(ix.by_words.get(" ".join(sorted(tk[i:i + size])), []))
            if len(hits) == 1:
                return hits[0], "nome nel testo", False
    for m in TITOLO_COGNOME_RE.finditer(t["descr"] + " " + t["breve"]):
        hits = ok(ix.by_cognome.get(fold(m.group(1)), []))
        if len(hits) == 1:
            return hits[0], "cognome nel testo", True
    return None, "", False


def plan_import(tickets, users, esistenti, sedi_note=(), tecnico=""):
    """Prepara l'importazione senza scrivere niente.
    tickets: group_tickets(...); users: db.users_basic(); esistenti: db.interventi_per_import();
    sedi_note: le sedi già nell'app; tecnico: il nome di chi lavora i ticket (nei ticket aperti da
    lui l'utente vero non è «Aperto da»)."""
    ix = _Index(users)
    by_id = {u["id"]: u for u in users}
    tel_tec = telefono_tecnico(t for t in tickets if tecnico and words_key(t.get("apertoda") or "") == words_key(tecnico))
    tec_key = words_key(tecnico) if tecnico else ""
    escludi = {u["id"] for u in ix.by_name(tecnico)} if tecnico else set()
    presenti = {}
    for e in esistenti:
        for tok in TICKET_IN_TEXT_RE.findall((e.get("ticket") or "").upper()):
            presenti.setdefault(tok, e)
    sedi_k = {fold(s) for s in sedi_note}
    st = {"righe": 0, "ticket": 0, "annullati": 0, "gia_presenti": 0, "nuovi": 0,
          "soluzioni_aggiunte": 0, "riassegnati": 0, "sede_scritta": 0, "sede_diversa": 0,
          "da_richiedente": 0, "da_chiamante": 0, "dedotti": 0, "senza_utente": 0, "righe_file": 0,
          "per_come": {}}
    plan = {"nuovi": [], "utenti_nuovi": [], "sede_utenti": [], "soluzioni": [], "riassegna": [],
            "controllare": [], "sedi_nuove": [], "stats": st}
    nuovi_utenti = {}                    # parole del nome -> {"nome","sede","matricola"}
    sede_set = {}
    sedi_create = []
    st["ticket"] = len(tickets)
    st["righe_file"] = sum(t.get("righe", 1) for t in tickets)

    def nome_vero(u):
        return u.get("ordine") or u["nome"]

    def registra(nome, sede, matricola=""):
        """Una persona da creare; se compare in più ticket (anche col nome scritto al contrario) è una sola."""
        rec = nuovi_utenti.setdefault(words_key(nome), {"nome": nome, "sede": sede, "matricola": matricola})
        if matricola and not rec["matricola"]:
            rec["matricola"] = matricola
        if sede and not rec["sede"]:
            rec["sede"] = sede
        return rec["nome"]

    def trova_persona(nome, incerto=False):
        """(persona esistente | None, nome nuovo «Cognome Nome» | '') per un nome scritto da ServiceNow."""
        hits = _uniq(ix.by_name(nome)) or ix.by_name_loose(nome)
        if len(hits) == 1:
            return hits[0], ""
        if len(hits) > 1:
            return None, ""
        return None, cognome_nome(nome, incerto)

    for t in tickets:
        stato = map_stato(t["stato"])
        if stato is None:
            st["annullati"] += 1
            continue
        sede = map_sede(t["sede"], sedi_note)
        if sede and fold(sede) not in sedi_k and sede not in sedi_create:
            sedi_create.append(sede)
        user, nuovo, come, dubbio = None, "", "", False
        for campo, etichetta in (("chiamante", "chiamante"), ("apertoda", "richiedente")):
            nome = t.get(campo) or ""
            if not nome or (tec_key and words_key(nome) == tec_key):
                continue
            user, nuovo = trova_persona(nome)
            if user or nuovo:
                come = etichetta
                break
        mat_nuovo = ""
        if not user and not nuovo:
            user, come, dubbio = _find_by_clues(ix, t, escludi)
        if not user and not nuovo:                  # un nome di persona scritto nel testo, non ancora nell'app
            for nm, mat in names_in_text(t["descr"] + " " + t["breve"]):
                if tec_key and words_key(nm) == tec_key:
                    continue
                u2, n2 = trova_persona(nm, True)
                if u2:
                    user, come, dubbio = u2, "nome nel testo", True
                elif n2:
                    nuovo, mat_nuovo, come, dubbio = n2, mat, "nome nel testo (persona nuova)", True
                if u2 or n2:
                    break
        ex = presenti.get(t["ticket"])
        if ex:
            st["gia_presenti"] += 1
            sol = clean_text(t["soluzione"], tecnico, tel_tec) if t.get("soluzione") else ""
            if sol and not (ex.get("soluzione") or "").strip():
                plan["soluzioni"].append({"id": ex["id"], "soluzione": sol, "ticket": t["ticket"]})
                st["soluzioni_aggiunte"] += 1
            # un intervento importato con l'utente dedotto: se ora il file dice chi è, lo si corregge
            nota = ex.get("note") or ""
            if (come in ("chiamante", "richiedente") and (user or nuovo) and SEGNO in nota
                    and (DEDOTTO in nota or "utente non riconosciuto" in nota)):
                plan["riassegna"].append({"id": ex["id"], "ticket": t["ticket"],
                                          "utente": user["nome"] if user else nuovo,
                                          "nuovo": bool(nuovo)})
                if nuovo:
                    nuovo = registra(nuovo, sede)
                    plan["riassegna"][-1]["utente"] = nuovo
                st["riassegnati"] += 1
            continue
        # la persona
        if user:
            nome_utente = user["nome"]
            if sede and not (user.get("sede") or "").strip():
                if user["id"] not in sede_set:
                    sede_set[user["id"]] = sede
                    plan["sede_utenti"].append({"id": user["id"], "sede": sede,
                                                "nome": nome_vero(user)})
                    st["sede_scritta"] += 1
            elif sede and fold(user.get("sede") or "") != fold(sede) and user["id"] not in sede_set:
                sede_set[user["id"]] = None
                st["sede_diversa"] += 1
        elif nuovo:
            nome_utente = registra(nuovo, sede, mat_nuovo)
            if come == "nome nel testo (persona nuova)":
                plan["controllare"].append("%s: persona creata dal testo del ticket: %s (controlla cognome e nome)"
                                           % (t["ticket"], nome_utente))
        else:
            nome_utente = DA_ASSEGNARE
            come = ""
            st["senza_utente"] += 1
            plan["controllare"].append("%s: utente non riconosciuto (messo su «%s»)"
                                       % (t["ticket"], DA_ASSEGNARE))
        if come in ("chiamante", "richiedente"):
            st["da_" + ("chiamante" if come == "chiamante" else "richiedente")] += 1
        elif come:
            st["dedotti"] += 1
            k = come.replace(" (persona nuova)", "")
            st["per_come"][k] = st["per_come"].get(k, 0) + 1
        if dubbio:
            plan["controllare"].append("%s: utente %s dedotto da %s, ma incerto"
                                       % (t["ticket"], nome_utente, come))
        richiesta = clean_text(t["descr"], tecnico, tel_tec) or clean_text(t["breve"], tecnico, tel_tec)
        note = []
        if t["ci"]:
            note.append("Dispositivo: " + " ".join(t["ci"].split()))
        breve = " ".join(t["breve"].split())
        if breve and fold(breve).startswith("ambito"):
            note.append("Categoria: " + breve)
        sign = SEGNO + " (%s)" % t["ticket"]
        if come not in ("chiamante", "richiedente") and not user and not nuovo:
            sign += " - utente non riconosciuto"
        elif come not in ("chiamante", "richiedente"):
            sign += " - %s da %s%s" % (DEDOTTO, come, ", da controllare" if dubbio else "")
        note.append(sign)
        plan["nuovi"].append({
            "ticket": t["ticket"], "data_ora": t["aperto"] or "", "utente": nome_utente,
            "sede": sede or (user.get("sede") if user else "") or "", "richiesta": richiesta,
            "soluzione": clean_text(t["soluzione"], tecnico, tel_tec) if t.get("soluzione") else "",
            "note": "\n".join(note), "stato": stato, "come": come})
        st["nuovi"] += 1
    plan["utenti_nuovi"] = sorted(nuovi_utenti.values(), key=lambda u: u["nome"].casefold())
    plan["sedi_nuove"] = sorted(sedi_create, key=str.casefold)
    if any(i["utente"] == DA_ASSEGNARE for i in plan["nuovi"]):
        plan["utenti_nuovi"].append({"nome": DA_ASSEGNARE, "sede": "", "matricola": ""})
    plan["nuovi"].sort(key=lambda i: (i["data_ora"], i["ticket"]))
    return plan


def key(name):
    return " ".join(str(name or "").split()).casefold()


# ------------------------------------------------------------------------------ si applica
def apply_plan(db, plan):
    """Scrive nel database quello che ha preparato plan_import. Restituisce quanti interventi."""
    for u in plan["utenti_nuovi"]:
        if not db.find_user(u["nome"]):
            db.add_user(u["nome"], u["sede"], "", "", "", u.get("matricola") or "")
    for s in plan["sede_utenti"]:
        if s["sede"]:
            db.set_user_basic(s["id"], sede=s["sede"])
    for i in plan["nuovi"]:
        db.save_phone_intervento(None, "pc-" + uuid.uuid4().hex, i["data_ora"], i["utente"],
                                 i["sede"], i["richiesta"], i["soluzione"], i["note"],
                                 i["ticket"], _now_iso(), i["stato"])
    for s in plan["soluzioni"]:
        db.update_intervento_import(s["id"], soluzione=s["soluzione"])
    for r in plan["riassegna"]:
        db.update_intervento_import(r["id"], utente=r["utente"], togli_dedotto=True)
    return len(plan["nuovi"])
