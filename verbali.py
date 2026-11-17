# -*- coding: utf-8 -*-
"""
Verbali Word per l'help desk (trasferimento, consegna/presa in carico, assegnazione, censimento).

I modelli sono normali file Word con dei segnaposto tra doppie graffe, per esempio {{UTENTE}}.
Il programma apre il modello, sostituisce i segnaposto con i dati inseriti e salva un nuovo file:
loghi, caratteri e impaginazione restano quelli del modello, che si può modificare con Word.

Regole dei segnaposto
  {{NOME}}   sostituito dal valore; se il valore è vuoto resta una riga ______ da scrivere a mano
  {{?NOME}}  come sopra, ma se il valore è vuoto l'intero paragrafo viene tolto
  {{-NOME}}  come sopra, ma se il valore è vuoto non scrive niente
  {{FIRMA_TECNICO}} {{FIRMA_UTENTE}}  messi davanti alla riga della firma: se c'è la firma la
             immagine compare sopra la riga, altrimenti non succede niente (la riga resta da firmare
             a mano). {{FIRMA_GAP_...}} e {{FIRMA_SOLO_UTENTE}} sono righe che compaiono solo se
             c'è la firma.
Questo file non usa tkinter: si può provare anche senza finestre.
"""
import base64
import copy
import json
import os
import re
import shutil
import tempfile
import uuid

MARKER_RE = re.compile(r"\{\{([?-])?([A-Z0-9_]+)\}\}")
BLANK = "_" * 30
CHECK_ON = "☒"    # ☒
CHECK_OFF = "☐"   # ☐
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MAX_FOTO_PX = 1800
FIRMA_RE = re.compile(r"\{\{(FIRMA_[A-Z_]+)\}\}")
FIRMA_LARG_CM = 4.2       # larghezza massima della firma nel Word
FIRMA_ALT_CM = 1.05       # altezza massima
FIRMA_PX = 900            # lato lungo massimo dell'immagine della firma

# chiave -> (titolo, nome del file modello, foto obbligatoria)
TIPI = {
    "trasferimento": ("Trasferimento apparato", "Modello verbale trasferimento apparato.docx", False),
    "presa_in_carico": ("Consegna e presa in carico", "Modello verbale consegna e presa in carico.docx", False),
    "assegnazione": ("Assegnazione materiale", "Modello verbale assegnazione materiale.docx", False),
    "censimento": ("Censimento materiale", "Modello verbale censimento materiale.docx", True),
}
TIPI_ORDINE = ["trasferimento", "presa_in_carico", "assegnazione", "censimento"]

MOTIVI_TRASFERIMENTO = ["Apparato non utilizzato", "Riallocazione risorse", "Sostituzione apparato",
                        "Riparazione/Manutenzione", "Pensionamento Utente"]
MOTIVI_PRESA_IN_CARICO = ["Sostituzione dispositivo", "Guasto tecnico", "Licenziamento",
                          "Pensionamento utente", "Dismissione"]
TIPOLOGIE = ["Notebook", "PC fisso", "Monitor", "Smartphone", "Tablet", "Stampante", "Altro"]
TIPI_DOTAZIONE = ["Notebook aziendale", "Cuffie professionali", "Smartphone aziendale",
                  "Scheda SIM aziendale", "Altro"]
COLONNE_CODICE = ["IMEI", "Numero di serie", "IMEI / Seriale"]

# titolo -> preposizione articolata ("all'avv.", "al dott.", "alla dott.ssa" ...)
TITOLI = {
    "": "a", "avv.": "all’", "dott.": "al", "dott.ssa": "alla", "ing.": "all’",
    "arch.": "all’", "geom.": "al", "rag.": "al", "sig.": "al", "sig.ra": "alla",
    "prof.": "al", "prof.ssa": "alla", "on.": "all’", "dr.": "al", "dr.ssa": "alla",
}


class VerbaleError(Exception):
    """Errore da mostrare così com'è all'utente."""


def _docx():
    try:
        import docx
        return docx
    except ImportError as e:  # pragma: no cover
        raise VerbaleError("Manca il componente per creare i file Word (python-docx). "
                           "Ricostruisci il programma con requirements.txt aggiornato.\n(%s)" % e)


def _w(tag):
    from docx.oxml.ns import qn
    return qn("w:" + tag)


# --------------------------------------------------------------------------
# Modelli
# --------------------------------------------------------------------------
def ensure_templates(folder):
    """Scrive nella cartella i modelli che mancano (non tocca quelli che hai modificato)."""
    from verbali_modelli import MODELLI
    os.makedirs(folder, exist_ok=True)
    for tipo, (_, nome, _) in TIPI.items():
        path = os.path.join(folder, nome)
        if not os.path.exists(path):
            with open(path, "wb") as f:
                f.write(base64.b64decode(MODELLI[tipo]))
    return folder


def restore_templates(folder):
    """Rimette i modelli originali; quelli che c'erano restano come copia."""
    from verbali_modelli import MODELLI
    os.makedirs(folder, exist_ok=True)
    kept = []
    for tipo, (_, nome, _) in TIPI.items():
        path = os.path.join(folder, nome)
        if os.path.exists(path):
            stem, ext = os.path.splitext(nome)
            old = os.path.join(folder, "%s (prima del ripristino)%s" % (stem, ext))
            shutil.copy2(path, old)
            kept.append(old)
        with open(path, "wb") as f:
            f.write(base64.b64decode(MODELLI[tipo]))
    return kept


def template_path(folder, tipo):
    ensure_templates(folder)
    return os.path.join(folder, TIPI[tipo][1])


# --------------------------------------------------------------------------
# Sostituzione dei segnaposto dentro ai paragrafi
# --------------------------------------------------------------------------
def _t_nodes(p_el):
    out = []
    for t in p_el.iter(_w("t")):
        node, inside_box = t.getparent(), False
        while node is not None and node is not p_el:
            if node.tag == _w("txbxContent"):
                inside_box = True
                break
            node = node.getparent()
        if not inside_box:
            out.append(t)
    return out


def paragraph_text(p_el):
    return "".join(t.text or "" for t in _t_nodes(p_el))


def _set_text(t, text):
    t.text = text
    t.set(XML_SPACE, "preserve")


def replace_markers(p_el, values, blank=BLANK):
    """Sostituisce i segnaposto di un paragrafo (anche se Word li ha spezzati in più parti).
    Restituisce True se il paragrafo deve essere tolto ({{?NOME}} con valore vuoto)."""
    guard = 0
    while guard < 200:
        guard += 1
        nodes = _t_nodes(p_el)
        text = "".join(t.text or "" for t in nodes)
        m = MARKER_RE.search(text)
        if not m:
            return False
        mode, name = m.group(1), m.group(2)
        val = values.get(name)
        val = "" if val is None else str(val)
        if mode == "?" and not val.strip():
            return True
        repl = val if val.strip() else ("" if mode == "-" else blank)
        s, e = m.span()
        spans, pos = [], 0
        for n in nodes:
            ln = len(n.text or "")
            spans.append((n, pos, pos + ln))
            pos += ln
        hit = [(n, a, b) for n, a, b in spans if b > s and a < e]
        first, fa, _ = hit[0]
        last, la, _ = hit[-1]
        prefix = (first.text or "")[:s - fa]
        suffix = (last.text or "")[e - la:]
        lines = repl.split("\n")
        for n, _, _ in hit[1:-1]:
            _set_text(n, "")
        if first is last:
            _set_text(first, prefix + lines[0] + (suffix if len(lines) == 1 else ""))
        else:
            _set_text(first, prefix + lines[0])
            _set_text(last, suffix)
        if len(lines) > 1:
            from lxml import etree
            anchor = last_t = first
            for line in lines[1:]:
                br = etree.Element(_w("br"))
                nt = etree.Element(_w("t"))
                anchor.addnext(br)
                br.addnext(nt)
                _set_text(nt, line)
                anchor = last_t = nt
            if first is last:
                _set_text(last_t, (last_t.text or "") + suffix)
    return False


def _all_paragraphs(doc):
    return list(doc.element.body.iter(_w("p")))


def _find_paragraph(doc, name):
    for p in _all_paragraphs(doc):
        txt = paragraph_text(p)
        if any(("{{%s%s}}" % (m, name)) in txt for m in ("", "?", "-")):
            return p
    return None


def _repeat_paragraph(doc, name, items):
    """Ripete il paragrafo che contiene il segnaposto {{name}} una volta per ogni elemento
    (items = lista di dizionari di valori). Senza elementi il paragrafo sparisce."""
    proto = _find_paragraph(doc, name)
    if proto is None:
        return None
    for values in items:
        clone = copy.deepcopy(proto)
        proto.addprevious(clone)
        if replace_markers(clone, values):
            clone.getparent().remove(clone)
    following = proto.getnext()
    proto.getparent().remove(proto)
    return following   # il paragrafo che viene dopo l'elenco (serve per le righe da compilare a mano)


def _repeat_table_row(doc, marker, rows):
    """Nella tabella la riga che contiene {{marker}} fa da esempio: viene ripetuta per ogni riga."""
    for tr in doc.element.body.iter(_w("tr")):
        if any("{{%s}}" % marker in paragraph_text(p) for p in tr.iter(_w("p"))):
            for values in rows:
                clone = copy.deepcopy(tr)
                tr.addprevious(clone)
                for p in clone.iter(_w("p")):
                    replace_markers(p, values, blank="")
            tr.getparent().remove(tr)
            return True
    return False


# --------------------------------------------------------------------------
# Controlli sui codici (IMEI, doppioni)
# --------------------------------------------------------------------------
def luhn_ok(digits):
    total = 0
    for i, ch in enumerate(reversed(digits)):
        d = int(ch)
        if i % 2 == 1:
            d *= 2
            if d > 9:
                d -= 9
        total += d
    return total % 10 == 0


def imei_problem(value):
    """Spiega cosa non torna in un IMEI (None se va bene o se non è un numero)."""
    s = re.sub(r"[\s-]", "", value or "")
    if not s or not s.isdigit():
        return None
    if len(s) != 15:
        return "ha %d cifre (un IMEI ne ha 15)" % len(s)
    if not luhn_ok(s):
        return "non supera il controllo di validità, forse c'è un errore di battitura"
    return None


def controlla_apparati(rows, colonna="IMEI"):
    """Elenco di avvisi per le righe di un censimento: codici ripetuti, IMEI strani, righe vuote."""
    avvisi, seen = [], {}
    check_imei = colonna.strip().upper() == "IMEI"
    for i, r in enumerate(rows, 1):
        code = (r.get("codice") or "").strip()
        if not code:
            avvisi.append("Riga %d: manca il codice (%s)." % (i, colonna))
            continue
        key = re.sub(r"[\s-]", "", code).casefold()
        if key in seen:
            avvisi.append("Riga %d: il codice %s è uguale a quello della riga %d." % (i, code, seen[key]))
        else:
            seen[key] = i
        if check_imei:
            why = imei_problem(code)
            if why:
                avvisi.append("Riga %d: l'IMEI %s %s." % (i, code, why))
    return avvisi


def parse_apparati_text(text, marca="", modello=""):
    """Legge un elenco incollato (anche da Excel): una riga per apparecchio, colonne separate da
    tabulazione, punto e virgola o virgola. Con una sola colonna è il codice: marca e modello
    vengono presi da quelli indicati."""
    rows = []
    for line in (text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        parts = [p.strip() for p in re.split(r"\t|;", line)]
        if len(parts) == 1 and "," in line:
            parts = [p.strip() for p in line.split(",")]
        if len(parts) == 1:
            rows.append({"marca": marca, "modello": modello, "codice": parts[0]})
        elif len(parts) == 2:
            rows.append({"marca": marca, "modello": parts[0], "codice": parts[1]})
        else:
            rows.append({"marca": parts[0], "modello": parts[1], "codice": parts[2]})
    return rows


# --------------------------------------------------------------------------
# Memoria dei verbali precedenti: suggerimenti e compilazione automatica
# --------------------------------------------------------------------------
CAMPI_STORICO = ("tipologia", "marca_modello", "provenienza", "destinazione", "contatti", "titolo",
                 "oggetto", "descrizione", "sede_citta", "indirizzo", "cap", "tecnico", "collega",
                 "materiale", "accessori", "dot_modello", "ap_marca", "ap_modello")


def nome_pulito(text):
    """'Mario Rossi (matricola M1)' -> 'Mario Rossi'."""
    return re.sub(r"\s*\(matricola[^)]*\)\s*$", "", text or "", flags=re.I).strip()


def tipologia_da_dotazione(tipo):
    """Da 'Smartphone aziendale' a 'Smartphone' (vuoto se non e' un apparato da trasferire)."""
    t = (tipo or "").casefold()
    if "notebook" in t or "portatile" in t:
        return "Notebook"
    if "smartphone" in t or "cellulare" in t or "telefono" in t:
        return "Smartphone"
    if "tablet" in t:
        return "Tablet"
    if "pc" in t or "desktop" in t or "computer" in t:
        return "PC fisso"
    return ""


def _codice_key(text):
    return re.sub(r"[\s-]", "", text or "").casefold()


def storico(rows):
    """rows = [(tipo, dati), ...] dal piu' recente al piu' vecchio (i verbali gia' fatti).
    Restituisce quello che serve per suggerire i valori e compilare i campi da soli:
      valori[campo]   testi gia' usati, i piu' recenti per primi
      dispositivi     apparati noti (tipologia, marca_modello, seriale, utente)
      titoli          nome -> titolo usato nelle assegnazioni
      sedi            citta' -> indirizzo e CAP dei censimenti
      oggetti         oggetto del censimento -> descrizione
      modelli         modello -> marca (censimenti)"""
    out = {"valori": {k: [] for k in CAMPI_STORICO}, "dispositivi": [], "titoli": {}, "sedi": {},
           "oggetti": {}, "modelli": {}}
    seen = {k: set() for k in CAMPI_STORICO}
    dev_seen = set()

    def valore(campo, text):
        text = (text or "").strip() if isinstance(text, str) else ""
        if text and text.casefold() not in seen[campo]:
            seen[campo].add(text.casefold())
            out["valori"][campo].append(text)

    def dispositivo(tipologia, mm, seriale, utente):
        tipologia, mm, seriale, utente = [x.strip() if isinstance(x, str) else ""
                                          for x in (tipologia, mm, seriale, utente)]
        if not (seriale or mm):
            return
        key = _codice_key(seriale) if seriale else (tipologia.casefold(), mm.casefold(),
                                                    utente.casefold())
        if key in dev_seen:
            return
        dev_seen.add(key)
        out["dispositivi"].append({"tipologia": tipologia, "marca_modello": mm,
                                   "seriale": seriale, "utente": utente})

    for tipo, d in rows:
        d = d or {}
        g = lambda k: d.get(k).strip() if isinstance(d.get(k), str) else ""
        for k in ("tipologia", "marca_modello", "provenienza", "destinazione", "contatti", "titolo",
                  "oggetto", "descrizione", "sede_citta", "indirizzo", "cap", "tecnico", "collega"):
            valore(k, g(k))
        if tipo == "trasferimento":
            dispositivo(g("tipologia"), g("marca_modello"), g("seriale"), g("assegnatario"))
        elif tipo == "presa_in_carico":
            dispositivo(g("tipologia"), g("marca_modello"), g("seriale"), nome_pulito(g("utente")))
            for x in d.get("materiale") or []:
                valore("materiale", x)
        elif tipo == "assegnazione":
            nome = g("utente")
            if g("titolo") and nome and nome.casefold() not in out["titoli"]:
                out["titoli"][nome.casefold()] = g("titolo")
            for it in d.get("dotazione") or []:
                if not isinstance(it, dict):
                    continue
                valore("dot_modello", it.get("modello"))
                tip = tipologia_da_dotazione(it.get("tipo"))
                if tip:
                    dispositivo(tip, it.get("modello"), it.get("codice"), nome)
            for x in d.get("accessori") or []:
                valore("accessori", x)
        elif tipo == "censimento":
            citta = g("sede_citta")
            if citta and citta.casefold() not in out["sedi"] and (g("indirizzo") or g("cap")):
                out["sedi"][citta.casefold()] = {"indirizzo": g("indirizzo"), "cap": g("cap")}
            if g("oggetto") and g("oggetto").casefold() not in out["oggetti"]:
                out["oggetti"][g("oggetto").casefold()] = {"descrizione": g("descrizione")}
            for r in d.get("apparati") or []:
                if not isinstance(r, dict):
                    continue
                marca, modello = (r.get("marca") or "").strip(), (r.get("modello") or "").strip()
                valore("ap_marca", marca)
                valore("ap_modello", modello)
                if modello and marca and modello.casefold() not in out["modelli"]:
                    out["modelli"][modello.casefold()] = marca
    return out


def _dot_scheda(user):
    """Le righe della dotazione scritta nella scheda della persona: [{"tipo","modello","seriale"}]."""
    d = (user or {}).get("dotazione") or []
    if isinstance(d, (str, bytes)):
        try:
            d = json.loads(d or "[]")
        except ValueError:
            d = []
    return [r for r in d if isinstance(r, dict)] if isinstance(d, (list, tuple)) else []


def _non_trasferibile(r):
    """Telefono Teams o fisso, monitor, presa di rete, stampante, token: non si consegnano con
    un verbale di trasferimento o di assegnazione."""
    t = ((r.get("tipo") or "") + " " + (r.get("modello") or "")).casefold()
    return any(w in t for w in ("teams", "audiocodes", "telefono fisso", "monitor", "presa", "stampante",
                                "token", "smart card"))


def dispositivi_utente(st, nome, user=None):
    """Gli apparati di una persona: quelli dei verbali gia' fatti, poi il seriale PC e l'IMEI
    scritti nella sua scheda utente, poi i notebook, i cellulari e i tablet della sua dotazione."""
    key = nome_pulito(nome).casefold()
    res = [dict(x) for x in st["dispositivi"] if x["utente"].casefold() == key]
    seen = {_codice_key(x["seriale"]) for x in res if x["seriale"]}
    for tip, ser in (("Notebook", (user or {}).get("seriale_pc")), ("Smartphone", (user or {}).get("imei"))):
        ser = (ser or "").strip()
        if ser and _codice_key(ser) not in seen:
            seen.add(_codice_key(ser))
            res.append({"tipologia": tip, "marca_modello": "", "seriale": ser,
                        "utente": nome_pulito(nome)})
    for r in _dot_scheda(user):
        tip = tipologia_da_dotazione(r.get("tipo"))
        if not tip or _non_trasferibile(r):
            continue
        ser, mod = (r.get("seriale") or "").strip(), (r.get("modello") or "").strip()
        if ser:
            k = _codice_key(ser)
            if k in seen:                      # gia' in elenco: gli do il modello se manca
                for x in res:
                    if _codice_key(x["seriale"]) == k and not x["marca_modello"] and mod:
                        x["marca_modello"] = mod
                continue
            seen.add(k)
            res.append({"tipologia": tip, "marca_modello": mod, "seriale": ser,
                        "utente": nome_pulito(nome)})
        elif mod and not any(x["tipologia"] == tip for x in res):
            res.append({"tipologia": tip, "marca_modello": mod, "seriale": "",
                        "utente": nome_pulito(nome)})
    return res


def dotazione_da_scheda(user):
    """Le righe del verbale di assegnazione (tipo, modello, codice, iccid, numero) ricavate dalla
    scheda della persona: notebook, smartphone, SIM e cuffie con il loro seriale o IMEI."""
    out, has_nb, has_ph = [], False, False
    ser_pc, imei = ((user or {}).get("seriale_pc") or "").strip(), ((user or {}).get("imei") or "").strip()
    for r in _dot_scheda(user):
        t = (r.get("tipo") or "").casefold()
        mod, ser = (r.get("modello") or "").strip(), (r.get("seriale") or "").strip()
        if re.search(r"\bsim\b", t):
            out.append({"tipo": TIPI_DOTAZIONE[3], "modello": mod, "codice": "", "iccid": ser, "numero": ""})
        elif "cuffi" in t:
            out.append({"tipo": TIPI_DOTAZIONE[1], "modello": mod, "codice": ser, "iccid": "", "numero": ""})
        elif _non_trasferibile(r):
            continue
        elif tipologia_da_dotazione(r.get("tipo")) == "Notebook":
            out.append({"tipo": TIPI_DOTAZIONE[0], "modello": mod, "codice": ser or (ser_pc if not has_nb else ""),
                        "iccid": "", "numero": ""})
            has_nb = True
        elif tipologia_da_dotazione(r.get("tipo")) == "Smartphone":
            out.append({"tipo": TIPI_DOTAZIONE[2], "modello": mod, "codice": ser or (imei if not has_ph else ""),
                        "iccid": "", "numero": ""})
            has_ph = True
    if not has_nb and ser_pc:
        out.append({"tipo": TIPI_DOTAZIONE[0], "modello": "", "codice": ser_pc, "iccid": "", "numero": ""})
    if not has_ph and imei:
        out.append({"tipo": TIPI_DOTAZIONE[2], "modello": "", "codice": imei, "iccid": "", "numero": ""})
    return out


def trova_dispositivo(st, seriale, users=()):
    """L'apparato con questo numero di serie/IMEI (dai verbali o dalle schede utente)."""
    key = _codice_key(seriale)
    if not key:
        return None
    for x in st["dispositivi"]:
        if _codice_key(x["seriale"]) == key:
            return dict(x)
    for u in users:
        for tip, ser in (("Notebook", u.get("seriale_pc")), ("Smartphone", u.get("imei"))):
            if ser and _codice_key(ser) == key:
                mod = next((x["marca_modello"] for x in dispositivi_utente({"dispositivi": []}, u.get("nome", ""), u)
                            if _codice_key(x["seriale"]) == key), "")
                return {"tipologia": tip, "marca_modello": mod, "seriale": ser.strip(),
                        "utente": u.get("nome", "")}
        for x in dispositivi_utente({"dispositivi": []}, u.get("nome", ""), u):      # anche la dotazione
            if x["seriale"] and _codice_key(x["seriale"]) == key:
                return x
    return None


def etichetta_dispositivo(x):
    return " · ".join(p for p in (x.get("tipologia"), x.get("marca_modello"),
                                  ("S/N " + x["seriale"]) if x.get("seriale") else "",
                                  ("(%s)" % x["utente"]) if x.get("utente") else "") if p)


# --------------------------------------------------------------------------
# Valori da mettere nei modelli
# --------------------------------------------------------------------------
def frase_assegnatario(titolo, nome):
    """'all'avv. Mario Rossi', 'al dott. Mario Rossi', 'a Mario Rossi' ..."""
    titolo = (titolo or "").strip()
    nome = (nome or "").strip()
    prep = TITOLI.get(titolo.casefold(), "a")
    if not titolo:
        return ("a " + nome).strip() if nome else ""
    if prep.endswith("’"):
        return "%s%s %s" % (prep, titolo, nome)
    return "%s %s %s" % (prep, titolo, nome)


def righe_dotazione(item):
    """Righe sotto il nome di un oggetto assegnato: S/N, IMEI, ICCID, numero."""
    tipo = (item.get("tipo") or "").casefold()
    codice = (item.get("codice") or "").strip() or BLANK
    if "sim" in tipo:
        return ["ICCID: " + ((item.get("iccid") or "").strip() or BLANK),
                "Numero: " + ((item.get("numero") or "").strip() or BLANK)]
    if "teams" in tipo or "audiocodes" in tipo:      # il telefono da scrivania ha il S/N, non l'IMEI
        return ["S/N: " + codice]
    if "smartphone" in tipo or "cellulare" in tipo or "telefono" in tipo:
        return ["IMEI: " + codice]
    return ["S/N: " + codice]


def _checks(selected, options, altro):
    vals = {}
    for i, label in enumerate(options, 1):
        vals["C%d" % i] = CHECK_ON if label in (selected or []) else CHECK_OFF
    vals["C%d" % (len(options) + 1)] = CHECK_ON if (altro or "").strip() else CHECK_OFF
    return vals


def valori_verbale(tipo, d, imp=None):
    """Prepara i valori (e le liste) del verbale a partire dai dati del modulo."""
    imp = imp or {}
    s = lambda k: (d.get(k) or "").strip()
    if tipo == "trasferimento":
        v = {"DATA": s("data"), "TIPOLOGIA": s("tipologia"), "MARCA_MODELLO": s("marca_modello"),
             "SERIALE": s("seriale"), "PROVENIENZA": s("provenienza"),
             "DESTINAZIONE": s("destinazione"), "ASSEGNATARIO": s("assegnatario"),
             "CONTATTI": s("contatti"), "ALTRO": s("altro"), "NOTE": s("note"),
             "DATA_SPED": s("data_sped") or s("data")}
        v.update(_checks(d.get("motivi"), MOTIVI_TRASFERIMENTO, s("altro")))
        return v, {}
    if tipo == "presa_in_carico":
        v = {"DATA": s("data"), "TIPOLOGIA": s("tipologia"), "MARCA_MODELLO": s("marca_modello"),
             "SERIALE": s("seriale"), "UTENTE": s("utente"), "INC": s("inc"), "RIC": s("ric"),
             "ALTRO": s("altro")}
        v.update(_checks(d.get("motivi"), MOTIVI_PRESA_IN_CARICO, s("altro")))
        mat = [{"MATERIALE": x.strip()} for x in (d.get("materiale") or []) if x.strip()]
        return v, {"MATERIALE": mat}
    if tipo == "assegnazione":
        nome = s("utente")
        v = {"DATA": s("data"), "ASSEGNATARIO_FRASE": frase_assegnatario(s("titolo"), nome),
             "MATRICOLA": s("matricola")}
        dot = []
        for it in d.get("dotazione") or []:
            dot.append({"DOT_TIPO": (it.get("tipo") or "").strip(),
                        "DOT_MODELLO": (it.get("modello") or "").strip(),
                        "DOT_RIGHE": "\n".join(righe_dotazione(it))})
        acc = [{"ACCESSORIO": x.strip()} for x in (d.get("accessori") or []) if x.strip()]
        return v, {"DOT_TIPO": dot, "ACCESSORIO": acc}
    if tipo == "censimento":
        citta = s("sede_citta") or imp.get("sede_citta", "")
        indirizzo = s("indirizzo") or imp.get("indirizzo", "")
        cap = s("cap") or imp.get("cap", "")
        col = s("col_codice") or "IMEI"
        descr = s("descrizione") or (s("oggetto")[:1].lower() + s("oggetto")[1:])
        righe = []
        for i, r in enumerate(d.get("apparati") or [], 1):
            righe.append({"N": str(i), "MARCA": (r.get("marca") or "").strip(),
                          "MODELLO": (r.get("modello") or "").strip(),
                          "CODICE": (r.get("codice") or "").strip()})
        v = {"DATA": s("data"), "OGGETTO": s("oggetto"), "DESCRIZIONE": descr,
             "SEDE1": ("Sede INAIL " + citta).strip() if citta else "", "SEDE2": indirizzo,
             "SEDE3": ("%s, %s" % (cap, citta)).strip(", ") if (cap or citta) else "",
             "CITTA": citta, "TECNICO": s("tecnico") or imp.get("tecnico", ""),
             "COLLEGA": s("collega"), "TOTALE": str(len(righe)), "COL_CODICE": col}
        return v, {"N": righe}
    raise VerbaleError("Tipo di verbale sconosciuto: %s" % tipo)


# --------------------------------------------------------------------------
# Foto
# --------------------------------------------------------------------------
def prepara_foto(src, dest_dir, stem):
    """Converte una foto in JPEG (ridotta, già ruotata) dentro dest_dir. Restituisce il percorso."""
    try:
        from PIL import Image, ImageOps
    except ImportError as e:  # pragma: no cover
        raise VerbaleError("Manca il componente per le immagini (Pillow): %s" % e)
    os.makedirs(dest_dir, exist_ok=True)
    try:
        with Image.open(src) as im:
            im = ImageOps.exif_transpose(im)
            if max(im.size) > MAX_FOTO_PX:
                im.thumbnail((MAX_FOTO_PX, MAX_FOTO_PX))
            im = im.convert("RGB")
            target = os.path.join(dest_dir, "%s.jpg" % stem)
            im.save(target, "JPEG", quality=85)
    except Exception as e:
        raise VerbaleError("Non riesco a leggere la foto \"%s\" (%s)." % (os.path.basename(src), e))
    return target


def _aggiungi_foto(doc, photos):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Cm, Pt
    from PIL import Image
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ALLEGATO FOTOGRAFICO")
    r.bold = True
    n = len(photos)
    for i, path in enumerate(photos, 1):
        with Image.open(path) as im:
            w, h = im.size
        ratio = w / float(h)
        width = min(16.0, 10.0 * ratio)
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.add_run().add_picture(path, width=Cm(width))
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rc = cap.add_run("Foto %d di %d" % (i, n))
        rc.italic = True
        rc.font.size = Pt(9)


# --------------------------------------------------------------------------
# Firme
# --------------------------------------------------------------------------
def prepara_firma(src, dest_dir, stem):
    """Ritaglia la firma (PNG con sfondo trasparente o scansione su bianco) e la ridimensiona.
    Restituisce il percorso del PNG pronto."""
    try:
        from PIL import Image, ImageChops, ImageFilter
    except ImportError as e:  # pragma: no cover
        raise VerbaleError("Manca il componente per le immagini (Pillow): %s" % e)
    os.makedirs(dest_dir, exist_ok=True)
    try:
        with Image.open(src) as im:
            im = im.convert("RGBA")
            alpha = im.getchannel("A")
            if alpha.getextrema() == (255, 255):         # niente trasparenza: tolgo il fondo del foglio
                grey = im.convert("L")
                # il fondo di una foto non e' mai bianco puro (ombre, carta grigia): stimo la luminosita'
                # del foglio zona per zona (il valore piu' chiaro dei blocchi vicini: i tratti sottili della
                # firma non lo abbassano) e considero inchiostro solo cio' che e' piu' scuro del foglio
                bs = max(8, max(grey.size) // 48)
                small = grey.resize((max(1, -(-grey.width // bs)), max(1, -(-grey.height // bs))),
                                    getattr(Image, "Resampling", Image).BOX)
                small = small.filter(ImageFilter.MaxFilter(5))
                paper = small.resize(grey.size, getattr(Image, "Resampling", Image).BILINEAR)
                diff = ImageChops.subtract(paper, grey)
                lut = [0 if v <= 24 else 255 if v >= 90 else int(round((v - 24) * 255.0 / 66))
                       for v in range(256)]
                alpha = diff.point(lut)
                ink = Image.new("RGBA", im.size, (11, 31, 75, 255))   # blu scuro, come la firma col dito
                ink.putalpha(alpha)
                im = ink
            box = alpha.point(lambda a: 255 if a > 40 else 0).getbbox()   # ignoro i puntini leggeri
            if not box:
                raise VerbaleError("La firma è vuota.")
            m = 6
            box = (max(0, box[0] - m), max(0, box[1] - m), min(im.width, box[2] + m),
                   min(im.height, box[3] + m))
            im = im.crop(box)
            if max(im.size) > FIRMA_PX:
                im.thumbnail((FIRMA_PX, FIRMA_PX))
            target = os.path.join(dest_dir, "%s.png" % stem)
            im.save(target, "PNG")
    except VerbaleError:
        raise
    except Exception as e:
        raise VerbaleError("Non riesco a leggere la firma (%s)." % e)
    return target


def _disegno_firma(doc, path):
    """Immagine della firma 'fluttuante' (non sposta il testo), agganciata al punto del segnaposto
    e alzata sopra la riga della firma."""
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.shared import Cm
    from PIL import Image
    with Image.open(path) as im:
        w, h = im.size
    width = FIRMA_LARG_CM
    height = width * h / float(w)
    if height > FIRMA_ALT_CM:
        height = FIRMA_ALT_CM
        width = height * w / float(h)
    inline = doc.part.new_pic_inline(path, Cm(width), Cm(height))
    cx, cy = int(Cm(width)), int(Cm(height))
    up = -int(Cm(0.55))                      # parte nella riga vuota sopra e scende fino alla riga
    anchor = parse_xml(
        '<wp:anchor %s distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        'relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="character"><wp:posOffset>%d</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>%d</wp:posOffset></wp:positionV>'
        '<wp:extent cx="%d" cy="%d"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        '</wp:anchor>' % (nsdecls("wp"), int(Cm(0.25)), up, cx, cy))
    for tag in ("wp:docPr", "wp:cNvGraphicFramePr", "a:graphic"):
        anchor.append(inline.find(qn(tag)))
    drawing = parse_xml("<w:drawing %s/>" % nsdecls("w"))
    drawing.append(anchor)
    return drawing


def applica_firme(doc, firme, avvisi=None, cartaceo=None):
    """firme = {"tecnico": percorso PNG, "utente": percorso PNG} (le chiavi mancanti = niente firma).
    cartaceo = chi firmerà a mano sul foglio stampato (esempio {"utente"}): per lui non si mette
    nessuna immagine e il modello resta com'è per la firma a mano (la riga vuota); nella ricevuta
    facoltativa dell'assegnazione la riga dell'utente compare comunque, da firmare a mano.
    Mette le immagini al posto dei segnaposto {{FIRMA_...}} e toglie le righe che servono solo
    quando c'è la firma. Restituisce quante firme ha messo."""
    from lxml import etree
    img = {k: bool(v) for k, v in (firme or {}).items()}
    paper = {k for k in (cartaceo or ()) if k in FIRME_CHIAVI and not img.get(k)}
    has = dict(img)              # chi ha l'immagine: serve lo spazio sopra la riga per l'immagine
    anyone = any(has.values())
    placed, seen = 0, set()
    for p in _all_paragraphs(doc):
        if p.getparent() is None:
            continue
        names = FIRMA_RE.findall(paragraph_text(p))
        if not names:
            continue
        drop = (("FIRMA_GAP_TECNICO" in names and not has.get("tecnico")) or
                ("FIRMA_GAP_UTENTE" in names and not has.get("utente")) or
                ("FIRMA_GAP_ANY" in names and not anyone) or
                ("FIRMA_SOLO_UTENTE" in names and not has.get("utente") and "utente" not in paper))
        if drop:
            p.getparent().remove(p)
            continue
        for t in _t_nodes(p):
            txt = t.text or ""
            if "{{FIRMA_" not in txt:
                continue
            parts = FIRMA_RE.split(txt)          # testo, nome, testo, nome, testo ...
            _set_text(t, parts[0])
            cur = t
            for i in range(1, len(parts), 2):
                name, after = parts[i], parts[i + 1]
                who = {"FIRMA_TECNICO": "tecnico", "FIRMA_UTENTE": "utente"}.get(name)
                if who:
                    seen.add(who)
                    if img.get(who):
                        d = _disegno_firma(doc, firme[who])
                        cur.addnext(d)
                        cur = d
                        placed += 1
                if after:
                    nt = etree.Element(_w("t"))
                    _set_text(nt, after)
                    cur.addnext(nt)
                    cur = nt
    if avvisi is not None:
        for who, label in (("tecnico", "del tecnico"), ("utente", "dell'utente")):
            if img.get(who) and who not in seen:
                avvisi.append("Il modello Word non ha il posto per la firma %s: la firma non è stata "
                              "inserita. Usa «Modelli Word > Ripristina i modelli originali»." % label)
    return placed


# --------------------------------------------------------------------------
# Posti firma: chi può firmare e in che modo
# --------------------------------------------------------------------------
FIRME_CHIAVI = ("tecnico", "utente")
# chiave -> (etichetta, obbligatoria) per ogni verbale (stessi posti del modello Word e del telefono)
FIRME_POSTI = {
    "trasferimento": [("tecnico", "Firma Tecnico di Presidio", True)],
    "presa_in_carico": [("utente", "Firma utente consegnatario", True),
                        ("tecnico", "Firma tecnico CED/IT", True)],
    "assegnazione": [("tecnico", "Firma tecnico CED/IT", True),
                     ("utente", "Firma utente per ricevuta (facoltativa)", False)],
    "censimento": [("tecnico", "Firma Tecnico di PRESIDIO", True)],
}
# Modi di firma di un posto:
#   nessuna  = niente (la riga resta da firmare)
#   tenuta   = la firma che c'è già (fatta col dito sul telefono o scelta prima)
#   salvata  = l'immagine della firma del tecnico salvata nel programma (SOLO posto del tecnico)
#   cartaceo = firmerà a mano sul foglio stampato
MODI_FIRMA = {"tecnico": ("nessuna", "tenuta", "salvata", "cartaceo"),
              "utente": ("nessuna", "tenuta", "cartaceo")}


def modi_firma_ammessi(chiave):
    """Modi di firma permessi a un posto firma. L'immagine della firma (salvata) c'è SOLO per il
    tecnico: per l'utente e per ogni altra persona si firma col dito o a mano sul foglio."""
    return MODI_FIRMA.get(chiave, ("nessuna", "tenuta", "cartaceo"))


def applica_scelte_firme(cartella, nomi, cartaceo, scelte, salvata=None):
    """Applica le scelte fatte nel programma del PC sulle firme di un verbale.
    cartella = cartella del verbale; nomi = {posto: nome file della firma già presente};
    cartaceo = posti già segnati come «firma a mano»; scelte = {posto: modo} (vedi MODI_FIRMA);
    salvata = percorso dell'immagine di firma del tecnico (se c'è).
    Restituisce (nomi, cartaceo) aggiornati. Un modo non permesso al posto (esempio un'immagine per
    l'utente) viene rifiutato con VerbaleError e non cambia niente."""
    nomi = dict(nomi or {})
    cart = {k for k in (cartaceo or ()) if k in FIRME_CHIAVI}
    for chi, modo in (scelte or {}).items():
        if modo not in modi_firma_ammessi(chi):
            raise VerbaleError("Per questo posto firma il modo «%s» non è permesso: si può firmare "
                               "col dito sul telefono oppure a mano sul foglio." % modo)
    for chi, modo in (scelte or {}).items():
        if modo == "tenuta":
            if chi not in nomi:
                raise VerbaleError("Non c'è nessuna firma da tenere.")
            cart.discard(chi)
        elif modo == "salvata":
            if not salvata or not os.path.isfile(salvata):
                raise VerbaleError("Non c'è ancora un'immagine di firma salvata: scegline una in "
                                   "«Dati per i verbali».")
            name = "firma_%s.png" % chi
            os.makedirs(cartella, exist_ok=True)
            prepara_firma(salvata, cartella, "firma_%s" % chi)
            nomi[chi] = name
            cart.discard(chi)
        elif modo == "cartaceo":
            nomi.pop(chi, None)
            cart.add(chi)
        else:                                   # nessuna
            nomi.pop(chi, None)
            cart.discard(chi)
    return nomi, cart


# --------------------------------------------------------------------------
# Creazione del documento
# --------------------------------------------------------------------------
def nome_file_proposto(tipo, d):
    titolo = TIPI[tipo][0]
    chi = (d.get("assegnatario") or d.get("utente") or d.get("oggetto") or "").strip()
    chi = re.sub(r"\s*\(.*?\)", "", chi).strip()
    day = re.sub(r"[^0-9]", "-", (d.get("data") or "").strip())
    day = "-".join(reversed([x for x in day.split("-") if x])) if day else ""
    parts = ["Verbale " + titolo.lower()]
    if chi:
        parts.append(chi)
    if day:
        parts.append(day)
    name = " - ".join(parts)
    name = re.sub(r'[\\/:*?"<>|]+', " ", name)
    return re.sub(r"\s+", " ", name).strip()[:120] + ".docx"


def crea_verbale(tipo, d, out_path, templates_dir, photos=None, imp=None, firme=None,
                 avvisi=None, cartaceo=None):
    """Compila il modello e salva il Word in out_path. photos = elenco di file JPEG già pronti.
    firme = {"tecnico": file immagine, "utente": file immagine} (facoltative).
    cartaceo = chi firma a mano sul foglio (la riga resta vuota, con lo spazio per la firma).
    avvisi = lista a cui aggiungere le segnalazioni (esempio: modello senza posto per la firma)."""
    docx = _docx()
    photos = list(photos or [])
    if tipo not in TIPI:
        raise VerbaleError("Tipo di verbale sconosciuto: %s" % tipo)
    if TIPI[tipo][2] and not photos:
        raise VerbaleError("Per il %s la foto è obbligatoria: allega almeno una foto." %
                           TIPI[tipo][0].lower())
    values, lists = valori_verbale(tipo, d, imp)
    path = template_path(templates_dir, tipo)
    try:
        doc = docx.Document(path)
    except Exception as e:
        raise VerbaleError("Non riesco ad aprire il modello \"%s\":\n%s" % (path, e))

    if tipo == "presa_in_carico":
        following = _repeat_paragraph(doc, "MATERIALE", lists["MATERIALE"])
        # ogni voce scritta prende il posto di una delle righe vuote da compilare a mano
        n = min(len(lists["MATERIALE"]), 3)
        while following is not None and n > 0 and following.tag == _w("p") and (
                following.find(".//" + _w("drawing")) is not None or
                following.find(".//" + _w("pict")) is not None):
            nxt = following.getnext()
            following.getparent().remove(following)
            following, n = nxt, n - 1
    elif tipo == "assegnazione":
        _repeat_paragraph(doc, "DOT_TIPO", lists["DOT_TIPO"])
        acc = lists["ACCESSORIO"]
        proto = _find_paragraph(doc, "ACCESSORIO")
        if proto is not None:
            if not acc:
                prev = proto.getprevious()   # il titolo "Accessori inclusi"
                if prev is not None and prev.tag == _w("p"):
                    prev.getparent().remove(prev)
            _repeat_paragraph(doc, "ACCESSORIO", acc)
    elif tipo == "censimento":
        _repeat_table_row(doc, "N", lists["N"])

    tmp = None
    try:
        pronte = {}
        for chi, src in (firme or {}).items():
            if src and os.path.isfile(src):
                tmp = tmp or tempfile.mkdtemp(prefix="firme_")
                pronte[chi] = prepara_firma(src, tmp, chi)
        applica_firme(doc, pronte, avvisi, cartaceo)

        for p in _all_paragraphs(doc):
            if p.getparent() is None:
                continue
            if replace_markers(p, values):
                p.getparent().remove(p)

        if photos:
            _aggiungi_foto(doc, photos)
        out_dir = os.path.dirname(out_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        try:
            doc.save(out_path)
        except PermissionError:
            raise VerbaleError("Non riesco a salvare \"%s\": il file è aperto in Word? "
                               "Chiudilo e riprova." % out_path)
    finally:
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)
    return out_path


def nuovo_codice():
    return uuid.uuid4().hex[:10]
