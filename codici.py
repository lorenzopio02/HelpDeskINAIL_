"""Lettura di codici a barre e QR per il PC: fotocamera (webcam), foto da file, immagine dagli appunti.

Il lettore dei codici a barre (Code 128 e Code 39) e' lo stesso che l'app dell'iPhone ha in
docs/barcode.js, riscritto in Python: stesse tabelle, stessi controlli, stesse regole.
I QR code li legge OpenCV. Serve numpy (e opencv per la webcam e i QR): se non ci sono, il
programma funziona lo stesso e la lettura con la fotocamera dice che non e' disponibile.
Niente di questo modulo usa Tk: la finestra e' in helpdesk.py (ScanDialog)."""
import io
import math
import os
import re
import time
from urllib.parse import parse_qs, unquote, urlparse

# numpy e OpenCV si caricano solo alla prima lettura (non all'avvio del programma: sono pesanti).
np = None
cv2 = None
_caricato = False


def _load():
    global np, cv2, _caricato
    if _caricato:
        return
    _caricato = True
    try:
        import numpy as _np
        np = _np
    except Exception:                                # senza numpy niente lettura
        np = None
    try:
        import cv2 as _cv2
        cv2 = _cv2
    except Exception:                                # senza OpenCV niente webcam e niente QR
        cv2 = None


class CodiciError(Exception):
    pass


def disponibile():
    """True se la lettura dei codici a barre da immagini e' possibile."""
    _load()
    return np is not None


def webcam_disponibile():
    _load()
    return np is not None and cv2 is not None


# ---------------------------------------------------------------- tabelle
# Code 128: larghezze barra/spazio dei 107 simboli (0-105 dati e start, 106 = stop)
C128_TXT = ("212222,222122,222221,121223,121322,131222,122213,122312,132212,221213,221312,231212,112232,"
            "122132,122231,113222,123122,123221,223211,221132,221231,213212,223112,312131,311222,321122,"
            "321221,312212,322112,322211,212123,212321,232121,111323,131123,131321,112313,132113,132311,"
            "211313,231113,231311,112133,112331,132131,113123,113321,133121,313121,211331,231131,213113,"
            "213311,213131,311123,311321,331121,312113,312311,332111,314111,221411,431111,111224,111422,"
            "121124,121421,141122,141221,112214,112412,122114,122411,142112,142211,241211,221114,413111,"
            "241112,134111,111242,121142,121241,114212,124112,124211,411212,421112,421211,212141,214121,"
            "412121,111143,111341,131141,114113,114311,411113,411311,113141,114131,311141,411131,211412,"
            "211214,211232,2331112")
C128 = [[int(c) for c in s] for s in C128_TXT.split(",")]
# Code 39: per ogni carattere i 9 elementi (barra, spazio, ...), 1 = largo e 0 = stretto
C39 = {"0": "000110100", "1": "100100001", "2": "001100001", "3": "101100000", "4": "000110001",
       "5": "100110000", "6": "001110000", "7": "000100101", "8": "100100100", "9": "001100100",
       "A": "100001001", "B": "001001001", "C": "101001000", "D": "000011001", "E": "100011000",
       "F": "001011000", "G": "000001101", "H": "100001100", "I": "001001100", "J": "000011100",
       "K": "100000011", "L": "001000011", "M": "101000010", "N": "000010011", "O": "100010010",
       "P": "001010010", "Q": "000000111", "R": "100000110", "S": "001000110", "T": "000010110",
       "U": "110000001", "V": "011000001", "W": "111000000", "X": "010010001", "Y": "110010000",
       "Z": "011010000", "-": "010000101", ".": "110000100", " ": "011000100", "*": "010010100",
       "$": "010101000", "/": "010100010", "+": "010001010", "%": "000101010"}
C39R = {v: k for k, v in C39.items()}
_C39_STAR = int(C39["*"], 2)
_C39_INT = {int(v, 2): k for k, v in C39.items()}


def _tables():
    """Prepara (una volta) le tabelle numeriche."""
    _load()
    if _tables.cache is None:
        p128 = np.array(C128[:106], dtype=np.float64)          # 0..105 (6 elementi)
        stop = np.array(C128[106], dtype=np.float64)
        _tables.cache = (p128, stop)
    return _tables.cache


_tables.cache = None


# ---------------------------------------------------------------- da una riga di pixel alle larghezze
def line_runs(s):
    """s: luminosita' 0..255 lungo una riga. Restituisce (runs, c0): larghezze delle fasce e se la
    fascia 0 e' scura. La prima e l'ultima fascia sono tagliate dal bordo."""
    _load()
    n = len(s)
    if n < 40:
        return None
    s = np.asarray(s, dtype=np.float64)
    a = np.empty(n)
    a[0] = (3 * s[0] + s[1]) / 4
    a[-1] = (s[-2] + 3 * s[-1]) / 4
    a[1:-1] = (s[:-2] + 2 * s[1:-1] + s[2:]) / 4
    B, R = 16, 4
    nb = -(-n // B)
    pad = nb * B - n
    bmin = np.concatenate([a, np.full(pad, 255.0)]).reshape(nb, B).min(axis=1)
    bmax = np.concatenate([a, np.zeros(pad)]).reshape(nb, B).max(axis=1)
    pm = np.concatenate([np.full(R, 255.0), bmin, np.full(R, 255.0)])
    pM = np.concatenate([np.zeros(R), bmax, np.zeros(R)])
    lo = np.min([pm[k:k + nb] for k in range(2 * R + 1)], axis=0)
    hi = np.max([pM[k:k + nb] for k in range(2 * R + 1)], axis=0)
    thr, con = (lo + hi) / 2, hi - lo
    blk = np.arange(n) // B
    ok = con[blk] >= 28
    d = ok & (a < thr[blk])
    chg = np.nonzero(d[1:] != d[:-1])[0] + 1             # indici i dove cambia colore
    if len(chg) < 20:
        return None
    j = chg - 1
    bi, bj = blk[chg], blk[j]
    den = a[chg] - a[j]
    good = ok[chg] & ok[j] & (den != 0)
    t = np.where(good, (thr[bi] - a[j]) / np.where(den == 0, 1, den), 0.5)
    t = np.clip(t, 0.0, 1.0)
    pos = (j + t).tolist()
    ed = []
    for p in pos:                                        # disturbo: fascia sotto 1 pixel
        if ed and p - ed[-1] < 1.0:
            ed.pop()
        else:
            ed.append(p)
    m = len(ed)
    if m < 20:
        return None
    runs = np.empty(m + 1)
    runs[0] = ed[0]
    runs[1:m] = np.diff(ed)
    runs[m] = n - ed[-1]
    return runs, bool(d[0])


# ---------------------------------------------------------------- Code 128
def _cand128(r, i, lo_v, hi_v, tab):
    """Il simbolo lo_v..hi_v piu' simile alle 6 fasce da i; -1 se nessuno e' abbastanza simile."""
    w = np.asarray(r[i:i + 6], dtype=np.float64)
    T = w.sum()
    if T < 6:
        return -1
    m = T / 11
    diff = np.abs(tab[lo_v:hi_v + 1] - w / m)
    e, mx = diff.sum(axis=1), diff.max(axis=1)
    e = np.where((e <= 2.4) & (mx <= 0.9), e, np.inf)
    k = int(np.argmin(e))
    return lo_v + k if np.isfinite(e[k]) else -1


def _is_stop128(r, i, stop):
    w = np.asarray(r[i:i + 7], dtype=np.float64)
    m = w.sum() / 13
    if m < 0.5:
        return False
    d = np.abs(w / m - stop)
    return d.sum() <= 2.2 and d.max() <= 0.85 and r[i + 7] >= 2.5 * m


def _text128(vals):
    st = {103: "A", 104: "B"}.get(vals[0], "C")
    cur_set, shift, out = st, "", []
    for v in vals[1:]:
        cur = shift or cur_set
        shift = ""
        if cur == "C":
            if v < 100:
                out.append("%02d" % v)
            elif v == 100:
                cur_set = "B"
            elif v == 101:
                cur_set = "A"
            continue
        if v < 96:
            out.append(chr((v + 32 if v < 64 else v - 64) if cur == "A" else v + 32))
            continue
        if v == 98:
            shift = "B" if cur_set == "A" else "A"
            continue
        if v == 99:
            cur_set = "C"
            continue
        if cur == "A" and v == 100:
            cur_set = "B"
        elif cur == "B" and v == 101:
            cur_set = "A"
    return "".join(out)


def decode128(r, c0):
    """r: lista di larghezze (con la prima e l'ultima tagliate). Restituisce i codici Code 128 letti."""
    p128, stop = _tables()
    n = len(r)
    out = []
    first = 2 if c0 else 1
    if n < first + 25:
        return out
    arr = np.asarray(r, dtype=np.float64)
    # tutte le possibili partenze in un colpo solo (start A, B, C)
    last = n - 2 - 24
    idx = np.arange(first, last + 1, 2)
    if len(idx) == 0:
        return out
    win = np.lib.stride_tricks.sliding_window_view(arr, 6)[idx]
    T = win.sum(axis=1)
    okT = T >= 6
    norm = win * (11.0 / np.where(okT, T, 1))[:, None]
    diff = np.abs(norm[:, None, :] - p128[103:106][None, :, :])
    e, mx = diff.sum(axis=2), diff.max(axis=2)
    e = np.where((e <= 2.4) & (mx <= 0.9) & okT[:, None], e, np.inf)
    best = np.argmin(e, axis=1)
    cands = [(int(idx[k]), 103 + int(best[k])) for k in np.nonzero(np.isfinite(e[np.arange(len(idx)), best]))[0]]
    skip = -1
    for i, sv in cands:
        if i <= skip:
            continue
        vals, pos, found = [sv], i + 6, None
        while True:
            if len(vals) >= 5 and pos + 7 <= n - 1 and _is_stop128(arr, pos, stop):
                chk = vals[-1]
                total = vals[0] + sum(k * vals[k] for k in range(1, len(vals) - 1))
                if total % 103 == chk:
                    found = (vals[:-1], pos)
                    break
            if pos + 6 > n - 2 or len(vals) >= 110:
                break
            c = _cand128(arr, pos, 0, 102, p128)
            if c < 0:
                break
            vals.append(c)
            pos += 6
        if found:
            text = re.sub(r"[\x00-\x1f\x7f]", "", _text128(found[0]))
            if text:
                out.append({"text": text, "format": "code128", "n": len(found[0]) - 1})
            skip = found[1]
    return out


# ---------------------------------------------------------------- Code 39
def _char39(w):
    """Le 9 larghezze di un carattere -> (carattere, larghezza dello stretto) oppure None."""
    order = sorted(range(9), key=lambda k: -w[k])
    wide_min, narrow_max = w[order[2]], w[order[3]]
    if narrow_max <= 0 or wide_min < 1.12 * narrow_max:
        return None
    wa = sum(w[k] for k in order[:3]) / 3
    na = sum(w[k] for k in order[3:]) / 6
    if na < 1 or wa / na < 1.7 or wa / na > 3.8:
        return None
    mask = 0
    for k in order[:3]:
        mask |= 1 << (8 - k)
    ch = _C39_INT.get(mask)
    return (ch, na) if ch else None


def decode39(r, c0):
    n = len(r)
    out = []
    first = 2 if c0 else 1
    if n < first + 10:
        return out
    arr = np.asarray(r, dtype=np.float64)
    last = n - 2 - 8
    idx = np.arange(first, last + 1, 2)
    if len(idx) == 0:
        return out
    win = np.lib.stride_tricks.sliding_window_view(arr, 9)[idx]
    srt = -np.sort(-win, axis=1)
    with np.errstate(divide="ignore", invalid="ignore"):
        wa, na = srt[:, :3].mean(axis=1), srt[:, 3:].mean(axis=1)
        pre = (srt[:, 3] > 0) & (srt[:, 2] >= 1.12 * srt[:, 3]) & (na >= 1) & (wa / na >= 1.7) & (wa / na <= 3.8)
    star_w = win[:, [1, 4, 6]].min(axis=1) > np.delete(win, [1, 4, 6], axis=1).max(axis=1)
    cand = idx[pre & star_w]
    skip = -1
    for i in cand.tolist():
        if i <= skip:
            continue
        s0 = _char39(arr[i:i + 9].tolist())
        if not s0 or s0[0] != "*":
            continue
        nr = s0[1]
        if arr[i - 1] < 3 * nr:
            continue
        pos, txt, ok = i + 9, "", False
        while pos + 9 <= n - 2 and len(txt) <= 60:
            gap = arr[pos]
            if gap > 3 * nr or gap < 0.3 * nr:
                break
            c = _char39(arr[pos + 1:pos + 10].tolist())
            if not c or c[1] < 0.55 * nr or c[1] > 1.8 * nr:
                break
            if c[0] == "*":
                ok = len(txt) >= 4 and arr[pos + 10] >= 3 * nr
                break
            txt += c[0]
            pos += 10
        if ok:
            out.append({"text": txt, "format": "code39", "n": len(txt)})
            skip = pos
    return out


# ---------------------------------------------------------------- una riga, un'immagine
def decode_line(s, formats=None):
    lr = line_runs(s)
    if not lr:
        return []
    runs, c0 = lr
    want = lambda f: formats is None or f in formats
    fw = runs.tolist()
    rv = fw[::-1]
    c0r = ((len(fw) - 1) % 2 == 0) == c0
    out = []
    for r, c in ((fw, c0), (rv, c0r)):
        if want("code128"):
            out.extend(decode128(r, c))
        if want("code39"):
            out.extend(decode39(r, c))
    return out


def _clip(px, py, dx, dy, W, H):
    t0, t1 = -1e9, 1e9
    for p, d, M in ((px, dx, W), (py, dy, H)):
        if abs(d) < 1e-9:
            if p < 0 or p > M:
                return None
            continue
        ta, tb = (0 - p) / d, (M - p) / d
        if ta > tb:
            ta, tb = tb, ta
        t0, t1 = max(t0, ta), min(t1, tb)
    return (t0, t1) if t1 > t0 else None


ALL_ANGLES = [0, 90, 15, -15, 30, -30, 45, -45, 60, -60, 75, -75]


def decode_gray(gray, angles=None, lines=16, formats=None):
    """gray: matrice (alto x largo) di grigi. Restituisce i codici trovati ordinati: ognuno ha
    text, format, n (simboli) e votes (da quante righe e' stato letto)."""
    _load()
    if np is None:
        return []
    g = np.asarray(gray, dtype=np.float32)
    h, w = g.shape
    W, H = w - 1, h - 1
    cx, cy = W / 2, H / 2
    flat = g.ravel()
    found = {}
    for deg in (ALL_ANGLES if angles is None else angles):
        th = math.radians(deg)
        dx, dy = math.cos(th), math.sin(th)
        nx, ny = -dy, dx
        E = (w * abs(dy) + h * abs(dx)) / 2
        for j in range(lines):
            o = -E + (j + 0.5) * 2 * E / lines
            px, py = cx + o * nx, cy + o * ny
            c = _clip(px, py, dx, dy, W, H)
            if not c:
                continue
            x0, y0 = px + c[0] * dx, py + c[0] * dy
            n = int(math.floor(c[1] - c[0]))
            if n < 40:
                continue
            t = np.arange(n)
            x = np.clip(x0 + dx * t, 0, W)
            y = np.clip(y0 + dy * t, 0, H)
            xi, yi = x.astype(np.int32), y.astype(np.int32)
            fx, fy = x - xi, y - yi
            x1, y1 = np.minimum(xi + 1, W), np.minimum(yi + 1, H)
            s = ((flat[yi * w + xi] * (1 - fx) + flat[yi * w + x1] * fx) * (1 - fy)
                 + (flat[y1 * w + xi] * (1 - fx) + flat[y1 * w + x1] * fx) * fy)
            for r in decode_line(s, formats):
                key = (r["format"], r["text"])
                if key in found:
                    found[key]["votes"] += 1
                else:
                    found[key] = {"text": r["text"], "format": r["format"], "n": r["n"], "votes": 1}
    return sorted(found.values(), key=lambda r: (-r["votes"], -r["n"]))


def to_gray(rgb):
    """Immagine (PIL o matrice RGB/grigi) -> matrice di grigi uint8."""
    _load()
    if not isinstance(rgb, np.ndarray):
        rgb = np.asarray(rgb.convert("RGB"))
    if rgb.ndim == 2:
        return rgb.astype(np.uint8)
    if rgb.shape[2] == 4:
        rgb = rgb[:, :, :3]
    return ((rgb[:, :, 0].astype(np.uint32) * 77 + rgb[:, :, 1].astype(np.uint32) * 150
             + rgb[:, :, 2].astype(np.uint32) * 29) >> 8).astype(np.uint8)


def _resize(img, w, h):
    _load()
    if cv2 is not None:
        return cv2.resize(img, (w, h), interpolation=cv2.INTER_AREA if w < img.shape[1] else cv2.INTER_LINEAR)
    from PIL import Image
    return np.asarray(Image.fromarray(img).resize((w, h), Image.LANCZOS if w < img.shape[1] else Image.BILINEAR))


# ---------------------------------------------------------------- controllo dei codici letti
def luhn_ok(s):
    total = 0
    for i, ch in enumerate(reversed(s)):
        d = int(ch)
        if i % 2 == 1:
            d *= 2
            if d > 9:
                d -= 9
        total += d
    return total % 10 == 0


def imei_ok(s):
    return bool(re.fullmatch(r"\d{15}", s)) and luhn_ok(s)


_SERIAL_KEYS = ("servicetag", "service_tag", "service-tag", "st", "sn", "serial", "serialnumber", "serial_number",
                "serialno", "s/n", "seriale", "snr", "sno", "productserial")


def serial_from_qr(text):
    """Da un QR di un'etichetta ricava il seriale: dall'indirizzo web (parametro o ultimo pezzo) oppure da
    'S/N: ...'. Se non riesce restituisce ''. Euristica: non provata su etichette vere."""
    t = (text or "").strip()
    if not t:
        return ""
    if re.match(r"^https?://", t, re.I):
        try:
            u = urlparse(t)
            q = {k.lower(): v for k, v in parse_qs(u.query).items()}
            for k in _SERIAL_KEYS:
                if k in q and q[k] and q[k][0].strip():
                    return q[k][0].strip()
            seg = unquote(u.path.rstrip("/").split("/")[-1])
            if re.fullmatch(r"[A-Za-z0-9\-]{6,24}", seg) and re.search(r"\d", seg) and re.search(r"[A-Za-z]", seg):
                return seg
        except Exception:
            pass
        return ""
    m = re.search(r"(?:s/?n|serial(?:\s*number)?|seriale|service\s*tag)\s*[:=#]?\s*([A-Za-z0-9\-]{4,})", t, re.I)
    if m:
        return m.group(1)
    return t if len(t) <= 80 and "\n" not in t else ""


def norm(text, kind, fmt="code128"):
    """Cosa vale come risposta per quel campo ('serial', 'imei' o 'any'); '' = da scartare."""
    t = re.sub(r"[\x00-\x1f\x7f]", "", "" if text is None else str(text)).strip()
    if fmt == "qr" and kind in ("serial", "imei"):
        t = serial_from_qr(t)
        if not t and kind == "imei":
            t = re.sub(r"[\x00-\x1f\x7f]", "", str(text)).strip()
    if not t:
        return ""
    if kind == "imei":
        t = re.sub(r"\s+", "", re.sub(r"^imei[\s:\-]*", "", t, flags=re.I))
        return t if imei_ok(t) else ""
    if kind == "serial":
        t = re.sub(r"\s+", "", t)
        return t if 4 <= len(t) <= 40 else ""
    return t if len(t) <= 80 else ""


def pick(lst, kind, single=False):
    """Tra i codici letti in una foto o in un fotogramma sceglie quello giusto per il campo."""
    for r in lst:
        t = norm(r["text"], kind, r["format"])
        if not t:
            continue
        strong = r["votes"] >= 2 or (r["format"] == "code128" and r["n"] >= 4) or r["format"] == "qr"
        if strong or (single and r["n"] >= 4):
            return {"text": t, "format": r["format"], "votes": r["votes"]}
    return None


# ---------------------------------------------------------------- QR (OpenCV)
def new_qr_detector():
    _load()
    if cv2 is None:
        return None
    try:
        return cv2.QRCodeDetectorAruco()
    except Exception:
        return cv2.QRCodeDetector()


def decode_qr(img, det=None):
    """img: matrice RGB o grigi. Restituisce i testi dei QR trovati (lista, anche vuota)."""
    _load()
    if cv2 is None or np is None:
        return []
    det = det or new_qr_detector()
    bgr = cv2.cvtColor(img, cv2.COLOR_RGB2BGR) if img.ndim == 3 else img
    texts = []
    try:
        res = det.detectAndDecodeMulti(bgr)
        texts = [t for t in (res[1] or ()) if t]
    except Exception:
        texts = []
    if not texts:
        try:
            t = det.detectAndDecode(bgr)[0]
            if t:
                texts = [t]
        except Exception:
            pass
    return texts


def qr_results(img, det=None):
    return [{"text": t, "format": "qr", "n": 9, "votes": 2} for t in decode_qr(img, det)]


# ---------------------------------------------------------------- foto ferma
def decode_image(img, kind="serial", qr=True):
    """Cerca un codice in una foto (PIL o matrice), a piu' ingrandimenti. Restituisce
    {'text','format','votes'} oppure None."""
    _load()
    if np is None:
        raise CodiciError("manca numpy")
    if isinstance(img, np.ndarray):
        rgb = img if img.ndim == 3 else np.stack([img] * 3, axis=2)
    else:
        rgb = np.asarray(img.convert("RGB"))
    rgb = rgb[:, :, :3]
    h0, w0 = rgb.shape[:2]
    gray = to_gray(rgb)
    tried = set()
    det = new_qr_detector() if (qr and cv2 is not None) else None
    for L in (1600, 2400, 1000, 700):
        sc = min(1.0, L / max(w0, h0))
        w, h = max(40, round(w0 * sc)), max(40, round(h0 * sc))
        if (w, h) in tried:
            continue
        tried.add((w, h))
        g = gray if (w, h) == (w0, h0) else _resize(gray, w, h)
        got = pick(decode_gray(g, lines=22), kind, True)
        if got:
            return got
        if det is not None:
            im = rgb if (w, h) == (w0, h0) else _resize(rgb, w, h)
            got = pick(qr_results(im, det), kind, True)
            if got:
                return got
    return None


def load_image(path):
    from PIL import Image, ImageOps
    im = Image.open(path)
    try:
        im = ImageOps.exif_transpose(im)
    except Exception:
        pass
    return im.convert("RGB")


# ---------------------------------------------------------------- lettura dal vivo
class LiveReader:
    """Le stesse regole dell'app dell'iPhone: un fotogramma alla volta, e il codice viene accettato quando
    e' affidabile (Code 128 lungo subito, gli altri dopo due conferme). Con multi=True resta in ascolto del
    prossimo codice e non riaggiunge quello ancora in vista."""
    EXTRA = ([15, -15, 30, -30], [45, -45, 60, -60], [75, -75, 22, -22])

    def __init__(self, kind="serial", multi=False, qr=True, max_side=1600):
        _load()
        self.kind, self.multi, self.qr, self.max_side = kind, multi, qr, max_side
        self.seen = {}
        self.last = ""
        self.last_seen = 0.0
        self.count = 0
        self.rej = ""
        self.tick = 0
        self.t0 = time.time()
        self.det = new_qr_detector() if (qr and cv2 is not None) else None

    def read_frame(self, rgb):
        """Cerca i codici in un fotogramma RGB. Restituisce la lista dei codici (come decode_gray)."""
        h, w = rgb.shape[:2]
        sc = min(1.0, self.max_side / max(w, h))
        if sc < 1.0:
            rgb = _resize(rgb, round(w * sc), round(h * sc))
        gray = to_gray(rgb)
        extra = self.EXTRA[self.tick % len(self.EXTRA)]
        res = decode_gray(gray, angles=[0, 90] + list(extra), lines=18)
        if self.det is not None and (self.tick % 2 == 0 or not res):
            res = res + qr_results(rgb, self.det)
        self.tick += 1
        return res

    def feed(self, results, now=None):
        """Valuta i codici di un fotogramma. Restituisce il testo accettato oppure None."""
        now = time.time() if now is None else now
        for k in [k for k, v in self.seen.items() if now - v[1] > 2.0]:
            del self.seen[k]
        for r in results:
            t = norm(r["text"], self.kind, r["format"])
            if not t:
                if r["n"] >= 5 and r["votes"] >= 2:
                    self.rej = str(r["text"])[:30]
                continue
            n, _ = self.seen.get(t, (0, 0))
            self.seen[t] = (n + 1, now)
            strong = r["votes"] >= 3 or (r["format"] == "code128" and r["n"] >= 4 and r["votes"] >= 1) \
                or r["format"] == "qr"
            if not (strong or n + 1 >= 2):
                continue
            if self.multi and t == self.last and now - self.last_seen < 1.5:
                self.last_seen = now                     # lo stesso codice ancora in vista
                continue
            self.seen.clear()
            self.last, self.last_seen = t, now
            self.count += 1
            return t
        return None

    def hint(self, now=None):
        """Il messaggio da mostrare quando non e' stato accettato niente (o '')."""
        now = time.time() if now is None else now
        if self.kind == "imei" and self.rej and not self.multi:
            return "Letto «%s», ma non è un IMEI valido (15 cifre). Cerca il codice dell'IMEI." % self.rej
        if not self.count and now - self.t0 > 12:
            return ("Non lo leggo ancora: avvicina il codice, tienilo dritto, fermo e ben illuminato "
                    "(la webcam di un portatile mette a fuoco da lontano: provalo a 20-30 cm). "
                    "Oppure scegli una foto.")
        return ""


# ---------------------------------------------------------------- fotocamera
class Camera:
    """Una webcam (indice 0, 1, ...) oppure un file video (per le prove). read() dà un'immagine RGB."""

    def __init__(self, source=0, loop=None):
        if not webcam_disponibile():
            raise CodiciError("Questa versione del programma non ha la lettura con la fotocamera "
                              "(manca OpenCV). Puoi scegliere una foto oppure scrivere a mano.")
        self.source = source
        self.loop = isinstance(source, str) if loop is None else loop
        if isinstance(source, int) and os.name == "nt":
            cap = cv2.VideoCapture(source, cv2.CAP_DSHOW)
        else:
            cap = cv2.VideoCapture(source)
        if cap is None or not cap.isOpened():
            raise CodiciError("Non trovo una fotocamera%s." % (" (numero %d)" % source if isinstance(source, int) else ""))
        if isinstance(source, int):
            try:
                cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*"MJPG"))
                cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1920)
                cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 1080)
                cap.set(cv2.CAP_PROP_AUTOFOCUS, 1)
            except Exception:
                pass
        self.cap = cap

    def read(self):
        ok, frame = self.cap.read()
        if not ok and self.loop:
            self.cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
            ok, frame = self.cap.read()
        if not ok or frame is None:
            return None
        return cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

    def close(self):
        try:
            if self.cap is not None:
                self.cap.release()
        finally:
            self.cap = None


def preview_png_b64(rgb, width=640, guide=True):
    """Il fotogramma ridotto a 'width' pixel, con il riquadro di guida, come PNG in base64 (per tk.PhotoImage)."""
    import base64
    from PIL import Image
    _load()
    h0, w0 = rgb.shape[:2]
    h = max(1, round(h0 * width / w0))
    a = _resize(np.ascontiguousarray(rgb[:, :, :3]), width, h) if w0 != width else np.ascontiguousarray(rgb[:, :, :3]).copy()
    if guide:
        top, bot = int(h * 0.3), int(h * 0.7)
        left, right = int(width * 0.07), int(width * 0.93)
        a[top:top + 2, left:right] = 255
        a[bot - 2:bot, left:right] = 255
        a[top:bot, left:left + 2] = 255
        a[top:bot, right - 2:right] = 255
        mid = h // 2
        a[mid - 1:mid + 1, left + 12:right - 12] = (209, 34, 51)
    buf = io.BytesIO()
    Image.fromarray(a).save(buf, "PNG", compress_level=1)
    return base64.b64encode(buf.getvalue()).decode("ascii")
