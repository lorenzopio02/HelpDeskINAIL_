# -*- coding: utf-8 -*-
"""
Visite 3D per il PC (le stesse del telefono, adattate a mouse e tastiera):
  - Passeggiata nel piano: "Dall'alto" (il piano visto come una casa delle bambole) e "Cammina"
    (in prima persona nel corridoio: ti giri a destra e a sinistra e sulle porte vedi le targhette
    con il numero della stanza, l'ufficio e il nome, come quelle vere; mini-mappa e, se vuoi, la tua faccia);
  - Edificio in 3D: i piani impilati; si gira intorno al palazzo e si clicca un piano per camminarci dentro.

Tk non ha il 3D: qui c'e' un piccolo motore che disegna con Pillow (poligoni in prospettiva e targhette
deformate come sulla parete). Tutto e' disegnato dai dati che l'app ha gia' (stanze, persone, stampanti).

La foto del personaggio (facoltativa) resta SOLO nella cartella dei dati di questo PC: non finisce nel
backup, ne' nei file per il telefono, ne' su GitHub.
"""
import base64
import io
import math
import os
import re
import time

from PIL import Image, ImageDraw, ImageFont

try:                                     # senza schermo (prove) il motore di disegno funziona lo stesso
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
except Exception:                        # pragma: no cover
    tk = ttk = filedialog = messagebox = None

# ---------------------------------------------------------------- misure del mondo (come sul telefono)
CW, RW, RD, H, EYE, WH = 120, 190, 160, 140, 70, 30       # corridoio, stanza, muri, occhi
DOORW, DOORH, PLW, PLH = 62, 118, 64, 110                 # porta e targhetta
FD, FPL = 0.30, 0.74                                      # posizione lungo il muro (0 = entrata del piano)
SPEED, TURN = 170.0, 85.0                                 # camminata (unita'/s) e rotazione (gradi/s)
FLOOR_Y, CEIL_Y = -EYE, H - EYE                           # y verso l'alto; 0 = altezza degli occhi
NEAR = 3.0
FOV_WALK, FOV_DH, FOV_TW = 72.0, 46.0, 40.0
AV_PX = 96
PLATE_ANGLE, PLATE_DIST = 38.0, 230.0                     # la targhetta "che sto guardando"

AVATAR_NAME = "la_mia_faccia.png"


def clamp(v, a, b):
    return max(a, min(b, v))


def _mix(c1, c2, t):
    t = clamp(t, 0.0, 1.0)
    return (int(c1[0] + (c2[0] - c1[0]) * t), int(c1[1] + (c2[1] - c1[1]) * t),
            int(c1[2] + (c2[2] - c1[2]) * t))


# ---------------------------------------------------------------- caratteri
_FONTS = {}


def get_font(size, bold=False):
    key = (int(size), bool(bold))
    f = _FONTS.get(key)
    if f is None:
        names = (("arialbd.ttf", "Arial Bold.ttf", "segoeuib.ttf", "DejaVuSans-Bold.ttf",
                  "LiberationSans-Bold.ttf") if bold else
                 ("arial.ttf", "Arial.ttf", "segoeui.ttf", "DejaVuSans.ttf",
                  "LiberationSans-Regular.ttf"))
        for nm in names:
            try:
                f = ImageFont.truetype(nm, key[0])
                break
            except Exception:
                f = None
        if f is None:
            try:
                f = ImageFont.load_default(size=key[0])
            except TypeError:
                f = ImageFont.load_default()
        _FONTS[key] = f
    return f


def _tw(draw, text, font):
    try:
        return draw.textlength(text, font=font)
    except Exception:
        return draw.textbbox((0, 0), text, font=font)[2]


def _text_center(draw, cx, cy, text, font, fill):
    try:
        l, t, r, b = draw.textbbox((0, 0), text, font=font)
    except Exception:
        l, t, r, b = 0, 0, _tw(draw, text, font), 10
    draw.text((cx - (l + r) / 2.0, cy - (t + b) / 2.0), text, font=font, fill=fill)


def _fit_line(draw, text, font, maxw):
    if _tw(draw, text, font) <= maxw:
        return text
    while len(text) > 1 and _tw(draw, text + "…", font) > maxw:
        text = text[:-1]
    return text.rstrip() + "…"


def _wrap(draw, text, font, maxw, max_lines):
    words = str(text).split()
    lines, cur = [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if _tw(draw, t, font) <= maxw or not cur:
            cur = t
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] += "…"
    return [_fit_line(draw, ln, font, maxw) for ln in lines]


# ---------------------------------------------------------------- layout del piano (identico al telefono)
def layout(rooms):
    """Le stanze di un piano: numeri dispari da una parte, pari dall'altra (come nei corridoi veri);
    senza numeri si alterna. z = lungo il corridoio (0 = entrata, in fondo negativo)."""
    par = bool(rooms) and all(re.search(r"[0-9]", str(r["codice"])) for r in rooms)
    cnt = {-1: 0, 1: 0}
    items = []
    for i, r in enumerate(rooms):
        if par:
            m = re.search(r"([0-9]+)[^0-9]*$", str(r["codice"])) or re.search(r"([0-9]+)", str(r["codice"]))
            side = -1 if int(m.group(1)) % 2 else 1
        else:
            side = 1 if i % 2 else -1
        col = cnt[side]
        cnt[side] += 1
        items.append({"r": r, "side": side, "col": col, "z": -(col + 0.5) * RW,
                      "x": side * (CW / 2.0 + RD / 2.0)})
    cols = max(1, cnt[-1], cnt[1])
    return {"items": items, "cols": cols, "LC": cols * RW}


def plate_data(room):
    """Quello che c'e' scritto sulla targhetta di una stanza."""
    ppl = room.get("persone") or []
    uff = []
    for p in ppl:
        u = str(p.get("ufficio") or "").strip()
        if u and u not in uff:
            uff.append(u)
    names = [str(p.get("nome") or "") for p in ppl]
    return {"uid": room["uid"], "num": str(room["codice"]), "uff": uff[:2], "names": names[:3],
            "more": max(0, len(names) - 3), "n": len(names), "free": not names,
            "miss": int(room.get("miss") or 0), "tot": max(0, int(room.get("scrivanie") or 0))}


def plate_text(pd):
    s = pd["num"]
    if pd["uff"]:
        s += " · " + ", ".join(pd["uff"])
    s += " · libera" if pd["free"] else " · " + ", ".join(pd["names"]) + (" +%d" % pd["more"] if pd["more"] else "")
    return s


# ---------------------------------------------------------------- la targhetta (immagine)
PL_W, PL_H = 256, 440
C_NAVY, C_PLATE, C_TXT, C_GREY = (11, 26, 74), (227, 230, 234), (29, 43, 74), (106, 118, 134)
C_EDGE = (154, 165, 177)
C_OK, C_NO, C_NONE = (26, 154, 74), (211, 32, 47), (154, 165, 177)


def plate_image(pd):
    """La targhetta come quelle vere: sopra il pannello blu con il numero bianco in grassetto,
    sotto quello chiaro con l'ufficio in grassetto e i nomi. RGBA 256x440 (angoli arrotondati)."""
    W, Hh = PL_W, PL_H
    im = Image.new("RGBA", (W, Hh), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, W - 1, Hh - 1), radius=12, fill=C_PLATE + (255,))
    top = int(Hh * 4.2 / 11.0)
    d.rounded_rectangle((0, 0, W - 1, top + 12), radius=12, fill=C_NAVY + (255,))
    d.rectangle((0, top - 6, W - 1, top), fill=C_NAVY + (255,))
    # numero
    num = pd["num"]
    size = 132
    while size > 26:
        f = get_font(size, True)
        if _tw(d, num, f) <= W * 0.84 and size <= top * 0.74:
            break
        size -= 4
    _text_center(d, W / 2.0, top / 2.0 + 2, num, get_font(size, True), (255, 255, 255, 255))
    # pallino: tutto a posto (verde) / mancano codici (rosso) / stanza libera (grigio)
    col = C_NONE if pd["free"] else (C_NO if pd["miss"] else C_OK)
    d.ellipse((W - 42, 12, W - 14, 40), fill=col + (255,), outline=(255, 255, 255, 255), width=3)
    # ufficio e nomi (la scrittura si rimpicciolisce finche' ci sta tutto)
    y0, avail, maxw = top + 16, Hh - top - 16 - 14, W - 34
    lines = []
    for size in range(27, 11, -1):
        fb, fr = get_font(size, True), get_font(size - 1, False)
        lines = []                                   # (testo, carattere, colore, corpo, aria prima)
        for u in pd["uff"]:
            lines += [(t, fb, C_TXT, size, 0) for t in _wrap(d, u, fb, maxw, 2)]
        first = 10 if pd["uff"] else 0
        if pd["free"]:
            lines.append(("Libera", fr, C_GREY, size - 1, first))
        else:
            for nm in pd["names"]:
                for t in _wrap(d, nm, fr, maxw, 2):
                    lines.append((t, fr, C_TXT, size - 1, first))
                    first = 0
            if pd["more"]:
                lines.append(("+%d" % pd["more"], fr, C_TXT, size - 1, 0))
        if sum(int(ln[3] * 1.24) + ln[4] for ln in lines) <= avail:
            break
    y = y0
    for t, f, c, sz, gap in lines:
        y += gap
        if y + sz > Hh - 8:
            break
        d.text((17, y), t, font=f, fill=c + (255,))
        y += int(sz * 1.24)
    d.rounded_rectangle((0, 0, W - 1, Hh - 1), radius=12, outline=C_EDGE + (255,), width=5)
    return im


def flat(im, bg=C_EDGE):
    out = Image.new("RGB", im.size, bg)
    out.paste(im, (0, 0), im)
    return out


def make_mips(rgb, levels=4):
    out = [rgb]
    w, h = rgb.size
    for _ in range(levels - 1):
        w, h = max(1, w // 2), max(1, h // 2)
        out.append(rgb.resize((w, h), Image.LANCZOS))
    return out


def end_image(title, subtitle, kind="exit"):
    """La parete in fondo al corridoio (USCITA) o quella dell'entrata."""
    W, Hh = 240, 280
    im = Image.new("RGB", (W, Hh), (214, 227, 241))
    d = ImageDraw.Draw(im)
    for y in range(Hh):
        k = y / float(Hh)
        d.line((0, y, W, y), fill=_mix((228, 238, 248), (201, 219, 238), k))
    f = get_font(30, True)
    t = _fit_line(d, title, f, W - 20)
    _text_center(d, W / 2.0, Hh * 0.13, t, f, C_NAVY)
    f2 = get_font(17, False)
    _text_center(d, W / 2.0, Hh * 0.25, _fit_line(d, subtitle, f2, W - 20), f2, C_NAVY)
    if kind == "exit":
        d.rounded_rectangle((W / 2 - 62, Hh * 0.33, W / 2 + 62, Hh * 0.33 + 36), radius=6, fill=(26, 154, 74))
        _text_center(d, W / 2.0, Hh * 0.33 + 18, "USCITA ›", get_font(21, True), (255, 255, 255))
    d.rectangle((0, 0, W - 1, Hh - 1), outline=(120, 140, 165), width=3)
    return im


def label_image(text, bg=(11, 26, 74), fg=(255, 255, 255), size=30, pad=14):
    d0 = ImageDraw.Draw(Image.new("RGB", (4, 4)))
    f = get_font(size, True)
    w = int(_tw(d0, text, f)) + pad * 2
    h = int(size * 1.5)
    im = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((0, 0, w - 1, h - 1), radius=10, fill=bg + (255,), outline=(255, 255, 255, 255), width=3)
    _text_center(d, w / 2.0, h / 2.0, text, f, fg)
    return im


def circle_avatar(img, size):
    im = img.convert("RGBA").resize((size, size), Image.LANCZOS)
    mask = Image.new("L", (size * 4, size * 4), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, size * 4 - 1, size * 4 - 1), fill=255)
    mask = mask.resize((size, size), Image.LANCZOS)
    im.putalpha(mask)
    return im


def save_avatar(src_path, dest_path):
    """La foto scelta, ritagliata quadrata al centro e ridotta (96 px): resta solo sul PC."""
    with Image.open(src_path) as im:
        try:
            from PIL import ImageOps
            im = ImageOps.exif_transpose(im)
        except Exception:
            pass
        im = im.convert("RGB")
        s = min(im.size)
        l, t = (im.size[0] - s) // 2, (im.size[1] - s) // 2
        im = im.crop((l, t, l + s, t + s)).resize((AV_PX, AV_PX), Image.LANCZOS)
        im.save(dest_path, "PNG")
    return dest_path


def load_avatar(path):
    try:
        if path and os.path.isfile(path):
            with Image.open(path) as im:
                return im.convert("RGB")
    except Exception:
        pass
    return None


# ---------------------------------------------------------------- telecamera e scena
class Cam:
    """Telecamera: x, y, z nel mondo (y verso l'alto), yaw = giro a destra in gradi (0 = verso il fondo
    del corridoio), pitch = alzo (positivo = guardo in su)."""

    def __init__(self, x=0.0, y=0.0, z=0.0, yaw=0.0, pitch=0.0):
        self.x, self.y, self.z, self.yaw, self.pitch = x, y, z, yaw, pitch

    def basis(self):
        t, p = math.radians(self.yaw), math.radians(self.pitch)
        st, ct, sp, cp = math.sin(t), math.cos(t), math.sin(p), math.cos(p)
        return (st * cp, sp, -ct * cp), (ct, 0.0, st), (-st * sp, cp, ct * sp)


def orbit(target, yaw, el, dist):
    """Telecamera che gira intorno a un punto: el = quanto guardo dall'alto (gradi)."""
    c = Cam(yaw=yaw, pitch=-el)
    f, _, _ = c.basis()
    c.x, c.y, c.z = target[0] - f[0] * dist, target[1] - f[1] * dist, target[2] - f[2] * dist
    return c


class Scene:
    """Elenco di cose da disegnare. Ogni voce: (strato, centro, parti). Le voci si disegnano per strato e, a parita',
    dalla piu' lontana alla piu' vicina (algoritmo del pittore); le parti di una voce nell'ordine dato."""

    def __init__(self, bg=(27, 31, 38), fog=None):
        self.bg, self.fog, self.items = bg, fog, []

    def poly(self, pts, fill, outline=None, layer=0, key=None, normal=None, nofog=False):
        n = float(len(pts))
        c = (sum(p[0] for p in pts) / n, sum(p[1] for p in pts) / n, sum(p[2] for p in pts) / n)
        self.items.append((layer, c, [("p", pts, fill, outline, key, normal, nofog)]))

    def group(self, layer, center, parts):
        self.items.append((layer, center, parts))

    def tex(self, p0, a, b, mips, key=None, layer=1, edge=C_EDGE):
        c = (p0[0] + (a[0] + b[0]) / 2.0, p0[1] + (a[1] + b[1]) / 2.0, p0[2] + (a[2] + b[2]) / 2.0)
        self.items.append((layer, c, [("t", p0, a, b, mips, key, edge)]))

    def sprite(self, anchor, w, h, img, key=None, layer=5):
        c = (anchor[0], anchor[1] + h / 2.0, anchor[2])        # si ordina dal centro, non dai piedi
        self.items.append((layer, c, [("s", anchor, w, h, img, key)]))


def _inv3(m):
    a, b, c = m[0]
    d, e, f = m[1]
    g, h, i = m[2]
    det = a * (e * i - f * h) - b * (d * i - f * g) + c * (d * h - e * g)
    if abs(det) < 1e-18:
        return None
    k = 1.0 / det
    return [[(e * i - f * h) * k, (c * h - b * i) * k, (b * f - c * e) * k],
            [(f * g - d * i) * k, (a * i - c * g) * k, (c * d - a * f) * k],
            [(d * h - e * g) * k, (b * g - a * h) * k, (a * e - b * d) * k]]


_PERSP = getattr(Image, "Transform", Image).PERSPECTIVE
_BILIN = getattr(Image, "Resampling", Image).BILINEAR
_LANCZ = getattr(Image, "Resampling", Image).LANCZOS


def point_in_poly(x, y, pts):
    inside = False
    n = len(pts)
    j = n - 1
    for i in range(n):
        xi, yi = pts[i]
        xj, yj = pts[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / ((yj - yi) or 1e-12) + xi:
            inside = not inside
        j = i
    return inside


def pick(hits, x, y):
    """La chiave dell'oggetto piu' in primo piano sotto il punto (x, y), o None."""
    for kind, geom, key in reversed(hits):
        if kind == "r":
            if geom[0] <= x <= geom[2] and geom[1] <= y <= geom[3]:
                return key
        elif point_in_poly(x, y, geom):
            return key
    return None


def render(scene, cam, W, H, ss=1, fov=FOV_WALK, hot=None):
    """Disegna la scena. Restituisce (immagine RGB W x H, hits) dove hits e' l'elenco degli oggetti cliccabili
    in ordine di disegno (l'ultimo e' in primo piano). ss = supersampling (2 = bordi lisci, piu' lento)."""
    Wd, Hd = int(W * ss), int(H * ss)
    hw, hh = Wd / 2.0, Hd / 2.0
    F = hw / math.tan(math.radians(fov) / 2.0)
    f, r, u = cam.basis()
    cx0, cy0, cz0 = cam.x, cam.y, cam.z
    img = Image.new("RGB", (Wd, Hd), scene.bg)
    draw = ImageDraw.Draw(img)
    hits = []
    fog = scene.fog
    lim = 90000.0

    def tocam(p):
        vx, vy, vz = p[0] - cx0, p[1] - cy0, p[2] - cz0
        return (vx * r[0] + vy * r[1] + vz * r[2], vx * u[0] + vy * u[1] + vz * u[2],
                vx * f[0] + vy * f[1] + vz * f[2])

    def clip_near(pc):
        out = []
        n = len(pc)
        for i in range(n):
            a, b = pc[i], pc[(i + 1) % n]
            ina, inb = a[2] >= NEAR, b[2] >= NEAR
            if ina:
                out.append(a)
            if ina != inb:
                t = (NEAR - a[2]) / (b[2] - a[2])
                out.append((a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, NEAR))
        return out

    def proj(pc):
        return [(clamp(hw + F * x / z, -lim, lim), clamp(hh - F * y / z, -lim, lim)) for x, y, z in pc]

    def fogk(dist):
        if not fog:
            return 0.0
        return clamp((dist - fog["start"]) / float(fog["end"] - fog["start"]), 0.0, 1.0) * fog["k"]

    order = sorted(range(len(scene.items)), key=lambda i: (
        scene.items[i][0], -((scene.items[i][1][0] - cx0) ** 2 + (scene.items[i][1][1] - cy0) ** 2
                             + (scene.items[i][1][2] - cz0) ** 2)))
    for idx in order:
        _layer, c, parts = scene.items[idx]
        dist = math.sqrt((c[0] - cx0) ** 2 + (c[1] - cy0) ** 2 + (c[2] - cz0) ** 2)
        fk = fogk(dist)
        for part in parts:
            kind = part[0]
            if kind == "p":
                _, pts, fill, outline, key, normal, nofog = part
                if normal is not None:
                    p0 = pts[0]
                    if normal[0] * (cx0 - p0[0]) + normal[1] * (cy0 - p0[1]) + normal[2] * (cz0 - p0[2]) <= 0:
                        continue
                pc = clip_near([tocam(p) for p in pts])
                if len(pc) < 3:
                    continue
                scr = proj(pc)
                xs, ys = [p[0] for p in scr], [p[1] for p in scr]
                if max(xs) < 0 or min(xs) > Wd or max(ys) < 0 or min(ys) > Hd:
                    continue
                col = fill
                if hot is not None and key == hot:
                    col = _mix(col, (255, 255, 255), 0.38)
                if fk and not nofog:
                    col = _mix(col, fog["color"], fk)
                    if outline:
                        outline = _mix(outline, fog["color"], fk)
                draw.polygon(scr, fill=col, outline=outline)
                if key is not None:
                    hits.append(("p", [(x / ss, y / ss) for x, y in scr], key))
            elif kind == "t":
                _, p0, a, b, mips, key, edge = part
                p1 = (p0[0] + a[0], p0[1] + a[1], p0[2] + a[2])
                p2 = (p1[0] + b[0], p1[1] + b[1], p1[2] + b[2])
                p3 = (p0[0] + b[0], p0[1] + b[1], p0[2] + b[2])
                pcs = [tocam(p) for p in (p0, p1, p2, p3)]
                pc = clip_near(pcs)
                if len(pc) < 3:
                    continue
                scr = proj(pc)
                xs, ys = [p[0] for p in scr], [p[1] for p in scr]
                bx0, by0 = max(0, int(math.floor(min(xs)))), max(0, int(math.floor(min(ys))))
                bx1, by1 = min(Wd, int(math.ceil(max(xs)))), min(Hd, int(math.ceil(max(ys))))
                if bx1 - bx0 < 2 or by1 - by0 < 2:
                    continue
                # quale versione ridotta della targhetta: la piu' piccola che basta per la larghezza sullo schermo
                alen = math.sqrt(a[0] ** 2 + a[1] ** 2 + a[2] ** 2)
                wpx = F * alen / max(NEAR, (pcs[0][2] + pcs[1][2] + pcs[2][2] + pcs[3][2]) / 4.0)
                tex = mips[0]
                for m in mips:
                    if m.size[0] >= wpx:
                        tex = m
                tw, th = tex.size
                P = pcs[0]
                A = (pcs[1][0] - P[0], pcs[1][1] - P[1], pcs[1][2] - P[2])
                B = (pcs[3][0] - P[0], pcs[3][1] - P[1], pcs[3][2] - P[2])
                m = [[(hw * A[2] + F * A[0]) / tw, (hw * B[2] + F * B[0]) / th, hw * P[2] + F * P[0]],
                     [(hh * A[2] - F * A[1]) / tw, (hh * B[2] - F * B[1]) / th, hh * P[2] - F * P[1]],
                     [A[2] / tw, B[2] / th, P[2]]]
                for j in range(3):
                    m[0][j] -= bx0 * m[2][j]
                    m[1][j] -= by0 * m[2][j]
                inv = _inv3(m)
                if inv is None or abs(inv[2][2]) < 1e-14:
                    continue
                s = inv[2][2]
                co = (inv[0][0] / s, inv[0][1] / s, inv[0][2] / s, inv[1][0] / s, inv[1][1] / s,
                      inv[1][2] / s, inv[2][0] / s, inv[2][1] / s)
                bw, bh = bx1 - bx0, by1 - by0
                try:
                    warped = tex.transform((bw, bh), _PERSP, co, _BILIN, fillcolor=edge)
                except Exception:
                    continue
                if fk > 0.02:
                    warped = Image.blend(warped, Image.new("RGB", warped.size, fog["color"]), fk)
                if hot is not None and key == hot:
                    warped = Image.blend(warped, Image.new("RGB", warped.size, (255, 255, 255)), 0.25)
                mask = Image.new("L", (bw, bh), 0)
                ImageDraw.Draw(mask).polygon([(x - bx0, y - by0) for x, y in scr], fill=255)
                img.paste(warped, (bx0, by0), mask)
                if key is not None:
                    hits.append(("t", [(x / ss, y / ss) for x, y in scr], key))
            else:
                _, anchor, wu, hu, sim, key = part
                pc = tocam(anchor)
                if pc[2] < NEAR:
                    continue
                sc = F / pc[2]
                sw, sh = int(wu * sc), int(hu * sc)
                if sw < 2 or sh < 2 or sw > 5000 or sh > 5000:
                    continue
                rs = sim.resize((sw, sh), _LANCZ if sw < sim.size[0] else _BILIN)
                x0 = int(hw + F * pc[0] / pc[2] - sw / 2.0)
                y0 = int(hh - F * pc[1] / pc[2] - sh)
                if fk > 0.02:
                    rs = _fog_rgba(rs, fog["color"], fk)
                img.paste(rs, (x0, y0), rs)
                if key is not None:
                    hits.append(("r", (x0 / ss, y0 / ss, (x0 + sw) / ss, (y0 + sh) / ss), key))
    if ss != 1:
        img = img.resize((int(W), int(H)), _LANCZ)
    return img, hits


def _fog_rgba(im, color, k):
    a = im.split()[3]
    rgb = Image.blend(im.convert("RGB"), Image.new("RGB", im.size, color), k)
    rgb.putalpha(a)
    return rgb


# ---------------------------------------------------------------- il piano: dati, scene, mini-mappa
class FloorModel:
    """Un piano pronto da visitare. data = {"edificio", "piano", "label", "rooms": [{"uid","id","codice","scrivanie",
    "persone":[{"nome","ufficio"}], "miss"}], "printers": [{"uid","nome","modello","stanza","red","low","fuori"}]}.
    printer_img = funzione modello -> immagine Pillow (o None)."""

    def __init__(self, data, printer_img=None):
        self.data = data
        self.rooms = list(data.get("rooms") or [])
        self.printers = list(data.get("printers") or [])
        self.L = layout(self.rooms)
        self.label = data.get("label") or "Piano %s" % data.get("piano")
        self.edificio = data.get("edificio") or ""
        self.plates = {r["uid"]: plate_data(r) for r in self.rooms}
        self.room = {r["uid"]: r for r in self.rooms}
        self.item = {it["r"]["uid"]: it for it in self.L["items"]}
        self.printer_img = printer_img
        self._c = {}

    # --- posizioni
    @property
    def LC(self):
        return self.L["LC"]

    def clamp_pose(self, x, z):
        return clamp(x, -CW / 2.0 + 16, CW / 2.0 - 16), clamp(z, -self.LC + 18, -6)

    def go_pose(self, uid):
        """Mi metto nel corridoio, dal lato opposto alla stanza e un po' prima, girato verso il punto in mezzo
        tra la porta e la targhetta: tutte e due entrano bene nella vista. Restituisce (x, z, yaw)."""
        it = self.item.get(uid)
        if not it:
            return None
        side = it["side"]
        zmid = -(it["col"] + (FD + FPL) / 2.0) * RW
        x, z = self.clamp_pose(-side * 40.0, zmid + 140)
        yaw = math.degrees(math.atan2(side * (CW / 2.0) - x, -(zmid - z)))
        return x, z, yaw

    def nearest(self, x, z, yaw):
        """La targhetta che sto guardando: la piu' vicina al centro della vista, entro 38 gradi e 230 unita'."""
        r = math.radians(yaw)
        fx, fz = math.sin(r), -math.cos(r)
        best = None
        for it in self.L["items"]:
            px, pz = it["side"] * CW / 2.0, -(it["col"] + FPL) * RW
            vx, vz = px - x, pz - z
            dist = math.hypot(vx, vz)
            if dist > PLATE_DIST or dist < 1:
                continue
            ang = math.degrees(math.acos(clamp((vx * fx + vz * fz) / dist, -1.0, 1.0)))
            if ang < PLATE_ANGLE and (best is None or ang < best[1]):
                best = (it["r"]["uid"], ang, dist)
        return best

    # --- immagini
    def plate_pack(self, uid):
        k = ("pl", uid)
        if k not in self._c:
            pd = self.plates[uid]
            rgba = plate_image(pd)
            self._c[k] = {"rgba": rgba, "mips": make_mips(flat(rgba))}
        return self._c[k]

    def printer_sprite(self, p):
        k = ("pr", p.get("uid"), p.get("red"), bool(p.get("low")), bool(p.get("fuori")))
        if k in self._c:
            return self._c[k]
        base = None
        try:
            base = self.printer_img(p.get("modello")) if self.printer_img else None
        except Exception:
            base = None
        if base is None:
            base = Image.new("RGBA", (120, 130), (0, 0, 0, 0))
            bd = ImageDraw.Draw(base)
            bd.rounded_rectangle((8, 30, 112, 124), radius=8, fill=(205, 210, 218, 255), outline=(90, 100, 115, 255), width=3)
            bd.rectangle((24, 14, 96, 36), fill=(235, 238, 242, 255), outline=(90, 100, 115, 255), width=2)
            bd.rectangle((28, 86, 92, 100), fill=(120, 130, 145, 255))
        base = base.convert("RGBA")
        sc = 130.0 / base.size[1]
        base = base.resize((max(2, int(base.size[0] * sc)), 130), _LANCZ)
        im = Image.new("RGBA", (base.size[0] + 24, base.size[1] + 24), (0, 0, 0, 0))
        im.paste(base, (6, 24), base)
        d = ImageDraw.Draw(im)
        x = im.size[0] - 4
        if p.get("red"):
            d.ellipse((x - 30, 0, x, 30), fill=(211, 32, 47, 255), outline=(255, 255, 255, 255), width=3)
            _text_center(d, x - 15, 15, str(int(p["red"])), get_font(18, True), (255, 255, 255, 255))
            x -= 34
        if p.get("low"):
            d.ellipse((x - 30, 0, x, 30), fill=(242, 197, 0, 255), outline=(255, 255, 255, 255), width=3)
            _text_center(d, x - 15, 15, "!", get_font(20, True), (90, 58, 0, 255))
        if p.get("fuori"):
            d.line((8, 40, im.size[0] - 18, im.size[1] - 8), fill=(211, 32, 47, 255), width=6)
            d.line((im.size[0] - 18, 40, 8, im.size[1] - 8), fill=(211, 32, 47, 255), width=6)
        self._c[k] = im
        return im

    def corridor_printers(self):
        return [p for p in self.printers if not p.get("stanza")]

    # --- la scena in prima persona
    def walk_scene(self):
        if "walk" in self._c:
            return self._c["walk"]
        sc = Scene(bg=(30, 34, 42), fog={"start": 120.0, "end": 2000.0, "k": 0.62, "color": (36, 41, 52)})
        L = self.L
        cols, LC = L["cols"], L["LC"]
        fy, cy = FLOOR_Y, CEIL_Y
        seg_n = 4
        seg = RW / float(seg_n)
        fa, fb = (136, 150, 169), (123, 137, 157)
        for c in range(cols):
            for j in range(seg_n):
                z0, z1 = -(c * RW + j * seg), -(c * RW + (j + 1) * seg)
                col = fa if (c + j) % 2 == 0 else fb
                sc.poly([(-CW / 2.0, fy, z0), (CW / 2.0, fy, z0), (CW / 2.0, fy, z1), (-CW / 2.0, fy, z1)], col, layer=0)
                sc.poly([(-CW / 2.0, cy, z0), (CW / 2.0, cy, z0), (CW / 2.0, cy, z1), (-CW / 2.0, cy, z1)], (238, 242, 246), layer=0)
            zc = -(c + 0.5) * RW
            # plafoniera
            sc.poly([(-34, cy - 0.3, zc + 22), (34, cy - 0.3, zc + 22), (34, cy - 0.3, zc - 22), (-34, cy - 0.3, zc - 22)], (250, 249, 236), layer=1)
            sc.poly([(-26, cy - 0.5, zc + 15), (26, cy - 0.5, zc + 15), (26, cy - 0.5, zc - 15), (-26, cy - 0.5, zc - 15)], (255, 255, 255), layer=1)
        # pareti: tre strisce per ogni riquadro (cimasa chiara, filetto, zoccolo scuro)
        band = [(fy, fy + 0.13 * H, (93, 107, 124)), (fy + 0.13 * H, fy + 0.16 * H, (201, 194, 180)), (fy + 0.16 * H, cy, (244, 241, 234))]
        n3 = 3
        for side in (-1, 1):
            x = side * CW / 2.0
            for c in range(cols):
                for j in range(n3):
                    z0, z1 = -(c + j / float(n3)) * RW, -(c + (j + 1) / float(n3)) * RW
                    for y0, y1, col in band:
                        sc.poly([(x, y0, z0), (x, y0, z1), (x, y1, z1), (x, y1, z0)], col, layer=0)
        # porte e targhette
        for it in L["items"]:
            uid, side, col = it["r"]["uid"], it["side"], it["col"]
            x = side * CW / 2.0
            zd = -(col + FD) * RW
            zp = -(col + FPL) * RW
            key = ("room", uid)
            # porta: cornice, anta, specchiatura, maniglia
            d0, d1 = zd + DOORW / 2.0, zd - DOORW / 2.0
            db = fy
            dt = fy + DOORH

            def wq(za, zb, y0, y1, xx=x):
                return [(xx, y0, za), (xx, y0, zb), (xx, y1, zb), (xx, y1, za)]
            parts = [("p", wq(d0 + 3, d1 - 3, db, dt + 3), (93, 70, 48), None, key, None, False),
                     ("p", wq(d0, d1, db, dt), (164, 122, 82), (122, 90, 60), key, None, False),
                     ("p", wq(d0 - 9, d1 + 9, dt - 12, dt - 44), (178, 136, 94), (135, 98, 62), key, None, False),
                     ("p", wq(d0 - 9, d1 + 9, dt - 56, dt - 106), (178, 136, 94), (135, 98, 62), key, None, False),
                     ("p", wq(d1 + 12, d1 + 7, fy + 0.44 * DOORH + 7, fy + 0.44 * DOORH - 7), (217, 221, 226), (119, 119, 119), key, None, False)]
            sc.group(1, (x, (db + dt) / 2.0, zd), parts)
            # targhetta
            pk = self.plate_pack(uid)
            ytop = fy + 18 + PLH
            if side < 0:
                p0, a = (x, ytop, zp + PLW / 2.0), (0.0, 0.0, -PLW)
            else:
                p0, a = (x, ytop, zp - PLW / 2.0), (0.0, 0.0, PLW)
            b = (0.0, -PLH, 0.0)
            sh = [(x, fy + 16, zp + PLW / 2.0 + 2), (x, fy + 16, zp - PLW / 2.0 - 2), (x, ytop + 2, zp - PLW / 2.0 - 2), (x, ytop + 2, zp + PLW / 2.0 + 2)]
            sc.group(1, (x, ytop - PLH / 2.0, zp), [("p", sh, (190, 183, 168), None, None, None, False),
                                                     ("t", p0, a, b, pk["mips"], key, C_EDGE)])
        # pareti in fondo e all'entrata
        sc.tex((-CW / 2.0, cy, -LC), (CW, 0.0, 0.0), (0.0, -H, 0.0), make_mips(end_image(self.label, self.edificio, "exit")), None, 0, (190, 205, 225))
        sc.tex((CW / 2.0, cy, 0.0), (-CW, 0.0, 0.0), (0.0, -H, 0.0), make_mips(end_image(self.label, "Entrata del piano", "in")), None, 0, (190, 205, 225))
        # stampanti nel corridoio: in fondo, davanti alla parete
        free = self.corridor_printers()
        for i, p in enumerate(free):
            im = self.printer_sprite(p)
            hu = 66.0
            wu = hu * im.size[0] / float(im.size[1])
            x = (i - (len(free) - 1) / 2.0) * 58
            sc.sprite((x, fy, -LC + 10), wu, hu, im, ("prn", p.get("uid")), layer=2)
        self._c["walk"] = sc
        return sc

    # --- la scena dall'alto (casa delle bambole)
    def dh_scene(self):
        if "dh" in self._c:
            return self._c["dh"]
        sc = Scene(bg=(11, 42, 77))
        L = self.L
        LC, cols = L["LC"], L["cols"]
        fy = FLOOR_Y
        Wd = CW + 2 * RD
        sc.poly([(-Wd / 2 - 24, fy, 24), (Wd / 2 + 24, fy, 24), (Wd / 2 + 24, fy, -LC - 24), (-Wd / 2 - 24, fy, -LC - 24)], (86, 99, 119), layer=0)
        sc.poly([(-CW / 2.0, fy, 0), (CW / 2.0, fy, 0), (CW / 2.0, fy, -LC), (-CW / 2.0, fy, -LC)], (158, 170, 187), layer=1)
        k = 0
        z = -10.0
        while z > -LC + 12:
            sc.poly([(-3, fy, z), (3, fy, z), (3, fy, z - 12), (-3, fy, z - 12)], (232, 238, 246), layer=2)
            z -= 24.0
            k += 1
        wallc, wallo = (228, 236, 246), (157, 184, 230)
        for it in L["items"]:
            uid, side, col = it["r"]["uid"], it["side"], it["col"]
            pd = self.plates[uid]
            cx, cz = it["x"], it["z"]
            key = ("room", uid)
            sc.poly([(cx - (RD - 10) / 2, fy, cz + (RW - 10) / 2), (cx + (RD - 10) / 2, fy, cz + (RW - 10) / 2),
                     (cx + (RD - 10) / 2, fy, cz - (RW - 10) / 2), (cx - (RD - 10) / 2, fy, cz - (RW - 10) / 2)],
                    (231, 223, 202) if pd["free"] else (217, 201, 162), (185, 174, 140), layer=2, key=key)
            tot = max(pd["n"], pd["tot"])
            for i in range(min(tot, 8)):
                gx, gz = (i % 2 - 0.5) * 52, (i // 2 - 1) * 34
                dx, dz = cx + gx + side * 6, cz + gz + 18
                pts = [(dx - 19, fy, dz + 11), (dx + 19, fy, dz + 11), (dx + 19, fy, dz - 11), (dx - 19, fy, dz - 11)]
                if i < pd["n"]:
                    sc.poly(pts, (45, 79, 134), None, layer=3, key=key)
                else:
                    sc.poly(pts, (231, 223, 202), (139, 130, 102), layer=3, key=key)
            y1 = fy + WH
            # muro esterno, muro di fondo, muro verso il corridoio con il vano della porta
            xo = side * (CW / 2.0 + RD)
            sc.poly([(xo, fy, cz + (RW - 10) / 2), (xo, fy, cz - (RW - 10) / 2), (xo, y1, cz - (RW - 10) / 2), (xo, y1, cz + (RW - 10) / 2)], wallc, wallo, layer=4)
            zb = -(col + 1) * RW
            sc.poly([(cx - RD / 2, fy, zb), (cx + RD / 2, fy, zb), (cx + RD / 2, y1, zb), (cx - RD / 2, y1, zb)], wallc, wallo, layer=4)
            if col == 0:
                sc.poly([(cx - RD / 2, fy, 0), (cx + RD / 2, fy, 0), (cx + RD / 2, y1, 0), (cx - RD / 2, y1, 0)], wallc, wallo, layer=4)
            xc = side * CW / 2.0
            a = FD * RW
            sega, segb = a - DOORW / 2.0, RW - a - DOORW / 2.0
            if sega > 1:
                za, zb2 = -col * RW, -col * RW - sega
                sc.poly([(xc, fy, za), (xc, fy, zb2), (xc, y1, zb2), (xc, y1, za)], wallc, wallo, layer=4)
            if segb > 1:
                za, zb2 = -(col + 1) * RW + segb, -(col + 1) * RW
                sc.poly([(xc, fy, za), (xc, fy, zb2), (xc, y1, zb2), (xc, y1, za)], wallc, wallo, layer=4)
            # targhetta in piedi accanto alla porta
            pk = self.plate_pack(uid)
            sc.sprite((side * (CW / 2.0 - 4), fy, -(col + FPL) * RW), PLW * 0.7, PLH * 0.7, pk["rgba"], key, layer=4)
        free = self.corridor_printers()
        for i, p in enumerate(free):
            im = self.printer_sprite(p)
            hu = 70.0
            wu = hu * im.size[0] / float(im.size[1])
            sc.sprite((0, fy, -LC + 30 + i * 60), wu, hu, im, ("prn", p.get("uid")), layer=4)
        self._c["dh"] = sc
        return sc

    # --- mini-mappa
    def minimap(self, x, z, yaw, avatar=None, hot=None, maxw=420):
        """Il piano visto dall'alto, lungo in orizzontale: rettangoli per le stanze, cono di vista e il segnaposto
        (la faccia, se c'e'). Restituisce (immagine RGBA, [(x0,y0,x1,y1,uid)])."""
        LC = self.LC
        pad = 8
        ms = min(0.30, (maxw - 2 * pad) / float(LC))
        mw, mh = int(LC * ms) + 2 * pad, int((CW + 2 * RD) * ms) + 2 * pad
        im = Image.new("RGBA", (mw, mh), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        d.rounded_rectangle((0, 0, mw - 1, mh - 1), radius=10, fill=(255, 255, 255, 225), outline=(150, 165, 185, 255))

        def mx(zz):
            return pad - zz * ms

        def my(xx):
            return pad + (xx + CW / 2.0 + RD) * ms
        d.rectangle((pad, my(-CW / 2.0), pad + LC * ms, my(CW / 2.0)), fill=(201, 211, 223, 255))
        rects = []
        f = get_font(11, True)
        for it in self.L["items"]:
            uid = it["r"]["uid"]
            x0, x1 = mx(it["z"] + RW / 2.0), mx(it["z"] - RW / 2.0)
            y0, y1 = my(it["x"] - RD / 2.0), my(it["x"] + RD / 2.0)
            fillc = (238, 233, 219, 255) if self.plates[uid]["free"] else (217, 201, 162, 255)
            if hot == uid:
                fillc = (190, 222, 255, 255)
            d.rectangle((x0 + 1, y0 + 1, x1 - 1, y1 - 1), fill=fillc, outline=(139, 130, 102, 255))
            _text_center(d, (x0 + x1) / 2.0, (y0 + y1) / 2.0, self.plates[uid]["num"], f, (29, 43, 74, 255))
            rects.append((x0, y0, x1, y1, uid))
        ax, ay = mx(z), my(x)
        r = math.radians(yaw)
        dx, dy = math.cos(r), math.sin(r)          # dove guardo, sulla mappa (verso destra = in fondo al corridoio)
        ln = 34.0
        ov = Image.new("RGBA", im.size, (0, 0, 0, 0))
        od = ImageDraw.Draw(ov)
        a1, a2 = r - 0.55, r + 0.55
        od.polygon([(ax, ay), (ax + math.cos(a1) * ln, ay + math.sin(a1) * ln), (ax + math.cos(a2) * ln, ay + math.sin(a2) * ln)], fill=(26, 154, 74, 95))
        im = Image.alpha_composite(im, ov)
        d = ImageDraw.Draw(im)
        if avatar is not None:
            av = circle_avatar(avatar, 26)
            d.ellipse((ax - 14, ay - 14, ax + 14, ay + 14), fill=(26, 154, 74, 255))
            im.paste(av, (int(ax - 13), int(ay - 13)), av)
        else:
            d.ellipse((ax - 7, ay - 7, ax + 7, ay + 7), fill=(26, 154, 74, 255), outline=(255, 255, 255, 255), width=3)
            d.line((ax, ay, ax + dx * 12, ay + dy * 12), fill=(255, 255, 255, 255), width=3)
        return im, rects

    # --- l'edificio: scheda del piano (per l'elenco dei piani)
    def summary(self):
        ppl = sum(p["n"] for p in self.plates.values())
        return {"label": self.label, "rooms": len(self.rooms), "people": ppl,
                "printers": len(self.printers), "red": sum(int(p.get("red") or 0) for p in self.printers),
                "low": sum(1 for p in self.printers if p.get("low")), "off": sum(1 for p in self.printers if p.get("fuori"))}


# ---------------------------------------------------------------- telecamere pronte
def dh_camera(model, yaw, el, zoom, pz, W, H, fov=FOV_DH):
    """Telecamera per la vista dall'alto: gira intorno al punto del corridoio a distanza pz dall'entrata.
    All'inizio si vede tutta la larghezza e circa quattro riquadri di lunghezza (o tutto il piano se e' corto)."""
    half_h = math.tan(math.radians(fov) / 2.0)
    half_v = half_h * float(H) / max(1.0, float(W))
    d1 = ((CW + 2 * RD + 70) / 2.0) / half_h
    d2 = min(model.LC, 4 * RW) / 2.0 * math.sin(math.radians(max(10.0, el))) / half_v * 1.1
    return orbit((0.0, FLOOR_Y, -pz), yaw, el, max(d1, d2) / max(0.2, zoom))


class TowerModel:
    """L'edificio intero: i piani impilati (dal piu' basso). floors = [FloorModel, ...] in ordine di piano."""
    S, GAP, T = 0.42, 70.0, 10.0

    def __init__(self, nome, floors):
        self.nome = nome
        self.floors = list(floors)
        cols = max([1] + [f.L["cols"] for f in self.floors])
        self.cols = cols
        self.Wd = (CW + 2 * RD) * self.S
        self.Dd = cols * RW * self.S
        self.n = len(self.floors)
        self.top = (self.n - 1) * self.GAP
        self._scene = None

    def floor_at(self, key):
        """Il piano (FloorModel) a cui appartiene la chiave di un oggetto cliccato."""
        if isinstance(key, tuple) and key and key[0] == "floor":
            return next((f for f in self.floors if f.data.get("piano") == key[1]), None)
        return None

    def scene(self):
        if self._scene:
            return self._scene
        S, GAP, T, Wd, Dd = self.S, self.GAP, self.T, self.Wd, self.Dd
        sc = Scene(bg=(28, 63, 102))
        hw, hd = Wd / 2.0, Dd / 2.0
        for i, fl in enumerate(self.floors):
            y = i * GAP
            key = ("floor", fl.data.get("piano"))
            lay = 2 * i
            # lati (solo quelli che guardano la telecamera) e piano di sopra
            sc.poly([(-hw, y - T, hd), (hw, y - T, hd), (hw, y, hd), (-hw, y, hd)], (70, 86, 107), (52, 66, 85), lay, key, (0, 0, 1))
            sc.poly([(-hw, y - T, -hd), (hw, y - T, -hd), (hw, y, -hd), (-hw, y, -hd)], (62, 77, 97), (52, 66, 85), lay, key, (0, 0, -1))
            sc.poly([(hw, y - T, hd), (hw, y - T, -hd), (hw, y, -hd), (hw, y, hd)], (58, 72, 92), (52, 66, 85), lay, key, (1, 0, 0))
            sc.poly([(-hw, y - T, hd), (-hw, y - T, -hd), (-hw, y, -hd), (-hw, y, hd)], (58, 72, 92), (52, 66, 85), lay, key, (-1, 0, 0))
            sc.poly([(-hw, y, hd), (hw, y, hd), (hw, y, -hd), (-hw, y, -hd)], (110, 127, 149), (82, 98, 120), lay, key)
            cw2 = CW * S / 2.0
            sc.poly([(-cw2, y, hd), (cw2, y, hd), (cw2, y, -hd), (-cw2, y, -hd)], (169, 181, 198), None, lay + 1, key)
            for it in fl.L["items"]:
                pd = fl.plates[it["r"]["uid"]]
                col = (238, 233, 219) if pd["free"] else (243, 199, 196) if pd["miss"] else (191, 227, 201)
                cx, cz = it["x"] * S, (it["z"] + self.cols * RW / 2.0) * S
                a, b = (RD - 12) * S / 2.0, (RW - 12) * S / 2.0
                sc.poly([(cx - a, y, cz + b), (cx + a, y, cz + b), (cx + a, y, cz - b), (cx - a, y, cz - b)], col, (185, 174, 140), lay + 1, key)
        roof = label_image(self.nome, size=30)
        wu = min(Wd * 1.6, roof.size[0] * 0.5)
        sc.sprite((0, self.top + 46, 0), wu, wu * roof.size[1] / float(roof.size[0]), roof, None, layer=2 * self.n + 3)
        self._scene = sc
        return sc

    def camera(self, yaw, el, zoom, W, H, fov=FOV_TW):
        rs = math.sqrt((self.Wd / 2.0) ** 2 + (self.Dd / 2.0) ** 2 + ((self.top + 80) / 2.0) ** 2)
        half = math.tan(math.radians(fov) / 2.0)
        asp = float(H) / max(1.0, float(W))
        dist = rs / min(half, half * asp) * 1.0 / max(0.2, zoom)
        return orbit((0.0, self.top / 2.0 - 8, 0.0), yaw, el, dist)


# ---------------------------------------------------------------- movimenti (senza Tk: si provano da soli)
class WalkCtl:
    """Stato e movimenti della passeggiata: in prima persona (x, z, yaw, pitch) e dall'alto (giro, inclinazione,
    avvicinamento, punto del corridoio)."""

    def __init__(self, model, mode="fp"):
        self.m = model
        self.mode = mode
        self.x, self.z, self.yaw, self.pitch = 0.0, -40.0, 0.0, 0.0
        self.d_yaw, self.d_el, self.d_zoom, self.d_pz = 0.0, 58.0, 1.0, min(model.LC, 4 * RW) / 2.0

    # --- in prima persona
    def step(self, fwd, turn, dt):
        """fwd: -1 indietro .. +1 avanti; turn: -1 sinistra .. +1 destra; dt in secondi."""
        self.yaw = (self.yaw + TURN * dt * turn + 180.0) % 360.0 - 180.0
        r = math.radians(self.yaw)
        sp = SPEED * dt * fwd
        self.x, self.z = self.m.clamp_pose(self.x + math.sin(r) * sp, self.z - math.cos(r) * sp)

    def look(self, dx, dy):
        self.yaw = (self.yaw + dx * 0.25 + 180.0) % 360.0 - 180.0
        self.pitch = clamp(self.pitch + dy * 0.15, -25.0, 25.0)

    def go_room(self, uid):
        p = self.m.go_pose(uid)
        if not p:
            return False
        self.x, self.z, self.yaw = p
        self.pitch = 0.0
        return True

    def nearest(self):
        return self.m.nearest(self.x, self.z, self.yaw)

    def camera(self):
        return Cam(self.x, 0.0, self.z, self.yaw, self.pitch)

    # --- dall'alto
    def dh_drag(self, dx, dy):
        self.d_yaw = (self.d_yaw - dx * 0.4 + 180.0) % 360.0 - 180.0
        self.d_el = clamp(self.d_el - dy * 0.3, 12.0, 85.0)

    def dh_zoom(self, f):
        self.d_zoom = clamp(self.d_zoom * f, 0.35, 6.0)

    def dh_step(self, fwd, turn, dt):
        self.d_pz = clamp(self.d_pz + 260.0 * dt * fwd, 0.0, float(self.m.LC))
        self.d_yaw = (self.d_yaw + 60.0 * dt * turn + 180.0) % 360.0 - 180.0


class TowerCtl:
    def __init__(self, tower):
        self.t = tower
        self.yaw, self.el, self.zoom = -35.0, 28.0, 1.0

    def drag(self, dx, dy):
        self.yaw = (self.yaw - dx * 0.4 + 180.0) % 360.0 - 180.0
        self.el = clamp(self.el - dy * 0.3, 8.0, 80.0)

    def zoom_by(self, f):
        self.zoom = clamp(self.zoom * f, 0.4, 5.0)


# ---------------------------------------------------------------- la tela di Tk
def _png_b64(img):
    buf = io.BytesIO()
    img.save(buf, "PNG", compress_level=1)
    return base64.b64encode(buf.getvalue()).decode("ascii")


class Stage:
    """Mostra nel Canvas l'immagine disegnata. Mentre ti muovi disegna in fretta (bordi a scaletta), appena ti fermi
    ridisegna con i bordi lisci. draw(rw, rh, hq) -> (immagine rw x rh, hits)."""
    MAXW, MAXH = 1100, 720

    def __init__(self, win, canvas, draw):
        self.win, self.canvas, self.draw = win, canvas, draw
        self.item = canvas.create_image(0, 0, anchor="nw")
        self.photo = None
        self.hits = []
        self.scale = 1.0
        self.job = None
        self.n = 0                      # quante volte ho disegnato (per le prove)

    def size(self):
        try:
            w, h = int(self.canvas.winfo_width()), int(self.canvas.winfo_height())
        except Exception:
            w, h = 900, 560
        w, h = max(240, w), max(160, h)
        s = min(1.0, self.MAXW / float(w), self.MAXH / float(h))
        return w, h, max(160, int(w * s)), max(120, int(h * s)), s

    def to_render(self, x, y):
        return x * self.scale, y * self.scale

    def redraw(self, hq=False, later=True):
        if self.job is not None:
            try:
                self.win.after_cancel(self.job)
            except Exception:
                pass
            self.job = None
        w, h, rw, rh, s = self.size()
        self.scale = rw / float(w)
        img, self.hits = self.draw(rw, rh, hq)
        if img.size != (w, h):
            img = img.resize((w, h), _BILIN)
        self.photo = tk.PhotoImage(data=_png_b64(img))
        self.canvas.itemconfigure(self.item, image=self.photo)
        self.n += 1
        if later and not hq:
            self.job = self.win.after(220, lambda: self.redraw(True, False))

    def request(self, hq=False):
        """Come redraw, ma se arrivano tante richieste di fila (il mouse) ridisegna una volta sola, a coda vuota."""
        if getattr(self, "_pending", False):
            return
        self._pending = True

        def go():
            self._pending = False
            self.redraw(hq)
        try:
            self.win.after_idle(go)
        except Exception:
            go()

    def cancel(self):
        if self.job is not None:
            try:
                self.win.after_cancel(self.job)
            except Exception:
                pass
            self.job = None


def _fit_geometry(win, w, h):
    """Dimensione iniziale della finestra, ridotta se lo schermo e' piccolo."""
    try:
        sw, sh = int(win.winfo_screenwidth()), int(win.winfo_screenheight())
        w, h = min(w, max(640, sw - 80)), min(h, max(480, sh - 120))
    except Exception:
        pass
    win.geometry("%dx%d" % (w, h))


def _thumb_photo(im, height):
    t = im.resize((max(2, int(im.size[0] * height / float(im.size[1]))), height), _LANCZ)
    return tk.PhotoImage(data=_png_b64(t))


def _ev_delta(e):
    """+1 / -1 per la rotellina (Windows: delta, Linux: Button-4/5)."""
    num = getattr(e, "num", None)
    if num == 4:
        return 1
    if num == 5:
        return -1
    d = getattr(e, "delta", 0) or 0
    return 1 if d > 0 else -1 if d < 0 else 0


KEYS = {"up": ("f", 1), "w": ("f", 1), "down": ("f", -1), "s": ("f", -1),
        "left": ("t", -1), "a": ("t", -1), "right": ("t", 1), "d": ("t", 1)}
NAVY_BG = "#0b2a4d"


class AvatarDialog(tk.Toplevel if tk else object):
    """«La mia faccia»: la foto resta solo su questo PC (cartella dei dati)."""

    def __init__(self, parent, host, on_change):
        super().__init__(parent)
        self.host, self.on_change = host, on_change
        self.title("La mia faccia")
        self.resizable(False, False)
        self.transient(parent)
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="La mia faccia", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(body, wraplength=380, justify="left", text=(
            "La foto compare al posto del pallino verde sulla mini-mappa quando cammini nel corridoio. "
            "Resta SOLO su questo PC, nella cartella dei dati: non va nel backup, non va nei file per il "
            "telefono e non va su GitHub.")).pack(anchor="w", pady=(6, 10))
        self.lbl = ttk.Label(body)
        self.lbl.pack(pady=(0, 8))
        self.msg = ttk.Label(body, foreground="#a91b29", wraplength=380)
        self.msg.pack(anchor="w")
        row = ttk.Frame(body)
        row.pack(fill="x", pady=(8, 0))
        self.btn_pick = ttk.Button(row, command=self.pick)
        self.btn_pick.pack(side="left")
        self.btn_del = ttk.Button(row, text="Togli la foto", command=self.remove)
        self.btn_del.pack(side="left", padx=(8, 0))
        ttk.Button(row, text="Chiudi", command=self.destroy).pack(side="right")
        self.bind("<Escape>", lambda e: self.destroy())
        self.refresh()

    def current(self):
        return load_avatar(self.host.avatar_path())

    def refresh(self):
        av = self.current()
        self.btn_pick.configure(text="Cambia la foto…" if av else "Scegli una foto…")
        if av is not None:
            self.thumb = tk.PhotoImage(data=_png_b64(circle_avatar(av, 96)))
            self.lbl.configure(image=self.thumb)
            self.btn_del.state(["!disabled"]) if hasattr(self.btn_del, "state") else None
        else:
            self.lbl.configure(image="", text="(nessuna foto)")
            self.btn_del.state(["disabled"]) if hasattr(self.btn_del, "state") else None

    def pick(self):
        path = filedialog.askopenfilename(
            parent=self, title="Scegli la tua foto",
            filetypes=[("Immagini", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"), ("Tutti i file", "*.*")])
        if not path:
            return
        try:
            save_avatar(path, self.host.avatar_path())
        except Exception:
            self.msg.configure(text="Non riesco a leggere questa foto: provane un'altra (JPG o PNG).")
            return
        self.msg.configure(text="")
        self.refresh()
        self.on_change()

    def remove(self):
        try:
            if os.path.isfile(self.host.avatar_path()):
                os.remove(self.host.avatar_path())
        except OSError:
            pass
        self.refresh()
        self.on_change()


# ---------------------------------------------------------------- la finestra della passeggiata
class WalkWindow(tk.Toplevel if tk else object):
    """Un piano da visitare: «Dall'alto» (casa delle bambole) oppure «Cammina» (prima persona nel corridoio)."""

    def __init__(self, host, edificio_id, piano, mode=None):
        super().__init__(host.parent)
        self.host, self.edificio_id, self.piano = host, edificio_id, piano
        self.title("Passeggiata nel piano")
        _fit_geometry(self, 1100, 780)
        self.minsize(720, 520)
        self.configure(bg=NAVY_BG)
        self.model = None
        self.ctl = None
        self.avatar = load_avatar(host.avatar_path())
        self._held, self._rel, self._tick_job, self._last = {}, {}, None, 0.0
        self._drag = None
        self._mini_at = (0, 0)
        self._mini_rects = []
        self._hud_uid = None
        self._hud_photo = None
        self._hot = None
        self._plate_thumbs = {}
        self.reload(first=True)
        self.ctl.mode = mode or "fp"
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.close)
        for ev, fn in (("<KeyPress>", self._key_down), ("<KeyRelease>", self._key_up)):
            self.bind(ev, fn)
        self.bind("<Escape>", lambda e: self.close())
        self.after(60, self._first_draw)

    # --- dati
    def reload(self, first=False):
        try:
            if not first and not self.winfo_exists():
                return
        except Exception:
            return
        data = self.host.floor_data(self.edificio_id, self.piano)
        self.model = FloorModel(data, self.host.printer_image)
        self._plate_thumbs = {}
        if self.ctl is None:
            self.ctl = WalkCtl(self.model)
        else:
            self.ctl.m = self.model
            self.ctl.x, self.ctl.z = self.model.clamp_pose(self.ctl.x, self.ctl.z)
            self.ctl.d_pz = clamp(self.ctl.d_pz, 0, self.model.LC)
        self._hud_uid = None
        if not first:
            self._fill_go()
            self.title("%s · %s" % (self.model.label, self.model.edificio))
            self.stage.redraw()

    def _fill_go(self):
        rooms = sorted(self.model.rooms, key=lambda r: [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", str(r["codice"]))])
        self.go_map = {str(r["codice"]): r["uid"] for r in rooms}
        self.cb_go.configure(values=list(self.go_map))

    # --- finestra
    def _build_ui(self):
        m = self.model
        self.title("%s · %s" % (m.label, m.edificio))
        bar = ttk.Frame(self, padding=(10, 8))
        bar.pack(side="top", fill="x")
        self.var_mode = tk.StringVar(value=self.ctl.mode)
        ttk.Label(bar, text="%s · %s" % (m.label, m.edificio), font=("Segoe UI", 11, "bold")).pack(side="left", padx=(0, 16))
        ttk.Radiobutton(bar, text="Dall'alto", variable=self.var_mode, value="dh", command=self.on_mode).pack(side="left")
        ttk.Radiobutton(bar, text="Cammina", variable=self.var_mode, value="fp", command=self.on_mode).pack(side="left", padx=(8, 16))
        ttk.Label(bar, text="Vai a").pack(side="left")
        self.var_go = tk.StringVar()
        self.cb_go = ttk.Combobox(bar, textvariable=self.var_go, state="readonly", width=9)
        self.cb_go.pack(side="left", padx=(6, 12))
        self.cb_go.bind("<<ComboboxSelected>>", lambda e: self.on_go())
        ttk.Button(bar, text="La mia faccia…", command=self.avatar_dialog).pack(side="left")
        ttk.Button(bar, text="Chiudi", command=self.close).pack(side="right")
        self._fill_go()
        self.hint = ttk.Label(self, style="Hint.TLabel", wraplength=1000, justify="left", padding=(10, 0, 10, 8))
        self.hint.pack(side="bottom", fill="x")
        low = ttk.Frame(self, padding=(10, 6))
        low.pack(side="bottom", fill="x")
        self.low = low
        self.canvas = tk.Canvas(self, bg=NAVY_BG, highlightthickness=0, width=900, height=560)
        self.canvas.pack(side="top", fill="both", expand=True)
        # prima persona: la targhetta che sto guardando
        self.hud = ttk.Frame(low)
        self.hud_img = ttk.Label(self.hud)
        self.hud_img.pack(side="left", padx=(0, 10))
        self.hud_txt = ttk.Label(self.hud, justify="left", wraplength=620, text="Avvicinati a una targhetta per leggerla.")
        self.hud_txt.pack(side="left", fill="x", expand=True)
        self.btn_enter = ttk.Button(self.hud, text="Entra nella stanza", style="Save.TButton", command=self.enter_hud)
        self.btn_enter.pack(side="right")
        # dall'alto: punto del corridoio e avvicinamento
        self.top = ttk.Frame(low)
        ttk.Label(self.top, text="Lungo il corridoio").pack(side="left")
        self.sl = ttk.Scale(self.top, from_=0, to=max(1, self.model.LC), orient="horizontal", command=self.on_slide)
        self.sl.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(self.top, text="−", width=3, command=lambda: self.zoom_dh(1 / 1.2)).pack(side="left")
        ttk.Button(self.top, text="+", width=3, command=lambda: self.zoom_dh(1.2)).pack(side="left", padx=(4, 0))
        self.stage = Stage(self, self.canvas, self.draw)
        for ev, fn in (("<ButtonPress-1>", self._press), ("<B1-Motion>", self._motion),
                       ("<ButtonRelease-1>", self._release), ("<Motion>", self._hover),
                       ("<MouseWheel>", self._wheel), ("<Button-4>", self._wheel), ("<Button-5>", self._wheel),
                       ("<Configure>", self._resize)):
            self.canvas.bind(ev, fn)
        self._show_mode()

    def _show_mode(self):
        fp = self.ctl.mode == "fp"
        for w in (self.hud, self.top):
            try:
                w.pack_forget()
            except Exception:
                pass
        if fp:
            self.hud.pack(side="top", fill="x")
            self.cb_go.configure(state="readonly")
            self.hint.configure(text="Frecce o W A S D per camminare e girarti (tieni premuto) · trascina il mouse per guardare "
                                     "intorno · rotellina = avanti/indietro · clicca una porta o una targhetta per entrare nella stanza.")
        else:
            self.top.pack(side="top", fill="x")
            self.sl.set(self.ctl.d_pz)
            self.hint.configure(text="Trascina per girare e inclinare · rotellina o +/− per avvicinarti · cursore (o frecce su/giù) per "
                                     "spostarti lungo il corridoio · clicca una stanza o una targhetta per aprirla.")

    def _first_draw(self):
        try:
            self.focus_force()
            self.canvas.focus_set()
        except Exception:
            pass
        self.stage.redraw()
        self._update_hud()

    # --- disegno
    def draw(self, rw, rh, hq):
        m, c = self.model, self.ctl
        ss = 2 if (hq and rw * rh <= 900000) else 1
        if c.mode == "fp":
            img, hits = render(m.walk_scene(), c.camera(), rw, rh, ss, FOV_WALK, hot=self._hot)
            mw = min(440, int(rw * 0.42))
            mini, rects = m.minimap(c.x, c.z, c.yaw, self.avatar, hot=self._hot[1] if self._hot else None, maxw=mw)
            ox, oy = rw - mini.size[0] - 10, 10
            img.paste(mini, (ox, oy), mini)
            self._mini_at, self._mini_rects = (ox, oy), rects
        else:
            cam = dh_camera(m, c.d_yaw, c.d_el, c.d_zoom, c.d_pz, rw, rh)
            img, hits = render(m.dh_scene(), cam, rw, rh, ss, FOV_DH, hot=self._hot)
            self._mini_rects = []
        return img, hits

    def refresh(self, hq=False, coalesce=False):
        if coalesce:
            self.stage.request(hq)
        else:
            self.stage.redraw(hq)
        self._update_hud()

    def _resize(self, e=None):
        if getattr(self, "stage", None):
            self.stage.redraw()

    # --- targhetta che sto guardando (prima persona)
    def _plate_thumb(self, uid):
        if uid not in self._plate_thumbs:
            self._plate_thumbs[uid] = _thumb_photo(self.model.plate_pack(uid)["rgba"], 88)
        return self._plate_thumbs[uid]

    def _update_hud(self):
        if self.ctl.mode != "fp":
            return
        b = self.ctl.nearest()
        uid = b[0] if b else None
        if uid == self._hud_uid:
            return
        self._hud_uid = uid
        if not uid:
            self.hud_img.configure(image="")
            self.hud_txt.configure(text="Avvicinati a una targhetta per leggerla.")
            self.btn_enter.state(["disabled"]) if hasattr(self.btn_enter, "state") else None
            return
        pd = self.model.plates[uid]
        self._hud_photo = self._plate_thumb(uid)
        self.hud_img.configure(image=self._hud_photo)
        txt = "Stanza %s" % pd["num"]
        if pd["uff"]:
            txt += "  ·  " + " · ".join(pd["uff"])
        txt += "\n" + ("Nessuno è assegnato." if pd["free"] else ", ".join(pd["names"]) + (" e altri %d" % pd["more"] if pd["more"] else ""))
        if pd["miss"] and not pd["free"]:
            txt += "\nCodici mancanti: %d" % pd["miss"]
        self.hud_txt.configure(text=txt)
        self.btn_enter.state(["!disabled"]) if hasattr(self.btn_enter, "state") else None

    # --- comandi
    def on_mode(self):
        self.ctl.mode = self.var_mode.get()
        self._hot = None
        self._hud_uid = None
        self._show_mode()
        self.refresh()

    def on_go(self):
        uid = self.go_map.get(self.var_go.get())
        if uid:
            self.go(uid)

    def go(self, uid):
        if self.ctl.mode == "fp":
            self.ctl.go_room(uid)
        else:
            it = self.model.item.get(uid)
            if it:
                self.ctl.d_pz = clamp(-it["z"], 0, self.model.LC)
                self.sl.set(self.ctl.d_pz)
        self._hud_uid = None
        self.refresh()

    def on_slide(self, v):
        if self.ctl is None or self.ctl.mode != "dh":
            return
        try:
            val = clamp(float(v), 0, self.model.LC)
        except Exception:
            return
        if abs(val - self.ctl.d_pz) < 0.5:
            return
        self.ctl.d_pz = val
        self.refresh()

    def zoom_dh(self, f):
        self.ctl.dh_zoom(f)
        self.refresh()

    def enter(self, uid):
        room = self.model.room.get(uid)
        if not room:
            return
        dlg = self.host.open_room(room["id"])
        if dlg is not None and hasattr(dlg, "bind"):
            dlg.bind("<Destroy>", lambda e, d=dlg: self.after(120, self.reload) if e.widget is d else None, add="+")
        else:
            self.after(1500, self.reload)

    def enter_hud(self):
        if self._hud_uid:
            self.enter(self._hud_uid)

    def avatar_dialog(self):
        def changed():
            self.avatar = load_avatar(self.host.avatar_path())
            self.refresh()
        AvatarDialog(self, self.host, changed)

    # --- tastiera
    def _keyname(self, e):
        k = str(getattr(e, "keysym", "")).lower()
        return k

    def _key_down(self, e):
        k = self._keyname(e)
        if k in ("return", "kp_enter") and self.ctl.mode == "fp" and self._hud_uid:
            self.enter_hud()
            return
        mv = KEYS.get(k)
        if not mv:
            return
        if k in self._rel:
            try:
                self.after_cancel(self._rel.pop(k))
            except Exception:
                self._rel.pop(k, None)
        self._held[k] = mv
        if self._tick_job is None:
            self._last = time.monotonic()
            self._tick_job = self.after(1, self._tick)

    def _key_up(self, e):
        k = self._keyname(e)
        if k in self._held:
            self._rel[k] = self.after(40, lambda kk=k: self._release_key(kk))

    def _release_key(self, k):
        self._held.pop(k, None)
        self._rel.pop(k, None)

    def _tick(self):
        self._tick_job = None
        if not self._held:
            self.stage.redraw(True, False)
            return
        now = time.monotonic()
        dt = min(0.06, max(0.001, now - self._last))
        self._last = now
        f = sum(v for ax, v in self._held.values() if ax == "f")
        t = sum(v for ax, v in self._held.values() if ax == "t")
        f, t = clamp(f, -1, 1), clamp(t, -1, 1)
        if self.ctl.mode == "fp":
            self.ctl.step(f, t, dt)
        else:
            self.ctl.dh_step(f, t, dt)
            self.sl.set(self.ctl.d_pz)
        self.stage.redraw(False)
        self._update_hud()
        self._tick_job = self.after(20, self._tick)

    # --- mouse
    def _press(self, e):
        try:
            self.canvas.focus_set()
        except Exception:
            pass
        self._drag = [e.x, e.y, 0.0]

    def _motion(self, e):
        if not self._drag:
            return
        dx, dy = e.x - self._drag[0], e.y - self._drag[1]
        self._drag[0], self._drag[1] = e.x, e.y
        self._drag[2] += abs(dx) + abs(dy)
        if self._drag[2] < 5:
            return
        if self.ctl.mode == "fp":
            self.ctl.look(dx, dy)
        else:
            self.ctl.dh_drag(dx, dy)
        self.refresh(coalesce=True)

    def _release(self, e):
        d, self._drag = self._drag, None
        if d is None or d[2] >= 5:
            return
        self.click(e.x, e.y)

    def click(self, x, y):
        """Clic (senza trascinare): mini-mappa -> vado li'; porta, targhetta o stanza -> la apro."""
        rx, ry = self.stage.to_render(x, y)
        if self.ctl.mode == "fp" and self._mini_rects:
            ox, oy = self._mini_at
            for x0, y0, x1, y1, uid in self._mini_rects:
                if ox + x0 <= rx <= ox + x1 and oy + y0 <= ry <= oy + y1:
                    self.go(uid)
                    return
        key = pick(self.stage.hits, rx, ry)
        if isinstance(key, tuple) and key[0] == "room":
            self.enter(key[1])

    def _hover(self, e):
        if self._drag:
            return
        rx, ry = self.stage.to_render(e.x, e.y)
        key = pick(self.stage.hits, rx, ry)
        hot = key if isinstance(key, tuple) and key[0] == "room" else None
        if hot != self._hot:
            self._hot = hot
            try:
                self.canvas.configure(cursor="hand2" if hot else "arrow")
            except Exception:
                pass
            self.stage.redraw(False)

    def _wheel(self, e):
        d = _ev_delta(e)
        if not d:
            return
        if self.ctl.mode == "fp":
            self.ctl.step(d, 0, 0.16)
            self.refresh()
        else:
            self.ctl.dh_zoom(1.15 if d > 0 else 1 / 1.15)
            self.refresh()

    def close(self):
        self.stage.cancel()
        for j in list(self._rel.values()) + [self._tick_job]:
            if j:
                try:
                    self.after_cancel(j)
                except Exception:
                    pass
        self._held.clear()
        self.destroy()


# ---------------------------------------------------------------- la finestra dell'edificio
def build_tower(host, edificio_id):
    floors = [FloorModel(host.floor_data(edificio_id, p), host.printer_image) for p in host.floors(edificio_id)]
    return TowerModel(host.building_name(edificio_id), floors)


class TowerWindow(tk.Toplevel if tk else object):
    """L'edificio intero: i piani impilati. Si gira intorno, si clicca un piano (o il suo pulsante) per camminarci dentro."""

    def __init__(self, host, edificio_id):
        super().__init__(host.parent)
        self.host, self.edificio_id = host, edificio_id
        self.tower = build_tower(host, edificio_id)
        self.ctl = TowerCtl(self.tower)
        self.title("%s · edificio in 3D" % self.tower.nome)
        _fit_geometry(self, 1100, 760)
        self.minsize(720, 520)
        self.configure(bg=NAVY_BG)
        self._drag = None
        self._hot = None
        bar = ttk.Frame(self, padding=(10, 8))
        bar.pack(side="top", fill="x")
        ttk.Label(bar, text="%s · edificio in 3D" % self.tower.nome, font=("Segoe UI", 11, "bold")).pack(side="left")
        ttk.Button(bar, text="Chiudi", command=self.close).pack(side="right")
        ttk.Button(bar, text="+", width=3, command=lambda: self.zoom(1.2)).pack(side="right", padx=(0, 6))
        ttk.Button(bar, text="−", width=3, command=lambda: self.zoom(1 / 1.2)).pack(side="right", padx=(0, 4))
        self.hint = ttk.Label(self, style="Hint.TLabel", padding=(10, 4, 10, 8), wraplength=1000, justify="left",
                              text="Trascina per girare intorno all'edificio · rotellina o +/− per avvicinarti · "
                                   "clicca un piano (anche dall'elenco a sinistra) per camminarci dentro.")
        self.hint.pack(side="bottom", fill="x")
        body = ttk.Frame(self)
        body.pack(side="top", fill="both", expand=True)
        rail = ttk.Frame(body, padding=(10, 0, 4, 0))
        rail.pack(side="left", fill="y")
        self.rail_btns = {}
        for fl in reversed(self.tower.floors):
            s = fl.summary()
            txt = "%s\n%d %s · %d %s\nStampanti: %d%s%s" % (
                s["label"], s["rooms"], "stanza" if s["rooms"] == 1 else "stanze",
                s["people"], "persona" if s["people"] == 1 else "persone", s["printers"],
                "  ·  ticket aperti: %d" % s["red"] if s["red"] else "",
                "  ·  toner basso" if s["low"] else "")
            b = ttk.Button(rail, text=txt, width=30, command=lambda p=fl.data.get("piano"): self.open_floor(p))
            b.pack(fill="x", pady=3)
            piano = fl.data.get("piano")
            b.bind("<Enter>", lambda e, p=piano: self.set_hot(("floor", p)))
            b.bind("<Leave>", lambda e: self.set_hot(None))
            self.rail_btns[piano] = b
        self.canvas = tk.Canvas(body, bg=NAVY_BG, highlightthickness=0, width=800, height=560)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.stage = Stage(self, self.canvas, self.draw)
        for ev, fn in (("<ButtonPress-1>", self._press), ("<B1-Motion>", self._motion),
                       ("<ButtonRelease-1>", self._release), ("<Motion>", self._hover),
                       ("<MouseWheel>", self._wheel), ("<Button-4>", self._wheel), ("<Button-5>", self._wheel),
                       ("<Configure>", lambda e: self.stage.redraw())):
            self.canvas.bind(ev, fn)
        self.bind("<Escape>", lambda e: self.close())
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.after(60, lambda: self.stage.redraw())

    def draw(self, rw, rh, hq):
        ss = 2 if (hq and rw * rh <= 900000) else 1
        cam = self.tower.camera(self.ctl.yaw, self.ctl.el, self.ctl.zoom, rw, rh)
        return render(self.tower.scene(), cam, rw, rh, ss, FOV_TW, hot=self._hot)

    def set_hot(self, key):
        if key != self._hot:
            self._hot = key
            self.stage.redraw(False)

    def zoom(self, f):
        self.ctl.zoom_by(f)
        self.stage.redraw()

    def open_floor(self, piano):
        self.host.open_walk(self.edificio_id, piano)

    def _press(self, e):
        self._drag = [e.x, e.y, 0.0]

    def _motion(self, e):
        if not self._drag:
            return
        dx, dy = e.x - self._drag[0], e.y - self._drag[1]
        self._drag[0], self._drag[1] = e.x, e.y
        self._drag[2] += abs(dx) + abs(dy)
        if self._drag[2] >= 5:
            self.ctl.drag(dx, dy)
            self.stage.request()

    def _release(self, e):
        d, self._drag = self._drag, None
        if d is None or d[2] >= 5:
            return
        key = pick(self.stage.hits, *self.stage.to_render(e.x, e.y))
        fl = self.tower.floor_at(key)
        if fl is not None:
            self.open_floor(fl.data.get("piano"))

    def _hover(self, e):
        if self._drag:
            return
        key = pick(self.stage.hits, *self.stage.to_render(e.x, e.y))
        hot = key if isinstance(key, tuple) and key and key[0] == "floor" else None
        if hot != self._hot:
            self._hot = hot
            try:
                self.canvas.configure(cursor="hand2" if hot else "arrow")
            except Exception:
                pass
            self.stage.redraw()

    def _wheel(self, e):
        d = _ev_delta(e)
        if d:
            self.ctl.zoom_by(1.15 if d > 0 else 1 / 1.15)
            self.stage.redraw()

    def close(self):
        self.stage.cancel()
        self.destroy()
