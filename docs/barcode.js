/* barcode.js - lettura di codici a barre in puro JavaScript (Code 128 e Code 39).
   Nessuna libreria esterna: funziona anche senza rete. Serve a scan.js.
   Perche' esiste: il lettore del browser (BarcodeDetector) su iPhone e' spento di fabbrica,
   quindi la lettura dei codici si fa qui, dentro l'app.
   Legge i codici lineari (quelli a barre classiche) dei seriali, degli IMEI e delle etichette.
   NON legge i QR code. */
(function (root) {
  "use strict";

  /* ---- tabelle ---- */
  /* Code 128: larghezze barra/spazio dei 107 simboli (0-105 dati e start, 106 = stop) */
  const C128 = ("212222,222122,222221,121223,121322,131222,122213,122312,132212,221213,221312,231212,112232,122132,122231,113222,123122,123221,223211,221132,221231,213212,223112,312131,311222,321122,321221,312212,322112,322211,212123,212321,232121,111323,131123,131321,112313,132113,132311,211313,231113,231311,112133,112331,132131,113123,113321,133121,313121,211331,231131,213113,213311,213131,311123,311321,331121,312113,312311,332111,314111,221411,431111,111224,111422,121124,121421,141122,141221,112214,112412,122114,122411,142112,142211,241211,221114,413111,241112,134111,111242,121142,121241,114212,124112,124211,411212,421112,421211,212141,214121,412121,111143,111341,131141,114113,114311,411113,411311,113141,114131,311141,411131,211412,211214,211232,2331112").split(",").map(s => Array.from(s, Number));
  /* Code 39: per ogni carattere i 9 elementi (barra, spazio, ...) con 1 = largo e 0 = stretto */
  const C39 = { "0": "000110100", "1": "100100001", "2": "001100001", "3": "101100000", "4": "000110001", "5": "100110000", "6": "001110000", "7": "000100101", "8": "100100100", "9": "001100100", "A": "100001001", "B": "001001001", "C": "101001000", "D": "000011001", "E": "100011000", "F": "001011000", "G": "000001101", "H": "100001100", "I": "001001100", "J": "000011100", "K": "100000011", "L": "001000011", "M": "101000010", "N": "000010011", "O": "100010010", "P": "001010010", "Q": "000000111", "R": "100000110", "S": "001000110", "T": "000010110", "U": "110000001", "V": "011000001", "W": "111000000", "X": "010010001", "Y": "110010000", "Z": "011010000", "-": "010000101", ".": "110000100", " ": "011000100", "*": "010010100", "$": "010101000", "/": "010100010", "+": "010001010", "%": "000101010" };
  const C39R = {};
  for (const k in C39) C39R[C39[k]] = k;

  /* ---- da una riga di pixel alle larghezze di barre e spazi ---- */
  /* s: luminosita' 0..255, n: numero di campioni.
     Restituisce { runs, c0 }: runs[k] e' la larghezza (in pixel, con decimali) della k-esima fascia;
     c0 = true se la fascia 0 e' scura. runs[0] e l'ultima sono fasce tagliate dal bordo. */
  function lineRuns(s, n) {
    if (n < 40) return null;
    const a = new Float32Array(n);
    a[0] = (3 * s[0] + s[1]) / 4; a[n - 1] = (s[n - 2] + 3 * s[n - 1]) / 4;
    for (let i = 1; i < n - 1; i++) a[i] = (s[i - 1] + 2 * s[i] + s[i + 1]) / 4;
    /* soglia locale: meta' tra il piu' scuro e il piu' chiaro delle vicinanze (regge alle ombre) */
    const B = 16, nb = Math.ceil(n / B), R = 4;
    const bmin = new Float32Array(nb).fill(255), bmax = new Float32Array(nb);
    for (let i = 0; i < n; i++) { const b = (i / B) | 0, v = a[i]; if (v < bmin[b]) bmin[b] = v; if (v > bmax[b]) bmax[b] = v; }
    const thr = new Float32Array(nb), con = new Float32Array(nb);
    for (let b = 0; b < nb; b++) {
      let lo = 255, hi = 0;
      for (let k = Math.max(0, b - R), e = Math.min(nb - 1, b + R); k <= e; k++) { if (bmin[k] < lo) lo = bmin[k]; if (bmax[k] > hi) hi = bmax[k]; }
      thr[b] = (lo + hi) / 2; con[b] = hi - lo;
    }
    const MINC = 28;                       // sotto questo contrasto non c'e' nessun codice
    const ed = [];
    let pd = con[0] >= MINC && a[0] < thr[0];
    const c0 = pd;
    for (let i = 1; i < n; i++) {
      const b = (i / B) | 0;
      const d = con[b] >= MINC && a[i] < thr[b];
      if (d !== pd) {
        const j = i - 1, bj = (j / B) | 0, den = a[i] - a[j];
        let t = 0.5;
        if (con[b] >= MINC && con[bj] >= MINC && den !== 0) { t = (thr[b] - a[j]) / den; t = t < 0 ? 0 : t > 1 ? 1 : t; }
        const pos = j + t;
        if (ed.length && pos - ed[ed.length - 1] < 1.0) ed.pop(); else ed.push(pos);   // disturbo: fascia sotto 1 pixel
        pd = d;
      }
    }
    const m = ed.length;
    if (m < 20) return null;
    const runs = new Float32Array(m + 1);
    runs[0] = ed[0];
    for (let k = 1; k < m; k++) runs[k] = ed[k] - ed[k - 1];
    runs[m] = n - ed[m - 1];
    return { runs, c0 };
  }

  /* ---- Code 128 ---- */
  /* confronta 6 fasce con i simboli from..to: restituisce i candidati (i migliori per primi) */
  function cand128(r, i, from, to, maxn) {
    let T = 0;
    for (let k = 0; k < 6; k++) T += r[i + k];
    if (T < 6) return [];
    const m = T / 11, c = [];
    for (let v = from; v <= to; v++) {
      const p = C128[v];
      let e = 0, mx = 0;
      for (let k = 0; k < 6; k++) { const d = Math.abs(r[i + k] / m - p[k]); e += d; if (d > mx) mx = d; }
      if (e <= 2.4 && mx <= 0.9) c.push([e, v]);
    }
    c.sort((x, y) => x[0] - y[0]);
    const best = c.length ? c[0][0] : 0;
    return c.filter(x => x[0] <= best + 0.9 && x[0] <= (c[0][0] <= 2.4 ? 3.0 : 2.4)).slice(0, maxn).map(x => x[1]);
  }
  function isStop128(r, i) {
    let T = 0;
    for (let k = 0; k < 7; k++) T += r[i + k];
    const m = T / 13, p = C128[106];
    if (m < 0.5) return false;
    let e = 0, mx = 0;
    for (let k = 0; k < 7; k++) { const d = Math.abs(r[i + k] / m - p[k]); e += d; if (d > mx) mx = d; }
    return e <= 2.2 && mx <= 0.85 && r[i + 7] >= 2.5 * m;       // dopo lo stop deve esserci la zona bianca
  }
  function text128(vals) {
    let set = vals[0] === 103 ? "A" : vals[0] === 104 ? "B" : "C", shiftTo = "", s = "";
    for (let i = 1; i < vals.length; i++) {
      const v = vals[i], cur = shiftTo || set;
      shiftTo = "";
      if (cur === "C") {
        if (v < 100) s += (v < 10 ? "0" : "") + v;
        else if (v === 100) set = "B";
        else if (v === 101) set = "A";
        continue;                                    // 102 = FNC1: lo salto
      }
      if (v < 96) { s += String.fromCharCode(cur === "A" ? (v < 64 ? v + 32 : v - 64) : v + 32); continue; }
      if (v === 98) { shiftTo = (set === "A" ? "B" : "A"); continue; }        // SHIFT: solo il prossimo carattere
      if (v === 99) { set = "C"; continue; }
      if (cur === "A" && v === 100) { set = "B"; continue; }
      if (cur === "B" && v === 101) { set = "A"; continue; }
      /* 96, 97, 102 e FNC4: funzioni che non producono testo */
    }
    return s;
  }
  /* cerca, da pos in poi, la fila di simboli che chiude con uno stop e un checksum giusto */
  function walk128(r, pos, vals, n, bud) {
    if (vals.length >= 5 && pos + 7 <= n - 1 && isStop128(r, pos)) {          // almeno start + 3 dati + checksum
      const chk = vals[vals.length - 1];
      let sum = vals[0];
      for (let k = 1; k < vals.length - 1; k++) sum += k * vals[k];
      if (sum % 103 === chk) return { vals: vals.slice(0, -1), pos };
    }
    if (pos + 6 > n - 2 || vals.length >= 110 || bud.n-- <= 0) return null;
    for (const v of cand128(r, pos, 0, 102, 1)) {
      vals.push(v);
      const out = walk128(r, pos + 6, vals, n, bud);
      vals.pop();
      if (out) return out;
    }
    return null;
  }
  function decode128(r, c0) {
    const n = r.length, out = [];
    for (let i = c0 ? 2 : 1; i + 24 <= n - 2; i += 2) {
      for (const sv of cand128(r, i, 103, 105, 1)) {
        const f = walk128(r, i + 6, [sv], n, { n: 120 });
        if (!f) continue;
        const text = text128(f.vals).replace(/[\x00-\x1f\x7f]/g, "");
        if (text) out.push({ text, format: "code128", n: f.vals.length - 1 });
        i = f.pos;                                   // salto oltre questo codice
        break;
      }
    }
    return out;
  }

  /* ---- Code 39 ---- */
  function char39(r, i) {
    const w = new Array(9), idx = [0, 1, 2, 3, 4, 5, 6, 7, 8];
    for (let k = 0; k < 9; k++) w[k] = r[i + k];
    idx.sort((x, y) => w[y] - w[x]);
    const wideMin = w[idx[2]], narrowMax = w[idx[3]];
    if (narrowMax <= 0 || wideMin < 1.12 * narrowMax) return null;
    let wa = 0, na = 0;
    for (let k = 0; k < 3; k++) wa += w[idx[k]];
    for (let k = 3; k < 9; k++) na += w[idx[k]];
    wa /= 3; na /= 6;
    if (na < 1 || wa / na < 1.7 || wa / na > 3.8) return null;
    const bits = new Array(9).fill("0");
    for (let k = 0; k < 3; k++) bits[idx[k]] = "1";
    const ch = C39R[bits.join("")];
    return ch ? { ch, narrow: na } : null;
  }
  function decode39(r, c0) {
    const n = r.length, out = [];
    for (let i = c0 ? 2 : 1; i + 8 <= n - 2; i += 2) {
      const s0 = char39(r, i);
      if (!s0 || s0.ch !== "*") continue;
      const nr = s0.narrow;
      if (r[i - 1] < 3 * nr) continue;               // zona bianca prima dell'inizio
      let pos = i + 9, txt = "", ok = false;
      while (pos + 9 <= n - 2 && txt.length <= 60) {
        const gap = r[pos];
        if (gap > 3 * nr || gap < 0.3 * nr) break;
        const c = char39(r, pos + 1);
        if (!c || c.narrow < 0.55 * nr || c.narrow > 1.8 * nr) break;
        if (c.ch === "*") { ok = txt.length >= 4 && r[pos + 10] >= 3 * nr; break; }
        txt += c.ch; pos += 10;
      }
      if (ok) { out.push({ text: txt, format: "code39", n: txt.length }); i = pos; }
    }
    return out;
  }

  /* ---- una riga di pixel: prova i due versi e i due formati ---- */
  function decodeLine(s, n, formats) {
    const lr = lineRuns(s, n);
    if (!lr) return [];
    const want = f => !formats || formats.includes(f);
    const fw = lr.runs, len = fw.length, rv = Float32Array.from(fw).reverse();
    const c0r = ((len - 1) % 2 === 0) === lr.c0;
    const out = [];
    for (const [r, c] of [[fw, lr.c0], [rv, c0r]]) {
      if (want("code128")) out.push(...decode128(r, c));
      if (want("code39")) out.push(...decode39(r, c));
    }
    return out;
  }

  /* ---- un'immagine: tante righe, in tante direzioni ---- */
  function clip(px, py, dx, dy, W, H) {
    let t0 = -1e9, t1 = 1e9;
    for (const [p, d, M] of [[px, dx, W], [py, dy, H]]) {
      if (Math.abs(d) < 1e-9) { if (p < 0 || p > M) return null; continue; }
      let ta = (0 - p) / d, tb = (M - p) / d;
      if (ta > tb) { const t = ta; ta = tb; tb = t; }
      if (ta > t0) t0 = ta;
      if (tb < t1) t1 = tb;
    }
    return t1 > t0 ? [t0, t1] : null;
  }
  const ALL_ANGLES = [0, 90, 15, -15, 30, -30, 45, -45, 60, -60, 75, -75];

  /* gray: Uint8Array w*h in grigi. opt: { angles, lines, formats }.
     Restituisce i codici trovati: { text, format, n, votes } (votes = quante righe lo hanno letto). */
  function decodeGray(gray, w, h, opt) {
    opt = opt || {};
    const angles = opt.angles || ALL_ANGLES, N = opt.lines || 16, par = opt.par || 1, found = new Map();
    const W = w - 1, H = h - 1, cx = W / 2, cy = H / 2, buf = new Float32Array(Math.ceil(Math.hypot(w, h)) + 4);
    for (const deg of angles) {
      const th = deg * Math.PI / 180, dx = Math.cos(th), dy = Math.sin(th), nx = -dy, ny = dx;
      const E = (w * Math.abs(dy) + h * Math.abs(dx)) / 2;
      for (let j = 0; j < N; j++) {
        const o = -E + (j + 0.5) * 2 * E / N, px = cx + o * nx, py = cy + o * ny;
        const c = clip(px, py, dx, dy, W, H);
        if (!c) continue;
        const x0 = px + c[0] * dx, y0 = py + c[0] * dy, len = c[1] - c[0], n = Math.floor(len);
        if (n < 40) continue;
        for (let i = 0; i < n; i++) {
          let acc = 0;
          for (let k = 0; k < par; k++) {
            const dd = (k - (par - 1) / 2) * 1.2;
            let x = x0 + dx * i + nx * dd, y = y0 + dy * i + ny * dd;
            x = x < 0 ? 0 : x > W ? W : x; y = y < 0 ? 0 : y > H ? H : y;
            const xi = x | 0, yi = y | 0, fx = x - xi, fy = y - yi;
            const x1 = xi < W ? xi + 1 : xi, y1 = yi < H ? yi + 1 : yi;
            acc += (gray[yi * w + xi] * (1 - fx) + gray[yi * w + x1] * fx) * (1 - fy) + (gray[y1 * w + xi] * (1 - fx) + gray[y1 * w + x1] * fx) * fy;
          }
          buf[i] = acc / par;
        }
        for (const r of decodeLine(buf, n, opt.formats)) {
          const key = r.format + "|" + r.text, old = found.get(key);
          if (old) old.votes++; else found.set(key, { text: r.text, format: r.format, n: r.n, votes: 1 });
        }
      }
    }
    return [...found.values()].sort((a, b) => b.votes - a.votes || b.n - a.n);
  }

  /* RGBA (da canvas) -> grigi */
  function toGray(rgba, w, h, out) {
    const g = out && out.length === w * h ? out : new Uint8Array(w * h);
    for (let i = 0, j = 0; i < g.length; i++, j += 4) g[i] = (rgba[j] * 77 + rgba[j + 1] * 150 + rgba[j + 2] * 29) >> 8;
    return g;
  }

  const API = { decodeGray, decodeLine, toGray, lineRuns, ALL_ANGLES };
  if (typeof module !== "undefined" && module.exports) module.exports = API;
  else root.BARCODE = API;
})(typeof window !== "undefined" ? window : this);
