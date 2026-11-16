/* qr.js - lettura dei QR code (versioni 1-10) dentro l'app, senza librerie esterne.
   Serve all'iPhone: li' il lettore del browser e' spento e barcode.js legge solo i codici a barre.
   Uso: QR.decodeGray(gray, w, h) -> [{ text, format: "qr", votes: 3, n }]  (gray = Uint8Array con w*h grigi)
   Passi: bianco/nero (soglia che si adatta alla luce) -> i tre quadrati d'angolo -> griglia dei moduli
   (con il quadratino di allineamento, per correggere la prospettiva) -> informazioni di formato ->
   dati -> correzione degli errori (Reed-Solomon) -> testo. Se il codice e' rovinato torna una lista vuota. */
(function () {
  "use strict";

  /* ---------- campo di Galois a 256 elementi (x^8+x^4+x^3+x^2+1) ---------- */
  const EXP = new Uint8Array(512), LOG = new Uint8Array(256);
  (function () {
    let x = 1;
    for (let i = 0; i < 255; i++) { EXP[i] = x; LOG[x] = i; x <<= 1; if (x & 256) x ^= 0x11d; }
    for (let i = 255; i < 512; i++) EXP[i] = EXP[i - 255];
  })();
  const gmul = (a, b) => (a && b) ? EXP[LOG[a] + LOG[b]] : 0;
  const gdiv = (a, b) => a ? EXP[(LOG[a] + 255 - LOG[b]) % 255] : 0;

  /* corregge i codici di un blocco (dati + controllo); restituisce false se sono troppi errori */
  function rsCorrect(cw, nsym) {
    const n = cw.length, S = new Array(nsym);
    let bad = false;
    for (let j = 0; j < nsym; j++) {
      let s = 0;
      for (let i = 0; i < n; i++) s = gmul(s, EXP[j]) ^ cw[i];
      S[j] = s; if (s) bad = true;
    }
    if (!bad) return true;
    /* Berlekamp-Massey: polinomio localizzatore (grado minimo prima) */
    let C = [1], B = [1], L = 0, m = 1, b = 1;
    for (let k = 0; k < nsym; k++) {
      let d = S[k];
      for (let i = 1; i <= L; i++) d ^= gmul(C[i] || 0, S[k - i]);
      if (d === 0) { m++; continue; }
      const T = C.slice(), coef = gdiv(d, b);
      for (let i = 0; i < B.length; i++) C[i + m] = (C[i + m] || 0) ^ gmul(coef, B[i]);
      if (2 * L <= k) { L = k + 1 - L; B = T; b = d; m = 1; } else m++;
    }
    while (C.length && C[C.length - 1] === 0) C.pop();
    const deg = C.length - 1;
    if (deg < 1 || deg > nsym / 2) return false;
    /* ricerca di Chien: posizioni (potenza i = n-1-indice) in cui sigma(alfa^-i) = 0 */
    const pos = [];
    for (let i = 0; i < n; i++) {
      let v = 0;
      for (let t = 0; t < C.length; t++) if (C[t]) v ^= EXP[(LOG[C[t]] + ((255 - i) % 255) * t) % 255];
      if (v === 0) pos.push(i);
    }
    if (pos.length !== deg) return false;
    /* evaluatore Omega = S * sigma mod x^nsym, poi Forney */
    const om = new Array(nsym).fill(0);
    for (let i = 0; i < nsym; i++) for (let t = 0; t < C.length && t <= i; t++) om[i] ^= gmul(C[t], S[i - t]);
    for (const p of pos) {
      const xi = EXP[p], xinv = EXP[(255 - p) % 255];
      let o = 0, xp = 1;
      for (let i = 0; i < nsym; i++) { o ^= gmul(om[i], xp); xp = gmul(xp, xinv); }
      let dd = 0;
      for (let t = 1; t < C.length; t += 2) { let q = 1; for (let u = 0; u < t - 1; u++) q = gmul(q, xinv); dd ^= gmul(C[t], q); }
      if (!dd) return false;
      cw[n - 1 - p] ^= gmul(xi, gdiv(o, dd));
    }
    for (let j = 0; j < nsym; j++) {
      let s = 0;
      for (let i = 0; i < n; i++) s = gmul(s, EXP[j]) ^ cw[i];
      if (s) return false;
    }
    return true;
  }

  /* ---------- tabelle ---------- */
  /* per versione 1-10 e livello L,M,Q,H: [codici di controllo per blocco, [[quanti blocchi, codici dati per blocco], ...]] */
  const BLOCKS = [null,
    [[7, [[1, 19]]], [10, [[1, 16]]], [13, [[1, 13]]], [17, [[1, 9]]]],
    [[10, [[1, 34]]], [16, [[1, 28]]], [22, [[1, 22]]], [28, [[1, 16]]]],
    [[15, [[1, 55]]], [26, [[1, 44]]], [18, [[2, 17]]], [22, [[2, 13]]]],
    [[20, [[1, 80]]], [18, [[2, 32]]], [26, [[2, 24]]], [16, [[4, 9]]]],
    [[26, [[1, 108]]], [24, [[2, 43]]], [18, [[2, 15], [2, 16]]], [22, [[2, 11], [2, 12]]]],
    [[18, [[2, 68]]], [16, [[4, 27]]], [24, [[4, 19]]], [28, [[4, 15]]]],
    [[20, [[2, 78]]], [18, [[4, 31]]], [18, [[2, 14], [4, 15]]], [26, [[4, 13], [1, 14]]]],
    [[24, [[2, 97]]], [22, [[2, 38], [2, 39]]], [22, [[4, 18], [2, 19]]], [26, [[4, 14], [2, 15]]]],
    [[30, [[2, 116]]], [22, [[3, 36], [2, 37]]], [20, [[4, 16], [4, 17]]], [24, [[4, 12], [4, 13]]]],
    [[18, [[2, 68], [2, 69]]], [26, [[4, 43], [1, 44]]], [24, [[6, 19], [2, 20]]], [28, [[6, 15], [2, 16]]]]];
  const ALIGN = [null, [], [6, 18], [6, 22], [6, 26], [6, 30], [6, 34], [6, 22, 38], [6, 24, 42], [6, 26, 46], [6, 28, 50]];
  const EC_BITS = [1, 0, 3, 2];                      // livelli L, M, Q, H -> i due bit scritti nel formato
  const FMT = [];                                    // tutti i 32 formati validi, gia' con la maschera 0x5412
  for (let d = 0; d < 32; d++) {
    let rem = d << 10;
    for (let i = 14; i >= 10; i--) if ((rem >> i) & 1) rem ^= 0x537 << (i - 10);
    FMT.push([((d << 10) | rem) ^ 0x5412, d]);
  }
  const popcount = x => { let c = 0; while (x) { c += x & 1; x >>= 1; } return c; };
  const ALNUM = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:";
  const MASKS = [(r, c) => (r + c) % 2 === 0, (r) => r % 2 === 0, (r, c) => c % 3 === 0, (r, c) => (r + c) % 3 === 0,
    (r, c) => (((r >> 1) + Math.floor(c / 3)) % 2) === 0, (r, c) => ((r * c) % 2 + (r * c) % 3) === 0,
    (r, c) => (((r * c) % 2 + (r * c) % 3) % 2) === 0, (r, c) => ((((r + c) % 2) + (r * c) % 3) % 2) === 0];

  /* ---------- dai moduli al testo ---------- */
  function funcMap(dim, ver) {
    const f = new Uint8Array(dim * dim), set = (r, c) => { if (r >= 0 && c >= 0 && r < dim && c < dim) f[r * dim + c] = 1; };
    for (let r = 0; r < 9; r++) for (let c = 0; c < 9; c++) set(r, c);
    for (let r = 0; r < 9; r++) for (let c = dim - 8; c < dim; c++) set(r, c);
    for (let r = dim - 8; r < dim; r++) for (let c = 0; c < 9; c++) set(r, c);
    for (let i = 0; i < dim; i++) { set(6, i); set(i, 6); }
    const al = ALIGN[ver];
    for (const a of al) for (const b of al) {
      if ((a === 6 && b === 6) || (a === 6 && b === al[al.length - 1]) || (b === 6 && a === al[al.length - 1])) continue;
      for (let i = -2; i <= 2; i++) for (let j = -2; j <= 2; j++) set(a + i, b + j);
    }
    if (ver >= 7) for (let i = 0; i < 6; i++) for (let j = 0; j < 3; j++) { set(i, dim - 11 + j); set(dim - 11 + j, i); }
    return f;
  }
  function readFormat(mod, dim) {
    const g = (c, r) => mod[r * dim + c];
    let a = 0, b = 0;
    for (let i = 0; i <= 5; i++) a = (a << 1) | g(i, 8);
    a = (a << 1) | g(7, 8); a = (a << 1) | g(8, 8); a = (a << 1) | g(8, 7);
    for (let j = 5; j >= 0; j--) a = (a << 1) | g(8, j);
    for (let j = dim - 1; j >= dim - 7; j--) b = (b << 1) | g(8, j);
    for (let i = dim - 8; i < dim; i++) b = (b << 1) | g(i, 8);
    let best = null, bd = 99;
    for (const v of [a, b]) for (const [code, d] of FMT) { const dist = popcount(v ^ code); if (dist < bd) { bd = dist; best = d; } }
    return bd <= 3 ? { ec: (function (x) { return EC_BITS.indexOf(x); })((best >> 3) & 3), mask: best & 7 } : null;
  }
  function readText(mod, dim, ver, fmt) {
    const tab = BLOCKS[ver] && BLOCKS[ver][fmt.ec];
    if (!tab) return null;
    const func = funcMap(dim, ver), mf = MASKS[fmt.mask], bits = [];
    let up = true;
    for (let col = dim - 1; col >= 1; col -= 2) {
      if (col === 6) col = 5;
      for (let k = 0; k < dim; k++) {
        const row = up ? dim - 1 - k : k;
        for (let j = 0; j < 2; j++) {
          const c = col - j;
          if (func[row * dim + c]) continue;
          bits.push(mod[row * dim + c] ^ (mf(row, c) ? 1 : 0));
        }
      }
      up = !up;
    }
    const ecPer = tab[0], groups = tab[1], blocks = [];
    for (const [cnt, dcw] of groups) for (let i = 0; i < cnt; i++) blocks.push({ data: dcw, cw: [] });
    let total = 0;
    for (const bl of blocks) total += bl.data + ecPer;
    if (bits.length < total * 8) return null;
    const bytes = new Array(total);
    for (let i = 0; i < total; i++) { let v = 0; for (let j = 0; j < 8; j++) v = (v << 1) | bits[i * 8 + j]; bytes[i] = v; }
    let p = 0;
    const maxData = Math.max(...blocks.map(x => x.data));
    for (let i = 0; i < maxData; i++) for (const bl of blocks) if (i < bl.data) bl.cw.push(bytes[p++]);
    for (let i = 0; i < ecPer; i++) for (const bl of blocks) bl.cw.push(bytes[p++]);
    const out = [];
    for (const bl of blocks) {
      if (!rsCorrect(bl.cw, ecPer)) return null;
      for (let i = 0; i < bl.data; i++) out.push(bl.cw[i]);
    }
    return parseData(out, ver);
  }
  function parseData(cw, ver) {
    let pos = 0;
    const nb = cw.length * 8;
    const rd = n => { let v = 0; for (let i = 0; i < n; i++) { const q = pos++; v = (v << 1) | (q < nb ? (cw[q >> 3] >> (7 - (q & 7))) & 1 : 0); } return v; };
    const big = ver >= 10;
    let text = "";
    const raw = [];
    const flush = () => {
      if (!raw.length) return;
      let s = "";
      try { s = new TextDecoder("utf-8", { fatal: true }).decode(new Uint8Array(raw)); } catch (e) { s = String.fromCharCode.apply(null, raw); }
      text += s; raw.length = 0;
    };
    while (pos + 4 <= nb) {
      const mode = rd(4);
      if (mode === 0) break;
      if (mode === 1) {
        let n = rd(big ? 12 : 10);
        flush();
        while (n >= 3) { const v = rd(10); if (v > 999) return null; text += String(v).padStart(3, "0"); n -= 3; }
        if (n === 2) { const v = rd(7); if (v > 99) return null; text += String(v).padStart(2, "0"); }
        else if (n === 1) { const v = rd(4); if (v > 9) return null; text += String(v); }
      } else if (mode === 2) {
        let n = rd(big ? 11 : 9);
        flush();
        while (n >= 2) { const v = rd(11); if (v >= 45 * 45) return null; text += ALNUM[Math.floor(v / 45)] + ALNUM[v % 45]; n -= 2; }
        if (n === 1) { const v = rd(6); if (v >= 45) return null; text += ALNUM[v]; }
      } else if (mode === 4) {
        const n = rd(big ? 16 : 8);
        for (let i = 0; i < n; i++) raw.push(rd(8));
      } else if (mode === 7) {
        const f = rd(8);
        if (f & 0x80) { if (f & 0x40) rd(16); else rd(8); }
        flush();
      } else return null;                                 // kanji, FNC1, strutturato: non servono per i seriali
    }
    flush();
    return text;
  }

  /* ---------- bianco e nero ---------- */
  function binAdaptive(gray, w, h, win, k) {
    const iw = w + 1, ii = new Float64Array(iw * (h + 1));
    for (let y = 0; y < h; y++) {
      let row = 0;
      for (let x = 0; x < w; x++) { row += gray[y * w + x]; ii[(y + 1) * iw + x + 1] = ii[y * iw + x + 1] + row; }
    }
    const bin = new Uint8Array(w * h), r = win >> 1;
    for (let y = 0; y < h; y++) {
      const y0 = Math.max(0, y - r), y1 = Math.min(h, y + r + 1);
      for (let x = 0; x < w; x++) {
        const x0 = Math.max(0, x - r), x1 = Math.min(w, x + r + 1);
        const s = ii[y1 * iw + x1] - ii[y0 * iw + x1] - ii[y1 * iw + x0] + ii[y0 * iw + x0];
        bin[y * w + x] = gray[y * w + x] * (x1 - x0) * (y1 - y0) < s * k ? 1 : 0;
      }
    }
    return bin;
  }
  function binOtsu(gray, w, h) {
    const hist = new Float64Array(256);
    for (let i = 0; i < w * h; i++) hist[gray[i]]++;
    const tot = w * h;
    let sum = 0;
    for (let i = 0; i < 256; i++) sum += i * hist[i];
    let wb = 0, sb = 0, best = 0, th = 128;
    for (let t = 0; t < 256; t++) {
      wb += hist[t]; if (!wb) continue;
      const wf = tot - wb; if (!wf) break;
      sb += t * hist[t];
      const d = sb / wb - (sum - sb) / wf, v = wb * wf * d * d;
      if (v > best) { best = v; th = t; }
    }
    const bin = new Uint8Array(w * h);
    for (let i = 0; i < w * h; i++) bin[i] = gray[i] <= th ? 1 : 0;
    return bin;
  }

  /* ---------- i tre quadrati d'angolo ---------- */
  function checkRatio(r) {                             // r = 5 lunghezze di fila (scuro, chiaro, scuro, chiaro, scuro)
    const tot = r[0] + r[1] + r[2] + r[3] + r[4];
    if (tot < 7) return 0;
    const m = tot / 7, v = m * 0.62;
    if (Math.abs(r[0] - m) < v && Math.abs(r[1] - m) < v && Math.abs(r[2] - 3 * m) < 3 * v && Math.abs(r[3] - m) < v && Math.abs(r[4] - m) < v) return m;
    return 0;
  }
  /* conta i 5 tratti attorno a (x,y) lungo una direzione (dx,dy); restituisce [lunghezze, indice del tratto centrale scuro] o null */
  function runsAlong(bin, w, h, x, y, dx, dy) {
    const at = i => { const px = x + dx * i, py = y + dy * i; return (px < 0 || py < 0 || px >= w || py >= h) ? -1 : bin[py * w + px]; };
    if (at(0) !== 1) return null;
    const cap = Math.max(w, h);
    let a = 0; while (a < cap && at(-a - 1) === 1) a++;
    let b = 0; while (b < cap && at(b + 1) === 1) b++;
    const center = a + b + 1;
    let a1 = 0; { let i = -a - 1; while (at(i) === 0) { a1++; i--; } if (at(i) !== 1) return null; }
    let a2 = 0; { let i = -a - a1 - 1; while (at(i) === 1) { a2++; i--; } }
    let b1 = 0; { let i = b + 1; while (at(i) === 0) { b1++; i++; } if (at(i) !== 1) return null; }
    let b2 = 0; { let i = b + b1 + 1; while (at(i) === 1) { b2++; i++; } }
    return { r: [a2, a1, center, b1, b2], off: (b - a) / 2 };            // off: dove cade il centro rispetto al punto di partenza
  }
  function findFinders(bin, w, h) {
    const cand = [];
    const step = h > 1500 ? 2 : 1;
    for (let y = 0; y < h; y += step) {
      const rl = [];
      let cur = bin[y * w], len = 0;
      let runStart = 0;
      const starts = [];
      for (let x = 0; x <= w; x++) {
        const v = x < w ? bin[y * w + x] : 1 - cur;
        if (v === cur) { len++; continue; }
        rl.push(len); starts.push(runStart);
        if (cur === 1 && rl.length >= 5) {
          const r = rl.slice(-5), m = checkRatio(r);
          if (m) {
            const s0 = starts[starts.length - 5], tot = r[0] + r[1] + r[2] + r[3] + r[4];
            const cx = Math.round(s0 + tot / 2 - 0.5), vr = runsAlong(bin, w, h, cx, y, 0, 1);
            if (vr && checkRatio(vr.r)) {
              const cy = Math.round(y + vr.off), hr = runsAlong(bin, w, h, cx, cy, 1, 0);
              if (hr && checkRatio(hr.r)) {
                const fx = cx + hr.off, fy = cy, size = (checkRatio(hr.r) + checkRatio(vr.r)) / 2;
                const ex = cand.find(c => Math.abs(c.x - fx) < size * 2 && Math.abs(c.y - fy) < size * 2 && Math.abs(c.size - size) < size);
                if (ex) { ex.x = (ex.x * ex.n + fx) / (ex.n + 1); ex.y = (ex.y * ex.n + fy) / (ex.n + 1); ex.size = (ex.size * ex.n + size) / (ex.n + 1); ex.n++; }
                else cand.push({ x: fx, y: fy, size, n: 1 });
              }
            }
          }
        }
        cur = v; len = 1; runStart = x;
        if (rl.length > 8) { rl.splice(0, rl.length - 8); starts.splice(0, starts.length - 8); }
      }
    }
    return cand.filter(c => c.n >= 2).sort((a, b) => b.n - a.n).slice(0, 12);
  }
  function triples(c) {
    const out = [];
    for (let i = 0; i < c.length; i++) for (let j = i + 1; j < c.length; j++) for (let k = j + 1; k < c.length; k++) {
      const P = [c[i], c[j], c[k]], d = [];
      for (let a = 0; a < 3; a++) { const b = (a + 1) % 3; d.push(Math.hypot(P[a].x - P[b].x, P[a].y - P[b].y)); }   // d[a]: lato tra a e a+1
      const mx = Math.max(...d), idx = d.indexOf(mx), s1 = d[(idx + 1) % 3], s2 = d[(idx + 2) % 3];
      const sizes = P.map(p => p.size), smax = Math.max(...sizes), smin = Math.min(...sizes);
      if (smax > smin * 1.7) continue;
      if (Math.abs(s1 - s2) > 0.3 * Math.max(s1, s2)) continue;
      if (Math.abs(mx - Math.SQRT2 * (s1 + s2) / 2) > 0.3 * mx) continue;
      const corner = P[(idx + 2) % 3];                                  // il vertice opposto al lato lungo: l'angolo in alto a sinistra
      const o1 = P[idx], o2 = P[(idx + 1) % 3];
      const cr = (o1.x - corner.x) * (o2.y - corner.y) - (o1.y - corner.y) * (o2.x - corner.x);
      const TR = cr > 0 ? o1 : o2, BL = cr > 0 ? o2 : o1;
      const dev = Math.abs(s1 - s2) / Math.max(s1, s2) + Math.abs(mx - Math.SQRT2 * (s1 + s2) / 2) / mx;
      out.push({ TL: corner, TR, BL, dev, n: P[0].n + P[1].n + P[2].n });
    }
    return out.sort((a, b) => a.dev - b.dev || b.n - a.n).slice(0, 4);
  }

  /* ---------- geometria ---------- */
  function homography(src, dst) {                                       // 4 punti (u,v) -> (x,y); 8 incognite
    const A = [], B = [];
    for (let i = 0; i < 4; i++) {
      const [u, v] = src[i], [x, y] = dst[i];
      A.push([u, v, 1, 0, 0, 0, -u * x, -v * x]); B.push(x);
      A.push([0, 0, 0, u, v, 1, -u * y, -v * y]); B.push(y);
    }
    for (let i = 0; i < 8; i++) {
      let p = i;
      for (let r = i + 1; r < 8; r++) if (Math.abs(A[r][i]) > Math.abs(A[p][i])) p = r;
      if (Math.abs(A[p][i]) < 1e-9) return null;
      [A[i], A[p]] = [A[p], A[i]]; [B[i], B[p]] = [B[p], B[i]];
      for (let r = i + 1; r < 8; r++) {
        const f = A[r][i] / A[i][i];
        for (let c = i; c < 8; c++) A[r][c] -= f * A[i][c];
        B[r] -= f * B[i];
      }
    }
    const h = new Array(8);
    for (let i = 7; i >= 0; i--) {
      let s = B[i];
      for (let c = i + 1; c < 8; c++) s -= A[i][c] * h[c];
      h[i] = s / A[i][i];
    }
    return (u, v) => { const d = h[6] * u + h[7] * v + 1; return [(h[0] * u + h[1] * v + h[2]) / d, (h[3] * u + h[4] * v + h[5]) / d]; };
  }
  function affineMap(T, dim) {
    const k = dim - 7, sx = (T.TR.x - T.TL.x) / k, sy = (T.TR.y - T.TL.y) / k, tx = (T.BL.x - T.TL.x) / k, ty = (T.BL.y - T.TL.y) / k;
    return (u, v) => [T.TL.x + (u - 3.5) * sx + (v - 3.5) * tx, T.TL.y + (u - 3.5) * sy + (v - 3.5) * ty];
  }
  function sampleGrid(bin, w, h, map, dim) {
    const mod = new Uint8Array(dim * dim), at = (x, y) => { const px = Math.round(x), py = Math.round(y); return (px < 0 || py < 0 || px >= w || py >= h) ? 0 : bin[py * w + px]; };
    const o = 0.22, P = [[0, 0], [o, 0], [-o, 0], [0, o], [0, -o]];
    for (let r = 0; r < dim; r++) for (let c = 0; c < dim; c++) {
      let s = 0;
      for (const [du, dv] of P) { const q = map(c + 0.5 + du, r + 0.5 + dv); s += at(q[0], q[1]); }
      mod[r * dim + c] = s >= 3 ? 1 : 0;
    }
    return mod;
  }
  /* campionamento sul grigio (interpolato): regge molto meglio i moduli piccoli di 2-3 pixel; la soglia la fissa il simbolo stesso */
  function sampleGray(gray, w, h, map, dim) {
    const g = (x, y) => {
      if (x < 0 || y < 0 || x > w - 1 || y > h - 1) return 255;
      const x0 = x | 0, y0 = y | 0, x1 = Math.min(w - 1, x0 + 1), y1 = Math.min(h - 1, y0 + 1), fx = x - x0, fy = y - y0;
      return (gray[y0 * w + x0] * (1 - fx) + gray[y0 * w + x1] * fx) * (1 - fy) + (gray[y1 * w + x0] * (1 - fx) + gray[y1 * w + x1] * fx) * fy;
    };
    const o = 0.2, P = [[0, 0], [o, 0], [-o, 0], [0, o], [0, -o]], vals = new Float64Array(dim * dim);
    let lo = 255, hi = 0;
    for (let r = 0; r < dim; r++) for (let c = 0; c < dim; c++) {
      let s = 0;
      for (const [du, dv] of P) { const q = map(c + 0.5 + du, r + 0.5 + dv); s += g(q[0], q[1]); }
      s /= 5; vals[r * dim + c] = s;
      if (s < lo) lo = s; if (s > hi) hi = s;
    }
    const mod = new Uint8Array(dim * dim);
    if (hi - lo < 25) return mod;
    /* soglia di Otsu sui valori dei moduli */
    const hist = new Float64Array(64);
    for (let i = 0; i < vals.length; i++) hist[Math.min(63, Math.floor((vals[i] - lo) / (hi - lo + 1e-6) * 64))]++;
    let tot = vals.length, sum = 0;
    for (let i = 0; i < 64; i++) sum += i * hist[i];
    let wb = 0, sb = 0, best = -1, bt = 32;
    for (let t = 0; t < 63; t++) {
      wb += hist[t]; sb += t * hist[t];
      const wf = tot - wb; if (!wb || !wf) continue;
      const d = sb / wb - (sum - sb) / wf, v = wb * wf * d * d;
      if (v > best) { best = v; bt = t; }
    }
    const th = lo + (bt + 1) / 64 * (hi - lo);
    for (let i = 0; i < vals.length; i++) mod[i] = vals[i] < th ? 1 : 0;
    return mod;
  }
  function timingScore(mod, dim) {
    let ok = 0, n = 0;
    for (let i = 8; i < dim - 8; i++) {
      const want = i % 2 === 0 ? 1 : 0;
      ok += (mod[6 * dim + i] === want) + (mod[i * dim + 6] === want); n += 2;
    }
    return n ? ok / n : 0;
  }
  function findAlignment(bin, w, h, aff, dim) {
    const at = (x, y) => { const px = Math.round(x), py = Math.round(y); return (px < 0 || py < 0 || px >= w || py >= h) ? -1 : bin[py * w + px]; };
    const score = (du, dv) => {
      let sc = 0;
      for (let i = -2; i <= 2; i++) for (let j = -2; j <= 2; j++) {
        const q = aff(dim - 6.5 + du + i, dim - 6.5 + dv + j);
        if (at(q[0], q[1]) === (Math.max(Math.abs(i), Math.abs(j)) === 1 ? 0 : 1)) sc++;
      }
      return sc;
    };
    /* prima a passi larghi (la prospettiva puo' spostare il quadratino di qualche modulo), poi a passi fini */
    let best = -1, bu = 0, bv = 0;
    for (let du = -6; du <= 6.01; du += 0.5) for (let dv = -6; dv <= 6.01; dv += 0.5) {
      const sc = score(du, dv);
      if (sc > best || (sc === best && Math.abs(du) + Math.abs(dv) < Math.abs(bu) + Math.abs(bv))) { best = sc; bu = du; bv = dv; }
    }
    if (best < 19) return null;
    let fb = -1, su = 0, sv = 0, cnt = 0;
    for (let du = bu - 0.5; du <= bu + 0.51; du += 0.125) for (let dv = bv - 0.5; dv <= bv + 0.51; dv += 0.125) {
      const sc = score(du, dv);
      if (sc > fb) { fb = sc; su = du; sv = dv; cnt = 1; } else if (sc === fb) { su += du; sv += dv; cnt++; }
    }
    if (fb < 22) return null;
    return aff(dim - 6.5 + su / cnt, dim - 6.5 + sv / cnt);
  }

  /* lato di un modulo misurato lungo gli assi del simbolo (con il simbolo ruotato o in prospettiva le righe dell'immagine ingannano) */
  function moduleAlong(bin, w, h, T) {
    const ux = T.TR.x - T.TL.x, uy = T.TR.y - T.TL.y, vx = T.BL.x - T.TL.x, vy = T.BL.y - T.TL.y;
    const ul = Math.hypot(ux, uy), vl = Math.hypot(vx, vy), est = [];
    const at = (x, y) => { const px = Math.round(x), py = Math.round(y); return (px < 0 || py < 0 || px >= w || py >= h) ? 0 : bin[py * w + px]; };
    for (const P of [T.TL, T.TR, T.BL]) {
      for (const [dx, dy] of [[ux / ul, uy / ul], [-ux / ul, -uy / ul], [vx / vl, vy / vl], [-vx / vl, -vy / vl]]) {
        const lim = P.size * 2.2 * 5, tr = [];
        let cur = 1;
        for (let d = 0.25; d < lim && tr.length < 3; d += 0.25) {
          const v = at(P.x + dx * d, P.y + dy * d);
          if (v !== cur) { tr.push(d); cur = v; }
        }
        if (tr.length === 3) est.push((tr[0] / 1.5 + tr[1] / 2.5 + tr[2] / 3.5) / 3);
      }
    }
    if (est.length < 4) return 0;
    est.sort((a, b) => a - b);
    return est[est.length >> 1];
  }

  function tryTriple(bin, gray, w, h, T) {
    const ms = moduleAlong(bin, w, h, T) || (T.TL.size + T.TR.size + T.BL.size) / 3;
    const dist = (Math.hypot(T.TR.x - T.TL.x, T.TR.y - T.TL.y) + Math.hypot(T.BL.x - T.TL.x, T.BL.y - T.TL.y)) / 2;
    const raw = dist / ms + 7, ver0 = Math.round((raw - 17) / 4);
    const tries = [];
    for (const v of [ver0, ver0 - 1, ver0 + 1]) if (v >= 1 && v <= 10 && tries.indexOf(v) < 0) tries.push(v);
    for (const ver of tries) {
      const dim = ver * 4 + 17, aff = affineMap(T, dim), maps = [aff];
      if (ver >= 2) {
        const al = findAlignment(bin, w, h, aff, dim);
        if (al) {
          const hm = homography([[3.5, 3.5], [dim - 3.5, 3.5], [3.5, dim - 3.5], [dim - 6.5, dim - 6.5]], [[T.TL.x, T.TL.y], [T.TR.x, T.TR.y], [T.BL.x, T.BL.y], al]);
          if (hm) maps.unshift(hm);
        }
      }
      const cands = maps.map(m => { const mod = sampleGray(gray, w, h, m, dim); return { mod, s: timingScore(mod, dim) }; }).sort((a, b) => b.s - a.s);
      for (const c of cands) {
        if (c.s < 0.7) continue;
        const f = readFormat(c.mod, dim);
        if (!f) continue;
        const t = readText(c.mod, dim, ver, f);
        if (t !== null && t !== "") return t;
      }
      /* ultimo tentativo: il quarto punto di riferimento (quadratino o angolo in basso a destra) si sposta a piccoli passi */
      const uv4 = ver >= 2 ? [dim - 6.5, dim - 6.5] : [dim - 3.5, dim - 3.5], c4 = [[3.5, 3.5], [dim - 3.5, 3.5], [3.5, dim - 3.5]];
      const d4 = [[T.TL.x, T.TL.y], [T.TR.x, T.TR.y], [T.BL.x, T.BL.y]];
      for (const du of [0, -1, 1, -2, 2, -3, 3]) for (const dv of [0, -1, 1, -2, 2, -3, 3]) {
        if (!du && !dv) continue;
        const hm = homography([...c4, uv4], [...d4, aff(uv4[0] + du * 0.75, uv4[1] + dv * 0.75)]);
        if (!hm) continue;
        const mod = sampleGray(gray, w, h, hm, dim);
        if (timingScore(mod, dim) < 0.7) continue;
        const f = readFormat(mod, dim);
        if (!f) continue;
        const t = readText(mod, dim, ver, f);
        if (t !== null && t !== "") return t;
      }
    }
    return null;
  }

  function decodeBin(bin, gray, w, h, found, fails) {
    const fin = findFinders(bin, w, h);
    if (fin.length < 3) return;
    for (const T of triples(fin)) {
      const t = tryTriple(bin, gray, w, h, T);
      if (t && found.indexOf(t) < 0) found.push(t);
      else if (!t && fails) fails.push({ T, ms: moduleAlong(bin, w, h, T) || (T.TL.size + T.TR.size + T.BL.size) / 3 });
    }
  }
  function binTries(gray, w, h) {
    const base = Math.min(w, h);
    return [() => binAdaptive(gray, w, h, (Math.round(base / 8) | 1), 0.92), () => binOtsu(gray, w, h), () => binAdaptive(gray, w, h, (Math.round(base / 3) | 1), 0.85), () => binAdaptive(gray, w, h, (Math.round(base / 16) | 1), 0.95)];
  }
  function attempt(gray, w, h, found, fails) {
    for (const mk of binTries(gray, w, h)) {
      try { decodeBin(mk(), gray, w, h, found, fails); } catch (e) { /* un tentativo storto non ferma gli altri */ }
      if (found.length) break;
    }
  }
  /* ritaglia attorno al codice e ingrandisce: con moduli di 2-3 pixel il lavoro riesce molto meglio */
  function zoomOn(gray, w, h, T, ms) {
    const xs = [T.TL.x, T.TR.x, T.BL.x, T.TR.x + T.BL.x - T.TL.x], ys = [T.TL.y, T.TR.y, T.BL.y, T.TR.y + T.BL.y - T.TL.y], m = ms * 8;
    const x0 = Math.max(0, Math.floor(Math.min(...xs) - m)), x1 = Math.min(w, Math.ceil(Math.max(...xs) + m));
    const y0 = Math.max(0, Math.floor(Math.min(...ys) - m)), y1 = Math.min(h, Math.ceil(Math.max(...ys) + m));
    const cw = x1 - x0, ch = y1 - y0;
    if (cw < 20 || ch < 20) return null;
    const f = Math.max(2, Math.min(5, Math.ceil(10 / ms)));
    if (cw * f * ch * f > 6e6) return null;
    const W = cw * f, H = ch * f, out = new Uint8Array(W * H);
    for (let y = 0; y < H; y++) {
      const sy = Math.min(ch - 1.001, (y + 0.5) / f - 0.5), yy = Math.max(0, sy) | 0, fy = Math.max(0, sy) - yy;
      for (let x = 0; x < W; x++) {
        const sx = Math.min(cw - 1.001, (x + 0.5) / f - 0.5), xx = Math.max(0, sx) | 0, fx = Math.max(0, sx) - xx;
        const i = (y0 + yy) * w + x0 + xx;
        out[y * W + x] = (gray[i] * (1 - fx) + gray[i + 1] * fx) * (1 - fy) + (gray[i + w] * (1 - fx) + gray[i + w + 1] * fx) * fy;
      }
    }
    return { gray: out, w: W, h: H };
  }
  function decodeGray(gray, w, h) {
    const found = [], fails = [];
    attempt(gray, w, h, found, fails);
    if (!found.length && fails.length) {
      /* i quadrati d'angolo ci sono ma il testo no: se i moduli sono piccoli ingrandisco attorno al codice e riprovo */
      const seen = [];
      for (const { T, ms } of fails) {
        if (seen.length >= 3) break;
        if (ms >= 7 || seen.some(q => Math.hypot(q.x - T.TL.x, q.y - T.TL.y) < 3 * ms)) continue;
        seen.push(T.TL);
        const z = zoomOn(gray, w, h, T, ms);
        if (z) attempt(z.gray, z.w, z.h, found, null);
        if (found.length) break;
      }
    }
    return found.map(t => ({ text: t, format: "qr", votes: 3, n: t.length }));
  }
  window.QR = { decodeGray, rsCorrect, _d: { binAdaptive, binOtsu, findFinders, triples, tryTriple } };
})();
