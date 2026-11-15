"use strict";
/* Crea PDF ed Excel (.xlsx) direttamente sul telefono, senza librerie esterne. */
(function (root) {
  const WIDTHS = { r: [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584, 761, 556, 556, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 556, 611, 556, 556, 222, 222, 333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 556, 500, 667, 278, 333, 556, 556, 556, 556, 260, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 556, 537, 278, 333, 333, 365, 556, 834, 834, 834, 611, 667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278, 556, 556, 556, 556, 556, 556, 556, 584, 611, 556, 556, 556, 556, 500, 556, 500], b: [278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584, 761, 556, 556, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 556, 611, 556, 556, 278, 278, 500, 500, 350, 556, 1000, 333, 1000, 556, 333, 944, 556, 500, 667, 278, 333, 556, 556, 556, 556, 280, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 611, 556, 278, 333, 333, 365, 556, 834, 834, 834, 611, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278, 611, 611, 611, 611, 611, 611, 611, 584, 611, 611, 611, 611, 611, 556, 611, 556] };
  const SPECIAL = { 0x20AC: 0x80, 0x201A: 0x82, 0x0192: 0x83, 0x201E: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x02C6: 0x88, 0x2030: 0x89, 0x0160: 0x8A, 0x2039: 0x8B, 0x0152: 0x8C, 0x017D: 0x8E, 0x2018: 0x91, 0x2019: 0x92, 0x201C: 0x93, 0x201D: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x02DC: 0x98, 0x2122: 0x99, 0x0161: 0x9A, 0x203A: 0x9B, 0x0153: 0x9C, 0x017E: 0x9E, 0x0178: 0x9F };
  const NAVY = "#002D5C", GREEN = "#2c954b", RED = "#d12233", LIGHT = "#f2f9fd", GREY = "#5b6b7d";

  /* testo -> caratteri a 8 bit (Windows-1252), come vuole il PDF */
  function toWin(str) {
    let out = "";
    for (const ch of String(str == null ? "" : str)) {
      const c = ch.codePointAt(0);
      if (c === 9) { out += " "; continue; }
      if (c < 32 || c === 127) continue;
      if (c < 128 || (c >= 0xA0 && c <= 0xFF)) out += String.fromCharCode(c);
      else if (SPECIAL[c]) out += String.fromCharCode(SPECIAL[c]);
      else out += "?";
    }
    return out;
  }
  function tw(s, bold, size) {
    const t = bold ? WIDTHS.b : WIDTHS.r; let w = 0;
    for (let i = 0; i < s.length; i++) { const v = t[s.charCodeAt(i) - 32]; w += v == null ? 556 : v; }
    return w * size / 1000;
  }
  function wrap(text, maxW, bold, size, maxLines) {
    const lines = [];
    for (const para of String(text == null ? "" : text).replace(/\r/g, "").split("\n")) {
      const s = toWin(para);
      if (!s.trim()) { lines.push(""); continue; }
      let cur = "";
      for (const word of s.split(/ +/)) {
        if (!word) continue;
        const trial = cur ? cur + " " + word : word;
        if (tw(trial, bold, size) <= maxW) { cur = trial; continue; }
        if (cur) { lines.push(cur); cur = ""; }
        let w = word;
        while (tw(w, bold, size) > maxW && w.length > 1) {
          let k = w.length; while (k > 1 && tw(w.slice(0, k), bold, size) > maxW) k--;
          lines.push(w.slice(0, k)); w = w.slice(k);
        }
        cur = w;
      }
      lines.push(cur);
    }
    while (lines.length > 1 && lines[lines.length - 1] === "") lines.pop();
    if (lines.length > maxLines) { lines.length = maxLines; lines[maxLines - 1] = lines[maxLines - 1].replace(/ +$/, "") + "..."; }
    return lines.length ? lines : [""];
  }
  const esc = s => s.replace(/[\\()]/g, m => "\\" + m).replace(/\r/g, "\\r").replace(/\n/g, "\\n");
  function rgb(hex) {
    const n = parseInt(hex.replace("#", ""), 16);
    return [(n >> 16 & 255) / 255, (n >> 8 & 255) / 255, (n & 255) / 255].map(v => v.toFixed(3)).join(" ");
  }
  const cellText = c => (c && typeof c === "object") ? c.v : c;
  const cellColor = c => (c && typeof c === "object" && c.c) ? c.c : null;
  const pdate = () => { const d = new Date(), p = n => String(n).padStart(2, "0"); return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`; };

  /* ---------------------------------------- PDF ---------------------------------------- */
  function makePdf(o) {
    const PW = 842, PH = 595, M = 28, FOOT = 22, bottom = PH - M - FOOT;
    const pages = []; let ops, y;
    const fillRect = (x, top, w, h, hex) => ops.push(`${rgb(hex)} rg ${x.toFixed(2)} ${(PH - top - h).toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f`);
    const hline = (x1, x2, top, hex, lw) => ops.push(`${rgb(hex)} RG ${lw} w ${x1.toFixed(2)} ${(PH - top).toFixed(2)} m ${x2.toFixed(2)} ${(PH - top).toFixed(2)} l S`);
    const text = (x, base, s, bold, size, hex) => ops.push(`${rgb(hex)} rg BT /F${bold ? 2 : 1} ${size} Tf ${x.toFixed(2)} ${(PH - base).toFixed(2)} Td (${esc(s)}) Tj ET`);
    const newPage = () => { ops = []; pages.push(ops); y = M; };

    newPage();
    let hx = M;
    if (o.logo) { const lh = 32, lw = lh * o.logo.w / o.logo.h; ops.push(`q ${lw.toFixed(2)} 0 0 ${lh} ${M} ${(PH - M - lh).toFixed(2)} cm /Im1 Do Q`); hx = M + lw + 14; }
    text(hx, M + 14, toWin(o.title || "Help Desk INAIL"), true, 17, NAVY);
    let sy = M + 28;
    for (const line of [("Generato il " + pdate())].concat(o.info || [])) {
      const ls = wrap(line, PW - M - hx, false, 8.5, 2);
      for (const l of ls) { text(hx, sy, l, false, 8.5, GREY); sy += 11; }
    }
    y = Math.max(sy, M + 40) + 4;
    const tri = (PW - 2 * M) / 3;
    fillRect(M, y, tri, 3, GREEN); fillRect(M + tri, y, tri, 3, "#dfe9f2"); fillRect(M + 2 * tri, y, tri, 3, RED);
    y += 14;

    function drawTable(headers, rows, weights, width) {
      const sum = weights.reduce((a, b) => a + b, 0);
      const cw = weights.map(w => w / sum * width);
      const xs = []; let acc = M; cw.forEach(w => { xs.push(acc); acc += w; });
      const PAD = 3.5, FS = 8, LH = 10, HFS = 8.5;
      const headRow = () => {
        const hl = headers.map((h, i) => wrap(h, cw[i] - 2 * PAD, true, HFS, 3));
        const hh = Math.max(...hl.map(a => a.length)) * 11 + 2 * PAD;
        fillRect(M, y, width, hh, NAVY);
        hl.forEach((ls, i) => ls.forEach((l, k) => text(xs[i] + PAD, y + PAD + 7.2 + k * 11, l, true, HFS, "#ffffff")));
        y += hh;
      };
      if (y + 40 > bottom) newPage();
      headRow();
      rows.forEach((r, ri) => {
        const cl = r.map((c, i) => wrap(cellText(c), cw[i] - 2 * PAD, false, FS, 12));
        const rh = Math.max(...cl.map(a => a.length)) * LH + 2 * PAD;
        if (y + rh > bottom) { newPage(); headRow(); }
        if (ri % 2 === 1) fillRect(M, y, width, rh, LIGHT);
        cl.forEach((ls, i) => { const col = cellColor(r[i]) || "#1c2733"; ls.forEach((l, k) => text(xs[i] + PAD, y + PAD + 6.6 + k * LH, l, !!cellColor(r[i]), FS, col)); });
        hline(M, M + width, y + rh, "#c5d3e0", 0.4);
        y += rh;
      });
      if (!rows.length) { text(M + PAD, y + 12, "Nessun dato.", false, FS, GREY); y += 18; }
    }
    drawTable(o.headers, o.rows, o.widths, PW - 2 * M);

    if ((o.summary || []).length) {
      y += 18; if (y + 60 > bottom) newPage();
      text(M, y + 10, "Riepilogo", true, 13, NAVY); y += 18;
      text(M, y + 8, toWin((o.totalLabel || "Totale") + ": " + o.total), true, 9.5, "#1c2733"); y += 16;
      for (const t of o.summary) {
        if (!t.rows.length) continue;
        if (y + 50 > bottom) newPage();
        text(M, y + 8, toWin(t.title), true, 9.5, NAVY); y += 12;
        const w = Math.min(PW - 2 * M, Math.max(300, t.headers.length * 90));
        drawTable(t.headers, t.rows, t.weights || t.headers.map((h, i) => i === 0 ? 3 : 1), w);
        y += 12;
      }
    }
    pages.forEach((p, i) => { ops = p; text(M, PH - 14, "Help Desk INAIL", false, 7.5, GREY); const s = `Pagina ${i + 1} di ${pages.length}`; text(PW - M - tw(s, false, 7.5), PH - 14, s, false, 7.5, GREY); });

    /* assemblaggio file */
    const objs = []; const add = s => { objs.push(s); return objs.length; };
    add("<</Type/Catalog/Pages 2 0 R>>");
    add(""); /* 2: pagine (riempito dopo) */
    add("<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>");
    add("<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Bold/Encoding/WinAnsiEncoding>>");
    let imgRef = 0;
    if (o.logo) imgRef = add(`<</Type/XObject/Subtype/Image/Width ${o.logo.w}/Height ${o.logo.h}/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/DCTDecode/Length ${o.logo.bin.length}>>\nstream\n${o.logo.bin}\nendstream`);
    const kids = [];
    pages.forEach((p, i) => {
      const body = p.join("\n");
      const cRef = add(`<</Length ${body.length}>>\nstream\n${body}\nendstream`);
      const res = `<</Font<</F1 3 0 R/F2 4 0 R>>${imgRef && i === 0 ? `/XObject<</Im1 ${imgRef} 0 R>>` : ""}>>`;
      kids.push(add(`<</Type/Page/Parent 2 0 R/MediaBox[0 0 ${PW} ${PH}]/Resources${res}/Contents ${cRef} 0 R>>`));
    });
    objs[1] = `<</Type/Pages/Kids[${kids.map(k => k + " 0 R").join(" ")}]/Count ${kids.length}>>`;
    let out = "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n"; const offs = [];
    objs.forEach((b, i) => { offs.push(out.length); out += `${i + 1} 0 obj\n${b}\nendobj\n`; });
    const xr = out.length;
    out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map(v => String(v).padStart(10, "0") + " 00000 n \n").join("");
    out += `trailer\n<</Size ${objs.length + 1}/Root 1 0 R>>\nstartxref\n${xr}\n%%EOF\n`;
    const u = new Uint8Array(out.length); for (let i = 0; i < out.length; i++) u[i] = out.charCodeAt(i) & 255;
    return u;
  }

  /* ------------------------- verbale in PDF (A4 verticale) ------------------------- */
  /* o.blocks: title, p, h, kv, checks, list, table, sigs {left, slots:[{label,img}]}; o.logo; o.photos: [img] (img = {bin,w,h} JPEG) */
  function makeDocPdf(o) {
    const PW = 595, PH = 842, M = 44, FOOT = 28, bottom = PH - M - FOOT, W = PW - 2 * M;
    const pages = [], imgs = []; let ops, y;
    const imgNo = im => { let i = imgs.indexOf(im); if (i < 0) { imgs.push(im); i = imgs.length - 1; } return i + 1; };
    const fillRect = (x, top, w, h, hex) => ops.push(`${rgb(hex)} rg ${x.toFixed(2)} ${(PH - top - h).toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f`);
    const hline = (x1, x2, top, hex, lw) => ops.push(`${rgb(hex)} RG ${lw} w ${x1.toFixed(2)} ${(PH - top).toFixed(2)} m ${x2.toFixed(2)} ${(PH - top).toFixed(2)} l S`);
    const text = (x, base, s, bold, size, hex) => ops.push(`${rgb(hex)} rg BT /F${bold ? 2 : 1} ${size} Tf ${x.toFixed(2)} ${(PH - base).toFixed(2)} Td (${esc(s)}) Tj ET`);
    const picture = (im, x, top, w, h) => ops.push(`q ${w.toFixed(2)} 0 0 ${h.toFixed(2)} ${x.toFixed(2)} ${(PH - top - h).toFixed(2)} cm /Im${imgNo(im)} Do Q`);
    const newPage = () => { ops = []; pages.push(ops); y = M; };
    const need = h => { if (y + h > bottom) newPage(); };
    const INK = "#1c2733";

    newPage();
    if (o.logo) { const lh = 30, lw = lh * o.logo.w / o.logo.h; picture(o.logo, M, M, lw, lh); y = M + lh + 8; } else y = M + 4;
    const tri = W / 3;
    fillRect(M, y, tri, 3, GREEN); fillRect(M + tri, y, tri, 3, "#dfe9f2"); fillRect(M + 2 * tri, y, tri, 3, RED);
    y += 18;

    function table(b) {
      const sum = b.weights.reduce((a, c) => a + c, 0), cw = b.weights.map(w => w / sum * W), xs = []; let acc = M;
      cw.forEach(w => { xs.push(acc); acc += w; });
      const PAD = 4, FS = 9, LH = 12;
      const head = () => {
        need(26);
        fillRect(M, y, W, 18, NAVY);
        b.head.forEach((h, i) => text(xs[i] + PAD, y + 12.5, wrap(h, cw[i] - 2 * PAD, true, FS, 1)[0], true, FS, "#ffffff"));
        y += 18;
      };
      head();
      b.rows.forEach((r, ri) => {
        const cl = r.map((c, i) => wrap(c, cw[i] - 2 * PAD, false, FS, 4)), rh = Math.max(...cl.map(a => a.length)) * LH + 2 * PAD - 2;
        if (y + rh > bottom) { newPage(); head(); }
        if (ri % 2 === 1) fillRect(M, y, W, rh, LIGHT);
        cl.forEach((ls, i) => ls.forEach((l, k) => text(xs[i] + PAD, y + PAD + 8 + k * LH - 1, l, false, FS, INK)));
        hline(M, M + W, y + rh, "#c5d3e0", 0.4);
        y += rh;
      });
      y += 8;
    }
    function sigs(b) {
      const slots = b.slots || [], n = Math.max(1, slots.length), H = 96;
      need(H + 14);
      if (b.left) { text(M, y + 10, toWin(b.left), true, 10, INK); }
      y += 16;
      const sw = W / n;
      slots.forEach((s, i) => {
        const cx = M + sw * i + sw / 2, lw = Math.min(sw - 30, 190);
        text(cx - tw(wrap(s.label, 400, true, 9, 1)[0], true, 9) / 2, y + 8, wrap(s.label, 400, true, 9, 1)[0], true, 9, INK);
        if (s.img) {
          const maxW = Math.min(lw, 150), maxH = 46, k = Math.min(maxW / s.img.w, maxH / s.img.h), w = s.img.w * k, h = s.img.h * k;
          picture(s.img, cx - w / 2, y + 62 - h, w, h);
        }
        hline(cx - lw / 2, cx + lw / 2, y + 64, INK, 0.6);
      });
      y += H - 16 + 6;
    }

    for (const b of o.blocks || []) {
      if (b.t === "title") { const ls = wrap(b.s, W, true, 13, 4); need(ls.length * 17 + 14); ls.forEach(l => { text(M + (W - tw(l, true, 13)) / 2, y + 13, l, true, 13, NAVY); y += 17; }); y += 8; }
      else if (b.t === "h") { need(36); y += 8; text(M, y + 10, b.s, true, 10, NAVY); y += 15; }
      else if (b.t === "p") { const ls = wrap(b.s, W, !!b.bold, 10, 80); ls.forEach(l => { need(14); text(M, y + 10, l, !!b.bold, 10, INK); y += 13.5; }); y += 5; }
      else if (b.t === "kv") { const kk = toWin(b.k + ": "), kw = tw(kk, true, 10), ls = wrap(b.v, W - kw, false, 10, 10); need(14); text(M, y + 10, kk, true, 10, INK); ls.forEach((l, i) => { if (i) need(14); text(M + kw, y + 10, l, false, 10, INK); y += 13.5; }); }
      else if (b.t === "checks") { for (const it of b.items) { const ls = wrap((it.on ? "[X]  " : "[   ]  ") + it.s, W - 10, it.on, 10, 3); ls.forEach(l => { need(14); text(M + 6, y + 10, l, it.on, 10, INK); y += 13.5; }); } y += 3; }
      else if (b.t === "list") { for (const it of b.items) { const ls = wrap(it, W - 22, false, 10, 5); ls.forEach((l, i) => { need(14); if (!i) text(M + 8, y + 10, "-", false, 10, INK); text(M + 20, y + 10, l, false, 10, INK); y += 13.5; }); } y += 3; }
      else if (b.t === "table") table(b);
      else if (b.t === "sigs") sigs(b);
      else if (b.t === "gap") y += b.h || 8;
    }

    const ph = o.photos || [];
    if (ph.length) {
      newPage();
      const t = "ALLEGATO FOTOGRAFICO"; text(M + (W - tw(t, true, 12)) / 2, y + 12, t, true, 12, NAVY); y += 28;
      ph.forEach((im, i) => {
        const k = Math.min(W / im.w, 330 / im.h), w = im.w * k, h = im.h * k;
        need(h + 24);
        picture(im, M + (W - w) / 2, y, w, h); y += h + 12;
        const cap = `Foto ${i + 1} di ${ph.length}`; text(M + (W - tw(cap, false, 8.5)) / 2, y, cap, false, 8.5, GREY); y += 18;
      });
    }
    const stamp = toWin(o.footer || ("Copia generata dal telefono il " + pdate()));
    pages.forEach((p, i) => { ops = p; text(M, PH - 22, stamp, false, 7.5, GREY); const s = `Pagina ${i + 1} di ${pages.length}`; text(PW - M - tw(s, false, 7.5), PH - 22, s, false, 7.5, GREY); });

    const objs = []; const add = s => { objs.push(s); return objs.length; };
    add("<</Type/Catalog/Pages 2 0 R>>");
    add("");
    add("<</Type/Font/Subtype/Type1/BaseFont/Helvetica/Encoding/WinAnsiEncoding>>");
    add("<</Type/Font/Subtype/Type1/BaseFont/Helvetica-Bold/Encoding/WinAnsiEncoding>>");
    const refs = imgs.map(im => add(`<</Type/XObject/Subtype/Image/Width ${im.w}/Height ${im.h}/ColorSpace/DeviceRGB/BitsPerComponent 8/Filter/DCTDecode/Length ${im.bin.length}>>\nstream\n${im.bin}\nendstream`));
    const xo = refs.length ? `/XObject<<${refs.map((r, i) => `/Im${i + 1} ${r} 0 R`).join("")}>>` : "";
    const kids = [];
    pages.forEach(p => {
      const body = p.join("\n"), cRef = add(`<</Length ${body.length}>>\nstream\n${body}\nendstream`);
      kids.push(add(`<</Type/Page/Parent 2 0 R/MediaBox[0 0 ${PW} ${PH}]/Resources<</Font<</F1 3 0 R/F2 4 0 R>>${xo}>>/Contents ${cRef} 0 R>>`));
    });
    objs[1] = `<</Type/Pages/Kids[${kids.map(k => k + " 0 R").join(" ")}]/Count ${kids.length}>>`;
    let out = "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n"; const offs = [];
    objs.forEach((b, i) => { offs.push(out.length); out += `${i + 1} 0 obj\n${b}\nendobj\n`; });
    const xr = out.length;
    out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map(v => String(v).padStart(10, "0") + " 00000 n \n").join("");
    out += `trailer\n<</Size ${objs.length + 1}/Root 1 0 R>>\nstartxref\n${xr}\n%%EOF\n`;
    const u = new Uint8Array(out.length); for (let i = 0; i < out.length; i++) u[i] = out.charCodeAt(i) & 255;
    return u;
  }

  /* immagine (foto, firma) -> JPEG pronto per il PDF, su fondo bianco */
  async function toJpegImg(src, max, q) {
    const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error("immagine non leggibile")); i.src = src; });
    const k = Math.min(1, (max || 1400) / Math.max(img.naturalWidth, img.naturalHeight));
    const c = document.createElement("canvas"); c.width = Math.max(1, Math.round(img.naturalWidth * k)); c.height = Math.max(1, Math.round(img.naturalHeight * k));
    const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0, c.width, c.height);
    return { bin: atob(c.toDataURL("image/jpeg", q || 0.85).split(",")[1]), w: c.width, h: c.height };
  }

  async function loadLogo() {
    try {
      const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = "logo_header.png"; });
      const c = document.createElement("canvas"); c.width = img.naturalWidth; c.height = img.naturalHeight;
      const g = c.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, c.width, c.height); g.drawImage(img, 0, 0);
      return { bin: atob(c.toDataURL("image/jpeg", 0.92).split(",")[1]), w: c.width, h: c.height };
    } catch (e) { return null; }
  }

  /* ------------------------------------- Excel (.xlsx) ------------------------------------- */
  const crcT = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u) { let c = 0xFFFFFFFF; for (let i = 0; i < u.length; i++) c = crcT[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  const concat = parts => { const n = parts.reduce((a, b) => a + b.length, 0), o = new Uint8Array(n); let p = 0; parts.forEach(b => { o.set(b, p); p += b.length; }); return o; };
  function zip(files) {
    const enc = new TextEncoder(), parts = [], cen = []; let off = 0;
    const d = new Date(), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    for (const f of files) {
      const nm = enc.encode(f.name), crc = crc32(f.data), sz = f.data.length;
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true); lh.setUint16(10, time, true); lh.setUint16(12, date, true);
      lh.setUint32(14, crc, true); lh.setUint32(18, sz, true); lh.setUint32(22, sz, true); lh.setUint16(26, nm.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), nm, f.data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true); ch.setUint16(12, time, true); ch.setUint16(14, date, true);
      ch.setUint32(16, crc, true); ch.setUint32(20, sz, true); ch.setUint32(24, sz, true); ch.setUint16(28, nm.length, true); ch.setUint32(42, off, true);
      cen.push(new Uint8Array(ch.buffer), nm);
      off += 30 + nm.length + sz;
    }
    const cenSize = cen.reduce((a, b) => a + b.length, 0), end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true); end.setUint32(12, cenSize, true); end.setUint32(16, off, true);
    return concat([...parts, ...cen, new Uint8Array(end.buffer)]);
  }
  const xe = s => String(s == null ? "" : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const colName = i => { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  /* stili: 0 normale, 1 intestazione, 2 testo a capo, 3 titolo, 4 grassetto */
  function sheetXml(sh) {
    const rows = [];
    if (sh.headers) rows.push(sh.headers.map(h => ({ v: h, s: 1 })));
    (sh.rows || []).forEach(r => rows.push(r));
    const sd = rows.map((r, ri) => {
      const cells = r.map((c, ci) => {
        const isObj = c && typeof c === "object", v = isObj ? c.v : c, s = isObj && c.s != null ? c.s : (sh.headers && ri > 0 ? 2 : 0);
        const ref = colName(ci) + (ri + 1);
        if (v === "" || v == null) return s ? `<c r="${ref}" s="${s}"/>` : "";
        if (typeof v === "number") return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
        return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xe(v)}</t></is></c>`;
      }).join("");
      return `<row r="${ri + 1}">${cells}</row>`;
    }).join("");
    const cols = (sh.widths || []).map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("");
    const pane = sh.headers ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : `<sheetViews><sheetView workbookViewId="0"/></sheetViews>`;
    const filter = sh.headers && rows.length > 1 ? `<autoFilter ref="A1:${colName(sh.headers.length - 1)}${rows.length}"/>` : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${pane}<sheetFormatPr defaultRowHeight="15"/>${cols ? `<cols>${cols}</cols>` : ""}<sheetData>${sd}</sheetData>${filter}</worksheet>`;
  }
  function makeXlsx(sheets) {
    const enc = new TextEncoder(), f = (name, xml) => ({ name, data: enc.encode(xml) });
    const names = sheets.map((s, i) => xe(String(s.name || "Foglio" + (i + 1)).replace(/[\[\]:*?\/\\]/g, " ").slice(0, 31)));
    const files = [
      f("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`),
      f("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      f("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n, i) => `<sheet name="${n}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`),
      f("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`),
      f("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="14"/><color rgb="FF002D5C"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF002D5C"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`),
      ...sheets.map((s, i) => f(`xl/worksheets/sheet${i + 1}.xml`, sheetXml(s)))
    ];
    return zip(files);
  }

  root.HDExport = { makePdf, makeXlsx, loadLogo, makeDocPdf, toJpegImg };
})(window);
