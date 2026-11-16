/* Mappa della Toscana: dove sono le sedi e quanti interventi hanno. Funziona senza campo (tutto è qui dentro).
   Si apre da Registro › Mappa. Il contorno è SEMPLIFICATO (disegnato da punti di riferimento, non è un confine
   ufficiale); le sedi si riconoscono dal nome (Firenze, «Sede di Pisa», DR Toscana...). Quelle che non si
   riconoscono si posizionano toccando la mappa: la posizione resta su questo telefono. */
"use strict";
window.TOSCANA = (function () {
  /*DATI-INIZIO*/
  const D = {
    "K": 0.72537, "S": 492.0, "LON0": 9.65, "LAT1": 44.55, "W": 1000, "H": 1181,
    "outline": [[10.03,44.055], [10.045,44.03], [10.1,43.99], [10.15,43.97], [10.18,43.95], [10.2,43.92], [10.23,43.87], [10.26,43.82],
      [10.28,43.75], [10.27,43.68], [10.29,43.62], [10.31,43.57], [10.3,43.54], [10.33,43.49], [10.37,43.43], [10.4,43.4], [10.44,43.35],
      [10.47,43.31], [10.5,43.25], [10.53,43.18], [10.54,43.12], [10.55,43.04], [10.5,42.99], [10.49,42.96], [10.52,42.93], [10.56,42.92],
      [10.62,42.97], [10.7,42.95], [10.76,42.92], [10.78,42.89], [10.77,42.84], [10.74,42.8], [10.8,42.78], [10.87,42.77], [10.93,42.72],
      [10.97,42.69], [11.03,42.64], [11.07,42.6], [11.1,42.57], [11.13,42.55], [11.17,42.53], [11.19,42.5], [11.13,42.48], [11.1,42.45],
      [11.08,42.42], [11.1,42.39], [11.17,42.375], [11.22,42.385], [11.25,42.4], [11.29,42.415], [11.38,42.385], [11.47,42.355],
      [11.52,42.39], [11.57,42.43], [11.62,42.5], [11.69,42.56], [11.74,42.62], [11.78,42.7], [11.82,42.77], [11.88,42.78], [11.95,42.82],
      [11.97,42.88], [12.03,42.93], [12.03,42.97], [12.06,43.02], [12.1,43.07], [12.05,43.12], [12.03,43.2], [12.05,43.26], [12.1,43.31],
      [12.12,43.36], [12.13,43.42], [12.12,43.47], [12.15,43.53], [12.17,43.58], [12.2,43.62], [12.25,43.65], [12.33,43.65], [12.38,43.7],
      [12.37,43.76], [12.3,43.8], [12.22,43.83], [12.15,43.84], [12.05,43.8], [11.98,43.82], [11.9,43.83], [11.83,43.86], [11.78,43.87],
      [11.7,43.9], [11.68,43.94], [11.62,43.98], [11.66,44.04], [11.7,44.09], [11.66,44.14], [11.57,44.16], [11.5,44.17], [11.46,44.2],
      [11.38,44.22], [11.3,44.2], [11.24,44.17], [11.17,44.14], [11.1,44.12], [11.0,44.08], [10.9,44.15], [10.82,44.17], [10.75,44.15],
      [10.68,44.17], [10.62,44.15], [10.55,44.16], [10.5,44.2], [10.45,44.25], [10.35,44.26], [10.28,44.3], [10.22,44.31], [10.15,44.33],
      [10.12,44.36], [10.05,44.4], [10.0,44.43], [9.93,44.46], [9.86,44.43], [9.8,44.4], [9.78,44.35], [9.82,44.3], [9.88,44.25],
      [9.93,44.2], [9.97,44.16], [9.99,44.13], [10.0,44.09]],
    "islands": {
      "Elba": [[10.11,42.8],[10.17,42.82],[10.26,42.83],[10.34,42.83],[10.4,42.86],[10.45,42.84],[10.45,42.81],[10.42,42.78],[10.41,42.73],[10.34,42.73],[10.27,42.75],[10.19,42.74],[10.12,42.76]],
      "Capraia": [[9.865,43.05],[9.862,43.076],[9.852,43.097],[9.838,43.108],[9.822,43.108],[9.808,43.097],[9.798,43.076],[9.795,43.05],[9.798,43.024],[9.808,43.003],[9.822,42.992],[9.838,42.992],[9.852,43.003],[9.862,43.024]],
      "Gorgona": [[9.914,43.43],[9.91,43.441],[9.9,43.446],[9.89,43.441],[9.886,43.43],[9.89,43.419],[9.9,43.414],[9.91,43.419]],
      "Pianosa": [[10.1,42.585],[10.094,42.595],[10.08,42.599],[10.066,42.595],[10.06,42.585],[10.066,42.575],[10.08,42.571],[10.094,42.575]],
      "Montecristo": [[10.33,42.335],[10.324,42.349],[10.31,42.355],[10.296,42.349],[10.29,42.335],[10.296,42.321],[10.31,42.315],[10.324,42.321]],
      "Giglio": [[10.95,42.36],[10.947,42.382],[10.937,42.399],[10.923,42.409],[10.907,42.409],[10.893,42.399],[10.883,42.382],[10.88,42.36],[10.883,42.338],[10.893,42.321],[10.907,42.311],[10.923,42.311],[10.937,42.321],[10.947,42.338]],
      "Giannutri": [[11.125,42.255],[11.118,42.262],[11.1,42.265],[11.082,42.262],[11.075,42.255],[11.082,42.248],[11.1,42.245],[11.118,42.248]]},
    "comuni": {"Firenze": [43.7696,11.2558,1], "Pisa": [43.7228,10.4017,1], "Siena": [43.3188,11.3308,1], "Arezzo": [43.4633,11.8797,1],
      "Livorno": [43.5485,10.3106,1], "Lucca": [43.843,10.5027,1], "Grosseto": [42.7635,11.1124,1], "Massa": [44.0354,10.1393,1],
      "Carrara": [44.0793,10.0977,0], "Pistoia": [43.9331,10.9175,1], "Prato": [43.8777,11.1022,1], "Empoli": [43.7192,10.9456,0],
      "Viareggio": [43.8681,10.2511,0], "Piombino": [42.9254,10.5266,0], "Cecina": [43.3085,10.5169,0], "Pontedera": [43.6619,10.6313,0],
      "Montecatini Terme": [43.8836,10.7727,0], "Sansepolcro": [43.571,12.14,0], "Cortona": [43.2756,11.9855,0],
      "Montepulciano": [43.0988,11.7826,0], "Poggibonsi": [43.4693,11.1472,0], "Colle di Val d'Elsa": [43.4228,11.1197,0],
      "San Gimignano": [43.4677,11.0432,0], "Volterra": [43.4002,10.8605,0], "Follonica": [42.9233,10.7606,0],
      "Orbetello": [42.4423,11.214,0], "Porto Santo Stefano": [42.4372,11.118,0], "Portoferraio": [42.8143,10.3263,0],
      "Massa Marittima": [43.0505,10.8933,0], "Castiglione della Pescaia": [42.7647,10.8808,0], "Aulla": [44.2108,9.9689,0],
      "Pontremoli": [44.3764,9.8806,0], "Fivizzano": [44.2373,10.1288,0], "Castelnuovo di Garfagnana": [44.11,10.4097,0],
      "Barga": [44.0772,10.4778,0], "Borgo a Mozzano": [43.9993,10.5421,0], "Pescia": [43.9022,10.689,0],
      "Pietrasanta": [43.9588,10.2282,0], "Forte dei Marmi": [43.9626,10.1711,0], "Camaiore": [43.9378,10.3052,0],
      "San Miniato": [43.6823,10.8505,0], "Cascina": [43.6767,10.5547,0], "Pontassieve": [43.7667,11.4333,0],
      "Sesto Fiorentino": [43.8326,11.1999,0], "Scandicci": [43.7547,11.1856,0], "Bagno a Ripoli": [43.7428,11.3036,0],
      "Figline Valdarno": [43.6175,11.4672,0], "Borgo San Lorenzo": [43.9548,11.3872,0], "Barberino di Mugello": [44.0,11.2167,0],
      "Firenzuola": [44.1228,11.3789,0], "Montevarchi": [43.5236,11.5681,0], "San Giovanni Valdarno": [43.5708,11.5276,0],
      "Bibbiena": [43.697,11.816,0], "Poppi": [43.7194,11.7672,0], "Sinalunga": [43.2088,11.7347,0], "Chiusi": [43.0172,11.9494,0],
      "Abbadia San Salvatore": [42.8793,11.6786,0], "Radicofani": [42.8999,11.7679,0], "Pitigliano": [42.6336,11.6683,0],
      "Manciano": [42.587,11.5167,0], "Capalbio": [42.4167,11.3667,0], "Scansano": [42.6819,11.3253,0], "Roccastrada": [42.9988,11.1608,0],
      "Rosignano Marittimo": [43.4044,10.4752,0], "Castagneto Carducci": [43.1633,10.6053,0], "San Vincenzo": [43.1,10.539,0],
      "Campiglia Marittima": [43.0636,10.6097,0], "Suvereto": [43.0742,10.6753,0], "Capraia Isola": [43.05,9.8333,0],
      "Marciana Marina": [42.8,10.1964,0], "Porto Azzurro": [42.764,10.3965,0], "Rio Marina": [42.8167,10.4283,0],
      "Capoliveri": [42.7466,10.3797,0], "Isola del Giglio": [42.3647,10.9156,0], "Ponsacco": [43.6264,10.6203,0],
      "Altopascio": [43.8142,10.6717,0], "Capannori": [43.8419,10.5698,0], "Fucecchio": [43.7289,10.8074,0],
      "Castelfiorentino": [43.6073,10.9646,0], "Certaldo": [43.547,11.04,0], "Montelupo Fiorentino": [43.7333,11.0167,0],
      "Greve in Chianti": [43.5855,11.3162,0], "Impruneta": [43.6972,11.2528,0], "Fiesole": [43.8056,11.2938,0],
      "Reggello": [43.6864,11.5336,0], "Rufina": [43.8131,11.48,0], "Vicchio": [43.9179,11.4504,0], "Marradi": [44.0731,11.6228,0],
      "Abetone": [44.1433,10.6553,0], "San Marcello Piteglio": [44.0554,10.7959,0], "Quarrata": [43.85,10.98,0],
      "Monsummano Terme": [43.87,10.82,0], "Vernio": [44.0633,11.1511,0], "Montemurlo": [43.9199,11.0153,0],
      "Campi Bisenzio": [43.8167,11.1333,0], "Piancastagnaio": [42.845,11.702,0], "Castel del Piano": [42.8944,11.5333,0],
      "Arcidosso": [42.85,11.5333,0], "Sarteano": [43.0,11.8667,0], "Cetona": [43.0,11.9,0], "Montalcino": [43.0575,11.49,0],
      "Pienza": [43.0766,11.6786,0], "Chianciano Terme": [43.0617,11.8219,0], "Monte San Savino": [43.3333,11.7333,0],
      "Castiglion Fiorentino": [43.3333,11.9167,0], "Anghiari": [43.54,12.06,0], "Pieve Santo Stefano": [43.6667,12.0333,0],
      "Stia": [43.8,11.7,0], "Sorano": [42.6833,11.7167,0], "Gavorrano": [42.9167,10.9,0], "Scarlino": [42.9167,10.8667,0],
      "Pomarance": [43.2967,10.8718,0], "Peccioli": [43.55,10.7167,0], "Santa Croce sull'Arno": [43.7167,10.7667,0],
      "San Giuliano Terme": [43.75,10.4333,0], "Massarosa": [43.8667,10.3333,0], "Seravezza": [43.9961,10.2292,0],
      "Bagni di Lucca": [44.0167,10.5833,0], "Camporgiano": [44.15,10.3833,0], "Villafranca in Lunigiana": [44.3,9.95,0],
      "Licciana Nardi": [44.2,10.0,0], "Albinia": [42.5,11.19,0], "Talamone": [42.55,11.13,0], "Cavo": [42.85,10.4167,0]},
    "alias": {"dr toscana": "Firenze", "direzione regionale": "Firenze", "direzione regionale toscana": "Firenze",
      "massa carrara": "Massa", "massa-carrara": "Massa", "mc": "Massa", "elba": "Portoferraio", "isola d'elba": "Portoferraio",
      "valdarno": "Montevarchi", "mugello": "Borgo San Lorenzo", "versilia": "Viareggio", "lunigiana": "Aulla",
      "garfagnana": "Castelnuovo di Garfagnana", "val di cornia": "Piombino", "argentario": "Porto Santo Stefano",
      "chianti": "Greve in Chianti", "valdichiana": "Sinalunga", "casentino": "Poppi", "amiata": "Abbadia San Salvatore"}
  };
  /*DATI-FINE*/
  const K = D.K, S = D.S, W = D.W, H = D.H;
  const PERIODI = [["tutto", "Sempre"], ["anno", "Quest'anno"], ["mese", "Questo mese"], ["30g", "Ultimi 30 giorni"]];
  const ci = (a, b) => { a = String(a).toLowerCase(); b = String(b).toLowerCase(); return a < b ? -1 : a > b ? 1 : 0; };
  const p2 = n => String(n).padStart(2, "0");
  const isoDay = d => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;

  /* ---------- geografia ---------- */
  const proj = (lon, lat) => [(lon - D.LON0) * K * S, (D.LAT1 - lat) * S];
  const unproj = (x, y) => [D.LAT1 - y / S, x / (K * S) + D.LON0];          // [latitudine, longitudine]
  const norm = s => String(s == null ? "" : s).normalize("NFKD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[’`´]/g, "'").replace(/[^a-z0-9']+/g, " ").trim();
  let IDX = null;
  function index() {
    if (!IDX) {
      const idx = new Map();
      Object.keys(D.comuni).forEach(n => idx.set(norm(n), n));
      Object.keys(D.alias).forEach(k => { const nk = norm(k); if (!idx.has(nk)) idx.set(nk, D.alias[k]); });
      IDX = { idx, keys: [...idx.keys()].sort((a, b) => b.length - a.length || (a < b ? -1 : a > b ? 1 : 0)) };
    }
    return IDX;
  }
  /* il comune di cui parla il nome della sede */
  function find(name) {
    const n = norm(name);
    if (!n) return null;
    const { idx, keys } = index();
    if (idx.has(n)) return idx.get(n);
    const pad = " " + n + " ";
    for (const k of keys) if (pad.includes(" " + k + " ")) return idx.get(k);
    return null;
  }
  function position(name, ov) {
    const k = norm(name);
    if (!k) return null;
    if (ov && ov[k]) return [ov[k][0], ov[k][1], "manuale"];
    const c = find(name);
    return c ? [D.comuni[c][0], D.comuni[c][1], c] : null;
  }
  function cleanPos(raw) {
    const out = {};
    if (raw && typeof raw === "object" && !Array.isArray(raw)) for (const k of Object.keys(raw)) {
      const v = raw[k], lat = Array.isArray(v) ? parseFloat(v[0]) : NaN, lon = Array.isArray(v) ? parseFloat(v[1]) : NaN;
      if (lat > 41.5 && lat < 45 && lon > 9 && lon < 13) out[k] = [Math.round(lat * 1e4) / 1e4, Math.round(lon * 1e4) / 1e4];
    }
    return out;
  }
  function periodFrom(kind, now) {
    const d = now || new Date();
    if (kind === "anno") return `${d.getFullYear()}-01-01`;
    if (kind === "mese") return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-01`;
    if (kind === "30g") { const s = new Date(d); s.setDate(s.getDate() - 29); return isoDay(s); }
    return "";
  }

  /* ---------- conteggi per sede ---------- */
  async function records(kind, now) {
    const dal = periodFrom(kind, now), recs = new Map(), senza = { interventi: 0, da_seguire: 0 };
    const rec = (raw, peso) => {
      raw = String(raw || "").trim();
      const k = norm(raw);
      if (!k) return null;
      let r = recs.get(k);
      if (!r) recs.set(k, r = { key: k, raws: new Map(), interventi: 0, da_seguire: 0, utenti: 0, stampanti: 0 });
      r.raws.set(raw, (r.raws.get(raw) || 0) + (peso == null ? 1 : peso));
      return r;
    };
    (typeof SEDI !== "undefined" ? SEDI : []).forEach(s => rec(s, 0.5));
    for (const i of await getAll("interventi")) {
      const r = rec(i.sede) || senza;
      if (!dal || String(i.data_ora || "").slice(0, 10) >= dal) r.interventi++;
      if (stOf(i) !== "Risolto") r.da_seguire++;
    }
    for (const u of await getAll("utenti")) { const r = rec(u.sede); if (r) r.utenti++; }
    const ps = await getAll("stampanti");
    for (const e of await getAll("edifici")) { const r = rec(e.sede); if (r) r.stampanti += ps.filter(s => s.edificio_uid === e.uid).length; }
    const ov = cleanPos(await getMeta("mappa_pos"));
    const sedi = [...recs.values()].map(r => {
      const raws = [...r.raws.entries()].sort((a, b) => b[1] - a[1] || (a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0)).map(x => x[0]);
      const pos = position(raws[0], ov);
      return Object.assign(r, { raws, nome: raws[0], lat: pos ? pos[0] : null, lon: pos ? pos[1] : null, via: pos ? pos[2] : "" });
    });
    sedi.sort((a, b) => b.interventi - a.interventi || b.da_seguire - a.da_seguire || ci(a.nome, b.nome));
    return { sedi, senza_sede: senza, overrides: ov };
  }

  /* ---------- pallini e nomi ---------- */
  const pinR = n => Math.min(46, 18 + 3.5 * Math.sqrt(n));
  /* allontana i pallini che si sovrapporrebbero (Carrara e Massa sono a pochi chilometri) */
  function separate(pins, gap, rounds) {
    gap = gap == null ? 2 : gap; rounds = rounds == null ? 60 : rounds;
    for (let t = 0; t < rounds; t++) {
      let moved = false;
      for (let i = 0; i < pins.length; i++) for (let j = i + 1; j < pins.length; j++) {
        const a = pins[i], b = pins[j];
        let dx = b.x - a.x, dy = b.y - a.y, d = Math.hypot(dx, dy);
        const need = a.r + b.r + gap;
        if (d < need - 0.01) {
          if (d < 1e-6) { dx = 1; dy = 0; d = 1; }
          const push = (need - d) / 2, ux = dx / d, uy = dy / d;
          a.x -= ux * push; a.y -= uy * push; b.x += ux * push; b.y += uy * push;
          moved = true;
        }
      }
      if (!moved) break;
    }
  }
  /* le sedi con posizione, unite quando cadono nello stesso punto (Firenze e DR Toscana) */
  function makePins(recs) {
    const groups = new Map();
    for (const r of recs) if (r.lat != null) {
      const k = r.lat.toFixed(4) + "," + r.lon.toFixed(4);
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(r);
    }
    const pins = [];
    for (const rs of groups.values()) {
      rs.sort((a, b) => b.interventi - a.interventi || ci(a.nome, b.nome));
      const n = rs.reduce((s, r) => s + r.interventi, 0), [x, y] = proj(rs[0].lon, rs[0].lat);
      pins.push({ x, y, tx: x, ty: y, r: pinR(n), n, rosso: rs.reduce((s, r) => s + r.da_seguire, 0), label: rs.map(r => r.nome).join(" + "), sedi: rs.map(r => r.key) });
    }
    pins.sort((a, b) => b.n - a.n || ci(a.label, b.label));
    separate(pins);
    return pins;
  }
  /* i nomi sotto, sopra, a destra o a sinistra dei pallini, senza sovrapporsi */
  function placeLabels(pins, fs) {
    fs = fs || 28;
    const over = (a, b) => a[0] < b[2] && b[0] < a[2] && a[1] < b[3] && b[1] < a[3];
    const circles = pins.map(p => [p.x - p.r, p.y - p.r, p.x + p.r, p.y + p.r]), placed = [], out = [];
    pins.forEach((p, i) => {
      const tw = p.label.length * fs * 0.56;
      const cands = [["middle", p.x, p.y + p.r + fs * 0.95], ["middle", p.x, p.y - p.r - fs * 0.3],
                     ["start", p.x + p.r + 6, p.y + fs * 0.35], ["end", p.x - p.r - 6, p.y + fs * 0.35]];
      let first = null, chosen = null;
      for (const [anchor, tx0, ty] of cands) {
        let tx = tx0, x0 = anchor === "middle" ? tx - tw / 2 : anchor === "start" ? tx : tx - tw, shift = 0;
        if (x0 < 2) shift = 2 - x0; else if (x0 + tw > W - 2) shift = (W - 2) - (x0 + tw);
        x0 += shift; tx += shift;
        const box = [x0, ty - fs * 0.8, x0 + tw, ty + fs * 0.25], cand = { anchor, lx: tx, ly: ty, box };
        if (!first) first = cand;
        if (placed.some(b => over(box, b)) || circles.some((c, j) => j !== i && over(box, c))) continue;
        chosen = cand; break;
      }
      chosen = chosen || first;
      placed.push(chosen.box); out.push(chosen);
    });
    return out;
  }
  /* il pallino più vicino al punto toccato, se è abbastanza vicino */
  function pinAt(pins, x, y, margin) {
    margin = margin == null ? 14 : margin;
    let best = null, bd = null;
    pins.forEach((p, i) => { const d = Math.hypot(x - p.x, y - p.y); if (d <= p.r + margin && (bd === null || d < bd)) { best = i; bd = d; } });
    return best;
  }

  /* ---------- disegno ---------- */
  const pathOf = pts => "M" + pts.map(([lon, lat]) => proj(lon, lat).map(v => v.toFixed(1)).join(" ")).join("L") + "Z";
  const LAND = pathOf(D.outline) + Object.values(D.islands).map(pathOf).join("");
  let per = "tutto", zoom = 1, view = "l", pins = [], recs = [], selKeys = [], placing = null, last = null, busy = 0;
  const FS = 28, MARGIN = 26;
  const num = n => n === 1 ? "1 intervento" : `${n} interventi`;

  function svgOf() {
    const labs = placeLabels(pins, FS);
    let s = `<svg id="tzS" viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Mappa della Toscana con le sedi" style="display:block;width:100%;height:auto;${placing ? "cursor:crosshair" : ""}"><rect width="${W}" height="${H}" fill="#dfeaf5"/><path d="${LAND}" fill="#f4efe0" stroke="#7a8a6a" stroke-width="1.6" stroke-linejoin="round"/>`;
    pins.forEach((p, i) => {
      const sel = p.sedi.some(k => selKeys.includes(k)), fz = Math.round(Math.max(22, Math.min(32, p.r * 1.05)) * (p.n >= 100 ? 0.75 : 1));
      s += `<g class="tp" data-i="${i}"><circle cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="${p.r.toFixed(1)}" fill="${p.n ? "#002d5c" : "#8a99ab"}" stroke="${sel ? "#ffc20e" : "#fff"}" stroke-width="${sel ? 7 : 3}"/>${p.n ? `<text x="${p.x.toFixed(1)}" y="${p.y.toFixed(1)}" dy=".35em" text-anchor="middle" font-size="${fz}" font-weight="800" fill="#fff">${p.n}</text>` : ""}`;
      if (p.rosso) s += `<circle cx="${(p.x + p.r * 0.85).toFixed(1)}" cy="${(p.y - p.r * 0.85).toFixed(1)}" r="10" fill="#d12233" stroke="#fff" stroke-width="2.5"/>`;
      const l = labs[i];
      s += `<text x="${l.lx.toFixed(1)}" y="${l.ly.toFixed(1)}" text-anchor="${l.anchor}" font-size="${FS}" font-weight="700" fill="#1c2733" stroke="#fff" stroke-width="6" paint-order="stroke" stroke-linejoin="round">${esc(p.label)}</text></g>`;
    });
    return s + "</svg>";
  }
  function sedeCard(r) {
    const k = esc(r.key), where = r.via === "manuale" ? "scelta da te" : (r.via || "non indicata");
    return `<div class="card"><h3>${esc(r.nome)}</h3>
      <div class="meta">${num(r.interventi)}${per === "tutto" ? "" : " nel periodo"} · <b${r.da_seguire ? ' style="color:#d12233"' : ""}>${r.da_seguire} da seguire</b> · ${r.utenti} utenti · ${r.stampanti} stampanti</div>
      <div class="meta">Posizione: ${esc(where)}</div>
      <div class="row"><div><button class="btn navy" type="button" data-a="int" data-k="${k}">Vedi gli interventi</button></div><div><button class="btn" type="button" data-a="open" data-k="${k}">Solo da seguire</button></div></div>
      <div class="row"><div><button class="btn" type="button" data-a="place" data-k="${k}">${r.lat == null ? "Posiziona sulla mappa" : "Sposta sulla mappa"}</button></div>${r.via === "manuale" ? `<div><button class="btn red" type="button" data-a="unplace" data-k="${k}">Togli la posizione scelta</button></div>` : ""}</div></div>`;
  }
  async function render() {
    const box = $("rMap");
    if (!box) return;
    const my = ++busy;
    const res = await records(per);
    if (my !== busy) return;
    recs = res.sedi; pins = makePins(recs); last = res;
    selKeys = selKeys.filter(k => recs.some(r => r.key === k));
    const w0 = $("tzW"), sl = w0 ? w0.scrollLeft : 0, st = w0 ? w0.scrollTop : 0;
    const sel = recs.filter(r => selKeys.includes(r.key)), un = recs.filter(r => r.lat == null);
    box.innerHTML = `<label for="tzPer">Interventi</label><select id="tzPer">${PERIODI.map(([k, t]) => `<option value="${k}"${k === per ? " selected" : ""}>${esc(t)}</option>`).join("")}</select>
      <div id="tzMsg"></div>
      <div class="plw" id="tzW"><div class="pli" id="tzI" style="width:${Math.round(zoom * 100)}%">${svgOf()}</div></div>
      <div class="zr"><button class="btn sm" id="tzMinus" type="button" aria-label="Riduci la mappa">−</button><span>Zoom <b id="tzZ">${Math.round(zoom * 100)}%</b></span><button class="btn sm" id="tzPlus" type="button" aria-label="Ingrandisci la mappa">＋</button></div>
      <div id="tzCard">${sel.length ? sel.map(sedeCard).join("") : `<div class="hint">Tocca un pallino per vedere la sede. Il numero è quello degli interventi${per === "tutto" ? "" : " nel periodo scelto"}; il pallino rosso dice che ce ne sono ancora da seguire (aperti o in attesa, di qualsiasi data). Grigio: nessun intervento.</div>`}</div>
      ${un.length ? `<div class="card"><h3>Sedi senza posizione</h3><div class="meta">Non so dove metterle: toccale e poi tocca la mappa.</div>${un.map(r => `<div class="row" style="margin-top:8px;align-items:center"><div style="flex:2"><b>${esc(r.nome)}</b><div class="meta">${num(r.interventi)} · ${r.da_seguire} da seguire</div></div><div><button class="btn sm" type="button" data-a="place" data-k="${esc(r.key)}">Posiziona</button></div></div>`).join("")}</div>` : ""}
      ${res.senza_sede.interventi || res.senza_sede.da_seguire ? `<div class="hint">Senza sede indicata: ${num(res.senza_sede.interventi)} (${res.senza_sede.da_seguire} da seguire).</div>` : ""}
      <div class="hint">Il contorno della Toscana è semplificato; le sedi sono messe dove si trova il comune. Se una sede è nel posto sbagliato, sceglila e usa «Sposta sulla mappa».</div>`;
    const w1 = $("tzW"); w1.scrollLeft = sl; w1.scrollTop = st;
    if (placing) showPlacing();
    wire();
  }
  function showPlacing() {
    const r = recs.find(x => x.key === placing);
    if (!r) { placing = null; return; }
    show($("tzMsg"), `Tocca la mappa dove si trova «${r.nome}».`);
    $("tzMsg").insertAdjacentHTML("beforeend", `<button class="btn" id="tzCancel" type="button">Annulla</button>`);
    $("tzCancel").onclick = () => { placing = null; render(); };
  }
  function svgPoint(ev) {
    const s = $("tzS"), m = s && s.getScreenCTM();
    if (!m) return null;
    const pt = s.createSVGPoint(); pt.x = ev.clientX; pt.y = ev.clientY;
    const q = pt.matrixTransform(m.inverse());
    return [q.x, q.y];
  }
  async function savePos(key, lat, lon) {
    const ov = cleanPos(await getMeta("mappa_pos"));
    if (lat == null) delete ov[key]; else ov[key] = [Math.round(lat * 1e4) / 1e4, Math.round(lon * 1e4) / 1e4];
    await setMeta("mappa_pos", ov);
  }
  function wire() {
    $("tzPer").onchange = () => { per = $("tzPer").value; render(); };
    const again = () => { $("tzI").style.width = Math.round(zoom * 100) + "%"; $("tzZ").textContent = Math.round(zoom * 100) + "%"; };
    $("tzMinus").onclick = () => { zoom = Math.max(1, zoom - 0.5); again(); };
    $("tzPlus").onclick = () => { zoom = Math.min(4, zoom + 0.5); again(); };
    $("tzS").onclick = async ev => {
      const pt = svgPoint(ev);
      if (!pt) return;
      if (placing) {
        const [lat, lon] = unproj(pt[0], pt[1]);
        if (!(lat > 41.9 && lat < 44.7 && lon > 9.5 && lon < 12.6)) { show($("tzMsg"), "Tocca dentro la mappa.", true); return; }
        const key = placing; placing = null;
        await savePos(key, lat, lon);
        selKeys = [key];
        return render();
      }
      const i = pinAt(pins, pt[0], pt[1], MARGIN);
      if (i === null) return;
      selKeys = pins[i].sedi.slice();      // se il pallino riunisce più sedi, le mostra tutte
      await render();
      const c = $("tzCard"); if (c) c.scrollIntoView({ block: "nearest", behavior: "smooth" });
    };
    document.querySelectorAll("#rMap [data-a]").forEach(b => b.onclick = async () => {
      const r = recs.find(x => x.key === b.dataset.k);
      if (!r) return;
      const a = b.dataset.a;
      if (a === "int" || a === "open") return toRegistry(r.nome, a === "open");
      if (a === "place") { placing = r.key; selKeys = [r.key]; await render(); const m = $("tzMsg"); if (m) m.scrollIntoView({ block: "nearest", behavior: "smooth" }); return; }
      if (a === "unplace") { await savePos(r.key, null); selKeys = [r.key]; return render(); }
    });
  }

  /* ---------- dalla mappa al Registro ---------- */
  function toRegistry(nome, soloAperti) {
    ["rUser", "rSede", "rTicket", "rTicketNo", "rDal", "rAl", "rSearch", "rStato"].forEach(id => $(id).value = "");
    const sel = $("rSede");
    if ([...sel.options].every(o => o.value !== nome)) sel.add(new Option(nome, nome));
    sel.value = nome;
    if (soloAperti) $("rStato").value = "aperti";
    else if (per !== "tutto") { $("rDal").value = periodFrom(per); $("rAl").value = isoDay(new Date()); }
    setView("l");
    renderReg();
    window.scrollTo(0, 0);
  }
  const hide = (el, yes) => { if (el) { el.classList.toggle("hidden", yes); el.style.display = yes ? "none" : ""; } };
  function setView(v) {
    view = v;
    const seg = $("rSeg");
    if (!seg) return;
    seg.querySelectorAll("button").forEach(b => b.classList.toggle("on", b.dataset.v === v));
    ["rFlt", "rCount", "rExport", "rList"].forEach(id => hide($(id), v === "m"));
    hide($("rMap"), v !== "m");
    if (v === "m") render();
  }
  function init() {
    const seg = $("rSeg");
    if (!seg) return;
    seg.querySelectorAll("button").forEach(b => b.onclick = () => setView(b.dataset.v));
    setView("l");
  }
  init();

  return { D, PERIODI, norm, find, position, proj, unproj, cleanPos, periodFrom, records, pinR, separate, makePins, placeLabels, pinAt,
           render, setView, toRegistry, entering: () => { if (view === "m") render(); },
           _state: () => ({ per, zoom, view, pins, recs, selKeys, placing, last }) };
})();
