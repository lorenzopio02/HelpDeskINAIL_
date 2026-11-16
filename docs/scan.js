/* scan.js - lettura di codici a barre e seriali con la fotocamera del telefono.
   Usa il lettore del browser (BarcodeDetector) quando c'e', altrimenti il lettore dentro l'app (barcode.js):
   su iPhone il lettore del browser e' spento, quindi li' lavora sempre barcode.js.
   Funziona con i campi che hanno l'attributo data-scan ("serial", "imei" o "any"): accanto compare il pulsante 📷.
   Se la fotocamera non parte (permesso negato, niente rete https...) restano "Scatta una foto" e "Scrivi a mano". */
(function () {
  "use strict";
  const BC = window.BARCODE;
  const $ = id => document.getElementById(id);

  /* ---- controllo dei codici ---- */
  function luhn(s) {
    let sum = 0;
    for (let i = 0; i < s.length; i++) {
      let d = +s[s.length - 1 - i];
      if (i % 2 === 1) { d *= 2; if (d > 9) d -= 9; }
      sum += d;
    }
    return sum % 10 === 0;
  }
  const imeiOk = s => /^\d{15}$/.test(s) && luhn(s);
  /* cosa vale come risposta per quel campo; "" = da scartare */
  function norm(text, kind) {
    let t = String(text == null ? "" : text).replace(/[\x00-\x1f\x7f]/g, "").trim();
    if (!t) return "";
    if (kind === "imei") { t = t.replace(/^imei[\s:\-]*/i, "").replace(/\s+/g, ""); return imeiOk(t) ? t : ""; }
    if (kind === "serial") { t = t.replace(/\s+/g, ""); return t.length >= 4 && t.length <= 40 ? t : ""; }
    return t.length <= 80 ? t : "";
  }
  /* tra i codici letti in una foto o in un fotogramma sceglie quello giusto per il campo */
  function pick(list, kind, single) {
    const ok = [];
    for (const r of list) {
      const t = norm(r.text, kind);
      if (!t) continue;
      const strong = r.votes >= 2 || (r.format === "code128" && r.n >= 4);
      if (strong || (single && r.n >= 4)) ok.push({ text: t, format: r.format, votes: r.votes });
    }
    return ok.length ? ok[0] : null;
  }

  /* ---- foto ferma: si prova a piu' risoluzioni ---- */
  async function loadBitmap(file) {
    try { return await createImageBitmap(file, { imageOrientation: "from-image" }); } catch (e) { /* vecchio Safari: sotto */ }
    return await new Promise((res, rej) => {
      const url = URL.createObjectURL(file), im = new Image();
      im.onload = () => { res(im); setTimeout(() => URL.revokeObjectURL(url), 2000); };
      im.onerror = () => { URL.revokeObjectURL(url); rej(new Error("immagine non leggibile")); };
      im.src = url;
    });
  }
  const pause = () => new Promise(r => setTimeout(r, 15));
  async function decodeSource(src, sw, sh, kind) {
    const tried = new Set();
    for (const L of [1600, 2400, 1000, 700]) {
      const sc = Math.min(1, L / Math.max(sw, sh)), w = Math.max(40, Math.round(sw * sc)), h = Math.max(40, Math.round(sh * sc)), key = w + "x" + h;
      if (tried.has(key)) continue;
      tried.add(key);
      const cv = document.createElement("canvas"); cv.width = w; cv.height = h;
      const cx = cv.getContext("2d", { willReadFrequently: true });
      cx.drawImage(src, 0, 0, w, h);
      const list = BC.decodeGray(BC.toGray(cx.getImageData(0, 0, w, h).data, w, h), w, h, { lines: 22 });
      const got = pick(list, kind, true);
      if (got) return got;
      await pause();
    }
    return null;
  }
  async function decodeFile(file, kind) {
    const bmp = await loadBitmap(file);
    const w = bmp.naturalWidth || bmp.width, h = bmp.naturalHeight || bmp.height;
    const got = await decodeSource(bmp, w, h, kind);
    if (bmp.close) bmp.close();
    return got;
  }

  /* ---- schermata della fotocamera ---- */
  const S = { opts: null, stream: null, timer: 0, busy: false, seen: new Map(), last: "", lastSeen: 0, count: 0, tickN: 0, det: null, rej: "", t0: 0, list: [], track: null, torch: false };
  const EXTRA = [[15, -15, 30, -30], [45, -45, 60, -60], [75, -75, 22, -22]];      // dopo 0 e 90, ogni giro prova quattro direzioni in piu'
  let dom = null, ctx = null, cv = null, grayBuf = null;

  let cssDone = false;
  function css() {
    if (cssDone) return;
    cssDone = true;
    const st = document.createElement("style");
    st.textContent = `.acw,.scw{position:relative}
.scbtn{position:absolute;right:5px;top:50%;transform:translateY(-50%);width:40px;height:36px;border:0;border-radius:8px;background:var(--navy,#002D5C);color:#fff;font-size:18px;line-height:1;padding:0;margin:0;cursor:pointer}
input[type=search]~.scbtn{right:30px}
.scflash{box-shadow:0 0 0 3px #2c954b;transition:box-shadow .3s}
#scanOv{position:fixed;inset:0;z-index:100;background:#000;color:#fff;display:flex;flex-direction:column;font-family:inherit}
#scanOv.hidden{display:none}
#scanOv .sbar{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:max(10px,env(safe-area-inset-top)) 12px 10px;background:#002D5C}
#scanOv .sbar b{font-size:16px}
#scanOv .sbar button{font:inherit;font-weight:700;border:0;border-radius:10px;padding:9px 14px;background:#fff;color:#002D5C}
#scanOv .sview{position:relative;flex:1;min-height:0;background:#111;overflow:hidden}
#scanOv video{position:absolute;inset:0;width:100%;height:100%;object-fit:cover}
#scanOv .sframe{position:absolute;left:7%;right:7%;top:50%;height:30%;transform:translateY(-50%);border:3px solid rgba(255,255,255,.9);border-radius:14px;box-shadow:0 0 0 100vmax rgba(0,0,0,.45);pointer-events:none}
#scanOv .sframe i{position:absolute;left:4%;right:4%;top:50%;height:2px;background:#d12233;opacity:.9}
#scanOv .smsg{padding:10px 14px;font-size:14px;line-height:1.35;background:#001f40;min-height:42px}
#scanOv .smsg.err{background:#d12233}#scanOv .smsg.ok{background:#2c954b}
#scanOv .slist{padding:0 14px 6px;background:#001f40;font-size:13px;max-height:84px;overflow:auto}
#scanOv .sbtns{display:flex;flex-wrap:wrap;gap:8px;padding:10px 12px max(12px,env(safe-area-inset-bottom));background:#001f40}
#scanOv .sbtns button{flex:1 1 140px;font:inherit;font-weight:700;border:0;border-radius:10px;padding:13px 10px;background:#dde8f3;color:#002D5C}
#scanOv .sbtns button.hidden{display:none}`;
    document.head.appendChild(st);
  }
  function ensureDom() {
    if (dom) return dom;
    css();
    const d = document.createElement("div");
    d.id = "scanOv"; d.className = "hidden"; d.setAttribute("role", "dialog"); d.setAttribute("aria-label", "Lettura codice con la fotocamera");
    d.innerHTML = `<div class="sbar"><b id="scanTitle">Leggi il codice</b><button type="button" id="scanClose">Chiudi</button></div>
      <div class="sview"><video id="scanVid" playsinline muted autoplay></video><div class="sframe"><i></i></div></div>
      <div class="smsg" id="scanMsg"></div><div class="slist hidden" id="scanList"></div>
      <div class="sbtns"><button type="button" id="scanPhoto">📷 Scatta una foto</button><button type="button" id="scanTorch" class="hidden">🔦 Torcia</button><button type="button" id="scanManual">⌨ Scrivi a mano</button></div>
      <input type="file" id="scanFile" accept="image/*" style="display:none">`;
    document.body.appendChild(d);
    dom = { root: d, vid: $("scanVid"), msg: $("scanMsg"), title: $("scanTitle"), list: $("scanList"), file: $("scanFile"), torch: $("scanTorch"), close: $("scanClose") };
    $("scanClose").onclick = () => close();
    $("scanManual").onclick = () => { const f = S.opts && S.opts.onManual; close(); if (f) f(); };
    $("scanPhoto").onclick = () => { dom.file.value = ""; dom.file.click(); };
    dom.torch.onclick = async () => {
      if (!S.track) return;
      try { S.torch = !S.torch; await S.track.applyConstraints({ advanced: [{ torch: S.torch }] }); } catch (e) { S.torch = false; }
    };
    dom.file.onchange = async e => {
      const f = e.target.files && e.target.files[0];
      if (!f || !S.opts) return;
      say("Cerco il codice nella foto…");
      let got = null, fail = "";
      try { got = await decodeFile(f, S.kind); } catch (er) { fail = "Non riesco ad aprire questa foto."; }
      if (got) return accept(got.text, true);
      say(fail || (S.kind === "imei" ? "Nella foto non trovo un IMEI valido (15 cifre). Riprova più vicino e dritto, oppure scrivilo a mano." : "Nella foto non trovo nessun codice. Riprova più vicino, con il codice dritto e bene illuminato, oppure scrivilo a mano."), "err");
    };
    document.addEventListener("visibilitychange", () => { if (document.hidden && S.opts) close(); });
    return dom;
  }
  function say(t, cls) { dom.msg.textContent = t; dom.msg.className = "smsg" + (cls ? " " + cls : ""); }

  function stopCam() {
    clearTimeout(S.timer); S.timer = 0;
    if (S.stream) { S.stream.getTracks().forEach(t => t.stop()); S.stream = null; }
    S.track = null; S.torch = false;
    if (dom && dom.vid) { try { dom.vid.pause(); } catch (e) { /* niente */ } dom.vid.srcObject = null; }
  }
  function close() {
    stopCam(); S.opts = null;
    if (dom) dom.root.classList.add("hidden");
    document.documentElement.style.overflow = "";
  }
  function unlockAudio() {
    try {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (!AC) return;
      const ac = beep.ac || (beep.ac = new AC());
      if (ac.state === "suspended") ac.resume();
    } catch (e) { /* niente suono */ }
  }
  function beep() {
    try {
      const AC = window.AudioContext || window.webkitAudioContext;
      if (!AC) return;
      const ac = beep.ac || (beep.ac = new AC()), o = ac.createOscillator(), g = ac.createGain();
      o.type = "square"; o.frequency.value = 1500; g.gain.value = 0.08; o.connect(g); g.connect(ac.destination); o.start(); o.stop(ac.currentTime + 0.12);
    } catch (e) { /* niente suono */ }
    try { if (navigator.vibrate) navigator.vibrate(60); } catch (e) { /* niente */ }
  }
  function accept(text, fromPhoto) {
    const o = S.opts;
    if (!o) return;
    beep();
    if (o.multi && !fromPhoto) {
      S.last = text; S.lastSeen = Date.now(); S.count++; S.list.unshift(text);
      dom.list.classList.remove("hidden");
      dom.list.textContent = `Letti: ${S.count} · ultimo ${S.list[0]}${S.list.length > 1 ? " · prima " + S.list.slice(1, 4).join(", ") : ""}`;
      dom.close.textContent = `Fine (${S.count})`;
      say("Letto: " + text + ". Inquadra il prossimo codice.", "ok");
      try { o.onCode(text); } catch (e) { /* il campo non c'e' piu' */ }
      return;
    }
    const f = o.onCode; close();
    try { f(text); } catch (e) { /* il campo non c'e' piu' */ }
  }

  /* un fotogramma della fotocamera */
  async function tick() {
    S.timer = 0;
    if (!S.stream || !S.opts) return;
    const v = dom.vid;
    if (v.readyState >= 2 && v.videoWidth) {
      try {
        const sc = Math.min(1, 1600 / Math.max(v.videoWidth, v.videoHeight)), w = Math.round(v.videoWidth * sc), h = Math.round(v.videoHeight * sc);
        if (!cv) { cv = document.createElement("canvas"); ctx = cv.getContext("2d", { willReadFrequently: true }); }
        if (cv.width !== w || cv.height !== h) { cv.width = w; cv.height = h; }
        ctx.drawImage(v, 0, 0, w, h);
        let list = [];
        if (S.det) {
          try { list = (await S.det.detect(cv)).map(x => ({ text: x.rawValue, format: x.format === "code_128" ? "code128" : x.format, n: 9, votes: 2 })); } catch (e) { S.det = null; }
        }
        if (!list.length && BC) {
          const ex = EXTRA[S.tickN++ % EXTRA.length];
          grayBuf = BC.toGray(ctx.getImageData(0, 0, w, h).data, w, h, grayBuf);
          list = BC.decodeGray(grayBuf, w, h, { angles: [0, 90, ...ex], lines: 18 });
        }
        consider(list);
      } catch (e) { /* fotogramma saltato */ }
    }
    if (S.stream && S.opts) S.timer = setTimeout(tick, 90);
  }
  function consider(list) {
    const now = Date.now();
    for (const [k, v] of S.seen) if (now - v.t > 2000) S.seen.delete(k);
    for (const r of list) {
      const t = norm(r.text, S.kind);
      if (!t) { if (r.n >= 5 && r.votes >= 2) S.rej = String(r.text).slice(0, 30); continue; }
      const p = S.seen.get(t) || { n: 0, t: 0 };
      p.n++; p.t = now; S.seen.set(t, p);
      const strong = r.votes >= 3 || (r.format === "code128" && r.n >= 4 && r.votes >= 1);
      if (!(strong || p.n >= 2)) continue;
      if (S.opts.multi && t === S.last && now - S.lastSeen < 1500) { S.lastSeen = now; continue; }     // lo stesso codice ancora in vista: non lo riaggiungo
      S.seen.clear(); accept(t, false);
      return;
    }
    if (S.opts && S.kind === "imei" && S.rej && !S.opts.multi) say(`Letto «${S.rej}», ma non è un IMEI valido (15 cifre). Cerca il codice dell'IMEI.`, "err");
    else if (S.opts && !S.count && now - S.t0 > 12000) say("Non lo leggo ancora: avvicinati un po', tieni il codice dritto e ben illuminato. Oppure scatta una foto.");
  }
  async function startCam() {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) { say("Questa pagina non può usare la fotocamera in diretta. Usa «Scatta una foto» oppure scrivi il codice.", "err"); return; }
    try {
      S.stream = await navigator.mediaDevices.getUserMedia({ audio: false, video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 } } });
    } catch (e) {
      const n = e && e.name;
      say(n === "NotAllowedError" || n === "SecurityError"
        ? "La fotocamera è bloccata. Consenti l'uso della fotocamera a questa pagina (Impostazioni di Safari) oppure usa «Scatta una foto»."
        : n === "NotFoundError" || n === "OverconstrainedError" ? "Non trovo la fotocamera. Usa «Scatta una foto» oppure scrivi il codice." : "La fotocamera non parte. Usa «Scatta una foto» oppure scrivi il codice.", "err");
      return;
    }
    if (!S.opts) { stopCam(); return; }                       // intanto hai chiuso
    const v = dom.vid; v.srcObject = S.stream; v.muted = true; v.setAttribute("playsinline", "");
    try { await v.play(); } catch (e) { /* parte comunque dal tocco */ }
    S.track = S.stream.getVideoTracks()[0] || null;
    try {
      const caps = S.track && S.track.getCapabilities ? S.track.getCapabilities() : {};
      if (caps.focusMode && caps.focusMode.includes("continuous")) await S.track.applyConstraints({ advanced: [{ focusMode: "continuous" }] });
      dom.torch.classList.toggle("hidden", !caps.torch);
    } catch (e) { /* alcune fotocamere non lo permettono */ }
    S.det = null;
    if ("BarcodeDetector" in window) {
      try {
        const want = S.kind === "any" ? ["code_128", "code_39", "qr_code", "ean_13", "ean_8", "itf", "codabar", "upc_a"] : ["code_128", "code_39"];
        const have = BarcodeDetector.getSupportedFormats ? await BarcodeDetector.getSupportedFormats() : want;
        const use = want.filter(f => have.includes(f));
        if (use.length) S.det = new BarcodeDetector({ formats: use });
      } catch (e) { S.det = null; }
    }
    S.t0 = Date.now();
    say(S.kind === "imei" ? "Inquadra il codice a barre dell'IMEI (15 cifre)." : "Inquadra il codice a barre e tienilo fermo dentro il riquadro.");
    S.timer = setTimeout(tick, 150);
  }

  /* opts: { title, kind: "serial"|"imei"|"any", multi, onCode(testo), onManual() } */
  function open(opts) {
    ensureDom();
    stopCam();
    unlockAudio();
    S.opts = opts; S.kind = opts.kind || "serial"; S.seen = new Map(); S.last = ""; S.lastSeen = 0; S.count = 0; S.list = []; S.rej = ""; S.tickN = 0; S.t0 = Date.now();
    dom.title.textContent = opts.title || "Leggi il codice";
    dom.close.textContent = "Chiudi"; dom.list.classList.add("hidden"); dom.list.textContent = "";
    dom.root.classList.remove("hidden");
    document.documentElement.style.overflow = "hidden";
    say("Avvio la fotocamera…");
    if (!BC) say("Manca il lettore di codici: usa «Scrivi a mano».", "err");
    return startCam();
  }

  /* ---- il pulsante 📷 accanto ai campi con data-scan ---- */
  function kindOf(inp) {
    const k = inp.dataset.scan;
    if (k === "imei" || k === "serial" || k === "any") {
      if (k === "serial" && inp.dataset.scanImeiField) {         // il campo diventa IMEI se l'altro campo lo dice
        const f = $(inp.dataset.scanImeiField), re = new RegExp(inp.dataset.scanImeiRe || "^imei$", "i");
        if (f && re.test(String(f.value || "").trim())) return "imei";
      }
      return k;
    }
    return "serial";
  }
  function fill(inp, text) {
    inp.value = text;
    inp.dispatchEvent(new Event("input", { bubbles: true }));
    inp.dispatchEvent(new Event("change", { bubbles: true }));
    inp.classList.add("scflash"); setTimeout(() => inp.classList.remove("scflash"), 1600);
  }
  function labelOf(inp) {
    const l = inp.id && document.querySelector(`label[for="${inp.id}"]`);
    return l ? l.textContent.replace(/[*:]/g, "").trim() : "";
  }
  function attach(inp) {
    if (!inp || inp._scanOn) return;
    inp._scanOn = true; inp.setAttribute("data-scan-on", "1");
    let wrap = inp.parentNode;
    if (!wrap.classList.contains("acw")) { wrap = document.createElement("div"); wrap.className = "scw"; inp.parentNode.insertBefore(wrap, inp); wrap.appendChild(inp); }
    const b = document.createElement("button");
    b.type = "button"; b.className = "scbtn"; b.textContent = "📷"; b.setAttribute("aria-label", "Leggi con la fotocamera");
    wrap.appendChild(b);
    inp.style.paddingRight = inp.type === "search" ? "78px" : "52px";
    b.addEventListener("click", ev => {
      ev.preventDefault(); ev.stopPropagation();
      const kind = kindOf(inp), multi = inp.dataset.scanMulti === "1", what = labelOf(inp);
      open({
        kind, multi,
        title: multi ? "Leggi un codice dopo l'altro" : kind === "imei" ? "Leggi l'IMEI" : what ? "Leggi: " + what : "Leggi il codice",
        onCode: text => {
          fill(inp, text);
          if (multi) inp.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
        },
        onManual: () => { try { inp.focus(); } catch (e) { /* niente */ } }
      });
    });
  }
  function scanDom() { document.querySelectorAll("input[data-scan]:not([data-scan-on])").forEach(attach); }
  function init() {
    css();
    scanDom();
    let pending = false;
    new MutationObserver(() => { if (pending) return; pending = true; Promise.resolve().then(() => { pending = false; scanDom(); }); }).observe(document.body, { childList: true, subtree: true });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init); else init();

  window.SCAN = { open, close, attach, luhn, imeiOk, norm, pick, decodeFile, decodeSource, scanDom, _state: S };
})();
