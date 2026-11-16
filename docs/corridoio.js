/* Passeggiata nel piano e edificio in 3D (telefono).
   - Passeggiata: un piano visto dall'alto (scorri lungo il corridoio, giri col dito) oppure IN PRIMA PERSONA:
     cammini nel corridoio, ti giri a destra e a sinistra e sulle porte vedi le targhette (blu col numero della
     stanza, sotto quella chiara con ufficio e nome), come quelle vere. Toccando una porta o una targhetta entri nella stanza 3D.
   - Edificio: i piani impilati, si gira intorno al palazzo e si tocca un piano per camminarci dentro.
   Tutto disegnato in 3D con la CSS (nessuna libreria), dai dati che ci sono già (stanze, persone, stampanti).
   La foto del personaggio (facoltativa) resta solo sul telefono: non viaggia nei file di scambio. */
"use strict";
window.CORR = (function () {
  /* misure del mondo (px): corridoio, stanza, altezza dei muri, altezza degli occhi */
  const CW = 120, RW = 190, RD = 160, H = 140, EYE = 70, WH = 30;
  const DOORW = 62, DOORH = 118, PLW = 64, PLH = 110, FD = 0.30, FPL = 0.74;   /* porta e targhetta: posizione lungo il muro (0 = entrata) */
  const PERSP = 360, SPEED = 170, TURN = 85, AV_PX = 96;
  const clampN = (v, a, b) => Math.max(a, Math.min(b, v));
  const X = () => ROOMS._x;
  const cmpN = (a, b) => String(a).localeCompare(String(b), "it", { numeric: true, sensitivity: "base" });

  /* ---------- stile ---------- */
  const css = `
.wkPage,.twPage{position:fixed;inset:0;z-index:19;display:flex;flex-direction:column;background:#0b2a4d;color:#fff;overscroll-behavior:contain}
.twPage{z-index:18}.wkPage.hidden,.twPage.hidden{display:none}
.wkSegRow{display:flex;justify-content:center;padding:0 12px 8px}.wkSeg{display:flex;gap:6px}.wkSeg button{border:0;border-radius:10px;background:#dde8f3;color:var(--navy);font:700 13px/1.1 -apple-system,"Segoe UI",sans-serif;padding:9px 10px}.wkSeg button.on{background:#fff;box-shadow:0 0 0 2px var(--green) inset}
.wkView{flex:1;min-height:0;position:relative;overflow:hidden;touch-action:none;-webkit-user-select:none;user-select:none}
.wkView.dh,.wkView.tw{perspective:1000px;perspective-origin:50% 38%;background:linear-gradient(#0b2a4d,#1c3f66 55%,#27537f)}
.wkView.fp{perspective:${PERSP}px;perspective-origin:50% 50%;background:linear-gradient(#1b1f26,#3a3f48)}
.wkWorld{position:absolute;left:50%;top:50%;width:0;height:0;transform-style:preserve-3d}
.wkWorld.dh{transform:scale3d(var(--zm,1),var(--zm,1),var(--zm,1)) rotateX(var(--tilt,-58deg)) rotateY(var(--yaw,0deg)) translate3d(0,0,calc(var(--pz,0) * 1px))}
.wkWorld.tw{transform:scale3d(var(--zm,1),var(--zm,1),var(--zm,1)) rotateX(var(--tilt,-28deg)) rotateY(var(--yaw,-35deg)) translate3d(0,calc(var(--ty,0) * 1px),0)}
.wkWorld.fp{transform:translateZ(${PERSP}px) rotateX(var(--pitch,0deg)) rotateY(var(--yaw,0deg)) translate3d(calc(var(--cx,0) * -1px),0,calc(var(--cz,0) * -1px))}
.c3p{position:absolute;left:50%;top:50%;backface-visibility:hidden;-webkit-backface-visibility:hidden;transform-style:preserve-3d}
.c3p.bs{backface-visibility:visible;-webkit-backface-visibility:visible}
.c3floor{background:repeating-linear-gradient(0deg,rgba(0,0,0,.10) 0 2px,transparent 2px 47.5px),linear-gradient(90deg,#76859a 0 4%,#8896a9 4% 96%,#76859a 96%)}
.c3ceil{background:#eef2f6}.c3ceil i{position:absolute;left:28%;right:28%;top:42%;bottom:42%;background:#fff;border-radius:6px;box-shadow:0 0 18px 6px rgba(255,255,255,.9)}
.c3wall{background:linear-gradient(#f4f1ea 0 84%,#c9c2b4 84% 87%,#5d6b7c 87%);overflow:visible}
.c3end{background:linear-gradient(#e4eef8,#c9dbee);display:flex;flex-direction:column;align-items:center;justify-content:center;color:var(--navy);text-align:center}
.c3end b{font-size:22px}.c3end span{font-size:12px;margin-top:4px;padding:0 8px}
.c3end .exit{background:#1a9a4a;color:#fff;font-weight:800;font-size:12px;padding:3px 8px;border-radius:4px;margin-top:8px}
.cdoor{position:absolute;bottom:12px;width:${DOORW}px;height:${DOORH}px;padding:0;border:3px solid #5d4630;border-radius:3px 3px 0 0;background:linear-gradient(90deg,#a47a52,#b58a60 50%,#a47a52)}
.cdoor::after{content:"";position:absolute;right:7px;top:56%;width:5px;height:14px;border-radius:3px;background:#d9dde2;box-shadow:0 0 0 1px #777}
.cdoor::before{content:"";position:absolute;left:10px;right:10px;top:10px;height:30px;border:2px solid rgba(60,40,20,.35);border-radius:2px}
.cplw{position:absolute;bottom:18px;width:${PLW}px;height:${PLH}px;cursor:pointer}
.cpl{font-size:10px;width:6.4em;height:11em;box-shadow:0 2px 3px rgba(0,0,0,.4);border-radius:.25em;overflow:hidden;background:#e3e6ea;text-align:left;color:#1d2b4a;font-family:-apple-system,"Segoe UI",Arial,sans-serif}
.cpn{position:relative;height:4.2em;background:#0b1a4a;color:#fff;display:flex;align-items:center;justify-content:center}
.cpn b{font-size:2.1em;line-height:1;letter-spacing:-.02em;max-width:96%;overflow:hidden;text-overflow:clip;white-space:nowrap}
.cps{position:absolute;right:.35em;top:.35em;width:.85em;height:.85em;border-radius:50%;border:.12em solid #fff;background:#9aa5b1}.cps.ok{background:#1a9a4a}.cps.no{background:#d3202f}
.cpt{padding:.45em .5em;font-size:.88em;line-height:1.18}.cpt b{display:block;font-weight:700;margin-bottom:.25em}.cpt span{display:block;overflow-wrap:anywhere}.cpt em{display:block;color:#6a7686}
.cpl.big{font-size:12px}
.wkBar{display:flex;gap:8px;align-items:center;justify-content:center;flex-wrap:wrap;padding:8px 12px calc(10px + env(safe-area-inset-bottom));background:#0b2a4d}
.wkB{width:56px;height:56px;border:0;border-radius:14px;background:#dde8f3;color:var(--navy);font:700 24px/1 -apple-system,sans-serif;padding:0;touch-action:none;-webkit-user-select:none;user-select:none}
.wkB:active{background:#fff}
.wkBar select{width:auto;max-width:130px;padding:9px;font-size:14px}
.wkBar input[type=range]{flex:1;min-width:120px;width:auto;padding:0}
.wkMini{position:absolute;top:8px;right:8px;background:rgba(255,255,255,.9);border-radius:10px;padding:4px;touch-action:none}
.wkMini svg{display:block}
.wkHud{position:absolute;left:10px;right:10px;bottom:10px;background:rgba(0,20,45,.9);border:1px solid rgba(255,255,255,.3);border-radius:14px;padding:8px 10px;display:none;align-items:center;gap:12px}
.wkHud.on{display:flex}.wkHud .cpl{flex:0 0 auto}.wkHud .wkHt{flex:1;min-width:0;font-size:13px;line-height:1.35}.wkHud .wkHt b{font-size:16px;display:block}
.wkHud .btn{width:auto;margin:0;padding:10px 12px;font-size:14px}
@media (max-height:700px){.wkHud .cpl{display:none}}
.wkHint{padding:6px 14px;font-size:12px;color:#c4d4e6;text-align:center;background:#0b2a4d}
.wkAv{position:fixed;inset:0;z-index:30;background:rgba(0,20,45,.78);display:flex;align-items:flex-end;justify-content:center}
.wkAv>div{background:#fff;color:var(--text);width:100%;max-width:560px;border-radius:16px 16px 0 0;padding:16px 16px calc(18px + env(safe-area-inset-bottom))}
.c3bb{position:absolute;left:50%;top:50%;transform-origin:50% 100%;transform:translate3d(var(--x,0px),var(--y,0px),var(--z,0px)) rotateY(calc(var(--yaw,0deg) * -1)) rotateX(calc(var(--tilt,0deg) * -1));padding:0;border:0;background:none;color:#fff;font:inherit}
.c3bb .cpl{font-size:11px}
.c3bb.pr{width:56px;margin:-70px 0 0 -28px}.c3bb.pr img{width:56px;display:block}
.c3bb.pr .dots{position:absolute;top:-4px;right:-4px;display:flex;gap:2px}.c3bb.pr .dot{width:16px;height:16px;border-radius:50%;border:2px solid #fff;color:#fff;font:800 10px/12px sans-serif;text-align:center}.c3bb.pr .dot.red{background:#d3202f}.c3bb.pr .dot.yel{background:#f2c500;color:#5a3a00}
.c3lab{background:rgba(0,20,45,.85);border:1px solid rgba(255,255,255,.4);border-radius:8px;padding:3px 7px;font-size:12px;font-weight:700;white-space:nowrap}
.dhTile{background:#d9c9a2;border:1px solid #b9ae8c;border-radius:4px}.dhTile.free{background:#e7dfca;border-style:dashed}
.dhDesk{background:#2d4f86;border-radius:3px}.dhDesk.free{background:transparent;border:1px dashed #8b8266}
.dhWall{background:rgba(228,236,246,.92);border:1px solid #9db8e6}
.dhCorr{background:#9eaabb}.dhGround{background:#566377}
.dhLine{background:repeating-linear-gradient(0deg,#e8eef6 0 12px,transparent 12px 24px)}
.twSlab{background:#6e7f95;cursor:pointer}.twSide{background:#46566b}.twCorr{background:#a9b5c6}.twTile{background:#e8dfc8;border:1px solid #b9ae8c}.twTile.ok{background:#bfe3c9}.twTile.no{background:#f3c7c4}.twTile.free{background:#eee9db}
.twRail{position:absolute;left:6px;top:6px;bottom:6px;width:112px;overflow-y:auto;display:flex;flex-direction:column;justify-content:center;gap:6px;touch-action:pan-y}.wkWorld.tw{left:calc(50% + 56px)}
.twBb{display:block;width:100%;flex:0 0 auto;background:rgba(0,20,45,.9);border:2px solid rgba(255,255,255,.55);border-radius:10px;padding:5px 9px;text-align:left;color:#fff;font:inherit;cursor:pointer}
.twBb b{display:block;font-size:13px}.twBb span{display:block;font-size:11px;color:#c4d4e6;line-height:1.3}.twBb i{font-style:normal;font-weight:800;border-radius:9px;padding:0 6px;margin-right:4px;font-size:11px}.twBb i.red{background:#d3202f}.twBb i.yel{background:#f2c500;color:#5a3a00}
.twRoof{background:#0b1a4a;border:2px solid #fff;border-radius:8px;padding:4px 10px;font-weight:800;font-size:14px;text-align:center;color:#fff}`;
  (function () { const s = document.createElement("style"); s.textContent = css; document.head.appendChild(s); })();

  /* ---------- dati del piano ---------- */
  /* le stanze di un piano: numeri dispari da una parte, pari dall'altra (come nei corridoi veri); senza numeri si alterna */
  function layout(rooms) {
    const par = rooms.length > 0 && rooms.every(r => /\d/.test(r.codice));
    const cnt = { "-1": 0, "1": 0 };
    const items = rooms.map((r, i) => {
      let side;
      if (par) { const m = String(r.codice).match(/(\d+)\D*$/) || String(r.codice).match(/(\d+)/); side = (parseInt(m[1], 10) % 2) ? -1 : 1; } else side = i % 2 ? 1 : -1;
      const col = cnt[side]++;
      return { r, side, col, z: -(col + 0.5) * RW, x: side * (CW / 2 + RD / 2) };
    });
    const cols = Math.max(1, cnt["-1"], cnt["1"]);
    return { items, cols, LC: cols * RW };
  }
  /* la targhetta di una stanza: numero, ufficio e nomi */
  function plateData(r) {
    const ppl = ROOMS.peopleIn(r.uid), uff = [];
    let miss = 0;
    ppl.forEach(p => {
      const u = X().userOf(p.k), f = u && u.ufficio ? String(u.ufficio).trim() : "";
      if (f && !uff.includes(f)) uff.push(f);
      if (u) { const dot = X().dotList(u), ph = X().phoneInfo(dot, u.imei); miss += X().deskDevices(u, dot, ph).filter(d => !d.value).length; } else miss++;
    });
    const names = ppl.map(p => p.utente);
    return { uid: r.uid, num: r.codice, uff: uff.slice(0, 2), names: names.slice(0, 3), more: Math.max(0, names.length - 3), n: names.length, free: !names.length, miss, tot: Math.max(0, parseInt(r.scrivanie, 10) || 0) };
  }
  const plateHtml = (pd, big) => `<div class="cpl${big ? " big" : ""}"><div class="cpn"><b>${esc(pd.num)}</b><i class="cps ${pd.free ? "" : pd.miss ? "no" : "ok"}"></i></div><div class="cpt">${pd.uff.map(u => `<b>${esc(u)}</b>`).join("")}${pd.free ? `<em>Libera</em>` : pd.names.map(nm => `<span>${esc(nm)}</span>`).join("") + (pd.more ? `<span>+${pd.more}</span>` : "")}</div></div>`;
  const plateText = pd => `${pd.num}${pd.uff.length ? " · " + pd.uff.join(", ") : ""}${pd.free ? " · libera" : " · " + pd.names.join(", ") + (pd.more ? " +" + pd.more : "")}`;

  /* un piano (rettangolo) del mondo 3D */
  const plane = o => `<div class="c3p ${o.cls || ""}" ${o.attrs || ""} style="width:${o.w}px;height:${o.h}px;margin:${-o.h / 2}px 0 0 ${-o.w / 2}px;transform:translate3d(${o.x || 0}px,${o.y || 0}px,${o.z || 0}px) rotateY(${o.ry || 0}deg) rotateX(${o.rx || 0}deg);${o.style || ""}">${o.html || ""}</div>`;
  const bbd = (cls, x, y, z, html, attrs, style) => `<div class="c3bb ${cls}" ${attrs || ""} style="--x:${x}px;--y:${y}px;--z:${z}px;${style || ""}">${html}</div>`;

  /* ---------- stato ---------- */
  let pageEl = null, ctx = null;           /* ctx = il piano che sto visitando */
  const FPC = { x: 0, z: -40, yaw: 0, pitch: 0 };
  const DH = { yaw: 0, tilt: -58, zoom: 1, pz: 0, moved: 0 };
  let held = { f: 0, t: 0 }, raf = 0, lastT = 0, mode = "dh";
  let avatar = null;

  function floorData(buid, floor) {
    const rooms = ROOMS.roomsOf(buid, floor), L = layout(rooms);
    const sc = typeof scorteNow === "function" ? scorteNow() : [];
    const prs = stampanti.filter(s => s.edificio_uid === buid && s.piano === floor).map(s => ({ s, fl: typeof prnFlags === "function" ? prnFlags(s, sc) : { red: 0, low: [] } }));
    return { buid, floor, rooms, L, plates: new Map(L.items.map(it => [it.r.uid, plateData(it.r)])), prs, bname: X().bName(buid) };
  }
  const fl = d => (typeof floorLabel === "function" ? floorLabel(d.floor) : "Piano " + d.floor);

  /* ---------- scena in prima persona ---------- */
  function sceneFp(d) {
    const { L } = d, bySide = { "-1": new Map(), "1": new Map() };
    L.items.forEach(it => bySide[it.side].set(it.col, it));
    let h = "";
    for (let c = 0; c < L.cols; c++) {
      const zc = -(c + 0.5) * RW;
      h += plane({ w: CW, h: RW, y: EYE, z: zc, rx: 90, cls: "c3floor" });
      h += plane({ w: CW, h: RW, y: EYE - H, z: zc, rx: -90, cls: "c3ceil", html: "<i></i>" });
      for (const side of [-1, 1]) {
        const it = bySide[side].get(c), pd = it && d.plates.get(it.r.uid);
        let inner = "";
        if (it) {
          const dl = (side < 0 ? FD : 1 - FD) * RW, pl = (side < 0 ? FPL : 1 - FPL) * RW;
          inner = `<button type="button" class="cdoor" data-r="${esc(it.r.uid)}" style="left:${(dl - DOORW / 2).toFixed(1)}px" aria-label="Porta ${esc(it.r.codice)}"></button><div class="cplw" data-r="${esc(it.r.uid)}" style="left:${(pl - PLW / 2).toFixed(1)}px" title="${esc(plateText(pd))}">${plateHtml(pd)}</div>`;
        }
        h += plane({ w: RW, h: H, x: side * CW / 2, y: EYE - H / 2, z: zc, ry: side < 0 ? 90 : -90, cls: "c3wall", html: inner });
      }
    }
    h += plane({ w: CW, h: H, y: EYE - H / 2, z: -L.LC, cls: "c3end", html: `<b>${esc(fl(d))}</b><span>${esc(d.bname)}</span><div class="exit">USCITA ›</div>` });
    h += plane({ w: CW, h: H, y: EYE - H / 2, z: 0, ry: 180, cls: "c3end", html: `<b>${esc(fl(d))}</b><span>Entrata del piano</span>` });
    /* stampanti nel corridoio: in fondo, davanti alla parete */
    const free = d.prs.filter(p => !p.s.stanza_uid);
    free.forEach((p, i) => {
      const x = (i - (free.length - 1) / 2) * 52;
      h += plane({ w: 46, h: 62, x, y: EYE - 31, z: -L.LC + 8, cls: "bs", html: `<img src="${printerImg(p.s)}" alt="" style="width:100%;height:100%;object-fit:contain">`, attrs: `title="${esc(printerName(p.s))}"` });
    });
    return h;
  }

  /* ---------- scena dall'alto (casa delle bambole) ---------- */
  function sceneDh(d) {
    const { L } = d, Wd = CW + 2 * RD;
    let h = plane({ w: Wd + 40, h: L.LC + 40, y: EYE, z: -L.LC / 2, rx: 90, cls: "dhGround bs" });
    h += plane({ w: CW, h: L.LC, y: EYE - 0.6, z: -L.LC / 2, rx: 90, cls: "dhCorr bs" });
    h += plane({ w: 6, h: L.LC - 20, y: EYE - 1.2, z: -L.LC / 2, rx: 90, cls: "dhLine bs" });
    d.L.items.forEach(it => {
      const pd = d.plates.get(it.r.uid), cx = it.x, cz = it.z, side = it.side;
      h += plane({ w: RD - 10, h: RW - 10, x: cx, y: EYE - 0.8, z: cz, rx: 90, cls: `dhTile bs${pd.free ? " free" : ""}`, attrs: `data-r="${esc(it.r.uid)}"` });
      /* muro di fondo, muri laterali e muro verso il corridoio con il vano della porta */
      h += plane({ w: RW - 10, h: WH, x: side * (CW / 2 + RD), y: EYE - WH / 2, z: cz, ry: side < 0 ? 90 : -90, cls: "dhWall bs" });
      h += plane({ w: RD, h: WH, x: cx, y: EYE - WH / 2, z: -(it.col + 1) * RW, cls: "dhWall bs" });
      if (it.col === 0) h += plane({ w: RD, h: WH, x: cx, y: EYE - WH / 2, z: 0, cls: "dhWall bs" });
      const dz = -(it.col + FD) * RW, a = (dz - (-it.col * RW)), segA = Math.abs(a) - DOORW / 2, segB = RW - Math.abs(a) - DOORW / 2;
      if (segA > 1) h += plane({ w: segA, h: WH, x: side * CW / 2, y: EYE - WH / 2, z: -it.col * RW - segA / 2, ry: side < 0 ? 90 : -90, cls: "dhWall bs" });
      if (segB > 1) h += plane({ w: segB, h: WH, x: side * CW / 2, y: EYE - WH / 2, z: -(it.col + 1) * RW + segB / 2, ry: side < 0 ? 90 : -90, cls: "dhWall bs" });
      /* scrivanie: una per persona (blu) e le libere tratteggiate */
      const tot = Math.max(pd.n, pd.tot), cols2 = 2;
      for (let i = 0; i < Math.min(tot, 8); i++) {
        const gx = (i % cols2 - 0.5) * 52, gz = (Math.floor(i / cols2) - 1) * 34;
        h += plane({ w: 38, h: 22, x: cx + gx + side * 6, y: EYE - 1.2, z: cz + gz + 18, rx: 90, cls: `dhDesk bs${i < pd.n ? "" : " free"}` });
      }
      /* targhetta in piedi accanto alla porta */
      h += bbd("", side * (CW / 2 - 2), EYE - 1, -(it.col + FPL) * RW, `<button type="button" class="dhPl" data-r="${esc(it.r.uid)}" style="background:none;border:0;padding:0;color:inherit">${plateHtml(pd)}</button>`, "", "margin:-120px 0 0 -35px;width:70px");
    });
    /* stampanti in corridoio */
    const free = d.prs.filter(p => !p.s.stanza_uid);
    free.forEach((p, i) => {
      const z = -L.LC + 30 + i * 60;
      h += bbd("pr", 0, EYE - 1, z, `<img src="${printerImg(p.s)}" alt=""><span class="dots">${p.fl.red ? `<i class="dot red">${p.fl.red}</i>` : ""}${p.fl.low.length ? `<i class="dot yel">!</i>` : ""}</span>`, `data-p="${esc(p.s.uid)}" title="${esc(printerName(p.s))}"`);
    });
    return h;
  }

  /* ---------- pagina della passeggiata ---------- */
  function ensurePage() {
    if (!pageEl) { pageEl = document.createElement("div"); pageEl.id = "wkPage"; pageEl.className = "wkPage hidden"; document.body.appendChild(pageEl); }
    return pageEl;
  }
  const closeWalk = () => { stopLoop(); if (pageEl) { pageEl.classList.add("hidden"); pageEl.innerHTML = ""; } ctx = null; };
  function stopLoop() { held = { f: 0, t: 0 }; if (raf) cancelAnimationFrame(raf); raf = 0; }

  async function openWalk(buid, floor, opt) {
    await ROOMS.load();
    if (typeof loadPrn === "function" && !stampanti.length) { try { await loadPrn(); } catch (e) { /* nessuna stampante */ } }
    const d = floorData(buid, floor);
    if (!d.rooms.length) { alert("Su questo piano non ci sono ancora stanze: creale in «Piani e stanze»."); return; }
    ctx = d; avatar = (await getMeta("avatar")) || null;
    mode = (opt && opt.mode) || mode || "dh";
    FPC.x = 0; FPC.z = -40; FPC.yaw = 0; FPC.pitch = 0;
    DH.yaw = 0; DH.tilt = -58; DH.pz = d.L.LC / 2; DH.moved = 0;
    const vw = window.innerWidth, vh = window.innerHeight;
    DH.zoom = clampN(Math.min(vw * 0.92 / (CW + 2 * RD + 40), (vh - 190) * 1.25 / (d.L.LC + 60)), 0.32, 1.3);
    build();
  }

  function build(keep) {
    const d = ctx; if (!d) return;
    const pg = ensurePage(), fp = mode === "fp";
    stopLoop();
    pg.innerHTML = `<div class="r3Top"><button class="dkBack" id="wkClose" type="button">← Indietro</button><b>${esc(fl(d))} · ${esc(d.bname)}</b></div><div class="wkSegRow"><span class="wkSeg"><button type="button" id="wkMd" class="${fp ? "" : "on"}">🏠 Dall'alto</button><button type="button" id="wkMf" class="${fp ? "on" : ""}">🚶 Cammina</button></span></div>
      <div class="wkView ${fp ? "fp" : "dh"}" id="wkView"><div class="wkWorld ${fp ? "fp" : "dh"}" id="wkWorld">${fp ? sceneFp(d) : sceneDh(d)}</div>${fp ? `<div class="wkMini" id="wkMini"></div><div class="wkHud" id="wkHud"></div>` : ""}</div>
      ${fp ? `<div class="wkBar"><button class="wkB" id="wkL" type="button" aria-label="Gira a sinistra">↶</button><button class="wkB" id="wkF" type="button" aria-label="Avanti">▲</button><button class="wkB" id="wkBk" type="button" aria-label="Indietro">▼</button><button class="wkB" id="wkR" type="button" aria-label="Gira a destra">↷</button>
        <select id="wkGo" aria-label="Vai alla stanza"><option value="">Vai a…</option>${d.L.items.slice().sort((a, b) => cmpN(a.r.codice, b.r.codice)).map(it => `<option value="${esc(it.r.uid)}">${esc(it.r.codice)}</option>`).join("")}</select><button class="wkB" id="wkAvB" type="button" aria-label="La mia faccia" style="width:48px;height:48px;font-size:22px">🙂</button></div>
        <div class="wkHint">Tieni premute ▲ ▼ per camminare, ↶ ↷ per girarti (o trascina il dito per guardare intorno). Tocca una porta o una targhetta per entrare.</div>`
      : `<div class="wkBar"><input type="range" id="wkSl" min="0" max="${d.L.LC}" step="1" value="${DH.pz}" aria-label="Lungo il corridoio"><span class="r3Zm"><button id="wkOut" type="button" aria-label="Allontana">−</button><button id="wkIn" type="button" aria-label="Avvicina">+</button></span></div>
        <div class="wkHint">Trascina per girare · cursore o due dita per spostarti e avvicinarti · tocca una targhetta per entrare nella stanza</div>`}`;
    pg.classList.remove("hidden");
    $("wkClose").onclick = closeWalk;
    $("wkMd").onclick = () => { if (mode !== "dh") { mode = "dh"; build(); } };
    $("wkMf").onclick = () => { if (mode !== "fp") { mode = "fp"; build(); } };
    const world = $("wkWorld"), view = $("wkView");
    const pts = new Map(); let pinch = 0;
    const span = () => { const a = [...pts.values()]; return Math.hypot(a[0][0] - a[1][0], a[0][1] - a[1][1]); };
    const tapRoom = uid => { if (DH.moved > 8) { DH.moved = 0; return; } X().open3d(uid); };
    /* --- dall'alto --- */
    if (!fp) {
      const apply = () => { world.style.setProperty("--yaw", DH.yaw.toFixed(1) + "deg"); world.style.setProperty("--tilt", DH.tilt.toFixed(1) + "deg"); world.style.setProperty("--zm", DH.zoom.toFixed(3)); world.style.setProperty("--pz", DH.pz.toFixed(1)); if ($("wkSl")) $("wkSl").value = DH.pz; };
      apply();
      view.onpointerdown = e => { pts.set(e.pointerId, [e.clientX, e.clientY]); DH.moved = 0; if (pts.size === 2) pinch = span(); };
      view.onpointermove = e => {
        const old = pts.get(e.pointerId); if (!old) return;
        pts.set(e.pointerId, [e.clientX, e.clientY]);
        const dx = e.clientX - old[0], dy = e.clientY - old[1]; DH.moved += Math.abs(dx) + Math.abs(dy);
        if (pts.size === 2) { const s = span(); if (pinch) DH.zoom = clampN(DH.zoom * s / pinch, 0.25, 2.2); pinch = s; }
        else { DH.yaw += dx * 0.45; DH.tilt = clampN(DH.tilt + dy * 0.3, -82, -14); }
        apply();
      };
      view.onpointerup = view.onpointercancel = view.onpointerleave = e => { pts.delete(e.pointerId); pinch = 0; };
      $("wkIn").onclick = () => { DH.zoom = clampN(DH.zoom * 1.18, 0.25, 2.2); apply(); };
      $("wkOut").onclick = () => { DH.zoom = clampN(DH.zoom / 1.18, 0.25, 2.2); apply(); };
      $("wkSl").oninput = e => { DH.pz = +e.target.value; apply(); };
      pg.querySelectorAll(".dhPl,.dhTile").forEach(b => b.onclick = () => tapRoom(b.dataset.r));
      return;
    }
    /* --- in prima persona --- */
    const hud = $("wkHud");
    const apply = () => {
      world.style.setProperty("--yaw", FPC.yaw.toFixed(1) + "deg"); world.style.setProperty("--pitch", FPC.pitch.toFixed(1) + "deg");
      world.style.setProperty("--cx", FPC.x.toFixed(1)); world.style.setProperty("--cz", FPC.z.toFixed(1));
      drawMini(); drawHud();
    };
    const nearest = () => {      /* la targhetta che sto guardando: la piu' vicina al centro della vista, entro 40 gradi e 230 px */
      const r = FPC.yaw * Math.PI / 180, fx = Math.sin(r), fz = -Math.cos(r);
      let best = null;
      ctx.L.items.forEach(it => {
        const px = it.side * CW / 2, pz = -(it.col + FPL) * RW, vx = px - FPC.x, vz = pz - FPC.z, dist = Math.hypot(vx, vz);
        if (dist > 230 || dist < 1) return;
        const ang = Math.acos(clampN((vx * fx + vz * fz) / dist, -1, 1)) * 180 / Math.PI;
        if (ang < 40 && (!best || ang < best.ang)) best = { it, ang, dist };
      });
      return best;
    };
    function drawHud() {
      const b = nearest();
      if (!b) { hud.classList.remove("on"); hud.innerHTML = ""; hud.dataset.r = ""; return; }
      const pd = ctx.plates.get(b.it.r.uid);
      if (hud.dataset.r !== pd.uid) {
        hud.dataset.r = pd.uid;
        hud.innerHTML = `${plateHtml(pd, true)}<div class="wkHt"><b>Stanza ${esc(pd.num)}</b>${pd.free ? "Nessuno è assegnato." : esc(pd.names.join(", ")) + (pd.more ? " e altri " + pd.more : "")}${pd.uff.length ? "<br>" + esc(pd.uff.join(" · ")) : ""}${pd.miss && !pd.free ? `<br>Codici mancanti: ${pd.miss}` : ""}</div><button class="btn navy" id="wkEnter" type="button">Entra</button>`;
        $("wkEnter").onclick = () => X().open3d(pd.uid);
      }
      hud.classList.add("on");
    }
    /* mini-mappa: il piano dall'alto con il segnaposto (la mia faccia, se l'ho scelta) */
    const MS = Math.min(0.2, 170 / ctx.L.LC), MW = Math.round((CW + 2 * RD) * MS), MH = Math.round(ctx.L.LC * MS);
    const mx = x => (x + CW / 2 + RD) * MS, my = z => -z * MS;
    function drawMini() {
      const r = FPC.yaw * Math.PI / 180, ax = mx(FPC.x), ay = my(FPC.z), ln = 16;
      const cone = `<polygon points="${ax},${ay} ${(ax + Math.sin(r - 0.5) * ln).toFixed(1)},${(ay + Math.cos(r - 0.5) * ln).toFixed(1)} ${(ax + Math.sin(r + 0.5) * ln).toFixed(1)},${(ay + Math.cos(r + 0.5) * ln).toFixed(1)}" fill="rgba(26,154,74,.35)"/>`;
      const me = avatar ? `<defs><clipPath id="wkCl"><circle cx="${ax}" cy="${ay}" r="9"/></clipPath></defs><image href="${avatar}" x="${ax - 9}" y="${ay - 9}" width="18" height="18" clip-path="url(#wkCl)" preserveAspectRatio="xMidYMid slice"/><circle cx="${ax}" cy="${ay}" r="9" fill="none" stroke="#1a9a4a" stroke-width="2"/>` : `<circle cx="${ax}" cy="${ay}" r="6" fill="#1a9a4a" stroke="#fff" stroke-width="2"/>`;
      $("wkMini").innerHTML = `<svg width="${MW}" height="${MH}" viewBox="0 0 ${MW} ${MH}" role="img" aria-label="Mini-mappa del piano"><rect x="${mx(-CW / 2)}" y="0" width="${CW * MS}" height="${MH}" fill="#c9d3df"/>${ctx.L.items.map(it => `<rect data-r="${esc(it.r.uid)}" x="${mx(it.x - RD / 2)}" y="${my(it.z + RW / 2)}" width="${RD * MS - 1}" height="${RW * MS - 1}" fill="${ctx.plates.get(it.r.uid).free ? "#eee9db" : "#d9c9a2"}" stroke="#8b8266" stroke-width="1"/><text x="${mx(it.x)}" y="${my(it.z) + 3}" font-size="9" text-anchor="middle" fill="#1d2b4a" style="pointer-events:none">${esc(it.r.codice)}</text>`).join("")}${cone}${me}</svg>`;
    }
    $("wkMini").onpointerdown = e => {
      e.stopPropagation();
      const t = e.target.closest && e.target.closest("rect[data-r]"); if (t) goRoom(t.getAttribute("data-r"));
    };
    function goRoom(uid) {
      const it = ctx.L.items.find(x => x.r.uid === uid); if (!it) return;
      /* mi metto in mezzo al corridoio, un po' prima della stanza, girato verso il muro: porta e targhetta entrano tutte e due nella vista */
      FPC.x = 0; FPC.z = clampN(-(it.col + (FD + FPL) / 2) * RW + 90, -ctx.L.LC + 18, -6); FPC.yaw = it.side * 34; FPC.pitch = 0; apply();
    }
    $("wkGo").onchange = e => { if (e.target.value) goRoom(e.target.value); e.target.value = ""; e.target.blur(); };
    /* movimento */
    const loop = t => {
      raf = 0;
      const dt = Math.min(0.05, (t - lastT) / 1000 || 0.016); lastT = t;
      if (!held.f && !held.t) return;
      FPC.yaw += TURN * dt * held.t;
      const r = FPC.yaw * Math.PI / 180, sp = SPEED * dt * held.f;
      FPC.x = clampN(FPC.x + Math.sin(r) * sp, -CW / 2 + 16, CW / 2 - 16);
      FPC.z = clampN(FPC.z - Math.cos(r) * sp, -ctx.L.LC + 18, -6);
      apply(); raf = requestAnimationFrame(loop);
    };
    const hold = (id, k, v) => {
      const b = $(id), on = e => { e.preventDefault(); try { b.setPointerCapture(e.pointerId); } catch (x) { /* ok */ } held[k] = v; if (!raf) { lastT = performance.now(); raf = requestAnimationFrame(loop); } };
      const off = () => { held[k] = 0; };
      b.onpointerdown = on; b.onpointerup = b.onpointercancel = b.onlostpointercapture = off;
    };
    hold("wkF", "f", 1); hold("wkBk", "f", -1); hold("wkL", "t", -1); hold("wkR", "t", 1);
    /* guardare intorno col dito; toccare porta o targhetta = entrare */
    view.onpointerdown = e => { if (e.target.closest && e.target.closest(".wkMini,.wkHud")) return; pts.set(e.pointerId, [e.clientX, e.clientY]); DH.moved = 0; };
    view.onpointermove = e => {
      const old = pts.get(e.pointerId); if (!old) return;
      pts.set(e.pointerId, [e.clientX, e.clientY]);
      const dx = e.clientX - old[0], dy = e.clientY - old[1]; DH.moved += Math.abs(dx) + Math.abs(dy);
      FPC.yaw += dx * 0.28; FPC.pitch = clampN(FPC.pitch + dy * 0.12, -22, 22); apply();
    };
    view.onpointerup = view.onpointercancel = view.onpointerleave = e => { pts.delete(e.pointerId); };
    pg.querySelectorAll(".cdoor,.cplw").forEach(b => b.onclick = () => tapRoom(b.dataset.r));
    document.onkeydown = e => {
      if (pageEl.classList.contains("hidden") || mode !== "fp") return;
      const k = e.key.toLowerCase(), m = { arrowup: ["f", 1], w: ["f", 1], arrowdown: ["f", -1], s: ["f", -1], arrowleft: ["t", -1], a: ["t", -1], arrowright: ["t", 1], d: ["t", 1] }[k];
      if (!m) return; e.preventDefault(); held[m[0]] = m[1]; if (!raf) { lastT = performance.now(); raf = requestAnimationFrame(loop); }
    };
    document.onkeyup = e => { const k = e.key.toLowerCase(); if (["arrowup", "w", "arrowdown", "s"].includes(k)) held.f = 0; if (["arrowleft", "a", "arrowright", "d"].includes(k)) held.t = 0; };
    $("wkAvB").onclick = avatarSheet;
    apply();
  }

  /* ---------- la mia faccia (solo sul telefono) ---------- */
  function squarePhoto(file) {
    return new Promise((res, rej) => {
      const url = URL.createObjectURL(file), im = new Image();
      im.onload = () => {
        const s = Math.min(im.naturalWidth, im.naturalHeight), c = document.createElement("canvas"); c.width = c.height = AV_PX;
        c.getContext("2d").drawImage(im, (im.naturalWidth - s) / 2, (im.naturalHeight - s) / 2, s, s, 0, 0, AV_PX, AV_PX);
        URL.revokeObjectURL(url); res(c.toDataURL("image/jpeg", 0.85));
      };
      im.onerror = () => { URL.revokeObjectURL(url); rej(new Error("foto non leggibile")); };
      im.src = url;
    });
  }
  function avatarSheet() {
    const el = document.createElement("div"); el.className = "wkAv"; el.id = "wkAv";
    el.innerHTML = `<div><h3 style="margin:0 0 6px;color:var(--navy)">La mia faccia</h3><div class="meta">La foto compare al posto del pallino verde sulla mini-mappa. Resta SOLO su questo telefono: non va nei file di scambio, non va sul PC e non va su GitHub.</div>
      ${avatar ? `<div style="text-align:center;margin:10px 0"><img src="${avatar}" alt="La mia faccia" style="width:84px;height:84px;border-radius:50%;border:3px solid #1a9a4a"></div>` : ""}<div id="wkAvMsg"></div>
      <input id="wkAvF" type="file" accept="image/*" class="hidden">
      <button class="btn green" id="wkAvPick" type="button">${avatar ? "Cambia la foto" : "Scegli una foto"}</button>
      ${avatar ? `<button class="btn red" id="wkAvDel" type="button">Togli la foto</button>` : ""}
      <button class="btn" id="wkAvClose" type="button">Chiudi</button></div>`;
    document.body.appendChild(el);
    const shut = () => { el.remove(); };
    $("wkAvClose").onclick = shut;
    $("wkAvPick").onclick = () => $("wkAvF").click();
    $("wkAvF").onchange = async e => {
      const f = e.target.files[0]; e.target.value = ""; if (!f) return;
      try { avatar = await squarePhoto(f); await setMeta("avatar", avatar); shut(); if (ctx && mode === "fp") { const hudKeep = $("wkHud") && $("wkHud").dataset.r; void hudKeep; build(); } }
      catch (er) { $("wkAvMsg").innerHTML = `<div class="msg err">Non riesco a leggere questa foto: provane un'altra.</div>`; }
    };
    if ($("wkAvDel")) $("wkAvDel").onclick = async () => { avatar = null; await setMeta("avatar", null); shut(); if (ctx && mode === "fp") build(); };
  }

  /* ---------- l'edificio intero ---------- */
  let twEl = null, TW = null;
  const TWS = { yaw: -35, tilt: -28, zoom: 1, ty: 0, moved: 0 };
  const closeTower = () => { if (twEl) { twEl.classList.add("hidden"); twEl.innerHTML = ""; } TW = null; };
  async function openTower(buid) {
    await ROOMS.load();
    if (typeof loadPrn === "function" && !stampanti.length) { try { await loadPrn(); } catch (e) { /* ok */ } }
    const b = edifici.find(x => x.uid === buid); if (!b) return;
    const floors = new Set(); for (let f = Math.min(b.piano_min, b.piano_max); f <= Math.max(b.piano_min, b.piano_max); f++) floors.add(f);
    stampanti.filter(s => s.edificio_uid === buid).forEach(s => floors.add(s.piano));
    const list = [...floors].sort((a, c) => a - c), data = list.map(f => floorData(buid, f));
    TW = { b, data };
    build_tower();
  }
  function build_tower() {
    const { b, data } = TW;
    if (!twEl) { twEl = document.createElement("div"); twEl.id = "twPage"; twEl.className = "twPage hidden"; document.body.appendChild(twEl); }
    const S = 0.42, GAP = 70, T = 10, cols = Math.max(1, ...data.map(d => d.L.cols)), Wd = (CW + 2 * RD) * S, Dd = cols * RW * S;
    const n = data.length, top = -(n - 1) * GAP;
    let h = "";
    const rail = [];
    data.forEach((d, i) => {
      const y = -i * GAP;
      h += plane({ w: Wd, h: Dd, y, z: 0, rx: 90, cls: "twSlab", attrs: `data-f="${d.floor}"` });
      h += plane({ w: Wd, h: T, y: y + T / 2, z: Dd / 2, cls: "twSide", attrs: `data-f="${d.floor}"` });
      h += plane({ w: Dd, h: T, x: Wd / 2, y: y + T / 2, z: 0, ry: 90, cls: "twSide", attrs: `data-f="${d.floor}"` });
      h += plane({ w: Dd, h: T, x: -Wd / 2, y: y + T / 2, z: 0, ry: -90, cls: "twSide", attrs: `data-f="${d.floor}"` });
      h += plane({ w: CW * S, h: Dd, y: y - 0.6, z: 0, rx: 90, cls: "twCorr", attrs: `data-f="${d.floor}"` });
      d.L.items.forEach(it => {
        const pd = d.plates.get(it.r.uid), st = pd.free ? "free" : pd.miss ? "no" : "ok";
        h += plane({ w: (RD - 12) * S, h: (RW - 12) * S, x: it.x * S, y: y - 0.9, z: (it.z + cols * RW / 2) * S, rx: 90, cls: `twTile ${st}`, attrs: `data-f="${d.floor}"` });
      });
      const rooms = d.rooms.length, ppl = d.rooms.reduce((t, r) => t + ROOMS.peopleIn(r.uid).length, 0), red = d.prs.reduce((t, p) => t + p.fl.red, 0), low = d.prs.filter(p => p.fl.low.length).length, off = d.prs.filter(p => p.s.fuori_servizio).length;
      rail.unshift(`<button type="button" class="twBb" data-f="${d.floor}"><b>${esc(fl(d))}</b><span>${rooms} ${rooms === 1 ? "stanza" : "stanze"} · ${ppl} ${ppl === 1 ? "persona" : "persone"}</span><span>🖨 ${d.prs.length}${red ? ` <i class="red">${red}</i>` : ""}${low ? ` <i class="yel">!</i>` : ""}${off ? ` <i class="red">✕</i>` : ""}</span></button>`);
    });
    h += bbd("", 0, top - 40, 0, `<div class="twRoof">${esc(b.nome)}</div>`, "", "margin:-34px 0 0 -100px;width:200px");
    twEl.innerHTML = `<div class="r3Top"><button class="dkBack" id="twClose" type="button">← Indietro</button><b>${esc(b.nome)} · edificio 3D</b><span class="r3Zm"><button id="twOut" type="button" aria-label="Allontana">−</button><button id="twIn" type="button" aria-label="Avvicina">+</button></span></div>
      <div class="wkView tw" id="twView"><div class="wkWorld tw" id="twWorld">${h}</div><div class="twRail" id="twRail">${rail.join("")}</div></div>
      <div class="wkHint">Trascina per girare intorno all'edificio · due dita o +/− per avvicinarti · tocca un piano per camminarci dentro</div>`;
    twEl.classList.remove("hidden");
    const world = $("twWorld"), view = $("twView");
    const totalH = (n - 1) * GAP + 80;
    TWS.yaw = -35; TWS.tilt = -28; TWS.ty = -top / 2 + 6; TWS.moved = 0;
    TWS.zoom = clampN(Math.min((window.innerWidth - 24 - 112) / (Math.hypot(Wd, Dd) + 60), (window.innerHeight - 200) / (totalH + Dd * 0.45)), 0.3, 1.5);
    const apply = () => { world.style.setProperty("--yaw", TWS.yaw.toFixed(1) + "deg"); world.style.setProperty("--tilt", TWS.tilt.toFixed(1) + "deg"); world.style.setProperty("--zm", TWS.zoom.toFixed(3)); world.style.setProperty("--ty", TWS.ty.toFixed(1)); };
    apply();
    const pts = new Map(); let pinch = 0;
    const span = () => { const a = [...pts.values()]; return Math.hypot(a[0][0] - a[1][0], a[0][1] - a[1][1]); };
    view.onpointerdown = e => { TWS.moved = 0; if (e.target.closest && e.target.closest(".twRail")) return; pts.set(e.pointerId, [e.clientX, e.clientY]); if (pts.size === 2) pinch = span(); };
    view.onpointermove = e => {
      const old = pts.get(e.pointerId); if (!old) return;
      pts.set(e.pointerId, [e.clientX, e.clientY]);
      const dx = e.clientX - old[0], dy = e.clientY - old[1]; TWS.moved += Math.abs(dx) + Math.abs(dy);
      if (pts.size === 2) { const s = span(); if (pinch) TWS.zoom = clampN(TWS.zoom * s / pinch, 0.25, 2.4); pinch = s; }
      else { TWS.yaw += dx * 0.45; TWS.tilt = clampN(TWS.tilt + dy * 0.3, -80, -8); }
      apply();
    };
    view.onpointerup = view.onpointercancel = view.onpointerleave = e => { pts.delete(e.pointerId); pinch = 0; };
    $("twIn").onclick = () => { TWS.zoom = clampN(TWS.zoom * 1.18, 0.25, 2.4); apply(); };
    $("twOut").onclick = () => { TWS.zoom = clampN(TWS.zoom / 1.18, 0.25, 2.4); apply(); };
    $("twClose").onclick = closeTower;
    twEl.querySelectorAll("[data-f]").forEach(e => e.onclick = ev => { if (TWS.moved > 8) { TWS.moved = 0; return; } ev.stopPropagation(); openWalk(TW.b.uid, +e.dataset.f); });
  }

  /* quando chiudo una scheda (per esempio dopo aver cambiato un nome) ridisegno la passeggiata con i dati nuovi */
  const mo = new MutationObserver(() => {
    const m = $("modal");
    if (m && m.classList.contains("hidden") && ctx && pageEl && !pageEl.classList.contains("hidden")) {
      const keep = { ...FPC };
      ROOMS.load().then(() => { if (!ctx) return; ctx = floorData(ctx.buid, ctx.floor); Object.assign(FPC, keep); build(true); });
    }
  });
  document.addEventListener("DOMContentLoaded", () => { const m = $("modal"); if (m) mo.observe(m, { attributes: true, attributeFilter: ["class"] }); });
  if (document.readyState !== "loading") { const m = $("modal"); if (m) mo.observe(m, { attributes: true, attributeFilter: ["class"] }); }

  return { openWalk, openTower, closeWalk, closeTower, layout, plateData, _fp: FPC, _dh: DH, _mode: () => mode, _ctx: () => ctx, setMode: m => { mode = m; } };
})();
