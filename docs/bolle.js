/* Bolle dei toner: registro il toner quando arriva (entra nell'inventario del piano) e lo "monto" quando lo cambio.
   - Arrivo: leggo il codice a barre della bolla (o la foto), scelgo colore, edificio e piano. Il toner e' "da montare".
   - Montaggio: inquadro il QR (o il codice a barre) del seriale della stampante: l'app la riconosce e controlla che
     il seriale coincida con quello scritto sulla bolla. "Montato adesso" scrive il cambio toner con data e ora,
     tiene il numero della bolla e la foto, e toglie il toner dall'inventario del piano.
   - La scorta (magazzino) sale all'arrivo e scende al montaggio, come gli altri cambi toner.
   Si usa da Stampanti › Cambi toner (elenco), Stampanti › Scorta (inventario per piano) e dalla scheda stampante. */
"use strict";
window.BOL = (function () {
  const LETTER = { K: "Nero", C: "Azzurro", M: "Magenta", Y: "Giallo" };
  const colorOfCode = c => { const m = /^[A-Z]{2,4}\d{2,4}([KCMY])$/i.exec(String(c || "").replace(/[\s-]+/g, "")); return m ? LETTER[m[1].toUpperCase()] : ""; };
  const snNorm = s => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  const cmpDt = (a, b) => String(b.data_arrivo || "").localeCompare(String(a.data_arrivo || ""));
  let all = [];

  async function load() { all = await getAll("bolle"); }
  const pending = () => all.filter(b => b.stato !== "montato");
  const edName = uid => { const e = edifici.find(x => x.uid === uid); return e ? e.nome : "?"; };
  const placeOf = b => `${edName(b.edificio_uid)} · ${floorLabel(b.piano)}`;
  const label = b => b.numero ? `bolla ${b.numero}` : b.codice ? `codice ${b.codice}` : "bolla";
  const floorsOf = e => {
    const set = new Set();
    for (let f = Math.max(e.piano_min, e.piano_max); f >= Math.min(e.piano_min, e.piano_max); f--) set.add(f);
    stampanti.filter(s => s.edificio_uid === e.uid).forEach(s => set.add(s.piano));
    return [...set].sort((a, c) => c - a);
  };
  const prnOption = s => `<option value="${esc(s.uid)}">${esc(printerName(s))}${s.modello ? " · " + esc(s.modello) : ""}${s.seriale ? " · " + esc(s.seriale) : ""}</option>`;
  const colChips = (id, on) => `<div class="chips" id="${id}">${COLORI.map(c => `<button type="button" data-c="${c[0]}"${c[0] === on ? ' class="on"' : ""}><span class="sw" style="background:${c[1]}"></span>${c[0]}</button>`).join("")}</div>`;
  const tag = b => b.stato === "montato" ? `<span class="tag bolOk">Montato</span>` : `<span class="tag bolWait">Da montare</span>`;
  function chipsPick(id, onChange) {
    let cur = (document.querySelector(`#${id} button.on`) || {}).dataset;
    cur = cur ? cur.c : "";
    document.querySelectorAll(`#${id} button`).forEach(b => b.onclick = () => {
      cur = b.dataset.c;
      document.querySelectorAll(`#${id} button`).forEach(x => x.classList.toggle("on", x === b));
      if (onChange) onChange(cur);
    });
    return { get: () => cur, set: c => { cur = c; document.querySelectorAll(`#${id} button`).forEach(x => x.classList.toggle("on", x.dataset.c === c)); } };
  }

  /* ---------------- numeri per la scorta e per la vista a piani ---------------- */
  /* toner arrivati sul telefono dopo l'ultimo invio di scorta dal PC: si sommano al magazzino */
  function arrivalsByKey(agg) {
    const m = {};
    all.forEach(b => {
      if (b.daPC || !b.modello || !((b.arrivo_mod || "") > (agg || ""))) return;
      const k = scKey(b.modello, b.colore); m[k] = (m[k] || 0) + 1;
    });
    return m;
  }
  const pendingAt = (euid, piano) => pending().filter(b => b.edificio_uid === euid && b.piano === piano);
  function badge(euid, piano) {
    const n = pendingAt(euid, piano).length;
    return n ? `<span class="bolB" title="${n} toner arrivati, ancora da montare">📦${n}</span>` : "";
  }
  const countText = list => COLORI.map(c => [c[0], list.filter(b => b.colore === c[0]).length]).filter(x => x[1]).map(x => `${x[0]} ×${x[1]}`).join(" · ");

  /* ---------------- elenco (Stampanti › Cambi toner) ---------------- */
  function renderList(el) {
    if (!el) return;
    const p = pending().sort(cmpDt), done = all.filter(b => b.stato === "montato").sort((a, c) => String(c.data_montaggio || "").localeCompare(String(a.data_montaggio || "")));
    const line = b => `<div class="card" data-b="${esc(b.uid)}"><h3><span class="sw" style="background:${colHex(b.colore)}"></span>${esc(b.colore)} ${tag(b)}</h3><div class="meta">${esc(placeOf(b))} · arrivato ${fmtDt(b.data_arrivo)}</div><div class="meta">${esc(label(b))}${b.stato === "montato" ? " · montato " + fmtDt(b.data_montaggio) : ""}</div></div>`;
    el.innerHTML = `<div class="card bolBox"><h3>📦 Toner in arrivo (bolle)</h3>
      <div class="meta">${p.length ? `Da montare: <b>${p.length}</b>` : "Nessun toner da montare."}</div>
      <button class="btn green" id="bolNew" type="button">📦 Nuova bolla (toner arrivato)</button>
      <div id="bolPend">${p.map(line).join("")}</div>
      ${done.length ? `<details class="flt"><summary>Montati di recente (${Math.min(done.length, 8)})</summary><div>${done.slice(0, 8).map(line).join("")}</div></details>` : ""}</div>`;
    $("bolNew").onclick = () => openNew();
    el.querySelectorAll("[data-b]").forEach(c => c.onclick = () => open(c.dataset.b));
  }
  /* ---------------- inventario per piano (Stampanti › Scorta) ---------------- */
  function renderStock(el) {
    if (!el) return;
    const p = pending();
    if (!p.length) { el.innerHTML = `<div class="card"><h3>📦 Toner nei piani</h3><div class="hint">Nessun toner da montare nei piani. Quando registri una bolla, il toner compare qui diviso per piano.</div></div>`; return; }
    const groups = new Map();
    p.forEach(b => { const k = b.edificio_uid + "|" + b.piano; if (!groups.has(k)) groups.set(k, []); groups.get(k).push(b); });
    const rows = [...groups.values()].sort((a, c) => edName(a[0].edificio_uid).localeCompare(edName(c[0].edificio_uid), "it") || c[0].piano - a[0].piano);
    el.innerHTML = `<div class="card"><h3>📦 Toner nei piani (da montare): ${p.length}</h3>` + rows.map(g => `<div class="bolFloor"><b>${esc(placeOf(g[0]))}</b> · ${g.length}<div class="meta">${esc(countText(g))}</div>${g.sort(cmpDt).map(b => `<div class="bolRow" data-b="${esc(b.uid)}"><span class="sw" style="background:${colHex(b.colore)}"></span>${esc(b.colore)} · ${esc(label(b))} · ${fmtDt(b.data_arrivo)}</div>`).join("")}</div>`).join("") + `</div>`;
    el.querySelectorAll("[data-b]").forEach(c => c.onclick = () => open(c.dataset.b));
  }

  /* ---------------- arrivo: nuova bolla ---------------- */
  function openNew() {
    const m = $("modal"), last = all.slice().sort(cmpDt)[0];
    if (!edifici.length) { m.innerHTML = `<div class="card full"><h3>📦 Toner arrivato</h3><div class="hint">Non ci sono ancora edifici e stampanti sul telefono. Nel programma del PC usa File › Esporta dati per il telefono, poi qui in Scambio › Ricevi dal PC.</div><button class="btn" id="bnClose" type="button">Chiudi</button></div>`; m.classList.remove("hidden"); $("bnClose").onclick = () => m.classList.add("hidden"); return; }
    let photo = "", dirty = false;
    m.innerHTML = `<div class="card full"><h3>📦 Toner arrivato (bolla)</h3>
      <div class="hint">Inquadra il codice a barre in alto sulla bolla (tasto 📷) oppure scatta una foto della bolla: se il codice si legge, lo scrivo io. Poi scegli colore e piano: il toner entra nell'inventario di quel piano.</div>
      <label for="bnCode">Codice a barre della bolla</label><input id="bnCode" data-scan="any" autocapitalize="characters" autocomplete="off" placeholder="es. LE:920245766">
      <label for="bnNum">Numero della bolla (Delivery Note, facoltativo)</label><input id="bnNum" inputmode="numeric" autocomplete="off" placeholder="es. 0024953032">
      <button class="btn navy" id="bnPhotoBtn" type="button">📷 Scatta la foto della bolla</button>
      <input type="file" id="bnFile" accept="image/*" capture="environment" style="display:none"><div id="bnPrev"></div>
      <label for="bnTn">Codice del toner (facoltativo, es. TNP79M)</label><input id="bnTn" autocapitalize="characters" autocomplete="off">
      <label>Colore del toner</label>${colChips("bnCols", "")}
      <div class="row"><div><label for="bnEd">Edificio</label><select id="bnEd"></select></div><div><label for="bnPi">Piano</label><select id="bnPi"></select></div></div>
      <label for="bnPr">Stampante a cui serve (se la sai)</label><select id="bnPr"></select>
      <label for="bnSn">Seriale scritto sulla bolla (facoltativo)</label><input id="bnSn" data-scan="serial" autocapitalize="characters" autocomplete="off">
      <label for="bnDt">Arrivato il</label><input id="bnDt" type="datetime-local" value="${localInput()}">
      <label for="bnNote">Note</label><input id="bnNote">
      <button class="btn green" id="bnSave" type="button">Metti nell'inventario del piano</button>
      <button class="btn" id="bnClose" type="button">Chiudi</button><div id="bnMsg"></div></div>`;
    m.classList.remove("hidden");
    const cols = chipsPick("bnCols");
    const fillFloors = (euid, piano) => {
      const e = edifici.find(x => x.uid === euid) || edifici[0], fl = floorsOf(e);
      $("bnPi").innerHTML = fl.map(f => `<option value="${f}">${esc(floorLabel(f))}</option>`).join("");
      $("bnPi").value = String(fl.includes(piano) ? piano : fl[0]);
      fillPrinters();
    };
    const fillPrinters = (sel) => {
      const ps = stampanti.filter(s => s.edificio_uid === $("bnEd").value && String(s.piano) === $("bnPi").value);
      $("bnPr").innerHTML = `<option value="">(non so / lo scelgo al montaggio)</option>` + ps.map(prnOption).join("");
      if (sel) $("bnPr").value = sel;
    };
    $("bnEd").innerHTML = edifici.map(e => `<option value="${esc(e.uid)}">${esc(e.nome)}</option>`).join("");
    $("bnEd").value = last && edifici.some(e => e.uid === last.edificio_uid) ? last.edificio_uid : edifici[0].uid;
    fillFloors($("bnEd").value, last ? last.piano : null);
    $("bnEd").onchange = () => fillFloors($("bnEd").value, null);
    $("bnPi").onchange = () => fillPrinters();
    $("bnDt").addEventListener("input", () => { dirty = true; });
    $("bnTn").addEventListener("input", () => { const c = colorOfCode($("bnTn").value); if (c) cols.set(c); });
    const hint = (t, err) => show($("bnMsg"), t, err);
    /* il seriale scritto sulla bolla riconosce la stampante (e quindi piano ed edificio) */
    const bySerial = sn => { const k = snNorm(sn); return k ? stampanti.find(s => snNorm(s.seriale) === k) : null; };
    const onSn = () => {
      const s = bySerial($("bnSn").value);
      if (!s) { if ($("bnSn").value.trim()) hint("Nessuna stampante ha questo seriale: scegli edificio e piano a mano."); return; }
      $("bnEd").value = s.edificio_uid; fillFloors(s.edificio_uid, s.piano); fillPrinters(s.uid);
      hint("Stampante riconosciuta: " + prnLabel(s) + ".");
    };
    $("bnSn").addEventListener("change", onSn);
    $("bnPhotoBtn").onclick = () => { $("bnFile").value = ""; $("bnFile").click(); };
    $("bnFile").onchange = async e => {
      const f = e.target.files && e.target.files[0]; if (!f) return;
      hint("Leggo la foto…");
      photo = await shrink(f, 1600);
      $("bnPrev").innerHTML = `<img src="${photo}" alt="Foto della bolla" style="margin-top:8px">`;
      let got = null;
      if (!$("bnCode").value.trim() && window.SCAN) { try { got = await SCAN.decodeFile(f, "any"); } catch (er) { got = null; } }
      if (got && got.text) { $("bnCode").value = got.text; hint("Foto salvata. Ho letto il codice a barre: " + got.text); }
      else hint($("bnCode").value.trim() ? "Foto salvata." : "Foto salvata. Il codice a barre non l'ho letto: usa il tasto 📷 accanto al codice oppure scrivilo.");
    };
    $("bnClose").onclick = () => m.classList.add("hidden");
    $("bnSave").onclick = async () => {
      const code = $("bnCode").value.trim(), num = $("bnNum").value.trim(), col = cols.get(), euid = $("bnEd").value, piano = parseInt($("bnPi").value, 10);
      if (!code && !num && !photo) return hint("Serve almeno il codice a barre, il numero o la foto della bolla.", true);
      if (!col) return hint("Tocca il colore del toner (o scrivi il codice del toner).", true);
      if (!euid || Number.isNaN(piano)) return hint("Scegli edificio e piano.", true);
      if (!$("bnDt").value) return hint("Controlla data e ora.", true);
      const twin = all.find(b => (code && b.codice === code) || (num && b.numero === num));
      if (twin && !confirm(`Questa bolla è già registrata (${fmtDt(twin.data_arrivo)}, ${twin.colore}). La registro di nuovo?`)) return;
      const pr = stampanti.find(s => s.uid === $("bnPr").value) || null;
      const dt = dirty ? toDataOra($("bnDt").value) : toDataOra(localInput());
      const rec = { uid: uuid(), codice: code, numero: num, toner_codice: $("bnTn").value.trim().toUpperCase(), colore: col, edificio_uid: euid, piano, stampante_uid: pr ? pr.uid : "", modello: pr ? (pr.modello || "") : "", seriale_bolla: $("bnSn").value.trim().toUpperCase(), data_arrivo: dt, arrivo_mod: nowIso(), stato: "da_montare", data_montaggio: "", stampante_montata_uid: "", seriale_letto: "", toner_uid: "", foto: photo, note: $("bnNote").value.trim(), modificato: nowIso() };
      await put("bolle", rec);
      await load(); if (typeof renderPrn === "function") renderPrn();
      m.classList.add("hidden");
      open(rec.uid, null, `Registrato: ${col} · ${placeOf(rec)}. Ora è nell'inventario del piano, da montare.`);
    };
  }

  /* ---------------- scheda della bolla e montaggio ---------------- */
  function open(uid, presel, msg, err) {
    const b = all.find(x => x.uid === uid); if (!b) return;
    const m = $("modal"), done = b.stato === "montato";
    const intended = stampanti.find(s => s.uid === b.stampante_uid) || null;
    const mounted = stampanti.find(s => s.uid === b.stampante_montata_uid) || null;
    const expectSn = snNorm(b.seriale_bolla) || (intended ? snNorm(intended.seriale) : "");
    const sorted = stampanti.slice().sort((a, c) => (c.edificio_uid === b.edificio_uid && c.piano === b.piano) - (a.edificio_uid === b.edificio_uid && a.piano === b.piano) || prnLabel(a).localeCompare(prnLabel(c), "it"));
    m.innerHTML = `<div class="card full"><h3><span class="sw" style="background:${colHex(b.colore)}"></span>${esc(b.colore)} ${tag(b)}</h3>
      <div class="meta">${esc(placeOf(b))} · arrivato ${fmtDt(b.data_arrivo)}</div>
      <div class="meta">${b.numero ? "Bolla n. <b>" + esc(b.numero) + "</b>" : ""}${b.numero && b.codice ? " · " : ""}${b.codice ? "codice <b>" + esc(b.codice) + "</b>" : ""}${b.toner_codice ? " · toner " + esc(b.toner_codice) : ""}</div>
      ${b.seriale_bolla ? `<div class="meta">Seriale sulla bolla: <b>${esc(b.seriale_bolla)}</b></div>` : ""}${intended ? `<div class="meta">Stampante indicata: ${esc(prnLabel(intended))}</div>` : ""}${b.note ? `<div class="meta">Note: ${esc(b.note)}</div>` : ""}
      ${b.foto ? `<img src="${b.foto}" alt="Foto della bolla" style="margin-top:8px">` : ""}
      ${done ? `<label>Montato</label><div class="meta">${fmtDt(b.data_montaggio)} · ${mounted ? esc(prnLabel(mounted)) : "stampante non più presente"}${b.seriale_letto ? " · seriale letto " + esc(b.seriale_letto) : ""}</div>
        <button class="btn navy" id="bmUndo" type="button">↩ Annulla il montaggio</button>`
      : `<label for="bmSn">Seriale della stampante (QR o codice a barre)</label><input id="bmSn" data-scan="serial" autocapitalize="characters" autocomplete="off" placeholder="inquadra l'etichetta della stampante">
        <label for="bmPr">oppure scegli la stampante</label><select id="bmPr"><option value="">(scegli)</option>${sorted.map(s => `<option value="${esc(s.uid)}">${esc(prnLabel(s))}${s.seriale ? " · " + esc(s.seriale) : ""}</option>`).join("")}</select>
        <div id="bmCheck"></div>
        <label>Colore del toner</label>${colChips("bmCols", b.colore)}
        <label for="bmDt">Montato il (se non lo tocchi, uso l'ora di adesso)</label><input id="bmDt" type="datetime-local" value="${localInput()}">
        <button class="btn green" id="bmGo" type="button">✅ Montato adesso</button>`}
      <button class="btn red sm" id="bmDel" type="button" style="margin-top:10px">Elimina la bolla</button>
      <button class="btn" id="bmClose" type="button">Chiudi</button><div id="bmMsg"></div></div>`;
    m.classList.remove("hidden");
    if (msg) show($("bmMsg"), msg, err);
    $("bmClose").onclick = () => m.classList.add("hidden");
    $("bmDel").onclick = async () => {
      if (!confirm(done ? "Eliminare questa bolla? Il cambio toner già registrato resta nello storico." : "Eliminare questa bolla? Il toner esce dall'inventario del piano.")) return;
      await del("bolle", b.uid); await load(); renderPrn(); m.classList.add("hidden");
    };
    if (done) {
      $("bmUndo").onclick = async () => {
        if (!confirm("Annullo il montaggio? Il cambio toner collegato viene tolto dallo storico e il toner torna da montare.\nSe avevi già mandato i dati al PC, togli il cambio anche lì.")) return;
        if (b.toner_uid) await del("toner", b.toner_uid);
        await put("bolle", { ...b, stato: "da_montare", data_montaggio: "", stampante_montata_uid: "", seriale_letto: "", toner_uid: "", modificato: nowIso() });
        await load(); await loadPrn(); renderPrn(); open(uid, null, "Montaggio annullato: il toner è di nuovo da montare.");
      };
      return;
    }
    const cols = chipsPick("bmCols");
    let dirty = false;
    $("bmDt").addEventListener("input", () => { dirty = true; });
    /* controllo fra la bolla e la stampante scelta */
    const state = { pr: null, sn: "", bad: false };
    const check = () => {
      const pr = stampanti.find(s => s.uid === $("bmPr").value) || null, read = snNorm($("bmSn").value), actual = read || (pr ? snNorm(pr.seriale) : "");
      state.pr = pr; state.sn = read; state.bad = false;
      let html = "";
      if (read && !stampanti.some(s => snNorm(s.seriale) === read) && !pr) html = `<div class="msg">Ho letto il seriale ${esc(read)}, ma nessuna stampante lo ha. Scegli la stampante dalla lista.</div>`;
      else if (pr && expectSn && actual) {
        if (actual === expectSn) html = `<div class="msg">✔ Stampante giusta: il seriale coincide con la bolla (${esc(actual)}).</div>`;
        else { state.bad = true; html = `<div class="msg err">⚠ Il seriale non coincide: la bolla è per ${esc(b.seriale_bolla || (intended && intended.seriale) || expectSn)}, questa stampante è ${esc(actual)}.</div>`; }
      } else if (pr && expectSn) html = `<div class="msg">Di questa stampante non so il seriale: non posso controllarlo con la bolla. Inquadra l'etichetta oppure procedi.</div>`;
      else if (pr) html = pr.edificio_uid === b.edificio_uid && pr.piano === b.piano ? `<div class="msg">✔ Stesso piano della bolla (${esc(floorLabel(pr.piano))}).</div>` : (state.bad = true, `<div class="msg err">⚠ Questa stampante è in ${esc(prnLabel(pr))}, ma il toner è arrivato a ${esc(placeOf(b))}.</div>`);
      $("bmCheck").innerHTML = html;
    };
    const onSn = () => {
      const k = snNorm($("bmSn").value); if (!k) return check();
      const s = stampanti.find(x => snNorm(x.seriale) === k) || (expectSn === k ? intended : null);
      if (s) $("bmPr").value = s.uid;
      check();
    };
    $("bmSn").addEventListener("change", onSn);
    $("bmPr").onchange = check;
    if (presel && stampanti.some(s => s.uid === presel)) $("bmPr").value = presel;
    check();
    $("bmGo").onclick = async () => {
      if (!state.pr) return show($("bmMsg"), "Inquadra il seriale della stampante o sceglila dalla lista.", true);
      const col = cols.get(); if (!col) return show($("bmMsg"), "Tocca il colore del toner.", true);
      if (!$("bmDt").value) return show($("bmMsg"), "Controlla data e ora.", true);
      if (state.bad && !confirm("Attenzione: la stampante non corrisponde alla bolla. Lo registro lo stesso?")) return;
      const dt = dirty ? toDataOra($("bmDt").value) : toDataOra(localInput());
      const note = `Bolla ${b.numero || b.codice || ""}${b.numero && b.codice ? " (" + b.codice + ")" : ""}${state.bad ? " - stampante diversa da quella della bolla" : ""}`.trim();
      const tuid = uuid();
      await put("toner", { uid: tuid, stampante_uid: state.pr.uid, colore: col, data_ora: dt, note, bolla_uid: b.uid, modificato: nowIso() });
      await put("bolle", { ...b, colore: col, modello: b.modello || state.pr.modello || "", stato: "montato", data_montaggio: dt, stampante_montata_uid: state.pr.uid, seriale_letto: state.sn ? state.sn.toUpperCase() : "", toner_uid: tuid, modificato: nowIso() });
      await load(); await loadPrn(); renderPrn();
      open(uid, null, `Registrato: ${col} montato il ${fmtDt(dt)} su ${printerName(state.pr)}. Il toner è uscito dall'inventario del piano.`);
    };
  }

  /* ---------------- scambio col PC ---------------- */
  function exportItems(mode, last) {
    return all.filter(b => mode === "all" || (b.modificato || "") > (last || "")).filter(b => !(b.daPC && b.modificato === b.daPC)).map(b => ({
      uid: b.uid, codice: b.codice || "", numero: b.numero || "", toner_codice: b.toner_codice || "", colore: b.colore, edificio_uid: b.edificio_uid, piano: b.piano,
      stampante_uid: b.stampante_uid || "", modello: b.modello || "", seriale_bolla: b.seriale_bolla || "", data_arrivo: b.data_arrivo, arrivo_mod: b.arrivo_mod || "", stato: b.stato,
      data_montaggio: b.data_montaggio || "", stampante_montata_uid: b.stampante_montata_uid || "", seriale_letto: b.seriale_letto || "", toner_uid: b.toner_uid || "",
      note: b.note || "", foto: b.foto || "", modificato: b.modificato }));
  }
  async function merge(data, fromPc) {
    const out = { n: 0, a: 0, p: 0, s: 0 };
    for (const x of data.bolle || []) {
      if (!x.uid || !x.edificio_uid || !COLORI.some(c => c[0] === x.colore) || !/^\d{4}-\d\d-\d\d \d\d:\d\d$/.test(x.data_arrivo || "")) { out.s++; continue; }
      const cur = await getKey("bolle", x.uid);
      if (cur && !((x.modificato || "") > (cur.modificato || ""))) { out.p++; continue; }
      const rec = { uid: String(x.uid), codice: String(x.codice || ""), numero: String(x.numero || ""), toner_codice: String(x.toner_codice || ""), colore: x.colore, edificio_uid: String(x.edificio_uid), piano: toInt(x.piano),
        stampante_uid: String(x.stampante_uid || ""), modello: String(x.modello || ""), seriale_bolla: String(x.seriale_bolla || ""), data_arrivo: x.data_arrivo, arrivo_mod: String(x.arrivo_mod || x.modificato || ""),
        stato: x.stato === "montato" ? "montato" : "da_montare", data_montaggio: String(x.data_montaggio || ""), stampante_montata_uid: String(x.stampante_montata_uid || ""), seriale_letto: String(x.seriale_letto || ""),
        toner_uid: String(x.toner_uid || ""), foto: typeof x.foto === "string" && x.foto.startsWith("data:") ? x.foto : (cur ? cur.foto || "" : ""), note: String(x.note || ""), modificato: x.modificato || nowIso() };
      if (fromPc) rec.daPC = rec.modificato;
      await put("bolle", rec);
      cur ? out.a++ : out.n++;
    }
    await load();
    return out;
  }
  const exportCount = (mode, last) => exportItems(mode, last).length;

  return { load, renderList, renderStock, openNew, open, arrivalsByKey, badge, pendingAt, pending: () => pending(), exportItems, exportCount, merge, colorOfCode, snNorm, _all: () => all };
})();
