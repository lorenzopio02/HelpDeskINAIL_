/* Piani e stanze: stanze di ogni piano, persone alle postazioni, foto e panorami, pianta del piano.
   Si usa da Stampanti › Piani e stanze, dal piano nella vista Edificio e dalla scheda utente.
   Le postazioni le numera l'app (Postazione 1, 2, 3...): chi usa l'app non le sceglie mai.
   Lo scambio col PC (file JSON) porta stanze, piante, foto e postazioni; vince la modifica più recente. */
"use strict";
window.ROOMS = (function () {
  const CODE_MAX = 24, FOTO_PX = 1600, PANO_PX = 2600, PIANTA_PX = 1600;
  const coll = new Intl.Collator("it", { numeric: true, sensitivity: "base" });
  const cmp = (a, b) => coll.compare(a, b);
  const cf = s => String(s || "").toLowerCase();
  const cleanCode = s => String(s == null ? "" : s).split(/\s+/).filter(Boolean).join(" ").slice(0, CODE_MAX);
  const pianteUid = (euid, piano) => `pi:${euid}:${piano}`;
  const num = v => { const n = parseInt(v, 10); return Number.isNaN(n) ? 0 : Math.max(0, n); };

  /* un codice per riga (vanno bene anche virgole e punti e virgola) */
  function parseCodes(text) {
    const seen = new Set(), good = [], bad = [];
    for (const raw of String(text || "").split(/[\r\n;,\t]+/)) {
      const c = raw.split(/\s+/).filter(Boolean).join(" ");
      if (!c) continue;
      if (c.length > CODE_MAX) bad.push(c);
      else if (!seen.has(cf(c))) { seen.add(cf(c)); good.push(c); }
    }
    return { good, bad };
  }

  /* ---------- dati ---------- */
  let S = [], P = [];   // stanze e postazioni (piccole: si rileggono sempre)
  async function load() { S = await getAll("stanze"); P = await getAll("postazioni"); }
  const room = uid => S.find(r => r.uid === uid && !r.eliminata) || null;
  const roomsOf = (euid, piano) => S.filter(r => !r.eliminata && r.edificio_uid === euid && r.piano === piano).sort((a, b) => cmp(a.codice, b.codice));
  const peopleIn = ruid => P.filter(p => p.stanza_uid === ruid && room(ruid)).sort((a, b) => (a.numero - b.numero) || cmp(a.utente, b.utente));
  const seatOf = k => { const p = P.find(x => x.k === k && x.stanza_uid); return p && room(p.stanza_uid) ? p : null; };
  const userOf = k => users.find(u => u.key === k) || null;
  const bName = euid => (edifici.find(x => x.uid === euid) || {}).nome || "?";
  const fullName = r => `${bName(r.edificio_uid)} · ${floorLabel(r.piano)} · ${r.codice}`;
  function occ(r) { const n = peopleIn(r.uid).length, tot = num(r.scrivanie); return { n, tot, txt: tot > 0 ? `occupate ${n}/${tot}` : "", over: tot > 0 && n > tot }; }
  const lab = r => `${r.codice} · ${occ(r).n}`;
  const dispCmp = (a, b) => cmp(a.utente, b.utente);

  async function renumberRoom(ruid, now, bump) {
    const rows = P.filter(p => p.stanza_uid === ruid).sort((a, b) => (a.numero - b.numero) || dispCmp(a, b));
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].numero !== i + 1) { const nr = { ...rows[i], numero: i + 1 }; if (bump) nr.modificato = now; await put("postazioni", nr); }
    }
    await load();
  }

  /* ---------- operazioni sulle stanze ---------- */
  async function addRoom(euid, piano, codice, scrivanie) {
    await load();
    const c = cleanCode(codice);
    if (!c) return { err: "Scrivi il codice della stanza." };
    if (roomsOf(euid, piano).some(r => cf(r.codice) === cf(c))) return { err: `Su questo piano la stanza ${c} esiste già.` };
    const r = { uid: "ro-" + uuid(), edificio_uid: euid, piano, codice: c, scrivanie: num(scrivanie), note: "", come: "", px: null, py: null, eliminata: false, modificato: nowIso() };
    await put("stanze", r); await load();
    return { ok: true, r };
  }
  async function addRooms(euid, piano, text) {
    await load();
    const { good, bad } = parseCodes(text), have = new Set(roomsOf(euid, piano).map(r => cf(r.codice)));
    let nuove = 0, esistenti = 0;
    for (const c of good) {
      if (have.has(cf(c))) { esistenti++; continue; }
      await put("stanze", { uid: "ro-" + uuid(), edificio_uid: euid, piano, codice: c, scrivanie: 0, note: "", come: "", px: null, py: null, eliminata: false, modificato: nowIso() });
      nuove++;
    }
    await load();
    return { nuove, esistenti, scartate: bad };
  }
  async function saveRoom(uid, f) {
    await load();
    const r = room(uid); if (!r) return { err: "Questa stanza non c'è più." };
    const c = cleanCode(f.codice);
    if (!c) return { err: "Scrivi il codice della stanza." };
    if (roomsOf(r.edificio_uid, r.piano).some(x => x.uid !== uid && cf(x.codice) === cf(c))) return { err: `Su questo piano la stanza ${c} esiste già.` };
    const nr = { ...r, codice: c, scrivanie: num(f.scrivanie), note: String(f.note || "").trim(), come: String(f.come || "").trim(), modificato: nowIso() };
    await put("stanze", nr); await load();
    return { ok: true, r: nr };
  }
  async function setPos(uid, px, py) {
    await load();
    const r = room(uid); if (!r) return false;
    const c = v => v == null ? null : Math.min(1, Math.max(0, Math.round(v * 10000) / 10000));
    await put("stanze", { ...r, px: c(px), py: c(py), modificato: nowIso() }); await load();
    return true;
  }
  async function deleteRoom(uid) {
    await load();
    const r = S.find(x => x.uid === uid); if (!r || r.eliminata) return false;
    const now = nowIso();
    for (const p of P.filter(x => x.stanza_uid === uid)) await put("postazioni", { ...p, stanza_uid: "", numero: 0, modificato: now });
    for (const f of await req(tx("stanzefoto").index("stanza_uid").getAll(uid))) if (!f.eliminata) await put("stanzefoto", { ...f, data: "", eliminata: true, modificato: now });
    await put("stanze", { ...r, eliminata: true, modificato: now });
    await load();
    return true;
  }

  /* ---------- persone e postazioni ---------- */
  /* mette la persona alla prima postazione libera; se sedeva altrove la toglie da lì (e rinumera) */
  async function assign(ukey, ruid) {
    await load();
    const r = room(ruid), u = userOf(ukey);
    if (!r) throw new Error("Questa stanza non c'è più.");
    if (!u) throw new Error("Questa persona non c'è più.");
    const old = seatOf(ukey), now = nowIso();
    if (old && old.stanza_uid === ruid) return { numero: old.numero, gia: true, da: null, sopra: false };
    let da = null;
    if (old) {
      da = room(old.stanza_uid);
      await put("postazioni", { ...old, stanza_uid: "", numero: 0, modificato: now });
      await load(); await renumberRoom(old.stanza_uid, now, true);
    }
    const n = P.filter(p => p.stanza_uid === ruid).length + 1;
    await put("postazioni", { k: ukey, utente: u.nome, stanza_uid: ruid, numero: n, modificato: now });
    await load();
    const tot = num(r.scrivanie);
    return { numero: n, gia: false, da, sopra: !!(tot && n > tot) };
  }
  async function unassign(ukey) {
    await load();
    const old = seatOf(ukey); if (!old) return false;
    const now = nowIso();
    await put("postazioni", { ...old, stanza_uid: "", numero: 0, modificato: now });
    await load(); await renumberRoom(old.stanza_uid, now, true);
    return true;
  }
  /* utente eliminato dal telefono (vale solo qui): sparisce anche la sua postazione, senza segnalarlo al PC */
  async function dropUser(ukey) {
    await load();
    const p = P.find(x => x.k === ukey); if (!p) return;
    await del("postazioni", ukey); await load();
    if (p.stanza_uid) await renumberRoom(p.stanza_uid, nowIso(), false);
  }
  /* utente rinominato: la postazione passa al nuovo nome (e il vecchio nome risulta libero) */
  async function renameUser(oldKey, newKey, nome) {
    await load();
    const p = P.find(x => x.k === oldKey); if (!p || oldKey === newKey) return;
    const now = nowIso();
    await put("postazioni", { k: oldKey, utente: p.utente, stanza_uid: "", numero: 0, modificato: now });
    if (p.stanza_uid) await put("postazioni", { k: newKey, utente: nome, stanza_uid: p.stanza_uid, numero: p.numero, modificato: now });
    await load();
  }

  /* ---------- foto, panorami e piante ---------- */
  const fotosOf = async ruid => (await req(tx("stanzefoto").index("stanza_uid").getAll(ruid))).filter(f => !f.eliminata).sort((a, b) => (a.modificato || "").localeCompare(b.modificato || ""));
  async function addFoto(ruid, file, tipo) {
    const data = await shrink(file, tipo === "panorama" ? PANO_PX : FOTO_PX);
    const f = { uid: "sf-" + uuid(), stanza_uid: ruid, nome: `${tipo}_${localInput().replace(/[-:T]/g, "")}_${Math.random().toString(36).slice(2, 6)}.jpg`, tipo, data, eliminata: false, modificato: nowIso() };
    await put("stanzefoto", f);
    return f;
  }
  async function delFoto(uid) { const f = await getKey("stanzefoto", uid); if (f && !f.eliminata) await put("stanzefoto", { ...f, data: "", eliminata: true, modificato: nowIso() }); }
  async function setFotoTipo(uid, tipo) { const f = await getKey("stanzefoto", uid); if (f) await put("stanzefoto", { ...f, tipo, modificato: nowIso() }); }
  const pianta = (euid, piano) => getKey("piante", pianteUid(euid, piano));
  async function setPianta(euid, piano, file) {
    const data = await shrink(file, PIANTA_PX);
    await put("piante", { uid: pianteUid(euid, piano), edificio_uid: euid, piano, foto: data, modificato: nowIso() });
  }
  const delPianta = (euid, piano) => del("piante", pianteUid(euid, piano));

  /* ---------- scambio col PC ---------- */
  const isN = (r, last) => (!last || (r.modificato || "") > last) && r.modificato !== r.daPC;
  async function exportItems(mode, last) {
    await load();
    const all = mode === "all";
    const stanze = S.filter(r => all || isN(r, last)).map(r => ({ uid: r.uid, edificio_uid: r.edificio_uid, piano: r.piano, codice: r.codice, scrivanie: num(r.scrivanie), note: r.note || "", come: r.come || "", px: r.px == null ? null : r.px, py: r.py == null ? null : r.py, eliminata: !!r.eliminata, modificato: r.modificato }));
    const stanze_foto = (await getAll("stanzefoto")).filter(f => all || isN(f, last)).map(f => ({ uid: f.uid, stanza_uid: f.stanza_uid, nome: f.nome, tipo: f.tipo === "panorama" ? "panorama" : "foto", data: f.eliminata ? "" : (f.data || ""), eliminata: !!f.eliminata, modificato: f.modificato }));
    const piante = (await getAll("piante")).filter(p => all || isN(p, last)).map(p => ({ uid: p.uid, edificio_uid: p.edificio_uid, piano: p.piano, foto: p.foto || "", modificato: p.modificato }));
    const postazioni = P.filter(p => all || isN(p, last)).map(p => ({ utente: p.utente, stanza_uid: p.stanza_uid || "", postazione: p.numero || 0, modificato: p.modificato }));
    return { stanze, stanze_foto, piante, postazioni };
  }
  async function countNew(last) {
    await load();
    return S.filter(r => isN(r, last)).length + P.filter(p => isN(p, last)).length + (await getAll("stanzefoto")).filter(f => isN(f, last)).length + (await getAll("piante")).filter(p => isN(p, last)).length;
  }
  /* stanze, piante, foto e postazioni arrivate dal PC (o da un ripristino): vince la modifica più recente.
     Se la stessa stanza (stesso edificio, piano e codice) è stata creata due volte con identificativi
     diversi, resta una sola stanza: quella già presente. */
  async function importData(data, fromPc) {
    await load();
    const res = { nuove: 0, aggiornate: 0, presenti: 0, saltate: 0, piante: 0, foto: 0, posti: 0, postiSaltati: 0 };
    const eds = new Set((await getAll("edifici")).map(e => e.uid)), us = await getAll("utenti"), roomOf = {}, touched = new Set();
    const tagPc = o => { if (fromPc) o.daPC = o.modificato; return o; };
    const fl = v => { if (v == null || v === "") return null; const n = Number(v); return Number.isFinite(n) ? Math.min(1, Math.max(0, n)) : null; };
    for (const it of data.stanze || []) {
      if (!it || typeof it !== "object") continue;
      const uid = String(it.uid || ""), euid = String(it.edificio_uid || ""), piano = Number.isInteger(+it.piano) && it.piano !== "" && it.piano !== null ? +it.piano : null, codice = cleanCode(it.codice);
      if (!uid || !eds.has(euid) || piano === null || !codice) { res.saltate++; continue; }
      const mod = String(it.modificato || nowIso()), gone = !!it.eliminata;
      let px = fl(it.px), py = fl(it.py); if (px === null || py === null) px = py = null;
      const f = { codice, scrivanie: num(it.scrivanie), note: String(it.note || ""), come: String(it.come || ""), px, py };
      let ex = S.find(r => r.uid === uid) || S.find(r => !r.eliminata && r.edificio_uid === euid && r.piano === piano && cf(r.codice) === cf(codice));
      if (!ex) {
        if (gone) continue;
        const nr = tagPc({ uid, edificio_uid: euid, piano, ...f, eliminata: false, modificato: mod });
        await put("stanze", nr); S.push(nr); roomOf[uid] = uid; res.nuove++;
      } else {
        if (!ex.eliminata) roomOf[uid] = ex.uid;
        if (mod > (ex.modificato || "")) {
          if (gone) {
            if (!ex.eliminata) {
              for (const p of P.filter(x => x.stanza_uid === ex.uid)) { await put("postazioni", tagPc({ ...p, stanza_uid: "", numero: 0, modificato: mod })); }
              for (const ph of await req(tx("stanzefoto").index("stanza_uid").getAll(ex.uid))) if (!ph.eliminata) await put("stanzefoto", tagPc({ ...ph, data: "", eliminata: true, modificato: mod }));
              await put("stanze", tagPc({ ...ex, eliminata: true, modificato: mod }));
            }
            delete roomOf[uid];
          } else {
            const other = S.find(r => !r.eliminata && r.uid !== ex.uid && r.edificio_uid === ex.edificio_uid && r.piano === ex.piano && cf(r.codice) === cf(codice));
            await put("stanze", tagPc({ ...ex, ...f, codice: other ? ex.codice : codice, eliminata: false, modificato: mod }));
            roomOf[uid] = ex.uid;
          }
          res.aggiornate++;
          await load();
        } else res.presenti++;
      }
    }
    await load();
    const roomId = uid => roomOf[uid] || (room(uid) ? uid : null);
    for (const it of data.piante || []) {
      if (!it || typeof it !== "object") continue;
      const euid = String(it.edificio_uid || ""), piano = +it.piano, payload = String(it.foto || ""), mod = String(it.modificato || nowIso());
      if (!eds.has(euid) || !Number.isInteger(piano) || payload.indexOf(",") < 0) continue;
      const cur = await pianta(euid, piano);
      if (cur && !(mod > (cur.modificato || ""))) continue;
      await put("piante", tagPc({ uid: pianteUid(euid, piano), edificio_uid: euid, piano, foto: payload, modificato: mod })); res.piante++;
    }
    for (const it of data.stanze_foto || []) {
      if (!it || typeof it !== "object") continue;
      const uid = String(it.uid || ""), payload = String(it.data || ""), mod = String(it.modificato || nowIso()), tipo = it.tipo === "panorama" ? "panorama" : "foto";
      const have = uid ? await getKey("stanzefoto", uid) : null;
      if (it.eliminata) { if (have && !have.eliminata) await put("stanzefoto", tagPc({ ...have, data: "", eliminata: true, modificato: mod })); continue; }
      const rid = roomId(String(it.stanza_uid || ""));
      if (!uid || !rid || payload.indexOf(",") < 0) continue;
      if (have) { if (!have.eliminata && tipo !== have.tipo && mod > (have.modificato || "")) await put("stanzefoto", tagPc({ ...have, tipo, modificato: mod })); continue; }
      await put("stanzefoto", tagPc({ uid, stanza_uid: rid, nome: String(it.nome || "foto.jpg"), tipo, data: payload, eliminata: false, modificato: mod })); res.foto++;
    }
    for (const it of data.postazioni || []) {
      if (!it || typeof it !== "object") continue;
      const nome = normName(String(it.utente || "")), k = keyOf(nome), suid = String(it.stanza_uid || ""), rid = suid ? roomId(suid) : null;
      if (!nome || !us.some(u => u.key === k) || (suid && !rid)) { res.postiSaltati++; continue; }
      const mod = String(it.modificato || nowIso()), old = P.find(p => p.k === k);
      if (old && (old.modificato || "") >= mod) continue;
      await put("postazioni", tagPc({ k, utente: us.find(u => u.key === k).nome, stanza_uid: rid || "", numero: rid ? num(it.postazione) : 0, modificato: mod }));
      if (old && old.stanza_uid) touched.add(old.stanza_uid);
      if (rid) touched.add(rid);
      res.posti++;
      await load();
    }
    for (const ruid of touched) await renumberRoom(ruid, "", false);
    await load();
    const parts = [];
    const n = (data.stanze || []).length;
    if (n) parts.push(`Stanze: ${res.nuove} nuove, ${res.aggiornate} aggiornate, ${res.presenti} già presenti` + (res.saltate ? `, ${res.saltate} ignorate` : ""));
    if (res.piante) parts.push(`Piante dei piani: ${res.piante}`);
    if (res.foto) parts.push(`Foto delle stanze: ${res.foto}`);
    if ((data.postazioni || []).length) parts.push(`Postazioni: ${res.posti} aggiornate` + (res.postiSaltati ? `, ${res.postiSaltati} ignorate (persona o stanza sconosciuta)` : ""));
    res.msg = parts.join("\n");
    return res;
  }

  /* ---------- interfaccia: sezione «Piani e stanze» ---------- */
  let curF = null, view = "", mark = false, markUid = null, zoom = 1, flash = null;
  const thisB = () => edifici.find(e => e.uid === bSel) || edifici[0] || null;
  function floorsOf(b) {
    const set = new Set();
    for (let f = Math.max(b.piano_min, b.piano_max); f >= Math.min(b.piano_min, b.piano_max); f--) set.add(f);
    stampanti.filter(s => s.edificio_uid === b.uid).forEach(s => set.add(s.piano));
    S.filter(r => !r.eliminata && r.edificio_uid === b.uid).forEach(r => set.add(r.piano));
    return [...set].sort((a, c) => c - a);
  }
  const phoneBits = u => [u && u.telefono ? "Tel " + esc(u.telefono) : "", u && u.seriale_pc ? "PC " + esc(u.seriale_pc) : "", u && u.imei ? "IMEI " + esc(u.imei) : ""].filter(Boolean).join(" · ");
  const note = (t, err) => { flash = { t, err: !!err }; };

  async function render() {
    const box = $("roomsPanel"); if (!box) return;
    await load();
    if (!edifici.length) { box.innerHTML = `<div class="card"><div class="hint">Non ci sono ancora edifici sul telefono.<br>Nel programma del PC usa File › Esporta dati per il telefono, poi qui in Scambio › Ricevi dal PC.</div></div>`; return; }
    const b = thisB(); bSel = b.uid;
    const floors = floorsOf(b);
    if (curF === null || !floors.includes(curF)) { curF = floors.includes(0) ? 0 : floors[0]; mark = false; markUid = null; }
    const rs = roomsOf(b.uid, curF), ppl = []; rs.forEach(r => peopleIn(r.uid).forEach(p => ppl.push({ r, p })));
    const pl = await pianta(b.uid, curF), hasPl = !!(pl && pl.foto);
    if (view !== "foto" && view !== "schema") view = hasPl ? "foto" : "schema";
    if (view === "foto" && !hasPl) view = "schema";
    const sc = scorteNow(), prs = stampanti.filter(s => s.edificio_uid === b.uid && s.piano === curF);
    let h = "";
    if (flash) { h += `<div class="msg${flash.err ? " err" : ""}" id="rsMsg">${esc(flash.t)}</div>`; flash = null; }
    h += edifici.length > 1 ? `<select id="rsB">${edifici.map(e => `<option value="${esc(e.uid)}"${e.uid === b.uid ? " selected" : ""}>${esc(e.nome)}${e.sede ? " (" + esc(e.sede) + ")" : ""}</option>`).join("")}</select>` : "";
    h += `<div class="chips" id="rsF">${floors.map(f => { const n = roomsOf(b.uid, f).length; return `<button type="button" data-f="${f}" class="${f === curF ? "on" : ""}">${floorLabel(f)}${n ? ` (${n})` : ""}</button>`; }).join("")}</div>`;
    h += `<div class="card"><h3>${esc(b.nome)} · ${floorLabel(curF)}</h3><div class="meta">${rs.length} ${rs.length === 1 ? "stanza" : "stanze"} · ${ppl.length} ${ppl.length === 1 ? "persona" : "persone"} · ${prs.length} ${prs.length === 1 ? "stampante" : "stampanti"}</div>
      <div class="row"><div><button class="btn green" id="rsNew" type="button">＋ Nuova stanza</button></div><div><button class="btn navy" id="rsBulk" type="button">Incolla più stanze</button></div></div></div>`;
    /* pianta */
    h += `<div class="card" id="rsPlan"><h3>Pianta del piano</h3>`;
    h += `<div class="seg" id="rsView"><button type="button" data-v="foto" class="${view === "foto" ? "on" : ""}"${hasPl ? "" : " disabled"}>Foto della pianta</button><button type="button" data-v="schema" class="${view === "schema" ? "on" : ""}">Schema</button></div>`;
    if (view === "foto") {
      h += `<div class="plw" id="plW"><div class="pli" id="plI" style="width:${Math.round(zoom * 100)}%"><img id="plImg" src="${esc(pl.foto)}" alt="Pianta del piano">${rs.filter(r => r.px != null && r.py != null).map(r => { const o = occ(r); return `<button type="button" class="pm${o.over ? " over" : ""}${mark && markUid === r.uid ? " sel" : ""}" data-r="${esc(r.uid)}" style="left:${(r.px * 100).toFixed(2)}%;top:${(r.py * 100).toFixed(2)}%">${esc(lab(r))}</button>`; }).join("")}</div></div>
        <div class="zr"><button class="btn sm" id="plMinus" type="button" aria-label="Riduci la pianta">−</button><span>Zoom <b id="plZ">${Math.round(zoom * 100)}%</b></span><button class="btn sm" id="plPlus" type="button" aria-label="Ingrandisci la pianta">＋</button></div>`;
    } else {
      h += rs.length ? `<div class="rgrid" id="rsGrid">${rs.map(r => { const o = occ(r); return `<button type="button" class="gt${o.over ? " over" : ""}" data-r="${esc(r.uid)}"><b>${esc(lab(r))}</b>${o.txt ? `<small>${o.txt}</small>` : ""}${r.px != null ? "" : ""}</button>`; }).join("")}</div>` : `<div class="hint">Nessuna stanza su questo piano: aggiungile con «Nuova stanza» o «Incolla più stanze».</div>`;
      if (!hasPl) h += `<div class="hint">Questa è la vista a schema (una casella per stanza). Se carichi la foto della pianta, potrai segnare le stanze direttamente sulla foto.</div>`;
    }
    h += `<input id="plFile" type="file" accept="image/*" class="hidden"><button class="btn" id="plUp" type="button">📷 ${hasPl ? "Sostituisci la foto della pianta" : "Carica la foto della pianta"}</button>`;
    if (hasPl) {
      h += `<button class="btn${mark ? " navy" : ""}" id="plMark" type="button">📍 ${mark ? "Finito di segnare" : "Segna le stanze sulla foto"}</button>`;
      if (mark) {
        h += `<div class="hint">Tocca una stanza qui sotto, poi tocca sulla foto il punto dove si trova.</div><div class="chips" id="plChips">${rs.map(r => `<button type="button" data-r="${esc(r.uid)}" class="${markUid === r.uid ? "on" : ""}">${esc(r.codice)}${r.px != null ? " ✓" : ""}</button>`).join("") || `<span class="meta">Prima aggiungi le stanze.</span>`}</div>`;
        const mr = markUid && room(markUid);
        if (mr && mr.px != null) h += `<button class="btn red" id="plUnmark" type="button">Togli il segno di ${esc(mr.codice)}</button>`;
      }
      h += `<button class="btn red" id="plDel" type="button">Togli la foto della pianta</button><div class="hint">La foto si toglie solo da questo telefono.</div>`;
    }
    h += `</div>`;
    /* stanze del piano */
    h += `<h2 style="margin-top:14px">Stanze</h2><div id="rsRooms">${rs.map(r => { const o = occ(r), nm = peopleIn(r.uid).map(p => p.utente.split(" ")[0]); return `<div class="card rc" data-r="${esc(r.uid)}"><h3>${esc(lab(r))}${o.txt ? ` <span class="tag${o.over ? " st-Aperto" : ""}">${o.txt}</span>` : ""}</h3>${r.come ? `<div class="meta">${esc(r.come.length > 90 ? r.come.slice(0, 90) + "…" : r.come)}</div>` : ""}${nm.length ? `<div class="meta">${esc(nm.join(", "))}</div>` : `<div class="meta">nessuno assegnato</div>`}${r.px != null ? "" : (hasPl ? `<div class="meta">non ancora segnata sulla pianta</div>` : "")}</div>`; }).join("") || `<div class="hint">Nessuna stanza.</div>`}</div>`;
    /* persone del piano */
    h += `<h2 style="margin-top:14px">Persone del piano</h2><div id="rsPpl">${ppl.map(({ r, p }) => `<div class="card rp" data-r="${esc(r.uid)}"><h3>${esc(p.utente)}</h3><div class="meta">${esc(r.codice)} · Postazione ${p.numero}</div>${phoneBits(userOf(p.k)) ? `<div class="meta">${phoneBits(userOf(p.k))}</div>` : ""}</div>`).join("") || `<div class="hint">Nessuno è assegnato a una stanza di questo piano.</div>`}</div>`;
    /* stampanti del piano, con i bollini */
    h += `<h2 style="margin-top:14px">Stampanti del piano</h2><div class="bld" style="margin-top:0"><div class="floor" style="border-top:0"><div class="ps" id="rsPrn">${prs.map(s => prCard(s, sc)).join("") || `<span class="meta" style="align-self:center">nessuna stampante</span>`}</div></div></div>`;
    h += `<div class="hint">Le stampanti sono del piano (non di una stanza). Rosso = ticket Konica aperto, giallo = scorta bassa.</div>`;
    box.innerHTML = h;
    wire(box, b, rs, pl);
  }

  function wire(box, b, rs, pl) {
    const again = () => render();
    if ($("rsB")) $("rsB").onchange = e => { bSel = e.target.value; curF = null; mark = false; markUid = null; zoom = 1; again(); };
    $("rsF").querySelectorAll("button").forEach(x => x.onclick = () => { curF = +x.dataset.f; mark = false; markUid = null; view = ""; zoom = 1; again(); });
    $("rsNew").onclick = () => newRoomDialog(b.uid, curF);
    $("rsBulk").onclick = () => bulkDialog(b.uid, curF);
    $("rsView").querySelectorAll("button").forEach(x => x.onclick = () => { if (x.disabled) return; view = x.dataset.v; again(); });
    $("rsRooms").querySelectorAll(".rc").forEach(c => c.onclick = () => openRoom(c.dataset.r));
    $("rsPpl").querySelectorAll(".rp").forEach(c => c.onclick = () => openRoom(c.dataset.r));
    $("rsPrn").querySelectorAll(".pr").forEach(c => c.onclick = () => openPrinter(c.dataset.u));
    if ($("rsGrid")) $("rsGrid").querySelectorAll(".gt").forEach(c => c.onclick = () => openRoom(c.dataset.r));
    $("plUp").onclick = () => $("plFile").click();
    $("plFile").onchange = async e => {
      const f = e.target.files[0]; e.target.value = ""; if (!f) return;
      try { await setPianta(b.uid, curF, f); view = "foto"; note("Pianta salvata. Ora puoi segnare le stanze sulla foto."); }
      catch (err) { note("Non riesco a leggere la foto: " + err.message, true); }
      again();
    };
    if (pl && pl.foto) {
      $("plMinus").onclick = () => { zoom = Math.max(1, zoom - 0.5); again(); };
      $("plPlus").onclick = () => { zoom = Math.min(4, zoom + 0.5); again(); };
      $("plMark").onclick = () => { mark = !mark; markUid = null; again(); };
      $("plDel").onclick = async () => { if (confirm("Togliere la foto della pianta da questo telefono? I segni delle stanze restano salvati.")) { await delPianta(b.uid, curF); view = ""; mark = false; note("Foto della pianta tolta."); again(); } };
      if ($("plChips")) $("plChips").querySelectorAll("button").forEach(x => x.onclick = () => { markUid = markUid === x.dataset.r ? null : x.dataset.r; again(); });
      if ($("plUnmark")) $("plUnmark").onclick = async () => { await setPos(markUid, null, null); again(); };
      $("plI").onclick = async ev => {
        const pm = ev.target.closest(".pm");
        if (pm) { if (mark) { markUid = pm.dataset.r; again(); } else openRoom(pm.dataset.r); return; }
        if (!mark) return;
        if (!markUid) { flash = { t: "Prima tocca, qui sotto, la stanza da segnare.", err: true }; again(); return; }
        const r = $("plImg").getBoundingClientRect();
        const px = (ev.clientX - r.left) / r.width, py = (ev.clientY - r.top) / r.height;
        if (px < 0 || px > 1 || py < 0 || py > 1) return;
        const ro = room(markUid); await setPos(markUid, px, py);
        const next = rs.find(x => x.px == null && x.uid !== markUid);   // passa da sola alla prossima non segnata
        note(`${ro ? ro.codice : "Stanza"} segnata sulla pianta.`); markUid = next ? next.uid : markUid; again();
      };
    }
  }

  function newRoomDialog(euid, piano) {
    const m = $("modal");
    m.innerHTML = `<div class="card"><h3>Nuova stanza</h3><div class="meta">${esc(bName(euid))} · ${floorLabel(piano)}</div>
      <label for="nrCod">Codice della stanza</label><input id="nrCod" maxlength="${CODE_MAX}" autocomplete="off" autocapitalize="characters" spellcheck="false" placeholder="Es. T12">
      <label for="nrScr">Scrivanie totali (facoltativo)</label><input id="nrScr" type="number" inputmode="numeric" min="0" placeholder="Lascia vuoto se non lo sai">
      <button class="btn green" id="nrOk" type="button">Crea la stanza</button><button class="btn" id="nrNo" type="button">Annulla</button><div id="nrMsg"></div></div>`;
    m.classList.remove("hidden");
    $("nrNo").onclick = () => m.classList.add("hidden");
    $("nrOk").onclick = async () => {
      const r = await addRoom(euid, piano, $("nrCod").value, $("nrScr").value);
      if (r.err) return show($("nrMsg"), r.err, true);
      m.classList.add("hidden"); note(`Stanza ${r.r.codice} creata.`); render();
    };
  }
  function bulkDialog(euid, piano) {
    const m = $("modal");
    m.innerHTML = `<div class="card"><h3>Incolla più stanze</h3><div class="meta">${esc(bName(euid))} · ${floorLabel(piano)}</div>
      <div class="hint">Un codice per riga (vanno bene anche virgole). Le stanze già presenti su questo piano non vengono duplicate.</div>
      <label for="bkTxt">Codici</label><textarea id="bkTxt" style="min-height:140px" autocapitalize="characters" spellcheck="false" placeholder="T1&#10;T2&#10;T12"></textarea>
      <div class="hint" id="bkInfo"></div>
      <button class="btn green" id="bkOk" type="button">Crea le stanze</button><button class="btn" id="bkNo" type="button">Annulla</button><div id="bkMsg"></div></div>`;
    m.classList.remove("hidden");
    const info = async () => {
      await load();
      const { good, bad } = parseCodes($("bkTxt").value), have = new Set(roomsOf(euid, piano).map(r => cf(r.codice)));
      const nu = good.filter(c => !have.has(cf(c))).length;
      $("bkInfo").textContent = good.length || bad.length ? `Verranno create ${nu} stanze nuove` + (good.length - nu ? `, ${good.length - nu} già presenti` : "") + (bad.length ? `, ${bad.length} scartate perché più lunghe di ${CODE_MAX} caratteri` : "") + "." : "";
    };
    $("bkTxt").addEventListener("input", info);
    $("bkNo").onclick = () => m.classList.add("hidden");
    $("bkOk").onclick = async () => {
      const r = await addRooms(euid, piano, $("bkTxt").value);
      if (!r.nuove && !r.esistenti && !r.scartate.length) return show($("bkMsg"), "Scrivi almeno un codice.", true);
      m.classList.add("hidden");
      note(`Stanze create: ${r.nuove}` + (r.esistenti ? `, già presenti: ${r.esistenti}` : "") + (r.scartate.length ? `, scartate perché troppo lunghe: ${r.scartate.join(", ")}` : "") + ".", !r.nuove && !!r.scartate.length);
      render();
    };
  }

  /* ---------- scheda della stanza ---------- */
  async function openRoom(uid, msg, err) {
    await load();
    const r = room(uid); if (!r) return;
    const m = $("modal"), o = occ(r), ppl = peopleIn(uid), fotos = await fotosOf(uid), prs = stampanti.filter(s => s.edificio_uid === r.edificio_uid && s.piano === r.piano);
    m.innerHTML = `<div class="card full"><h3>${esc(r.codice)}</h3>
      <div class="meta">${esc(bName(r.edificio_uid))} · ${floorLabel(r.piano)}</div><div id="roMsg"></div>
      <label for="roCod">Codice stanza</label><input id="roCod" maxlength="${CODE_MAX}" autocomplete="off" autocapitalize="characters" spellcheck="false" value="${esc(r.codice)}">
      <label for="roScr">Scrivanie totali (facoltativo)</label><input id="roScr" type="number" inputmode="numeric" min="0" placeholder="Vuoto = non indicate" value="${num(r.scrivanie) || ""}">
      <label for="roCome">Come arrivarci</label><textarea id="roCome" placeholder="Es. Ingresso principale, scale B, secondo corridoio a destra">${esc(r.come || "")}</textarea>
      <label for="roNote">Note</label><textarea id="roNote" style="min-height:56px">${esc(r.note || "")}</textarea>
      <button class="btn green" id="roSave" type="button">Salva la stanza</button>
      <label>Chi siede qui${o.txt ? ` · <span class="${o.over ? "tag st-Aperto" : "tag"}" id="roOcc">${o.txt}</span>` : ""}</label>
      <div id="roPpl">${ppl.map(p => `<div class="card"><b>Postazione ${p.numero}</b> · <span class="rpn" data-k="${esc(p.k)}" style="text-decoration:underline">${esc(p.utente)}</span>${phoneBits(userOf(p.k)) ? `<div class="meta">${phoneBits(userOf(p.k))}</div>` : ""}<button class="btn red sm" data-k="${esc(p.k)}" type="button" style="margin-top:6px">Togli da questa stanza</button></div>`).join("") || `<div class="hint">Nessuno è assegnato a questa stanza.</div>`}</div>
      <button class="btn navy" id="roAdd" type="button">＋ Aggiungi persona</button>
      <div id="roPick" class="hidden"><label for="roQ">Cerca nome, matricola o sede</label><input id="roQ" autocomplete="off" autocorrect="off" spellcheck="false" autocapitalize="words"><div id="roPL"></div></div>
      <label>Foto e panorami</label>
      <div class="thumbs" id="roTh">${fotos.map(f => `<div><img src="${esc(f.data)}" data-f="${esc(f.uid)}" alt="${f.tipo === "panorama" ? "Panorama" : "Foto"}" ${f.tipo === "panorama" ? 'style="width:150px"' : ""}>${f.tipo === "panorama" ? `<span class="tag" style="position:absolute;left:4px;bottom:4px">panorama</span>` : ""}<button class="x" data-x="${esc(f.uid)}" type="button" aria-label="Togli">×</button></div>`).join("") || `<span class="meta">Nessuna foto.</span>`}</div>
      <input id="roF1" type="file" accept="image/*" multiple class="hidden"><input id="roF2" type="file" accept="image/*" multiple class="hidden">
      <div class="row"><div><button class="btn" id="roB1" type="button">📷 Aggiungi foto</button></div><div><button class="btn" id="roB2" type="button">🌄 Aggiungi panorama</button></div></div>
      <div class="hint">Per il panorama usa la modalità «Pano» della fotocamera dell'iPhone, poi sceglila qui. Le foto vengono ridotte per non appesantire il file di scambio.</div>
      <label>Posizione sulla pianta</label>
      <div class="meta">${r.px != null ? "Segnata sulla pianta del piano." : "Non ancora segnata sulla pianta."}</div>
      <button class="btn" id="roMark" type="button">📍 ${r.px != null ? "Sposta il segno sulla pianta" : "Segna sulla pianta"}</button>
      <label>Stampanti di questo piano</label>
      <div id="roPrn">${prs.map(s => `<div class="card" data-u="${esc(s.uid)}"><b>${esc(printerName(s))}</b><div class="meta">${esc(s.modello || "")}</div></div>`).join("") || `<div class="hint">Nessuna stampante su questo piano.</div>`}</div>
      <button class="btn red" id="roDel" type="button">Elimina la stanza</button>
      <button class="btn" id="roClose" type="button">Chiudi</button></div>`;
    m.classList.remove("hidden");
    if (msg) show($("roMsg"), msg, err);
    const back = (t, e) => { render(); return openRoom(uid, t, e); };
    $("roClose").onclick = () => { m.classList.add("hidden"); render(); };
    $("roSave").onclick = async () => {
      const res = await saveRoom(uid, { codice: $("roCod").value, scrivanie: $("roScr").value, come: $("roCome").value, note: $("roNote").value });
      if (res.err) return show($("roMsg"), res.err, true);
      const oc = occ(res.r);
      back("Stanza salvata." + (oc.over ? `\nAttenzione: ci sono ${oc.n} persone ma le scrivanie sono ${oc.tot}.` : ""), oc.over);
    };
    $("roPpl").querySelectorAll("button[data-k]").forEach(b => b.onclick = async () => { await unassign(b.dataset.k); back("Persona tolta dalla stanza: le postazioni sono state rinumerate."); });
    $("roPpl").querySelectorAll(".rpn").forEach(s => s.onclick = () => editUser(s.dataset.k));
    $("roAdd").onclick = () => { $("roPick").classList.toggle("hidden"); if (!$("roPick").classList.contains("hidden")) { $("roQ").focus(); pick(); } };
    const pick = () => {
      const q = words($("roQ").value), here = new Set(peopleIn(uid).map(p => p.k));
      const list = users.filter(u => !here.has(u.key) && q.every(t => [u.nome, u.matricola, u.sede].join(" ").toLowerCase().includes(t))).slice(0, 8);
      const raw = normName($("roQ").value), canNew = raw.split(" ").length >= 2 && !users.some(u => u.key === keyOf(raw));   // «Nome Cognome»: così non si crea un utente per sbaglio con un nome a metà
      $("roPL").innerHTML = list.map(u => { const s = seatOf(u.key), sr = s && room(s.stanza_uid); return `<div class="card pk" data-k="${esc(u.key)}"><b>${esc(u.nome)}</b><div class="meta">${u.matricola ? "Matricola " + esc(u.matricola) + " · " : ""}${sr ? `siede in ${esc(sr.codice)} (${floorLabel(sr.piano)})` : "nessuna stanza"}</div></div>`; }).join("") + (canNew ? `<button class="btn" id="roNewU" type="button">＋ Crea «${esc(raw)}» come nuovo utente</button>` : "") || `<div class="hint">Nessun utente trovato.</div>`;
      $("roPL").querySelectorAll(".pk").forEach(c => c.onclick = () => doAssign(uid, c.dataset.k, back));
      if ($("roNewU")) $("roNewU").onclick = async () => {
        const b = edifici.find(x => x.uid === r.edificio_uid) || {};
        await put("utenti", { key: keyOf(raw), nome: raw, matricola: "", sede: b.sede || "", telefono: "", seriale_pc: "", imei: "", modificato: nowIso() });
        await loadUsers(); doAssign(uid, keyOf(raw), back);
      };
    };
    $("roQ").addEventListener("input", pick);
    $("roB1").onclick = () => $("roF1").click(); $("roB2").onclick = () => $("roF2").click();
    for (const [inp, tipo] of [["roF1", "foto"], ["roF2", "panorama"]]) $(inp).onchange = async e => {
      const files = [...e.target.files]; e.target.value = ""; if (!files.length) return;
      show($("roMsg"), "Salvo le foto…");
      for (const f of files) { try { await addFoto(uid, f, tipo); } catch (er) { /* foto non leggibile: salto */ } }
      back(`${files.length} ${tipo === "panorama" ? (files.length === 1 ? "panorama aggiunto" : "panorami aggiunti") : (files.length === 1 ? "foto aggiunta" : "foto aggiunte")}.`);
    };
    $("roTh").querySelectorAll("img").forEach(i => i.onclick = () => viewFoto(i.dataset.f, uid));
    $("roTh").querySelectorAll("button.x").forEach(x => x.onclick = async () => { if (confirm("Togliere questa foto?")) { await delFoto(x.dataset.x); back("Foto tolta."); } });
    $("roMark").onclick = () => { m.classList.add("hidden"); openFloor(r.edificio_uid, r.piano, { mark: true, markUid: uid }); };
    $("roPrn").querySelectorAll(".card").forEach(c => c.onclick = () => openPrinter(c.dataset.u));
    $("roDel").onclick = async () => {
      const n = peopleIn(uid).length;
      if (!confirm(`Eliminare la stanza ${r.codice}?` + (n ? `\nLe ${n} persone assegnate restano tra gli utenti ma non avranno più una stanza.` : "") + "\nLe foto della stanza vengono tolte.")) return;
      await deleteRoom(uid); m.classList.add("hidden"); note(`Stanza ${r.codice} eliminata.`); render();
    };
  }
  /* assegna una persona; se siede già in un'altra stanza, prima chiede conferma */
  async function doAssign(ruid, ukey, back) {
    const r = room(ruid), u = userOf(ukey), cur = seatOf(ukey);
    if (!r || !u) return;
    if (cur && cur.stanza_uid !== ruid) {
      const o = room(cur.stanza_uid);
      if (!confirm(`${u.nome} siede già in ${o.codice} (${floorLabel(o.piano)}, Postazione ${cur.numero}).\nSpostarlo in ${r.codice}?`)) return;
    }
    const res = await assign(ukey, ruid), oc = occ(room(ruid));
    back(res.gia ? `${u.nome} è già in questa stanza (Postazione ${res.numero}).` : `${u.nome}: ${r.codice} · Postazione ${res.numero}.` + (res.da ? `\nPrima era in ${res.da.codice}: le sue postazioni sono state rinumerate.` : "") + (res.sopra ? `\nAttenzione: ora ci sono ${oc.n} persone ma le scrivanie sono ${oc.tot}.` : ""), res.sopra);
  }
  function viewFoto(fuid, ruid) {
    getKey("stanzefoto", fuid).then(f => {
      if (!f) return;
      const m = $("modal"), pano = f.tipo === "panorama";
      m.innerHTML = `<div class="card full"><h3>${pano ? "Panorama" : "Foto"}</h3>${pano ? `<div style="overflow-x:auto;-webkit-overflow-scrolling:touch"><img src="${esc(f.data)}" alt="Panorama" style="width:auto;max-width:none;height:55vh;margin:0"></div><div class="hint">Scorri di lato per vedere tutto il panorama.</div>` : `<img src="${esc(f.data)}" alt="Foto">`}
        <button class="btn" id="vfTipo" type="button">${pano ? "Segna come foto normale" : "Segna come panorama"}</button><button class="btn" id="vfBack" type="button">Indietro</button></div>`;
      $("vfBack").onclick = () => openRoom(ruid);
      $("vfTipo").onclick = async () => { await setFotoTipo(fuid, pano ? "foto" : "panorama"); openRoom(ruid, pano ? "Segnata come foto." : "Segnata come panorama."); };
    });
  }

  /* ---------- scorciatoie: piano nella vista Edificio e scheda utente ---------- */
  function openFloor(euid, piano, opts) {
    bSel = euid; curF = piano; view = ""; zoom = 1; mark = !!(opts && opts.mark); markUid = (opts && opts.markUid) || null;
    prnTab = "e"; $("modal").classList.add("hidden"); go("prn");
  }
  async function floorMenu(euid, piano) {
    await load();
    const m = $("modal"), rs = roomsOf(euid, piano), n = rs.reduce((t, r) => t + peopleIn(r.uid).length, 0), pl = await pianta(euid, piano);
    m.innerHTML = `<div class="card"><h3>${esc(bName(euid))} · ${floorLabel(piano)}</h3><div class="meta">${rs.length} ${rs.length === 1 ? "stanza" : "stanze"} · ${n} ${n === 1 ? "persona" : "persone"}${pl && pl.foto ? " · pianta caricata" : " · pianta non ancora caricata"}</div>
      <button class="btn green" id="fmOpen" type="button">🗺 Apri pianta</button>
      <button class="btn navy" id="fmNew" type="button">＋ Nuova stanza</button>
      <button class="btn" id="fmNo" type="button">Chiudi</button></div>`;
    m.classList.remove("hidden");
    $("fmOpen").onclick = () => openFloor(euid, piano);
    $("fmNew").onclick = () => { bSel = euid; curF = piano; newRoomDialog(euid, piano); };
    $("fmNo").onclick = () => m.classList.add("hidden");
  }
  /* «Dove si siede» nella scheda utente */
  async function dove(ukey, el) {
    await load();
    const s = seatOf(ukey), r = s && room(s.stanza_uid), all = S.filter(x => !x.eliminata).sort((a, c) => cmp(fullName(a), fullName(c)));
    if (r) {
      el.innerHTML = `<div class="meta" id="eDoveTxt"><b>${esc(r.codice)} · Postazione ${s.numero}</b><br>${esc(bName(r.edificio_uid))} · ${floorLabel(r.piano)}</div><div class="row"><div><button class="btn sm" id="eDoveOpen" type="button" style="width:100%">Apri la stanza</button></div><div><button class="btn sm" id="eDoveOut" type="button" style="width:100%">Togli dalla stanza</button></div></div>`;
      $("eDoveOpen").onclick = () => openRoom(r.uid);
      $("eDoveOut").onclick = async () => { await unassign(ukey); dove(ukey, el); render(); };
    } else {
      el.innerHTML = `<div class="meta" id="eDoveTxt">Nessuna stanza assegnata.</div>` + (all.length ? `<div class="row"><div style="flex:3"><select id="eDoveSel"><option value="">Scegli la stanza…</option>${all.map(x => `<option value="${esc(x.uid)}">${esc(fullName(x))}</option>`).join("")}</select></div><div><button class="btn sm" id="eDoveSet" type="button" style="width:100%;height:100%">Assegna</button></div></div>` : `<div class="hint">Non ci sono stanze: creale in Stampanti › Piani e stanze.</div>`);
      if ($("eDoveSet")) $("eDoveSet").onclick = async () => {
        const v = $("eDoveSel").value; if (!v) return;
        const res = await assign(ukey, v), x = room(v), oc = occ(x);
        await dove(ukey, el); render();
        if (res.sopra) $("eDoveTxt").innerHTML += `<br><span style="color:var(--red);font-weight:700">Attenzione: ora ci sono ${oc.n} persone ma le scrivanie sono ${oc.tot}.</span>`;
      };
    }
  }

  return { load, render, openRoom, openFloor, floorMenu, dove, parseCodes, cleanCode,
           addRoom, addRooms, saveRoom, setPos, deleteRoom, assign, unassign, dropUser, renameUser,
           addFoto, delFoto, setPianta, delPianta, pianta, fotosOf, exportItems, countNew, importData,
           peopleIn, seatOf, roomsOf, occ, lab,
           _state: () => ({ S, P, curF, view, mark, markUid, zoom }) };
})();
