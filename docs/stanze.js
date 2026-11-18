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

  /* una stanza per riga, con le postazioni se scritte: "T12; 3", "T12 - 3", "T12 (3)", due colonne da Excel */
  const DESK = "(?:postazion[ei]|posti|scrivani[ae]|desks?)";
  const PAIR = [
    new RegExp("^([^;\\t]+?)\\s*(?:[;\\t:=]|\\s[-\u2013\u2014]\\s)\\s*(\\d{1,3})\\s*(?:" + DESK + ")?\\s*$", "i"),
    new RegExp("^([^;\\t()]+?)\\s*\\(\\s*(\\d{1,3})\\s*(?:" + DESK + ")?\\s*\\)\\s*$", "i")
  ];
  function parseRooms(text) {
    const order = [], desks = new Map(), bad = [];
    const add = (code, n) => {
      const c = String(code).split(/\s+/).filter(Boolean).join(" ");
      if (!c) return;
      if (c.length > CODE_MAX) { bad.push(c); return; }
      const k = cf(c);
      if (!desks.has(k)) { order.push(c); desks.set(k, n); } else if (n != null) desks.set(k, n);
    };
    for (let line of String(text || "").split(/[\r\n]+/)) {
      line = line.trim();
      if (!line) continue;
      let m = null;
      for (const rx of PAIR) { m = rx.exec(line); if (m) break; }
      if (m) add(m[1], parseInt(m[2], 10));
      else for (const part of line.split(/[;,\t]+/)) add(part, null);
    }
    return { rooms: order.map(c => ({ c, n: desks.get(cf(c)) })), bad };
  }

  /* ---------- dati ---------- */
  let S = [], P = [];   // stanze e postazioni (piccole: si rileggono sempre)
  async function load() { S = await getAll("stanze"); P = await getAll("postazioni"); }
  const room = uid => S.find(r => r.uid === uid && !r.eliminata) || null;
  const roomsOf = (euid, piano) => S.filter(r => !r.eliminata && r.edificio_uid === euid && r.piano === piano).sort((a, b) => cmp(a.codice, b.codice));
  const sortName = n => { const u = users.find(x => x.key === keyOf(n || "")); return (u && u.ordine) || n || ""; };   /* per cognome */
  const peopleIn = ruid => P.filter(p => p.stanza_uid === ruid && room(ruid)).sort((a, b) => (a.numero - b.numero) || cmp(sortName(a.utente), sortName(b.utente)));
  const seatOf = k => { const p = P.find(x => x.k === k && x.stanza_uid); return p && room(p.stanza_uid) ? p : null; };
  const userOf = k => users.find(u => u.key === k) || null;
  const bName = euid => (edifici.find(x => x.uid === euid) || {}).nome || "?";
  const fullName = r => `${bName(r.edificio_uid)} · ${floorLabel(r.piano)} · ${r.codice}`;
  function occ(r) { const n = peopleIn(r.uid).length, tot = num(r.scrivanie); return { n, tot, txt: tot > 0 ? `occupate ${n}/${tot}` : "", over: tot > 0 && n > tot }; }
  const lab = r => `${r.codice} · ${occ(r).n}`;
  const dispCmp = (a, b) => cmp(sortName(a.utente), sortName(b.utente));

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
    const { rooms, bad } = parseRooms(text), cur = new Map(roomsOf(euid, piano).map(r => [cf(r.codice), r]));
    let nuove = 0, esistenti = 0, aggiornate = 0;
    for (const { c, n } of rooms) {
      const ex = cur.get(cf(c));
      if (ex) {
        if (n != null && n !== num(ex.scrivanie)) { await put("stanze", { ...ex, scrivanie: n, modificato: nowIso() }); aggiornate++; }
        else esistenti++;
        continue;
      }
      await put("stanze", { uid: "ro-" + uuid(), edificio_uid: euid, piano, codice: c, scrivanie: n == null ? 0 : n, note: "", come: "", px: null, py: null, eliminata: false, modificato: nowIso() });
      nuove++;
    }
    await load();
    return { nuove, esistenti, aggiornate, scartate: bad };
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
    let da = null, presa = "";
    if (old) {
      da = room(old.stanza_uid);
      presa = await dropPresa(ukey, now);
      await put("postazioni", { ...old, stanza_uid: "", numero: 0, modificato: now });
      await load(); await renumberRoom(old.stanza_uid, now, true);
    }
    const n = P.filter(p => p.stanza_uid === ruid).length + 1;
    await put("postazioni", { k: ukey, utente: u.nome, stanza_uid: ruid, numero: n, modificato: now });
    await load();
    const tot = num(r.scrivanie);
    return { numero: n, gia: false, da, sopra: !!(tot && n > tot), presa_tolta: presa };
  }
  /* la presa di rete è della scrivania, non della persona: se cambia scrivania il codice di prima si toglie (ritorna il codice tolto) */
  async function dropPresa(ukey, now) {
    const cur = await getKey("utenti", ukey); if (!cur) return "";
    const dot = dotList(cur), p = dot.find(r => rowKind(r) === "presa"); if (!p) return "";
    await put("utenti", { ...cur, dotazione: dot.filter(r => rowKind(r) !== "presa"), modificato: now });
    if (typeof loadUsers === "function") await loadUsers();
    return p.seriale || p.modello || "";
  }
  async function unassign(ukey) {
    await load();
    const old = seatOf(ukey); if (!old) return false;
    const now = nowIso();
    await dropPresa(ukey, now);
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
  const dotBits = u => (u && typeof dotText === "function" && dotText(u.dotazione)) ? "Dotazione: " + esc(dotText(u.dotazione)) : "";
  const phoneBits = u => [u && u.ufficio ? esc(u.ufficio) : "", u && u.telefono ? "Tel " + esc(u.telefono) : "", u && u.seriale_pc ? "PC " + esc(u.seriale_pc) : "", u && u.imei ? "IMEI " + esc(u.imei) : "", dotBits(u)].filter(Boolean).join(" · ");
  const note = (t, err) => { flash = { t, err: !!err }; };

  /* ---------- scrivanie disegnate: una per postazione, col PC e il suo S/N ---------- */
  /* il telefono sulla scrivania viene dalla dotazione: «Cellulare»/«Smartphone» Samsung Galaxy S24, S24 Plus o S25,
     e/o «Telefono Teams» (AudioCodes) da scrivania. Con un IMEI e nessun modello: il cellulare classico (S24). */
  const LAPTOP_MODELLI = { "13": "Dell 13 pollici", "15": "Dell 15 pollici" };
  const CELL_MODELLI = ["S24", "S24 Plus", "S25"], TEAMS_TIPO = "Telefono Teams", TEAMS_MODELLO = "AudioCodes C435HD", PRESA_TIPO = "Presa di rete", PRESA_MAX = 30;
  /* nome normalizzato -> [disegno, nome mostrato]; S4, S4 Plus e S5 sono i nomi scritti per errore all'inizio */
  const CELL_KEYS = { s24: ["s24", "S24"], s24plus: ["s24plus", "S24 Plus"], s25: ["s25", "S25"], s4: ["s24", "S24"], s4plus: ["s24plus", "S24 Plus"], s5: ["s25", "S25"] };
  const normModel = t => String(t || "").toLowerCase().replace(/\+/g, " plus ").replace(/samsung|galaxy/g, " ").replace(/[^a-z0-9]/g, "");
  const rowKind = r => {
    const tipo = String((r && r.tipo) || "").toLowerCase(), both = tipo + " " + String((r && r.modello) || "").toLowerCase();
    if (tipo.trim().startsWith("presa")) return "presa";
    if (both.includes("teams") || both.includes("audiocodes")) return "teams";
    if (tipo.includes("cellulare") || tipo.includes("smartphone")) return "mobile";
    if (tipo.includes("notebook") || tipo.includes("portatile") || tipo.includes("laptop")) return "laptop";
    if (tipo.includes("monitor")) return "monitor";
    return "";
  };
  const laptopSize = m => { const x = /(?<!\d)(13|15)(?!\d)/.exec(String(m || "")); return x ? x[1] : ""; };
  function phoneInfo(dot, imei) {
    let mobile = null, label = "", teams = false, laptop = null, llabel = "";
    for (const r of dot) {
      const k = rowKind(r);
      if (k === "teams") teams = true;
      else if (k === "laptop" && laptop === null && r.modello) { laptop = laptopSize(r.modello) || "altro"; llabel = r.modello; }
      else if (k === "mobile" && mobile === null) {
        const n = normModel(r.modello);
        [mobile, label] = CELL_KEYS[n] || ["altro", r.modello || ""];
      }
    }
    if (mobile === null && String(imei || "").trim()) mobile = "s24";
    return { mobile, label, teams, laptop, laptop_label: llabel };
  }
  function withPhones(dot, mobile, teams, laptop) {   // mobile / laptop: "" = non indicato, null = lascia com'è
    const oldM = dot.find(r => rowKind(r) === "mobile"), oldT = dot.find(r => rowKind(r) === "teams");
    const laps = dot.filter(r => rowKind(r) === "laptop"), oldL = laps[0];
    const out = dot.filter(r => !rowKind(r) || rowKind(r) === "monitor" || rowKind(r) === "presa").concat(laps.slice(1));   // il monitor AOC e la presa restano sempre
    if (laptop === null || laptop === undefined) laps.slice(0, 1).forEach(r => out.push(r));
    else if (LAPTOP_MODELLI[laptop]) out.push({ tipo: oldL ? oldL.tipo : "Notebook", modello: LAPTOP_MODELLI[laptop], seriale: oldL ? oldL.seriale : "" });
    else if (oldL && oldL.seriale) out.push({ ...oldL, modello: "" });   // «non indicato»: il seriale scritto non si perde
    if (mobile === null) dot.filter(r => rowKind(r) === "mobile").forEach(r => out.push(r));
    else if (mobile) out.push({ tipo: oldM ? oldM.tipo : "Cellulare", modello: mobile, seriale: oldM ? oldM.seriale : "" });
    if (teams) out.push(oldT || { tipo: TEAMS_TIPO, modello: TEAMS_MODELLO, seriale: "" });
    return out;
  }
  /* misure (larghezza, altezza) delle foto nella scena 130 x 66: le stesse del programma sul PC */
  const IMG_SIZE = { lap13: [26, 18], lap15: [32, 22], s24: [10, 20], s24plus: [10, 22], s25: [10, 20], teams: [32, 24] };
  const imgUrl = k => "disp-" + k + ".png";
  const photoKeys = ph => ({ laptop: { "13": "lap13", "15": "lap15" }[ph && ph.laptop] || null, mobile: ["s24", "s24plus", "s25"].includes(ph && ph.mobile) ? ph.mobile : null, teams: ph && ph.teams ? "teams" : null });
  /* la scrivania vista da davanti, come nelle foto: pannello azzurro, monitor con webcam, tastiera, mouse e,
     con le FOTO dei dispositivi scelti, il portatile Dell (13 o 15 pollici, visto dall'alto), il cellulare
     Galaxy (S24, S24 Plus, S25) a destra e il telefono Teams AudioCodes a sinistra */
  function deskSvg(w, ph, faded, big) {   // big: scrivania a pagina intera (w = "100%"): scritta AOC e coperchio del PC sempre disegnato
    ph = ph || {};
    const R = (x, y, ww, hh, f) => `<rect x="${x}" y="${y}" width="${ww}" height="${hh}" fill="${f}"/>`, O = (x, y, ww, hh, f) => `<ellipse cx="${x + ww / 2}" cy="${y + hh / 2}" rx="${ww / 2}" ry="${hh / 2}" fill="${f}"/>`;
    const I = (k, x, y) => `<image data-k="${k}" href="${imgUrl(k)}" x="${x}" y="${y}" width="${IMG_SIZE[k][0]}" height="${IMG_SIZE[k][1]}" preserveAspectRatio="none"/>`;
    let g = R(0, 40, 130, 26, "#ecd3a8") + R(2, 28, 126, 14, "#9db8e6") + R(54, 41, 22, 3, "#15181d") + R(62, 35, 6, 6, "#15181d") + R(42, 6, 46, 30, "#15181d") + R(44.5, 8.5, 41, 22, "#2b313b") + R(60, 3, 10, 3.5, "#0b0d10") + R(44, 48, 40, 8, "#1d2026");
    for (let k = 0; k < 3; k++) g += R(47, 50 + k * 2.2, 34, 0.9, "#3a3f49");
    g += O(88.5, 47, 7, 9, "#1d2026");
    const pk = photoKeys(ph);
    if (pk.laptop) g += I(pk.laptop, 27 - IMG_SIZE[pk.laptop][0] / 2, 63 - IMG_SIZE[pk.laptop][1]);
    else if (ph.laptop === "altro" || big) g += R(13, 44, 28, 19, "#a9b2bd") + (big && ph.laptop !== "altro" ? `<text x="21" y="60" font-size="5" font-weight="700" text-anchor="middle" fill="#5b6b7d" font-family="sans-serif">PC</text>` : "");
    if (pk.teams) g += I("teams", 0, 66 - IMG_SIZE.teams[1]);
    if (pk.mobile) g += I(pk.mobile, 113 - IMG_SIZE[pk.mobile][0] / 2, 62 - IMG_SIZE[pk.mobile][1]);
    else if (ph.mobile === "altro") g += R(108.5, 45, 9, 17, "#14171c") + R(109.5, 46.5, 7, 14, "#2b313b");
    if (big) g += `<text x="65" y="34.3" font-size="3" font-weight="700" text-anchor="middle" fill="#c9ced6" font-family="sans-serif" letter-spacing=".3">AOC</text>` + R(108, 29, 10, 9, "#f4f6f8") + R(111, 32.5, 4, 3, "#2b313b") + R(112.2, 31.6, 1.6, 1, "#2b313b");   // la presa di rete a muro
    const size = big ? 'width="100%" style="display:block;height:auto"' : `width="${w}" height="${Math.round(w * 66 / 130)}"${faded ? ' style="opacity:.38"' : ""}`;
    return `<svg viewBox="0 0 130 66" ${size} aria-hidden="true" focusable="false">${g}</svg>`;
  }
  /* le foto dei tre apparecchi scelti, con la didascalia: cambiano quando cambia la scelta */
  const PIC_W = { lap13: 82, lap15: 100, s24: 40, s24plus: 43, s25: 40, teams: 100 };
  function picsHtml(ph) {
    const pk = photoKeys(ph), fig = (id, k, cap) => `<figure id="${id}">${k ? `<img src="${imgUrl(k)}" alt="" style="width:${PIC_W[k]}px">` : ""}<figcaption>${esc(cap)}</figcaption></figure>`;
    return fig("dkPicL", pk.laptop, LAPTOP_MODELLI[ph.laptop] || ph.laptop_label || "Portatile non indicato")
      + fig("dkPicM", pk.mobile, pk.mobile ? "Samsung Galaxy " + ph.label : (ph.label || "Cellulare non indicato"))
      + fig("dkPicT", pk.teams, ph.teams ? TEAMS_MODELLO : "Telefono Teams: no");
  }
  const phoneTxt = ph => [ph.laptop_label || "", [ph.label ? "Cell. " + ph.label : "", ph.teams ? "Telefono Teams" : ""].filter(Boolean).join(" · ")].filter(Boolean).join(" · ");
  const dotList = u => (typeof cleanDot === "function" ? cleanDot(u && u.dotazione) : []);
  const snOf = u => String((u && u.seriale_pc) || "").trim();
  function deskCards(r) {
    const ppl = peopleIn(r.uid), tot = num(r.scrivanie), out = [];
    for (const p of ppl) {
      const u = userOf(p.k), sn = snOf(u), dl = dotList(u), ex = dl.filter(r => !rowKind(r)).length, ph = phoneInfo(dl, u && u.imei);
      out.push(`<button type="button" class="dkc" data-r="${esc(r.uid)}" data-k="${esc(p.k)}" aria-label="Postazione ${p.numero}: ${esc(p.utente)}"><span class="dn">Postazione ${p.numero}</span>${deskSvg(94, ph)}<b>${esc(p.utente)}</b>${sn ? `<span class="sn">S/N ${esc(sn)}</span>` : `<span class="sn no">S/N non inserito</span>`}${phoneTxt(ph) ? `<span class="ex">${esc(phoneTxt(ph))}</span>` : ""}${ex ? `<span class="ex">+ ${ex} ${ex === 1 ? "apparecchio" : "apparecchi"}</span>` : ""}</button>`);
    }
    for (let i = ppl.length; i < tot; i++) out.push(`<button type="button" class="dkc free" data-r="${esc(r.uid)}" data-free="1" aria-label="Postazione ${i + 1} libera"><span class="dn">Postazione ${i + 1}</span>${deskSvg(94, {}, true)}<b>Libera</b></button>`);
    return out;
  }
  /* ---------- i dispositivi di una scrivania: PC, monitor AOC, cellulare (IMEI), telefono Teams ----------
     ognuno ha il suo bollino (verde = codice scritto, rosso = manca): si tocca e si scrive il codice a mano
     oppure si legge con la fotocamera */
  const BOLLINI = { pc: [36, 50], monitor: [65, 19], mobile: [113, 50], teams: [16, 58], presa: [113, 33] };   // posizioni nella scena 130 x 66
  const monitorRow = dot => dot.find(r => rowKind(r) === "monitor");
  function deskDevices(u, dot, ph) {
    const mon = monitorRow(dot), teams = dot.find(r => rowKind(r) === "teams"), out = [];
    out.push({ key: "pc", kind: "serial", name: "Computer", sub: ph.laptop_label || "", label: "S/N del computer", value: snOf(u) });
    out.push({ key: "monitor", kind: "serial", name: "Monitor AOC", sub: mon && mon.modello && !/^aoc$/i.test(mon.modello) ? mon.modello : "", label: "S/N del monitor AOC", value: mon ? mon.seriale || "" : "" });
    if (ph.mobile) out.push({ key: "mobile", kind: "imei", name: "Cellulare", sub: ph.mobile === "altro" ? ph.label : "Samsung Galaxy " + ph.label, label: "IMEI del cellulare", value: String((u && u.imei) || "").trim() });
    if (ph.teams) out.push({ key: "teams", kind: "serial", name: "Telefono Teams", sub: TEAMS_MODELLO, label: "S/N del telefono Teams", value: teams ? teams.seriale || "" : "" });
    const presa = dot.find(r => rowKind(r) === "presa");
    out.push({ key: "presa", kind: "presa", name: PRESA_TIPO, sub: presa ? presa.modello || "" : "", label: "Codice della presa di rete", value: presa ? presa.seriale || "" : "", port: presa ? presa.modello || "" : "" });
    return out;
  }
  /* la dotazione con il seriale del monitor o del telefono Teams cambiato (la riga si crea se manca) */
  function withSerial(dot, key, value, port) {
    if (key === "presa") {   // senza codice e senza porta la riga sparisce; port undefined = lascia la porta com'è
      const old = dot.find(r => rowKind(r) === "presa"), porta = port === undefined ? (old ? old.modello || "" : "") : String(port || "").replace(/\s+/g, " ").trim().slice(0, 80);
      const rows = dot.filter(r => rowKind(r) !== "presa");
      if (value || porta) rows.push({ tipo: PRESA_TIPO, modello: porta, seriale: value });
      return rows;
    }
    const kind = key === "teams" ? "teams" : "monitor", rows = dot.map(r => ({ ...r })), row = rows.find(r => rowKind(r) === kind);
    if (row) row.seriale = value;
    else if (value) rows.push(kind === "teams" ? { tipo: TEAMS_TIPO, modello: TEAMS_MODELLO, seriale: value } : { tipo: "Monitor", modello: "AOC", seriale: value });
    return rows;
  }
  /* gli altri apparecchi della dotazione (tablet, SIM, cuffie, PC fisso…): non il portatile, il monitor AOC, il cellulare,
     il telefono Teams e la presa, che hanno il loro bollino. «key» = dot:N, N = posizione nella dotazione */
  const deskOthers = dot => dot.map((r, i) => ({ r, i })).filter(x => !rowKind(x.r)).map(x => ({
    key: "dot:" + x.i, kind: "serial", other: true, idx: x.i, name: x.r.tipo || "Apparecchio", sub: x.r.modello || "",
    label: "S/N o IMEI" + (x.r.tipo ? " (" + x.r.tipo + ")" : ""), value: String(x.r.seriale || "").trim(), tipo: x.r.tipo || "", modello: x.r.modello || "" }));
  /* il cellulare e il telefono Teams che non ci sono ancora sulla scrivania: si aggiungono scrivendo (o leggendo) il codice */
  const deskAddable = (u, ph) => {
    const out = [];
    if (!ph.mobile) out.push({ key: "mobile", kind: "imei", add: true, name: "Cellulare", sub: "", label: "IMEI del cellulare", value: "" });
    if (!ph.teams) out.push({ key: "teams", kind: "serial", add: true, name: "Telefono Teams", sub: TEAMS_MODELLO, label: "S/N del telefono Teams", value: "" });
    return out;
  };
  /* la dotazione con un altro apparecchio aggiunto (idx = null) o cambiato; un tipo che ha già il suo bollino non va bene */
  function withOther(dot, idx, tipo, modello, seriale) {
    const row = { tipo: String(tipo || "").replace(/\s+/g, " ").trim().slice(0, 40), modello: String(modello || "").replace(/\s+/g, " ").trim().slice(0, 80), seriale: String(seriale || "") };
    if (!row.tipo) return { err: "Scrivi che apparecchio è (per esempio Tablet, Scheda SIM, Cuffie)." };
    const k = rowKind(row);
    if (k) return { err: k === "mobile" ? "Il cellulare ha il suo pulsante «Aggiungi: Cellulare» (con l'IMEI)." : k === "teams" ? "Il telefono Teams ha il suo pulsante «Aggiungi: Telefono Teams»." : k === "monitor" ? "Il monitor AOC ha il suo bollino sulla scrivania." : k === "presa" ? "La presa di rete ha il suo bollino sulla scrivania." : "Il portatile si cambia con «Portatile sulla scrivania» più in basso; il suo S/N è quello del Computer." };
    const rows = dot.map(r => ({ ...r }));
    if (idx == null) rows.push(row); else rows[idx] = row;
    return { rows };
  }
  /* come si scrive un codice: il S/N senza spazi e in maiuscolo, l'IMEI di 15 cifre con la cifra di controllo; "" = codice cancellato */
  function checkCode(kind, text) {
    const raw = String(text == null ? "" : text).replace(/[\x00-\x1f\x7f]/g, "");
    if (kind === "imei") {
      const t = raw.replace(/^\s*imei[\s:\-]*/i, "").replace(/[\s\-]/g, "");
      if (!t) return { v: "" };
      if (!/^\d{15}$/.test(t)) return { err: "L'IMEI ha 15 cifre: ne hai scritte " + t.length + ". Controlla di non aver saltato un numero." };
      if (!(window.SCAN ? window.SCAN.luhn(t) : true)) return { err: "Questo IMEI non torna con la cifra di controllo: c'è un numero sbagliato. Riguardalo, o leggilo con la fotocamera." };
      return { v: t };
    }
    if (kind === "presa") {
      const q = raw.replace(/\s+/g, " ").trim().toUpperCase();
      if (!q) return { v: "" };
      if (q.length > PRESA_MAX) return { err: "Il codice della presa può avere al massimo " + PRESA_MAX + " caratteri (ne hai scritti " + q.length + ")." };
      if (!/^[A-Z0-9][A-Z0-9 \-./_]*$/.test(q)) return { err: "Il codice della presa può avere lettere, numeri e i segni - . / _ (per esempio A12-34)." };
      return { v: q };
    }
    const t = raw.replace(/\s+/g, "").toUpperCase();
    if (!t) return { v: "" };
    if (t.length < 4 || t.length > 40) return { err: "Il S/N deve avere da 4 a 40 caratteri (ne hai scritti " + t.length + ")." };
    return { v: t };
  }
  async function setDeviceCode(k, key, value, port) {
    const cur = await getKey("utenti", k);
    if (!cur) return false;
    const patch = key === "pc" ? { seriale_pc: value } : key === "mobile" ? { imei: value } : { dotazione: withSerial(dotList(cur), key, value, port) };
    await put("utenti", { ...cur, ...patch, modificato: nowIso() });
    await loadUsers();
    return true;
  }
  function closeSheet() { const e = $("dvSheet"); if (e) e.remove(); }
  /* la schermata di un dispositivo: scrivi il codice o leggilo con la fotocamera */
  function openDevice(ruid, k, from, key) {
    const u = userOf(k), dot = dotList(u), ph = phoneInfo(dot, u && u.imei), d = deskDevices(u, dot, ph).find(x => x.key === key) || deskAddable(u, ph).find(x => x.key === key);
    if (!d) return;
    closeSheet();
    const el = document.createElement("div");
    el.id = "dvSheet"; el.className = "dvSheet";
    const imei = d.kind === "imei", presa = d.kind === "presa";
    el.innerHTML = `<div class="dvCard" role="dialog" aria-label="${esc(d.label)}">
      <h3>${esc(d.name)}${d.sub ? ` <small>${esc(d.sub)}</small>` : ""}</h3>
      <div class="meta">${esc((u && u.nome) || "")}</div>
      <label for="dvIn">${esc(d.label)}</label>
      <input id="dvIn" type="text" value="${esc(d.value)}" data-scan="${imei ? "imei" : presa ? "any" : "serial"}" autocomplete="off" autocapitalize="characters" autocorrect="off" spellcheck="false"${imei ? ' inputmode="numeric"' : ""} placeholder="${imei ? "15 cifre" : presa ? "per esempio A12-34" : "Numero di serie"}">
      <button class="btn navy" id="dvCam" type="button">📷 Leggi con la fotocamera</button>
      <div class="hint">${imei ? "L'IMEI sono 15 cifre: lo trovi nel codice a barre sulla scatola o digitando *#06# sul cellulare." : presa ? "Il codice della presa è scritto sulla placca a muro. Lo scrivi a mano, oppure lo leggi con la fotocamera se c'è un'etichetta con codice a barre o QR." : "Il S/N è sull'etichetta dell'apparecchio (di solito sotto o dietro), vicino al codice a barre."}</div>
      ${presa ? `<label for="dvPort">Porta dello switch (facoltativa)</label><input id="dvPort" type="text" value="${esc(d.port || "")}" maxlength="80" autocomplete="off" placeholder="per esempio SW1 porta 12">` : ""}
      <div id="dvMsg"></div>
      <button class="btn green" id="dvSave" type="button">Salva</button>
      ${d.value ? `<button class="btn red" id="dvClear" type="button">Cancella il codice</button>` : ""}
      <button class="btn" id="dvCancel" type="button">Annulla</button></div>`;
    document.body.appendChild(el);
    const inp = $("dvIn"), msg = (t, err) => { $("dvMsg").innerHTML = t ? `<div class="msg${err ? " err" : ""}">${esc(t)}</div>` : ""; };
    if (window.SCAN) window.SCAN.attach(inp);
    $("dvCam").onclick = () => { const b = inp.parentNode.querySelector(".scbtn"); if (b) b.click(); else msg("La fotocamera non è disponibile: scrivi il codice a mano.", true); };
    const done = async value => {
      await setDeviceCode(k, key, value, presa ? (value ? $("dvPort").value : "") : undefined);
      closeSheet();
      openDesk(ruid, k, from, { saved: d.name + (value ? ": codice salvato." : ": codice cancellato.") });
    };
    $("dvSave").onclick = async () => {
      const c = checkCode(d.kind, inp.value);
      if (c.err) { msg(c.err, true); inp.focus(); return; }
      if (!c.v) { msg("Scrivi il codice o leggilo con la fotocamera.", true); inp.focus(); return; }
      await done(c.v);
    };
    inp.addEventListener("keydown", ev => { if (ev.key === "Enter") { ev.preventDefault(); $("dvSave").click(); } });
    if ($("dvClear")) $("dvClear").onclick = () => done("");
    $("dvCancel").onclick = closeSheet;
  }
  /* un altro apparecchio (tablet, SIM, cuffie…): tipo, modello e codice a mano o con la fotocamera; idx = null = nuovo */
  function openOther(ruid, k, from, idx) {
    const u = userOf(k), dot = dotList(u), cur = idx == null ? null : dot[idx];
    if (idx != null && !cur) return;
    closeSheet();
    const el = document.createElement("div");
    el.id = "dvSheet"; el.className = "dvSheet";
    const tipi = typeof TIPI_DOT_PERS !== "undefined" ? TIPI_DOT_PERS.filter(t => !rowKind({ tipo: t })) : ["Tablet", "Scheda SIM", "Cuffie"];
    el.innerHTML = `<div class="dvCard" role="dialog" aria-label="${cur ? "Apparecchio" : "Nuovo apparecchio"}">
      <h3>${cur ? esc(cur.tipo || "Apparecchio") : "Aggiungi un apparecchio"}</h3>
      <div class="meta">${esc((u && u.nome) || "")}</div>
      <label for="doTipo">Che apparecchio è</label>
      <input id="doTipo" type="text" list="doTipi" value="${esc(cur ? cur.tipo : "")}" autocomplete="off" autocorrect="off" spellcheck="false" placeholder="per esempio Tablet, Scheda SIM, Cuffie">
      <datalist id="doTipi">${tipi.map(t => `<option value="${esc(t)}">`).join("")}</datalist>
      <label for="doMod">Modello (facoltativo)</label>
      <input id="doMod" type="text" value="${esc(cur ? cur.modello : "")}" autocomplete="off" autocorrect="off" spellcheck="false">
      <label for="dvIn">Seriale o IMEI</label>
      <input id="dvIn" type="text" value="${esc(cur ? cur.seriale : "")}" data-scan="any" autocomplete="off" autocapitalize="characters" autocorrect="off" spellcheck="false" placeholder="Numero di serie o IMEI">
      <button class="btn navy" id="dvCam" type="button">📷 Leggi con la fotocamera</button>
      <div class="hint">Il S/N è sull'etichetta dell'apparecchio; per una SIM c'è il numero stampato sulla scheda. Lo scrivi a mano oppure lo leggi con la fotocamera.</div>
      <div id="dvMsg"></div>
      <button class="btn green" id="dvSave" type="button">Salva</button>
      ${cur ? `<button class="btn red" id="dvDel" type="button">Togli l'apparecchio</button>` : ""}
      <button class="btn" id="dvCancel" type="button">Annulla</button></div>`;
    document.body.appendChild(el);
    const inp = $("dvIn"), msg = (t, err) => { $("dvMsg").innerHTML = t ? `<div class="msg${err ? " err" : ""}">${esc(t)}</div>` : ""; };
    if (window.SCAN) window.SCAN.attach(inp);
    $("dvCam").onclick = () => { const b = inp.parentNode.querySelector(".scbtn"); if (b) b.click(); else msg("La fotocamera non è disponibile: scrivi il codice a mano.", true); };
    const finish = async (rows, text) => {
      const fresh = await getKey("utenti", k);
      if (!fresh) return;
      await put("utenti", { ...fresh, dotazione: rows, modificato: nowIso() });
      await loadUsers();
      closeSheet();
      openDesk(ruid, k, from, { saved: text });
    };
    $("dvSave").onclick = async () => {
      const c = checkCode("serial", inp.value);
      if (c.err) { msg(c.err, true); inp.focus(); return; }
      const fresh = await getKey("utenti", k), base = dotList(fresh || u);
      const res = withOther(base, idx, $("doTipo").value, $("doMod").value, c.v);
      if (res.err) { msg(res.err, true); $("doTipo").focus(); return; }
      await finish(res.rows, (res.rows[idx == null ? res.rows.length - 1 : idx].tipo) + (idx == null ? ": apparecchio aggiunto." : ": salvato."));
    };
    if ($("dvDel")) $("dvDel").onclick = async () => {
      const fresh = await getKey("utenti", k), base = dotList(fresh || u).filter((r, i) => i !== idx);
      await finish(base, (cur.tipo || "Apparecchio") + ": tolto dalla dotazione.");
    };
    inp.addEventListener("keydown", ev => { if (ev.key === "Enter") { ev.preventDefault(); $("dvSave").click(); } });
    $("dvCancel").onclick = closeSheet;
  }
  /* la scrivania a pagina intera: la scena in grande con un bollino su ogni dispositivo */
  function openDesk(ruid, k, from, opt) {
    const r = room(ruid), p = peopleIn(ruid).find(x => x.k === k), u = userOf(k);
    if (!r || !p) return;
    const dot = dotList(u), ph = phoneInfo(dot, u && u.imei), devs = deskDevices(u, dot, ph);
    let pg = $("dkPage");
    if (!pg) { pg = document.createElement("div"); pg.id = "dkPage"; pg.className = "dkPage hidden"; document.body.appendChild(pg); }
    const line = (lab, v) => v ? `<div class="meta">${lab}: <b>${esc(v)}</b></div>` : "";
    const dots = devs.map(d => `<button type="button" class="dkDot ${d.value ? "ok" : "no"}" data-d="${d.key}" style="left:${(BOLLINI[d.key][0] / 130 * 100).toFixed(2)}%;top:${(BOLLINI[d.key][1] / 66 * 100).toFixed(2)}%" aria-label="${esc(d.name)}: ${d.value ? esc(d.value) : "codice mancante"}">${d.value ? "✓" : "+"}</button>`).join("");
    const rows = devs.map(d => `<button type="button" class="dvRow" data-d="${d.key}"><span class="dkDot ${d.value ? "ok" : "no"}" aria-hidden="true">${d.value ? "✓" : "+"}</span><span class="dvT"><b>${esc(d.name)}</b>${d.sub ? `<small>${esc(d.sub)}</small>` : ""}</span><span class="dvV${d.value ? "" : " no"}">${d.value ? esc(d.value) : (d.kind === "imei" ? "IMEI mancante" : d.kind === "presa" ? "Codice mancante" : "S/N mancante")}</span></button>`).join("");
    const addRows = deskAddable(u, ph).map(d => `<button type="button" class="dvRow dvAdd" data-d="${d.key}"><span class="dkDot add" aria-hidden="true">+</span><span class="dvT"><b>Aggiungi: ${esc(d.name)}</b>${d.sub ? `<small>${esc(d.sub)}</small>` : ""}</span><span class="dvV go">${d.kind === "imei" ? "scrivi o leggi l'IMEI" : "scrivi o leggi il S/N"}</span></button>`).join("");
    const othRows = deskOthers(dot).map(d => `<button type="button" class="dvRow" data-o="${d.idx}"><span class="dkDot ${d.value ? "ok" : "no"}" aria-hidden="true">${d.value ? "✓" : "+"}</span><span class="dvT"><b>${esc(d.name)}</b>${d.sub ? `<small>${esc(d.sub)}</small>` : ""}</span><span class="dvV${d.value ? "" : " no"}">${d.value ? esc(d.value) : "S/N mancante"}</span></button>`).join("");
    pg.innerHTML = `<div class="dkIn"><button class="dkBack" id="dkClose" type="button">← ${from === "room" ? "Torna alla stanza" : from === "room3d" ? "Torna alla stanza 3D" : "Indietro"}</button>
      <h3>${esc(r.codice)} · Postazione ${p.numero}</h3>
      <div class="meta">${esc(bName(r.edificio_uid))} · ${floorLabel(r.piano)}</div>
      <h2 class="dkWho">${esc(p.utente)}</h2>
      <div class="meta">${u && u.ufficio ? esc(u.ufficio) : "Ufficio non indicato"}</div>
      <div class="dkStage" id="dkScene">${deskSvg("100%", ph, false, true)}${dots}</div>
      <div class="hint" style="text-align:center">Tocca un bollino: <b style="color:var(--green)">✓</b> codice scritto, <b style="color:var(--red)">+</b> codice da inserire.</div>
      ${opt && opt.saved ? `<div class="msg" id="dkSaved">${esc(opt.saved)}</div>` : ""}
      <div class="dvList" id="dvList">${rows}${othRows}${addRows}</div>
      <button class="btn" id="dkOtherAdd" type="button">+ Aggiungi un altro apparecchio (tablet, SIM, cuffie…)</button>
      <button class="btn" id="dkCopy" type="button"${snOf(u) ? "" : " hidden"}>Copia il S/N del computer</button><div id="dkMsg"></div>
      ${line("Telefono", u && u.telefono)}${line("Matricola", u && u.matricola)}${line("Sede", u && u.sede)}
      <div class="dkPics" id="dkPics">${picsHtml(ph)}</div>
      <label for="dkMob">Cellulare sulla scrivania</label>
      <select id="dkMob">${[["", "Non indicato"]].concat(CELL_MODELLI.map(x => [x, x])).concat(ph.mobile === "altro" ? [["__alt", "Altro: " + ph.label]] : []).map(([v, t]) => `<option value="${esc(v)}"${(ph.mobile === "altro" ? v === "__alt" : v === (CELL_MODELLI.includes(ph.label) ? ph.label : "")) ? " selected" : ""}>${esc(t)}</option>`).join("")}</select>
      <label for="dkLap">Portatile sulla scrivania</label>
      <select id="dkLap">${[["", "Non indicato"], ["13", LAPTOP_MODELLI["13"]], ["15", LAPTOP_MODELLI["15"]]].concat(ph.laptop === "altro" ? [["__alt", "Altro: " + ph.laptop_label]] : []).map(([v, t]) => `<option value="${esc(v)}"${(ph.laptop === "altro" ? v === "__alt" : v === (ph.laptop || "")) ? " selected" : ""}>${esc(t)}</option>`).join("")}</select>
      <label class="chk" style="display:flex;gap:8px;align-items:center;margin:8px 0"><input type="checkbox" id="dkTeams"${ph.teams ? " checked" : ""} style="width:auto"> Telefono Teams (AudioCodes) sulla scrivania</label>
      <label>Dotazione</label>
      <div id="dkDot">${dot.length ? dot.map(d => `<div class="card"><b>${esc(d.tipo || "Apparecchio")}</b>${d.modello ? " · " + esc(d.modello) : ""}${d.seriale ? `<div class="meta">${rowKind(d) === "presa" ? "Codice presa" : "S/N o IMEI"}: <b>${esc(d.seriale)}</b></div>` : ""}</div>`).join("") : `<div class="hint">Nessun apparecchio in dotazione.</div>`}</div>
      <button class="btn navy" id="dkEdit" type="button">Modifica la scheda</button>
      <button class="btn" id="dkRoom" type="button">${from === "room" ? "Torna alla stanza" : "Apri la stanza"}</button>
      <button class="btn" id="dkClose2" type="button">Chiudi</button></div>`;
    pg.classList.remove("hidden"); pg.scrollTop = 0;
    const shut = () => { pg.classList.add("hidden"); pg.innerHTML = ""; closeSheet(); };
    pg.querySelectorAll("[data-d]").forEach(b => b.onclick = () => openDevice(ruid, k, from, b.dataset.d));
    pg.querySelectorAll("[data-o]").forEach(b => b.onclick = () => openOther(ruid, k, from, +b.dataset.o));
    $("dkOtherAdd").onclick = () => openOther(ruid, k, from, null);
    if ($("dkCopy")) $("dkCopy").onclick = async () => { const sn = snOf(userOf(k)), ok = await copyText(sn); $("dkMsg").innerHTML = ok ? `<div class="msg">Copiato: ${esc(sn)}</div>` : `<div class="msg err">Non riesco a copiare: tieni premuto sul numero e scegli Copia.</div>`; };
    const applyPhones = async () => {
      const v = $("dkMob").value, cur = (await getKey("utenti", k)) || u;
      const lv = $("dkLap").value, nd = withPhones(dotList(cur), v === "__alt" ? null : v, $("dkTeams").checked, lv === "__alt" ? null : lv);
      await put("utenti", { ...cur, dotazione: nd, modificato: nowIso() });
      await loadUsers(); openDesk(ruid, k, from);
    };
    $("dkMob").onchange = applyPhones; $("dkLap").onchange = applyPhones; $("dkTeams").onchange = applyPhones;
    $("dkEdit").onclick = () => { shut(); editUser(k); };
    $("dkRoom").onclick = () => { shut(); close3d(); openRoom(ruid); };
    const back = () => { shut(); if (from === "room3d") open3d(ruid, true); else render(); };
    $("dkClose").onclick = back; $("dkClose2").onclick = back;
  }
  /* ---------- visita 3D della stanza: la stanza con le sue scrivanie, si gira intorno col dito ---------- */
  const R3 = { yaw: -25, tilt: -52, zoom: 1, moved: 0 };
  const clampN = (v, a, b) => Math.max(a, Math.min(b, v));
  function close3d() { const e = $("r3Page"); if (e) { e.classList.add("hidden"); e.innerHTML = ""; } }
  function open3d(ruid, keep) {
    const r = room(ruid); if (!r) return;
    const ppl = peopleIn(ruid), tot = num(r.scrivanie), cards = [];
    ppl.forEach(p => { const u = userOf(p.k), dot = dotList(u), ph = phoneInfo(dot, u && u.imei); cards.push({ p, ph, devs: deskDevices(u, dot, ph) }); });
    for (let i = ppl.length; i < tot; i++) cards.push({ free: true, n: i + 1 });
    const n = cards.length, cols = n <= 2 ? Math.max(n, 1) : n <= 4 ? 2 : 3, rows = Math.max(1, Math.ceil(n / cols));
    const GX = 190, GZ = 150, W = Math.max(cols * GX + 30, 260), D = rows * GZ + 80, H = 120;
    if (!keep) { R3.yaw = -25; R3.tilt = -52; R3.zoom = clampN((window.innerWidth - 20) / (W * 1.2), 0.5, 1.4); }
    let pg = $("r3Page");
    if (!pg) { pg = document.createElement("div"); pg.id = "r3Page"; pg.className = "r3Page hidden"; document.body.appendChild(pg); }
    const pos = i => ({ x: Math.round(((i % cols) - (cols - 1) / 2) * GX), z: Math.round((Math.floor(i / cols) - (rows - 1) / 2) * GZ + 28) });
    const mats = cards.map((c, i) => { const q = pos(i); return `<i class="r3Mat${c.free ? " free" : ""}" style="left:${Math.round(W / 2 + q.x - 84)}px;top:${Math.round(D / 2 + q.z - 78)}px"></i>`; }).join("");
    const desks = cards.map((c, i) => {
      const q = pos(i), st = `style="--x:${q.x}px;--z:${q.z}px"`;
      if (c.free) return `<button type="button" class="r3Desk free" data-i="${i}" ${st} aria-label="Postazione ${c.n} libera"><span class="r3Name">Libera</span>${deskSvg(150, {}, true)}</button>`;
      return `<button type="button" class="r3Desk" data-i="${i}" ${st} aria-label="Postazione ${c.p.numero}: ${esc(c.p.utente)}"><span class="r3Name">${esc(c.p.utente)}</span>${deskSvg(150, c.ph)}<span class="r3Dots">${c.devs.map(d => `<i class="${d.value ? "ok" : "no"}" title="${esc(d.name)}"></i>`).join("")}</span></button>`;
    }).join("");
    pg.innerHTML = `<div class="r3Top"><button class="dkBack" id="r3Close" type="button">← Indietro</button><b>${esc(r.codice)} · visita 3D</b><span class="r3Zm"><button id="r3Out" type="button" aria-label="Allontana">−</button><button id="r3In" type="button" aria-label="Avvicina">+</button></span></div>
      <div class="r3View" id="r3View"><div class="r3World" id="r3World">
        <div class="r3Floor" style="width:${W}px;height:${D}px;margin:${-D / 2}px 0 0 ${-W / 2}px">${mats}</div>
        <div class="r3Wall" style="width:${W}px;height:${H}px;margin:${-H}px 0 0 ${-W / 2}px;transform:translateZ(${-D / 2}px)"><b>${esc(r.codice)}</b><span>${ppl.length} ${ppl.length === 1 ? "persona" : "persone"}${tot ? " · " + tot + " scrivanie" : ""}</span></div>
        <div class="r3Wall side" style="width:${D}px;height:${H}px;margin:${-H}px 0 0 ${-D / 2}px;transform:translateX(${-W / 2}px) rotateY(90deg)"></div>
        <div class="r3Wall side" style="width:${D}px;height:${H}px;margin:${-H}px 0 0 ${-D / 2}px;transform:translateX(${W / 2}px) rotateY(-90deg)"></div>
        ${desks}</div></div>
      <div class="r3Hint">Trascina per girare intorno alla stanza · due dita o +/− per avvicinarti · tocca una scrivania per aprirla</div>`;
    pg.classList.remove("hidden");
    const world = $("r3World"), view = $("r3View");
    const apply = () => { world.style.setProperty("--yaw", R3.yaw.toFixed(1) + "deg"); world.style.setProperty("--tilt", R3.tilt.toFixed(1) + "deg"); world.style.setProperty("--zm", R3.zoom.toFixed(3)); };
    apply();
    const pts = new Map(); let pinch = 0;
    const span = () => { const a = [...pts.values()]; return Math.hypot(a[0][0] - a[1][0], a[0][1] - a[1][1]); };
    view.onpointerdown = e => { pts.set(e.pointerId, [e.clientX, e.clientY]); R3.moved = 0; if (pts.size === 2) pinch = span(); };
    view.onpointermove = e => {
      const old = pts.get(e.pointerId); if (!old) return;
      pts.set(e.pointerId, [e.clientX, e.clientY]);
      const dx = e.clientX - old[0], dy = e.clientY - old[1]; R3.moved += Math.abs(dx) + Math.abs(dy);
      if (pts.size === 2) { const d = span(); if (pinch) R3.zoom = clampN(R3.zoom * d / pinch, 0.4, 2.2); pinch = d; }
      else { R3.yaw += dx * 0.45; R3.tilt = clampN(R3.tilt + dy * 0.3, -80, -12); }
      apply();
    };
    view.onpointerup = view.onpointercancel = view.onpointerleave = e => { pts.delete(e.pointerId); pinch = 0; };
    $("r3In").onclick = () => { R3.zoom = clampN(R3.zoom * 1.18, 0.4, 2.2); apply(); };
    $("r3Out").onclick = () => { R3.zoom = clampN(R3.zoom / 1.18, 0.4, 2.2); apply(); };
    $("r3Close").onclick = close3d;
    pg.querySelectorAll(".r3Desk").forEach(b => b.onclick = () => {
      if (R3.moved > 8) { R3.moved = 0; return; }               // era un trascinamento, non un tocco
      const c = cards[+b.dataset.i];
      if (c.free) { close3d(); openRoom(ruid); } else openDesk(ruid, c.p.k, "room3d");
    });
  }
  const wireDesks = (box, ruidOf, from) => box.querySelectorAll(".dkc").forEach(b => b.onclick = () => {
    if (b.dataset.free) { if (from === "room") { $("roPick").classList.remove("hidden"); $("roQ").focus(); $("roQ").dispatchEvent(new Event("input")); try { $("roAdd").scrollIntoView({ block: "center" }); } catch (e) { /* vecchi browser */ } } else openRoom(b.dataset.r); }
    else openDesk(b.dataset.r, b.dataset.k, from);
  });

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
    /* scrivanie del piano: una riga per stanza, con le scrivanie disegnate */
    h += `<h2 style="margin-top:14px">Scrivanie del piano</h2>`;
    h += rs.length ? `<div class="bld" style="margin-top:0" id="rsDesks">${rs.map(r => { const o = occ(r), cs = deskCards(r); return `<div class="dkRow${o.over ? " over" : ""}"><button type="button" class="dl" data-r="${esc(r.uid)}" aria-label="Apri la stanza ${esc(r.codice)}">${esc(r.codice)}${o.txt ? `<small>${o.txt}</small>` : ""}</button><div class="dkGrid">${cs.length ? cs.join("") : `<span class="meta" style="align-self:center">Nessuna postazione</span>`}</div></div>`; }).join("")}</div><div class="hint">Tocca una scrivania per vedere il S/N del PC. Tocca il codice della stanza per aprirla.</div>` : `<div class="hint">Nessuna stanza in questo piano.</div>`;
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
    if ($("rsDesks")) { wireDesks($("rsDesks"), null, "floor"); $("rsDesks").querySelectorAll(".dl").forEach(c => c.onclick = () => openRoom(c.dataset.r)); }
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
      <div class="hint">Una stanza per riga. Per indicare anche le postazioni scrivi per esempio <b>T12; 3</b> oppure <b>T12 - 3</b> (o due colonne copiate da Excel). Senza numero crei solo la stanza; più codici sulla stessa riga si separano con la virgola. Le stanze già presenti non si duplicano: se scrivi un numero diverso di postazioni lo aggiorno.</div>
      <label for="bkTxt">Stanze</label><textarea id="bkTxt" style="min-height:140px" autocapitalize="characters" spellcheck="false" placeholder="T1; 2&#10;T2 - 3&#10;T12; 3"></textarea>
      <div class="hint" id="bkInfo"></div>
      <button class="btn green" id="bkOk" type="button">Crea le stanze</button><button class="btn" id="bkNo" type="button">Annulla</button><div id="bkMsg"></div></div>`;
    m.classList.remove("hidden");
    const info = async () => {
      await load();
      const { rooms, bad } = parseRooms($("bkTxt").value), have = new Map(roomsOf(euid, piano).map(r => [cf(r.codice), r]));
      const nu = rooms.filter(x => !have.has(cf(x.c))).length, up = rooms.filter(x => have.has(cf(x.c)) && x.n != null && x.n !== num(have.get(cf(x.c)).scrivanie)).length;
      const con = rooms.filter(x => !have.has(cf(x.c)) && x.n != null).length, gia = rooms.length - nu - up;
      $("bkInfo").textContent = rooms.length || bad.length ? `Verranno create ${nu} stanze nuove` + (con ? ` (${con} con le postazioni indicate)` : "") + (up ? `, ${up} aggiornate nelle postazioni` : "") + (gia ? `, ${gia} già presenti` : "") + (bad.length ? `, ${bad.length} scartate perché più lunghe di ${CODE_MAX} caratteri` : "") + "." : "";
    };
    $("bkTxt").addEventListener("input", info);
    $("bkNo").onclick = () => m.classList.add("hidden");
    $("bkOk").onclick = async () => {
      const r = await addRooms(euid, piano, $("bkTxt").value);
      if (!r.nuove && !r.esistenti && !r.aggiornate && !r.scartate.length) return show($("bkMsg"), "Scrivi almeno un codice.", true);
      m.classList.add("hidden");
      note(`Stanze create: ${r.nuove}` + (r.aggiornate ? `, postazioni aggiornate: ${r.aggiornate}` : "") + (r.esistenti ? `, già presenti: ${r.esistenti}` : "") + (r.scartate.length ? `, scartate perché troppo lunghe: ${r.scartate.join(", ")}` : "") + ".", !r.nuove && !!r.scartate.length);
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
      <label>Scrivanie${o.txt ? ` · <span class="${o.over ? "tag st-Aperto" : "tag"}">${o.txt}</span>` : ""}</label>
      <div class="dkBox" id="roDesks">${deskCards(r).length ? `<div class="dkGrid">${deskCards(r).join("")}</div>` : `<div class="hint" style="padding:8px">Nessuna postazione: aggiungi una persona o scrivi le scrivanie totali.</div>`}</div>
      <div class="hint">Tocca una scrivania per vedere il numero di serie del PC (S/N), l'IMEI e la dotazione. Una scrivania tratteggiata è libera: toccala per aggiungere una persona.</div>
      <button class="btn navy" id="ro3d" type="button">🚪 Entra nella stanza (3D)</button>
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
    wireDesks($("roDesks"), null, "room");
    $("ro3d").onclick = () => open3d(r.uid);
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
        await put("utenti", { key: keyOf(raw), nome: raw, ufficio: "", dotazione: [], matricola: "", sede: b.sede || "", telefono: "", seriale_pc: "", imei: "", modificato: nowIso() });
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
    back(res.gia ? `${u.nome} è già in questa stanza (Postazione ${res.numero}).` : `${u.nome}: ${r.codice} · Postazione ${res.numero}.` + (res.da ? `\nPrima era in ${res.da.codice}: le sue postazioni sono state rinumerate.` : "") + (res.presa_tolta ? `\nLa presa di rete ${res.presa_tolta} era della scrivania di prima: l'ho tolta, scrivi quella nuova.` : "") + (res.sopra ? `\nAttenzione: ora ci sono ${oc.n} persone ma le scrivanie sono ${oc.tot}.` : ""), res.sopra);
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
        if (res.presa_tolta) $("eDoveTxt").innerHTML += `<br>La presa di rete ${esc(res.presa_tolta)} era della scrivania di prima: l'ho tolta.`;
      };
    }
  }

  return { load, render, openRoom, openDesk, open3d, close3d, openFloor, floorMenu, dove, parseCodes, parseRooms, cleanCode,
           addRoom, addRooms, saveRoom, setPos, deleteRoom, assign, unassign, dropUser, renameUser,
           addFoto, delFoto, setPianta, delPianta, pianta, fotosOf, exportItems, countNew, importData,
           peopleIn, seatOf, roomsOf, occ, lab, roomCode: uid => { const r = uid && room(uid); return r ? r.codice : ""; },
           _x: { room, userOf, dotList, phoneInfo, deskDevices, deskOthers, deskAddable, withOther, bName, open3d },
           _state: () => ({ S, P, curF, view, mark, markUid, zoom }) };
})();
