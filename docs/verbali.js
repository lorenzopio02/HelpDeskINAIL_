"use strict";
/* Verbali sul telefono: modulo, firma col dito, anteprima, PDF e scambio con il PC.
   Si appoggia alle funzioni di index.html (database, menu a tendina, finestre). */
(function () {
  /* ---------- i quattro verbali (stessi dati e stesse chiavi del programma del PC) ---------- */
  const TIPI = {
    trasferimento: { titolo: "Trasferimento apparato", breve: "Trasferimento", foto: false,
      firme: [["tecnico", "Firma Tecnico di Presidio", true]] },
    presa_in_carico: { titolo: "Consegna e presa in carico", breve: "Presa in carico", foto: false,
      firme: [["utente", "Firma utente consegnatario", true], ["tecnico", "Firma tecnico CED/IT", true]] },
    assegnazione: { titolo: "Assegnazione materiale", breve: "Assegnazione", foto: false,
      firme: [["tecnico", "Firma tecnico CED/IT", true], ["utente", "Firma utente per ricevuta (facoltativa)", false]] },
    censimento: { titolo: "Censimento materiale", breve: "Censimento", foto: true,
      firme: [["tecnico", "Firma Tecnico di PRESIDIO", true]] }
  };
  const ORDINE = ["trasferimento", "presa_in_carico", "assegnazione", "censimento"];
  const MOTIVI = {
    trasferimento: ["Apparato non utilizzato", "Riallocazione risorse", "Sostituzione apparato", "Riparazione/Manutenzione", "Pensionamento Utente"],
    presa_in_carico: ["Sostituzione dispositivo", "Guasto tecnico", "Licenziamento", "Pensionamento utente", "Dismissione"]
  };
  const TIPOLOGIE = ["Notebook", "PC fisso", "Monitor", "Smartphone", "Tablet", "Stampante", "Altro"];
  const TIPI_DOT = ["Notebook aziendale", "Cuffie professionali", "Smartphone aziendale", "Scheda SIM aziendale", "Altro"];
  const COLONNE = ["IMEI", "Numero di serie", "IMEI / Seriale"];
  const TITOLI_LISTA = ["", "avv.", "dott.", "dott.ssa", "ing.", "arch.", "geom.", "rag.", "sig.", "sig.ra", "prof.", "prof.ssa"];
  const TITOLI = { "": "a", "avv.": "all’", "dott.": "al", "dott.ssa": "alla", "ing.": "all’", "arch.": "all’", "geom.": "al", "rag.": "al", "sig.": "al", "sig.ra": "alla", "prof.": "al", "prof.ssa": "alla", "on.": "all’", "dr.": "al", "dr.ssa": "alla" };
  const PC_WORDS = ["notebook", "pc", "portatile", "computer", "desktop"], PHONE_WORDS = ["smartphone", "cellulare", "telefono", "tablet"];
  const LISTE = ["motivi", "materiale", "accessori"];
  const BLANK = "__________";
  const SET_KEYS = ["tecnico", "sede_citta", "indirizzo", "cap"];

  /* ---------- utilità ---------- */
  const isoToIt = v => { const m = /^(\d{4})-(\d\d)-(\d\d)/.exec(v || ""); return m ? `${m[3]}/${m[2]}/${m[1]}` : ""; };
  function itToIso(v) {
    v = (v || "").trim();
    let m = /^(\d{4})-(\d\d)-(\d\d)/.exec(v);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = /^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/.exec(v);
    if (!m) return "";
    const y = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${y}-${pad(+m[2])}-${pad(+m[1])}`;
  }
  const str = v => (v == null ? "" : String(v)).trim();
  const lowerFirst = s => s ? s[0].toLowerCase() + s.slice(1) : s;
  const isPcWord = t => PC_WORDS.some(w => t.includes(w)), isPhoneWord = t => PHONE_WORDS.some(w => t.includes(w));

  /* controlli sui codici (come nel programma del PC) */
  const luhn = s => { let t = 0; [...s].reverse().forEach((ch, i) => { let d = +ch; if (i % 2 === 1) { d *= 2; if (d > 9) d -= 9; } t += d; }); return t % 10 === 0; };
  function imeiProblem(v) {
    const s = (v || "").replace(/[\s-]/g, "");
    if (!s || !/^\d+$/.test(s)) return null;
    if (s.length !== 15) return `ha ${s.length} cifre (un IMEI ne ha 15)`;
    if (!luhn(s)) return "non supera il controllo di validità, forse c'è un errore di battitura";
    return null;
  }
  function controllaApparati(rows, colonna) {
    const out = [], seen = {}, isImei = (colonna || "IMEI").trim().toUpperCase() === "IMEI";
    rows.forEach((r, idx) => {
      const i = idx + 1, code = str(r.codice);
      if (!code) { out.push(`Riga ${i}: manca il codice (${colonna || "IMEI"}).`); return; }
      const key = code.replace(/[\s-]/g, "").toLowerCase();
      if (key in seen) out.push(`Riga ${i}: il codice ${code} è uguale a quello della riga ${seen[key]}.`); else seen[key] = i;
      if (isImei) { const why = imeiProblem(code); if (why) out.push(`Riga ${i}: l'IMEI ${code} ${why}.`); }
    });
    return out;
  }
  function parseApparati(text, marca, modello) {
    const rows = [];
    for (let line of (text || "").split(/\r?\n/)) {
      line = line.trim(); if (!line) continue;
      let parts = line.split(/\t|;/).map(p => p.trim());
      if (parts.length === 1 && line.includes(",")) parts = line.split(",").map(p => p.trim());
      if (parts.length === 1) rows.push({ marca, modello, codice: parts[0] });
      else if (parts.length === 2) rows.push({ marca, modello: parts[0], codice: parts[1] });
      else rows.push({ marca: parts[0], modello: parts[1], codice: parts[2] });
    }
    return rows;
  }
  function fraseAssegnatario(titolo, nome) {
    titolo = str(titolo); nome = str(nome);
    if (!titolo) return nome ? ("a " + nome) : "";
    const prep = TITOLI[titolo.toLowerCase()] || "a";
    return prep.endsWith("’") ? `${prep}${titolo} ${nome}` : `${prep} ${titolo} ${nome}`;
  }
  function righeDotazione(it) {
    const tipo = str(it.tipo).toLowerCase(), cod = str(it.codice) || BLANK;
    if (tipo.includes("sim")) return ["ICCID: " + (str(it.iccid) || BLANK), "Numero: " + (str(it.numero) || BLANK)];
    if (tipo.includes("smartphone") || tipo.includes("cellulare") || tipo.includes("telefono")) return ["IMEI: " + cod];
    return ["S/N: " + cod];
  }

  /* riga dell'elenco: (soggetto, dettaglio) */
  function riassunto(tipo, d) {
    const j = (...p) => p.map(str).filter(Boolean).join(" · ");
    const g = k => str(d[k]), sn = g("seriale") ? "S/N " + g("seriale") : "";
    if (tipo === "trasferimento") return [g("assegnatario") || g("tipologia"), j(g("tipologia"), g("marca_modello"), sn, [g("provenienza"), g("destinazione")].filter(Boolean).join(" → "))];
    if (tipo === "presa_in_carico") return [g("utente"), j(g("tipologia"), g("marca_modello"), sn, j(g("inc") ? "INC " + g("inc") : "", g("ric") ? "RIC " + g("ric") : ""))];
    if (tipo === "assegnazione") return [g("utente"), (d.dotazione || []).map(x => j(x.tipo, x.modello)).filter(Boolean).join(", ")];
    if (tipo === "censimento") return [g("oggetto"), `${(d.apparati || []).length} apparecchi` + (g("tecnico") ? " · " + g("tecnico") : "")];
    return ["", ""];
  }
  const slotsOf = tipo => TIPI[tipo].firme;
  const needSigned = tipo => slotsOf(tipo).filter(s => s[2]).map(s => s[0]);
  const isSigned = v => needSigned(v.tipo).every(k => v.firme && v.firme[k]);

  /* dati puliti (come fa il PC quando riceve un verbale) */
  function cleanDati(d) {
    const out = {};
    for (const [k, v] of Object.entries(d || {})) {
      if (k.startsWith("_")) continue;
      if (LISTE.includes(k)) out[k] = (Array.isArray(v) ? v : []).map(str).filter(Boolean);
      else if (k === "apparati") out[k] = (Array.isArray(v) ? v : []).filter(r => r && typeof r === "object").map(r => ({ marca: str(r.marca), modello: str(r.modello), codice: str(r.codice) }));
      else if (k === "dotazione") out[k] = (Array.isArray(v) ? v : []).filter(r => r && typeof r === "object").map(r => ({ tipo: str(r.tipo), modello: str(r.modello), codice: str(r.codice), iccid: str(r.iccid), numero: str(r.numero) }));
      else out[k] = str(v);
    }
    for (const k of ["data", "data_sped"]) if (out[k] && /^\d{4}-\d\d-\d\d/.test(out[k])) out[k] = isoToIt(out[k]);
    return out;
  }
  const normV = x => Array.isArray(x) ? (x.length ? JSON.stringify(x) : "") : (x == null ? "" : String(x));
  function sameDati(a, b) {
    const keys = new Set([...Object.keys(a || {}), ...Object.keys(b || {})].filter(k => !k.startsWith("_")));
    for (const k of keys) if (normV((a || {})[k]) !== normV((b || {})[k])) return false;
    return true;
  }
  const fileStem = v => {
    const [who] = riassunto(v.tipo, v.dati), day = (v.dati.data || "").replace(/\//g, "-");
    return ("Verbale " + TIPI[v.tipo].titolo.toLowerCase() + (who ? " - " + who.replace(/\s*\(.*?\)/g, "").trim() : "") + (day ? " - " + day : "")).replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 110);
  };

  /* ---------- impostazioni (tecnico, sede) e firma salvata del tecnico ---------- */
  let settings = {}, myFirma = "";
  async function loadSettings() {
    settings = (await getMeta("verbSet")) || {};
    myFirma = (await getMeta("firmaTecnico")) || "";
  }
  const saveSettings = () => setMeta("verbSet", settings);

  /* ---------- quello che hai già scritto nei verbali precedenti (stessa logica del programma del PC) ---------- */
  const CAMPI_STORICO = ["tipologia", "marca_modello", "provenienza", "destinazione", "contatti", "titolo", "oggetto", "descrizione", "sede_citta", "indirizzo", "cap", "tecnico", "collega", "materiale", "accessori", "dot_modello", "ap_marca", "ap_modello"];
  const nomePulito = x => String(x || "").replace(/\s*\(matricola[^)]*\)\s*$/i, "").trim();
  const codeKey = x => String(x || "").replace(/[\s-]/g, "").toLowerCase();
  function tipDaDot(tipo) {
    const x = String(tipo || "").toLowerCase();
    if (/notebook|portatile/.test(x)) return "Notebook";
    if (/smartphone|cellulare|telefono/.test(x)) return "Smartphone";
    if (x.includes("tablet")) return "Tablet";
    if (/pc|desktop|computer/.test(x)) return "PC fisso";
    return "";
  }
  const groupOf = tip => { tip = String(tip || "").toLowerCase(); return isPcWord(tip) ? "pc" : isPhoneWord(tip) ? "phone" : ""; };
  const uniq = arr => { const seen = new Set(), o = []; for (let x of arr) { x = str(x); const k = x.toLowerCase(); if (x && !seen.has(k)) { seen.add(k); o.push(x); } } return o; };
  /* rows = [{tipo, dati}] dal più recente al più vecchio */
  function storico(rows) {
    const out = { valori: {}, dispositivi: [], titoli: {}, sedi: {}, oggetti: {}, modelli: {} }, seen = {}, devSeen = new Set();
    CAMPI_STORICO.forEach(k => { out.valori[k] = []; seen[k] = new Set(); });
    const valore = (campo, text) => { text = typeof text === "string" ? text.trim() : ""; const k = text.toLowerCase(); if (text && !seen[campo].has(k)) { seen[campo].add(k); out.valori[campo].push(text); } };
    const dispositivo = (tip, mm, ser, ut) => {
      [tip, mm, ser, ut] = [tip, mm, ser, ut].map(x => typeof x === "string" ? x.trim() : "");
      if (!(ser || mm)) return;
      const key = ser ? codeKey(ser) : [tip, mm, ut].map(x => x.toLowerCase()).join("|");
      if (devSeen.has(key)) return;
      devSeen.add(key);
      out.dispositivi.push({ tipologia: tip, marca_modello: mm, seriale: ser, utente: ut });
    };
    for (const row of rows || []) {
      const tipo = row && row.tipo, d = row && row.dati && typeof row.dati === "object" ? row.dati : {};
      const g = k => typeof d[k] === "string" ? d[k].trim() : "", arr = k => Array.isArray(d[k]) ? d[k] : [];
      for (const k of ["tipologia", "marca_modello", "provenienza", "destinazione", "contatti", "titolo", "oggetto", "descrizione", "sede_citta", "indirizzo", "cap", "tecnico", "collega"]) valore(k, g(k));
      if (tipo === "trasferimento") dispositivo(g("tipologia"), g("marca_modello"), g("seriale"), g("assegnatario"));
      else if (tipo === "presa_in_carico") {
        dispositivo(g("tipologia"), g("marca_modello"), g("seriale"), nomePulito(g("utente")));
        arr("materiale").forEach(x => valore("materiale", x));
      } else if (tipo === "assegnazione") {
        const nome = g("utente");
        if (g("titolo") && nome && !(nome.toLowerCase() in out.titoli)) out.titoli[nome.toLowerCase()] = g("titolo");
        for (const it of arr("dotazione")) {
          if (!it || typeof it !== "object") continue;
          valore("dot_modello", it.modello);
          const tip = tipDaDot(it.tipo);
          if (tip) dispositivo(tip, it.modello, it.codice, nome);
        }
        arr("accessori").forEach(x => valore("accessori", x));
      } else if (tipo === "censimento") {
        const citta = g("sede_citta");
        if (citta && !(citta.toLowerCase() in out.sedi) && (g("indirizzo") || g("cap"))) out.sedi[citta.toLowerCase()] = { indirizzo: g("indirizzo"), cap: g("cap") };
        if (g("oggetto") && !(g("oggetto").toLowerCase() in out.oggetti)) out.oggetti[g("oggetto").toLowerCase()] = { descrizione: g("descrizione") };
        for (const r of arr("apparati")) {
          if (!r || typeof r !== "object") continue;
          const marca = str(r.marca), modello = str(r.modello);
          valore("ap_marca", marca); valore("ap_modello", modello);
          if (modello && marca && !(modello.toLowerCase() in out.modelli)) out.modelli[modello.toLowerCase()] = marca;
        }
      }
    }
    return out;
  }
  function dispositiviUtente(st, nome, u) {
    const key = nomePulito(nome).toLowerCase(), res = st.dispositivi.filter(x => x.utente.toLowerCase() === key).map(x => ({ ...x }));
    const seen = new Set(res.filter(x => x.seriale).map(x => codeKey(x.seriale)));
    for (const [tip, ser0] of [["Notebook", u && u.seriale_pc], ["Smartphone", u && u.imei]]) {
      const ser = str(ser0);
      if (ser && !seen.has(codeKey(ser))) { seen.add(codeKey(ser)); res.push({ tipologia: tip, marca_modello: "", seriale: ser, utente: nomePulito(nome) }); }
    }
    return res;
  }
  function trovaDispositivo(st, seriale, usrs) {
    const key = codeKey(seriale); if (!key) return null;
    const x = st.dispositivi.find(d => codeKey(d.seriale) === key); if (x) return { ...x };
    for (const u of usrs || []) for (const [tip, ser] of [["Notebook", u.seriale_pc], ["Smartphone", u.imei]]) if (ser && codeKey(ser) === key) return { tipologia: tip, marca_modello: "", seriale: str(ser), utente: u.nome || "" };
    return null;
  }
  const etichetta = x => [x.tipologia, x.marca_modello, x.seriale ? "S/N " + x.seriale : "", x.utente ? `(${x.utente})` : ""].filter(Boolean).join(" · ");
  let H = storico([]);
  const histVals = (campo, base = []) => uniq([...base, ...(H.valori[campo] || [])]);
  const histItems = (campo, base = []) => () => histVals(campo, base).map(text => ({ text }));

  /* ---------- il testo del verbale (come nei modelli Word) ---------- */
  function docOf(v) {
    const d = v.dati, g = k => str(d[k]), t = v.tipo, B = BLANK, kv = (k, x) => ({ t: "kv", k, v: x || B });
    const checks = (opts, sel) => ({ t: "checks", items: opts.map(o => ({ s: o, on: (sel || []).includes(o) })).concat([{ s: "Altro: " + (g("altro") || B), on: !!g("altro") }]) });
    const b = [];
    if (t === "trasferimento") {
      b.push({ t: "title", s: "VERBALE DI TRASFERIMENTO APPARATO INFORMATICO" },
        { t: "p", s: `In data ${g("data") || B} si dispone il trasferimento del seguente bene informatico:` },
        { t: "h", s: "DATI APPARATO" }, kv("Tipologia", g("tipologia")), kv("Marca/Modello", g("marca_modello")), kv("Numero di Serie", g("seriale")),
        { t: "h", s: "SEDE/UFFICIO DI PROVENIENZA" }, kv("Struttura", g("provenienza")),
        { t: "h", s: "SEDE/UFFICIO DI DESTINAZIONE" }, kv("Struttura", g("destinazione")),
        { t: "h", s: "ASSEGNATARIO (se presente)" }, kv("Nome e Cognome", g("assegnatario")), kv("Contatti", g("contatti")),
        { t: "h", s: "MOTIVAZIONE DEL TRASFERIMENTO" }, checks(MOTIVI.trasferimento, d.motivi));
      if (g("note")) b.push({ t: "h", s: "NOTE" }, { t: "p", s: g("note") });
      b.push({ t: "p", s: "L'apparato viene trasferito completo degli accessori eventualmente disponibili e nelle condizioni d'uso rilevate al momento della presa in carico per la spedizione." },
        { t: "sigs", left: "Data spedizione: " + (g("data_sped") || g("data") || B) });
    } else if (t === "presa_in_carico") {
      b.push({ t: "title", s: "VERBALE DI CONSEGNA E PRESA IN CARICO DISPOSITIVO INFORMATICO" },
        { t: "p", s: "Il sottoscritto dichiara di consegnare al Servizio CED/IT il seguente dispositivo aziendale per presa in carico tecnica." },
        { t: "h", s: "DATI DEL DISPOSITIVO" }, kv("Tipologia dispositivo", g("tipologia")), kv("Marca e modello", g("marca_modello")), kv("Numero seriale", g("seriale")),
        kv("Utente assegnatario", g("utente")), kv("N° INC", g("inc")), kv("N° RIC", g("ric")));
      if ((d.materiale || []).length) b.push({ t: "h", s: "ELENCO MATERIALE CONSEGNATO" }, { t: "list", items: d.materiale });
      b.push({ t: "h", s: "MOTIVAZIONE DELLA CONSEGNA" }, checks(MOTIVI.presa_in_carico, d.motivi),
        { t: "p", s: "Il dispositivo viene consegnato al personale CED/IT per la sostituzione, riconfigurazione, dismissione o gestione inventariale secondo le procedure aziendali previste." },
        { t: "p", s: "A seguito della presa in carico, il dispositivo verrà custodito e conservato presso i locali del CED/IT in ambiente chiuso e accessibile esclusivamente al personale autorizzato." },
        { t: "p", s: "Con la firma del presente verbale, l’utente dichiara di aver consegnato il dispositivo sopra indicato completo di eventuali accessori in dotazione e solleva il personale tecnico da responsabilità successive alla presa in carico." },
        { t: "p", s: "N.B. Il Servizio CED/IT prende in carico il dispositivo nello stato in cui viene consegnato." },
        { t: "sigs", left: "Data consegna: " + (g("data") || B) });
    } else if (t === "assegnazione") {
      b.push({ t: "title", s: "VERBALE DI ASSEGNAZIONE MATERIALE INFORMATICO" },
        { t: "p", s: `Con la presente si certifica che, in data odierna, ${fraseAssegnatario(g("titolo"), g("utente")) || B} con matricola: ${g("matricola") || B} è stato assegnato il seguente materiale aziendale:` },
        { t: "h", s: "Dotazione assegnata" });
      for (const it of d.dotazione || []) b.push({ t: "p", s: (str(it.tipo) + (str(it.modello) ? " " + str(it.modello) : "")).trim(), bold: true }, { t: "p", s: righeDotazione(it).join("\n") });
      if ((d.accessori || []).length) b.push({ t: "h", s: "Accessori inclusi" }, { t: "list", items: d.accessori });
      b.push({ t: "p", s: "Si attesta che il materiale sopra elencato è stato consegnato in buono stato, completo di tutti gli accessori e perfettamente funzionante." },
        { t: "sigs", left: "Data Assegnazione: " + (g("data") || B) });
    } else if (t === "censimento") {
      const citta = g("sede_citta") || settings.sede_citta || "", ind = g("indirizzo") || settings.indirizzo || "", cap = g("cap") || settings.cap || "";
      const tecnico = g("tecnico") || settings.tecnico || "", col = g("col_codice") || "IMEI";
      const descr = g("descrizione") || lowerFirst(g("oggetto"));
      const righe = d.apparati || [];
      b.push({ t: "p", s: "Oggetto: Verbale di censimento materiale informatico rinvenuto presso il CED - " + (g("oggetto") || B), bold: true });
      if (citta) b.push({ t: "p", s: "Sede INAIL " + citta });
      if (ind) b.push({ t: "p", s: ind });
      if (cap || citta) b.push({ t: "p", s: [cap, citta].filter(Boolean).join(", ") });
      b.push({ t: "p", s: `In data ${g("data") || B}, nell'ambito delle attività di riordino e ricognizione del locale CED della sede INAIL di ${citta || B}, il sottoscritto ${tecnico || B}, tecnico del presidio informatico, ha rinvenuto materiale elettronico non in uso costituito da ${descr || B}.` },
        { t: "p", s: "Al fine di garantire la tracciabilità dei beni presenti presso il locale CED, è stato effettuato il censimento dei dispositivi rinvenuti, con rilevazione dei relativi codici IMEI ove disponibili." },
        { t: "p", s: "Si precisa che il presente documento ha esclusivamente finalità di ricognizione e inventariazione del materiale rinvenuto durante le attività di riordino del locale tecnico." });
      if (g("collega")) b.push({ t: "p", s: `A seguito delle indicazioni ricevute dal collega ${g("collega")}, il materiale è stato individuato quale materiale destinato alle procedure di dismissione e smaltimento RAEE secondo le procedure vigenti.` });
      b.push({ t: "p", s: "Il presente verbale viene redatto al fine di documentare l'effettiva presenza dei dispositivi sottoelencati alla data del censimento. L'elenco allegato costituisce evidenza del materiale rinvenuto e censito durante le attività di riordino del CED." },
        { t: "p", s: "Il sottoscritto attesta esclusivamente la presenza dei beni sotto riportati alla data della ricognizione. Eventuali successive movimentazioni, trasferimenti, prelievi, dismissioni o mancanze riscontrate dopo la redazione del presente documento non sono riconducibili alle attività di censimento e inventariazione effettuate dal sottoscritto." },
        { t: "h", s: "Elenco materiale censito" },
        { t: "table", head: ["N.", "Marca", "Modello", col], rows: righe.map((r, i) => [String(i + 1), str(r.marca), str(r.modello), str(r.codice)]), weights: [1, 4, 4, 6] },
        { t: "p", s: `Totale dispositivi censiti: n. ${righe.length}`, bold: true },
        { t: "p", s: "Il presente verbale viene conservato agli atti quale evidenza delle attività di ricognizione e censimento svolte dal sottoscritto, attestando la presenza e la consistenza del materiale rinvenuto presso il CED alla data del rilevamento. La documentazione allegata costituisce riferimento ufficiale dello stato dei beni censiti al momento della verifica." },
        { t: "sigs", left: "Data: " + (g("data") || B) });
    }
    return b;
  }

  /* anteprima nella pagina */
  function docHtml(v) {
    const sl = slotsOf(v.tipo);
    return docOf(v).map(b => {
      if (b.t === "title") return `<div class="dt">${esc(b.s)}</div>`;
      if (b.t === "h") return `<div class="dh">${esc(b.s)}</div>`;
      if (b.t === "p") return `<p${b.bold ? ' class="bd"' : ""}>${esc(b.s).replace(/\n/g, "<br>")}</p>`;
      if (b.t === "kv") return `<div class="dk"><b>${esc(b.k)}:</b> ${esc(b.v)}</div>`;
      if (b.t === "checks") return `<div class="dm">${b.items.map(i => `<div>${i.on ? "☒" : "☐"} ${esc(i.s)}</div>`).join("")}</div>`;
      if (b.t === "list") return `<ul>${b.items.map(i => `<li>${esc(i)}</li>`).join("")}</ul>`;
      if (b.t === "table") return `<table class="dtb"><tr>${b.head.map(h => `<th>${esc(h)}</th>`).join("")}</tr>${b.rows.map(r => `<tr>${r.map(c => `<td>${esc(c)}</td>`).join("")}</tr>`).join("")}</table>`;
      if (b.t === "sigs") return `<div class="dk" style="margin-top:12px"><b>${esc(b.left)}</b></div><div class="sg">${sl.map(([k, label, req]) => {
        const img = v.firme && v.firme[k];
        return `<div class="slot" data-k="${k}"><div class="sgl">${esc(label)}</div>${img ? `<img src="${img}" alt="firma">` : `<div class="sgn">${req ? "da firmare" : "facoltativa"}</div>`}
          <button class="btn navy sm" type="button" data-sign="${k}">${img ? "Rifai la firma" : "✍ Firma qui"}</button>
          ${k === "tecnico" && !img && myFirma ? `<button class="btn sm" type="button" data-mine="1">Usa la mia firma</button>` : ""}
          ${img ? `<button class="btn sm" type="button" data-unsign="${k}">Togli</button>` : ""}</div>`;
      }).join("")}</div>`;
      return "";
    }).join("");
  }

  /* ---------- finestra della firma ---------- */
  function signPad(who, label) {
    return new Promise(resolve => {
      const ov = document.createElement("div");
      ov.className = "sigm";
      ov.innerHTML = `<div class="sigc"><h3>${esc(label)}</h3>
        <div class="hint" style="margin-top:0">${who === "utente" ? "Passa il telefono a chi deve firmare. " : ""}Firma con il dito dentro il riquadro.</div>
        <div class="sigbox"><canvas></canvas><span class="sigline"></span><span class="sigx">✕</span></div>
        ${who === "tecnico" ? `<label class="chk"><input type="checkbox" id="sgKeep" ${myFirma ? "" : "checked"}> Ricorda come mia firma (poi basta un tocco)</label>` : ""}
        <div id="sgMsg"></div>
        <div class="row"><div><button class="btn" type="button" id="sgClear">Cancella</button></div><div><button class="btn" type="button" id="sgCancel">Annulla</button></div></div>
        <button class="btn green" type="button" id="sgOk">✔ Conferma la firma</button></div>`;
      document.body.appendChild(ov);
      const c = ov.querySelector("canvas"), box = ov.querySelector(".sigbox"), g = c.getContext("2d");
      const dpr = Math.min(2, window.devicePixelRatio || 1);
      const r = box.getBoundingClientRect();
      c.width = Math.round(r.width * dpr); c.height = Math.round(r.height * dpr);
      g.scale(dpr, dpr); g.lineCap = "round"; g.lineJoin = "round"; g.strokeStyle = "#0b1f4b";
      let drawing = false, last = null, mid = null, ink = 0, bb = null, pid = null;
      const pos = e => { const q = c.getBoundingClientRect(); return { x: e.clientX - q.left, y: e.clientY - q.top }; };
      const grow = (p, w) => { bb = bb ? { a: Math.min(bb.a, p.x - w), b: Math.min(bb.b, p.y - w), c: Math.max(bb.c, p.x + w), d: Math.max(bb.d, p.y + w) } : { a: p.x - w, b: p.y - w, c: p.x + w, d: p.y + w }; };
      const width = e => e.pointerType === "pen" && e.pressure ? 1.6 + e.pressure * 2.2 : 2.7;
      c.addEventListener("pointerdown", e => {
        e.preventDefault(); drawing = true; pid = e.pointerId; { const mg = ov.querySelector("#sgMsg"); if (mg) mg.innerHTML = ""; } try { c.setPointerCapture(pid); } catch (x) { /* niente */ }
        last = pos(e); mid = last; g.lineWidth = width(e);
        g.beginPath(); g.arc(last.x, last.y, g.lineWidth / 2, 0, 7); g.fillStyle = "#0b1f4b"; g.fill(); grow(last, 3); ink += 2;
      });
      c.addEventListener("pointermove", e => {
        if (!drawing || e.pointerId !== pid) return;
        e.preventDefault();
        const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
        for (const ev of (evs.length ? evs : [e])) {
          const p = pos(ev), m = { x: (last.x + p.x) / 2, y: (last.y + p.y) / 2 };
          g.lineWidth = width(ev);
          g.beginPath(); g.moveTo(mid.x, mid.y); g.quadraticCurveTo(last.x, last.y, m.x, m.y); g.stroke();
          ink += Math.hypot(p.x - last.x, p.y - last.y); grow(p, 3);
          last = p; mid = m;
        }
      });
      const end = e => { if (!drawing || (e && e.pointerId !== pid)) return; drawing = false; try { c.releasePointerCapture(pid); } catch (x) { /* niente */ } g.beginPath(); g.moveTo(mid.x, mid.y); g.lineTo(last.x, last.y); g.stroke(); };
      c.addEventListener("pointerup", end); c.addEventListener("pointercancel", end); c.addEventListener("pointerleave", end);
      ["touchstart", "touchmove"].forEach(n => c.addEventListener(n, e => e.preventDefault(), { passive: false }));
      const done = v => { ov.remove(); resolve(v); };
      ov.querySelector("#sgClear").onclick = () => { g.clearRect(0, 0, c.width / dpr, c.height / dpr); ink = 0; bb = null; $("sgMsg").innerHTML = ""; };
      ov.querySelector("#sgCancel").onclick = () => done(null);
      ov.querySelector("#sgOk").onclick = () => {
        if (!bb || ink < 45) return show(ov.querySelector("#sgMsg"), "La firma è troppo breve: firma dentro il riquadro.", true);
        const m = 6, sx = Math.max(0, Math.floor((bb.a - m) * dpr)), sy = Math.max(0, Math.floor((bb.b - m) * dpr));
        const sw = Math.min(c.width - sx, Math.ceil((bb.c - bb.a + 2 * m) * dpr)), sh = Math.min(c.height - sy, Math.ceil((bb.d - bb.b + 2 * m) * dpr));
        const o = document.createElement("canvas"); o.width = Math.max(1, sw); o.height = Math.max(1, sh);
        o.getContext("2d").drawImage(c, sx, sy, sw, sh, 0, 0, sw, sh);
        const keep = who === "tecnico" && ov.querySelector("#sgKeep") && ov.querySelector("#sgKeep").checked;
        done({ data: o.toDataURL("image/png"), keep });
      };
    });
  }

  /* ---------- foto e PDF ---------- */
  async function pdfOf(v) {
    const sl = slotsOf(v.tipo), sigs = [];
    for (const [k, label] of sl) sigs.push({ label, img: v.firme && v.firme[k] ? await HDExport.toJpegImg(v.firme[k], 700, 0.92) : null });
    const photos = [];
    for (const f of v.foto || []) { try { photos.push(await HDExport.toJpegImg(f.data, 1400, 0.82)); } catch (e) { /* foto illeggibile: salto */ } }
    const blocks = docOf(v).map(b => b.t === "sigs" ? { ...b, slots: sigs } : b);
    return HDExport.makeDocPdf({ blocks, logo: await HDExport.loadLogo(), photos });
  }

  /* ---------- elenco ---------- */
  let cache = [];
  const fmtD = v => (v && /^\d{2}\/\d{2}\/\d{4}$/.test(v)) ? v : (isoToIt(v) || v || "");
  const sortKey = v => itToIso(v.dati.data) + v.modificato;
  const isNewV = (v, last) => (!last || v.modificato > last) && v.modificato !== v.daPC;
  async function render() {
    await loadSettings();
    const last = await getMeta("lastSent");
    cache = (await getAll("verbali")).sort((a, b) => sortKey(b).localeCompare(sortKey(a)));
    const t = $("vTipo").value, q = words($("vSearch").value);
    const list = cache.filter(v => {
      if (t && v.tipo !== t) return false;
      const [s, dt] = riassunto(v.tipo, v.dati);
      const hay = [TIPI[v.tipo].titolo, s, dt, v.dati.data, v.dati.matricola].join(" ").toLowerCase();
      return q.every(w => hay.includes(w));
    });
    $("vCount").textContent = `Verbali: ${list.length} su ${cache.length}. Da inviare al PC: ${cache.filter(v => isNewV(v, last)).length}.`;
    $("vList").innerHTML = list.map(v => {
      const [s, dt] = riassunto(v.tipo, v.dati), signed = isSigned(v), n = (v.foto || []).length;
      return `<div class="card" data-u="${esc(v.uid)}"><h3>${esc(s || TIPI[v.tipo].titolo)}</h3>
        <div class="meta">${esc(TIPI[v.tipo].titolo)} · ${esc(fmtD(v.dati.data))}</div>
        ${dt ? `<div>${esc(dt)}</div>` : ""}
        <div style="margin-top:5px"><span class="tag st-${signed ? "Risolto" : "Aperto"}">${signed ? "✍ Firmato" : "Da firmare"}</span>${n ? `<span class="tag">📷 ${n}</span>` : ""}${isNewV(v, last) ? `<span class="tag st-In-attesa">Da inviare al PC</span>` : `<span class="tag">Già nel PC</span>`}</div></div>`;
    }).join("") || `<div class="hint">Nessun verbale${cache.length ? " con questi filtri" : " ancora: usa uno dei quattro pulsanti qui sopra"}.</div>`;
    $("vList").querySelectorAll(".card").forEach(c => c.onclick = () => showVer(c.dataset.u));
    renderSet();
  }
  $("vTipo").innerHTML = `<option value="">Tutti i verbali</option>` + ORDINE.map(k => `<option value="${k}">${esc(TIPI[k].titolo)}</option>`).join("");
  $("vTipo").addEventListener("change", render); $("vSearch").addEventListener("input", render);
  $("vNew").querySelectorAll("button").forEach(b => b.onclick = () => openForm(b.dataset.t));

  /* miei dati e firma salvata */
  function renderSet() {
    for (const k of SET_KEYS) { const e = $("vs_" + k); if (e && document.activeElement !== e) e.value = settings[k] || ""; }
    $("vsFirma").innerHTML = myFirma ? `<img src="${myFirma}" alt="la mia firma" style="max-height:70px;max-width:100%;background:#fff;border:1px solid var(--border);border-radius:8px;padding:4px">` : `<div class="hint">Nessuna firma salvata.</div>`;
    $("vsSign").textContent = myFirma ? "Cambia la mia firma" : "✍ Fai la mia firma";
    $("vsDel").style.display = myFirma ? "" : "none";
  }
  SET_KEYS.forEach(k => $("vs_" + k).addEventListener("change", async () => { settings[k] = $("vs_" + k).value.trim(); await saveSettings(); }));
  $("vsSign").onclick = async () => {
    const r = await signPad("tecnico", "La mia firma (tecnico)");
    if (r) { myFirma = r.data; await setMeta("firmaTecnico", myFirma); renderSet(); }
  };
  $("vsDel").onclick = async () => { if (confirm("Togliere la firma salvata dal telefono?")) { myFirma = ""; await setMeta("firmaTecnico", ""); renderSet(); } };

  /* ---------- scelta del tipo (dalla scheda utente) ---------- */
  function chooseTipo(u) {
    const m = $("modal");
    m.innerHTML = `<div class="card"><h3>Nuovo verbale</h3><div class="meta">${u ? esc(u.nome) : ""}</div>${ORDINE.map(k => `<button class="btn navy" type="button" data-t="${k}">${esc(TIPI[k].titolo)}</button>`).join("")}<button class="btn" type="button" id="vcClose">Chiudi</button></div>`;
    m.classList.remove("hidden");
    m.querySelectorAll("[data-t]").forEach(b => b.onclick = () => openForm(b.dataset.t, null, u));
    $("vcClose").onclick = () => m.classList.add("hidden");
  }

  /* ---------- modulo ---------- */
  let S = null;
  const userKeyOf = t => ({ trasferimento: "assegnatario", presa_in_carico: "utente", assegnazione: "utente" })[t] || "";
  const luoghi = () => { const f = {}; [...edifici.map(e => e.nome), ...sediList].forEach(n => { n = (n || "").trim(); if (n) f[n.toLowerCase()] = f[n.toLowerCase()] || n; }); return Object.values(f).sort((a, b) => a.localeCompare(b, "it")).map(text => ({ text })); };
  const listItems = arr => () => arr.map(text => ({ text }));

  async function openForm(tipo, rec, prefillUser) {
    await loadSettings();
    H = storico((await getAll("verbali")).sort((a, b) => (b.modificato || "").localeCompare(a.modificato || "")).map(v => ({ tipo: v.tipo, dati: v.dati })));
    const d = rec ? JSON.parse(JSON.stringify(rec.dati)) : {};
    S = { uid: rec ? rec.uid : null, tipo, photos: rec ? (rec.foto || []).map(f => ({ ...f })) : [], apparati: (d.apparati || []).map(r => ({ ...r })), dotazione: (d.dotazione || []).map(r => ({ ...r })), editA: -1 };
    d.data = rec ? itToIso(d.data) : isoDay(new Date());
    d.data_sped = itToIso(d.data_sped);
    if (!rec) {
      if (tipo === "censimento") for (const k of SET_KEYS) if (settings[k]) d[k] = settings[k];
      if (tipo === "censimento") d.col_codice = "IMEI";
      if (prefillUser) {
        if (tipo === "trasferimento") { d.assegnatario = prefillUser.nome; d.contatti = prefillUser.telefono || ""; }
        if (tipo === "presa_in_carico") d.utente = prefillUser.matricola ? `${prefillUser.nome} (matricola ${prefillUser.matricola})` : prefillUser.nome;
        if (tipo === "assegnazione") { d.utente = prefillUser.nome; d.matricola = prefillUser.matricola || ""; }
      }
    }
    const val = k => esc(d[k] || "");
    const L = (k, label, o = {}) => `<label for="v_${k}">${label}</label><input id="v_${k}" ${o.attrs || ""} autocomplete="off" autocorrect="off" spellcheck="false" value="${val(k)}">${o.hint ? `<div class="hint">${o.hint}</div>` : ""}`;
    const T = (k, label, hint, rows) => `<label for="v_${k}">${label}</label><textarea id="v_${k}" style="min-height:${rows || 70}px">${esc(Array.isArray(d[k]) ? d[k].join("\n") : (d[k] || ""))}</textarea>${hint ? `<div class="hint">${hint}</div>` : ""}`;
    const motivi = opts => `<label>Motivazione</label>` + opts.map((o, i) => `<label class="chk"><input type="checkbox" data-m="${i}" ${(d.motivi || []).includes(o) ? "checked" : ""}> ${esc(o)}</label>`).join("") + L("altro", "Altro (motivo)", { hint: "Se lo scrivi, la casella «Altro» viene spuntata." });
    let body = "";
    if (tipo === "trasferimento") body = L("tipologia", "Tipologia *", { attrs: 'autocapitalize="sentences"' }) + `<div class="row"><div>${L("marca_modello", "Marca / Modello")}</div><div>${L("seriale", "Numero di serie", { attrs: 'autocapitalize="characters"' })}</div></div>` +
      L("provenienza", "Provenienza (struttura)") + L("destinazione", "Destinazione (struttura)") + L("assegnatario", "Assegnatario", { hint: "Scegli un utente in lista e compilo i suoi dati." }) + L("contatti", "Contatti", { attrs: 'inputmode="tel"' }) +
      motivi(MOTIVI.trasferimento) + T("note", "Note", "", 60) + `<label for="v_data_sped">Data spedizione</label><input id="v_data_sped" type="date" value="${esc(d.data_sped || "")}"><div class="hint">Se la lasci vuota uso la data del verbale.</div>`;
    else if (tipo === "presa_in_carico") body = L("tipologia", "Tipologia dispositivo *") + L("marca_modello", "Marca e modello") + L("seriale", "Numero seriale", { attrs: 'autocapitalize="characters"' }) +
      L("utente", "Utente assegnatario *", { hint: "Scegli un utente in lista e compilo i suoi dati." }) + `<div class="row"><div>${L("inc", "N° INC", { attrs: 'autocapitalize="characters"' })}</div><div>${L("ric", "N° RIC", { attrs: 'autocapitalize="characters"' })}</div></div>` +
      T("materiale", "Elenco materiale consegnato", "Una voce per riga (esempio: caricabatterie, cover).", 80) + motivi(MOTIVI.presa_in_carico);
    else if (tipo === "assegnazione") body = L("titolo", "Titolo", { hint: "Facoltativo: scrive «al dott.», «all'avv.», «alla dott.ssa»…" }) + L("utente", "Utente *") + L("matricola", "Matricola", { attrs: 'autocapitalize="characters"' }) +
      `<div class="card vbox"><h3>Dotazione assegnata *</h3><div id="v_dList"></div>
        <label for="v_dTipo">Oggetto</label><input id="v_dTipo" autocomplete="off" value="${esc(TIPI_DOT[0])}">
        <label for="v_dModello">Modello</label><input id="v_dModello" autocomplete="off" spellcheck="false">
        <label for="v_dCodice" id="v_dCodLab">S/N o IMEI</label><input id="v_dCodice" autocomplete="off" autocapitalize="characters" spellcheck="false">
        <div id="v_dSim" class="hidden"><div class="row"><div><label for="v_dIccid">ICCID</label><input id="v_dIccid" inputmode="numeric" autocomplete="off"></div><div><label for="v_dNumero">Numero</label><input id="v_dNumero" inputmode="tel" autocomplete="off"></div></div></div>
        <button class="btn navy" type="button" id="v_dAdd">＋ Aggiungi alla dotazione</button>
        <div class="hint">Se non scrivi il codice resta una riga ______ da compilare a mano nel Word.</div></div>` + T("accessori", "Accessori inclusi", "Una voce per riga (esempio: zaino porta notebook).", 70);
    else if (tipo === "censimento") body = L("oggetto", "Oggetto *", { hint: "Esempio: Dispositivi Microsoft Lumia 950" }) + L("descrizione", "Nel testo: «costituito da»", { hint: "Esempio: dispositivi smartphone Microsoft Lumia" }) +
      L("sede_citta", "Sede (città)") + L("indirizzo", "Indirizzo") + `<div class="row"><div>${L("cap", "CAP", { attrs: 'inputmode="numeric"' })}</div><div>${L("tecnico", "Tecnico *")}</div></div>` +
      L("collega", "Indicazioni ricevute da", { hint: "Facoltativo: se lo lasci vuoto la frase sulla dismissione non compare." }) + L("col_codice", "Titolo colonna codice") +
      `<div class="card vbox"><h3>Apparecchi censiti <span id="v_nApp" class="meta"></span></h3><div id="v_aList"></div>
        <div class="row"><div><label for="v_aMarca">Marca</label><input id="v_aMarca" autocomplete="off" spellcheck="false"></div><div><label for="v_aModello">Modello</label><input id="v_aModello" autocomplete="off" spellcheck="false"></div></div>
        <label for="v_aCodice" id="v_aCodLab">Codice</label><input id="v_aCodice" autocomplete="off" autocorrect="off" spellcheck="false" autocapitalize="characters" enterkeyhint="done">
        <button class="btn navy" type="button" id="v_aAdd">＋ Aggiungi apparecchio</button>
        <button class="btn" type="button" id="v_aPasteBtn">Incolla un elenco…</button>
        <div id="v_aPasteBox" class="hidden"><textarea id="v_aPaste" placeholder="Una riga per apparecchio. Da Excel: marca, modello, codice. Oppure solo i codici, uno per riga." style="min-height:110px"></textarea><button class="btn navy" type="button" id="v_aPasteOk">Aggiungi le righe</button></div>
        <div class="hint">Marca e modello restano scritti: per i prossimi basta il codice e Invio. Tocca una riga per correggerla.</div><div id="v_aWarn"></div></div>`;
    m_open(`<div class="card"><h3>${rec ? "Modifica verbale" : "Nuovo verbale"} · ${esc(TIPI[tipo].titolo)}</h3>
      <div class="hint">* = obbligatorio. Quello che lasci vuoto resta una riga da compilare a mano nel Word.</div>
      <label for="v_data">Data *</label><input id="v_data" type="date" value="${esc(d.data || "")}">
      ${body}
      <label>Foto ${TIPI[tipo].foto ? "(OBBLIGATORIA: almeno una)" : "(facoltativa)"}</label>
      <input id="v_foto" type="file" accept="image/*" multiple class="hidden">
      <button class="btn" type="button" id="v_btnFoto">📷 Scatta o scegli foto</button><div class="thumbs" id="v_thumbs"></div>
      <div class="hint">Finiscono in fondo al Word, nella pagina «Allegato fotografico».</div>
      <div class="hint"><b>Le firme si fanno dopo aver salvato</b>, nella pagina del verbale.</div>
      <button class="btn green" type="button" id="v_save">Salva verbale</button>
      <button class="btn" type="button" id="v_cancel">Annulla</button><div id="v_msg"></div></div>`);
    wireForm(d);
  }
  const m_open = html => { const m = $("modal"); m.innerHTML = html; m.classList.remove("hidden"); m.scrollTop = 0; };
  const closeModal = () => $("modal").classList.add("hidden");

  function wireForm(d) {
    const t = S.tipo, g = k => { const e = $("v_" + k); return e ? e.value : ""; };
    const ac = (k, fn) => { if ($("v_" + k)) autoComplete($("v_" + k), fn); };
    const setIfEmpty = (k, v) => { const e = $("v_" + k); if (e && v && !e.value.trim()) e.value = v; };
    const uk = userKeyOf(t);
    const lines = txt => txt.split(/\r?\n/).map(x => x.trim()).filter(Boolean);
    const luoghiItems = () => uniq([...histVals("provenienza"), ...histVals("destinazione"), ...luoghi().map(x => x.text)]).map(text => ({ text }));
    const collegaItems = () => { const base = userItems(), names = new Set(base.map(i => i.text.toLowerCase())); return [...base, ...histVals("collega").filter(x => !names.has(x.toLowerCase())).map(text => ({ text }))]; };
    const serialItems = () => {
      const grp = groupOf(g("tipologia")), seen = new Set(), items = [];
      const add = (ser, tip, mm, ut) => { ser = str(ser); const k = codeKey(ser); if (!ser || seen.has(k)) return; seen.add(k); items.push({ text: ser, hay: [ser, tip, mm, ut].join(" "), sub: [tip, mm, ut && `(${ut})`].filter(Boolean).join(" · "), grp: groupOf(tip) }); };
      H.dispositivi.forEach(x => add(x.seriale, x.tipologia, x.marca_modello, x.utente));
      users.forEach(u => { add(u.seriale_pc, "Notebook", "", u.nome); add(u.imei, "Smartphone", "", u.nome); });
      return grp ? [...items.filter(i => i.grp === grp), ...items.filter(i => i.grp !== grp)] : items;
    };
    ac("tipologia", histItems("tipologia", TIPOLOGIE)); ac("titolo", histItems("titolo", TITOLI_LISTA.filter(Boolean))); ac("sede_citta", histItems("sede_citta", sediList)); ac("col_codice", listItems(COLONNE));
    ac("marca_modello", histItems("marca_modello")); ac("seriale", serialItems); ac("contatti", histItems("contatti"));
    ac("provenienza", luoghiItems); ac("destinazione", luoghiItems);
    ac("oggetto", histItems("oggetto")); ac("descrizione", histItems("descrizione")); ac("indirizzo", histItems("indirizzo")); ac("cap", histItems("cap"));
    ac("tecnico", histItems("tecnico", [settings.tecnico])); ac("collega", collegaItems);
    if (uk) autoComplete($("v_" + uk), userItems, "Nessun utente in lista: scrivi il nome (oppure importa gli utenti dal PC in Scambio).");
    const curUser = () => {
      if (!uk) return null;
      const raw = g(uk).replace(/\s*\(matricola[^)]*\)\s*$/i, "");
      return raw.trim() ? users.find(u => u.key === keyOf(raw)) : null;
    };

    /* apparati già noti: tipologia, modello e seriale si compilano da soli */
    let devBox = null;
    if (uk && $("v_tipologia")) {
      devBox = document.createElement("div"); devBox.id = "v_devs";
      const wrap = $("v_" + uk).parentNode, hint = wrap.nextElementSibling;
      (hint && hint.classList.contains("hint") ? hint : wrap).insertAdjacentElement("afterend", devBox);
    }
    const applyDevice = (dev, override) => {
      for (const k of ["tipologia", "marca_modello", "seriale"]) { const e = $("v_" + k); if (e && (override || (dev[k] && !e.value.trim()))) e.value = dev[k] || ""; }
      if (uk && dev.utente && !$("v_" + uk).value.trim()) { $("v_" + uk).value = dev.utente; picked(); }
    };
    const drawDevs = devs => {
      if (!devBox) return;
      devBox.innerHTML = devs.length ? `<div class="hint">Apparati già noti di questa persona: tocca per compilare tipologia, modello e seriale.</div><div class="chips">${devs.map((x, i) => `<button type="button" data-i="${i}">${esc(etichetta({ ...x, utente: "" }))}</button>`).join("")}</div>` : "";
      devBox.querySelectorAll("button").forEach(b => b.onclick = () => applyDevice(devs[+b.dataset.i], true));
    };
    const userDevices = u => {
      const devs = dispositiviUtente(H, u.nome, u);
      drawDevs(devs);
      if (devs.length && !["tipologia", "marca_modello", "seriale"].some(k => g(k).trim())) applyDevice(devs[0]);
    };
    /* se cambi la tipologia (PC o cellulare) metto il seriale/IMEI giusto della persona */
    const deviceHint = () => {
      const u = curUser(); if (!u || !$("v_tipologia")) return;
      const grp = groupOf(g("tipologia")); if (!grp) return;
      const devs = dispositiviUtente(H, u.nome, u), cur = g("seriale").trim();
      const curDev = cur ? devs.find(x => codeKey(x.seriale) === codeKey(cur)) : null, wanted = devs.filter(x => groupOf(x.tipologia) === grp);
      if (cur && !(curDev && groupOf(curDev.tipologia) !== grp)) return;      // il seriale scritto va bene: non lo tocco
      if (curDev && g("marca_modello").trim() === curDev.marca_modello) $("v_marca_modello").value = "";
      $("v_seriale").value = wanted.length ? wanted[0].seriale : "";
      if (wanted.length && !g("marca_modello").trim()) $("v_marca_modello").value = wanted[0].marca_modello;
    };
    /* assegnazione: per un notebook o uno smartphone propongo il seriale/IMEI della persona */
    let dotAuto = null;
    const dotHint = () => {
      const u = t === "assegnazione" ? curUser() : null; if (!u || !$("v_dTipo")) return;
      const cur = $("v_dCodice").value.trim(); if (cur && cur !== dotAuto) return;      // l'hai scritto tu: non lo tocco
      const tip = tipDaDot($("v_dTipo").value), val = tip === "Notebook" ? str(u.seriale_pc) : tip === "Smartphone" ? str(u.imei) : "";
      dotAuto = val; $("v_dCodice").value = val;
    };
    function picked() {
      const u = curUser(); if (!u) return;
      const inp = $("v_" + uk);
      if (t === "trasferimento" || t === "assegnazione") { if (inp.value !== u.nome) inp.value = u.nome; }
      if (t === "trasferimento") setIfEmpty("contatti", u.telefono);
      else if (t === "presa_in_carico") { const w = u.matricola ? `${u.nome} (matricola ${u.matricola})` : u.nome; if (inp.value !== w) inp.value = w; }
      else if (t === "assegnazione") { setIfEmpty("matricola", u.matricola); setIfEmpty("titolo", H.titoli[u.nome.toLowerCase()] || ""); dotHint(); }
      if (t === "trasferimento" || t === "presa_in_carico") userDevices(u);
      deviceHint();
    }
    if (uk) $("v_" + uk).addEventListener("change", () => { if (!curUser()) drawDevs([]); picked(); });
    if ($("v_tipologia")) { $("v_tipologia").addEventListener("input", deviceHint); $("v_tipologia").addEventListener("change", deviceHint); }
    if ($("v_seriale")) $("v_seriale").addEventListener("change", () => { const dev = trovaDispositivo(H, g("seriale"), users); if (dev) applyDevice(dev); });

    /* verbale di censimento: oggetto → descrizione, città → indirizzo e CAP */
    const sedeAuto = {};      // indirizzo e CAP messi da me (non scritti da te): se cambi città li sostituisco
    if (t === "censimento" && !S.uid) for (const k of ["indirizzo", "cap"]) if (str(d[k])) sedeAuto[k] = str(d[k]);
    if (t === "censimento") {
      $("v_oggetto").addEventListener("change", () => { const info = H.oggetti[g("oggetto").trim().toLowerCase()]; if (info) setIfEmpty("descrizione", info.descrizione); });
      $("v_sede_citta").addEventListener("change", () => {
        const citta = g("sede_citta").trim().toLowerCase();
        let info = H.sedi[citta];
        if (!info && citta && citta === str(settings.sede_citta).toLowerCase()) info = { indirizzo: str(settings.indirizzo), cap: str(settings.cap) };
        for (const k of ["indirizzo", "cap"]) {
          const e = $("v_" + k);
          if (e.value.trim() && e.value.trim() === sedeAuto[k]) { e.value = ""; delete sedeAuto[k]; }      // era quello della città di prima
          const nv = info ? str(info[k]) : "";
          if (nv && !e.value.trim()) { e.value = nv; sedeAuto[k] = nv; }
        }
      });
    }

    /* materiale e accessori: voci dei verbali precedenti da aggiungere con un tocco */
    const suggest = k => {
      const ta = $("v_" + k), all = histVals(k); if (!ta || !all.length) return;
      const box = document.createElement("div"), chips = document.createElement("div"), more = document.createElement("input");
      box.id = "v_" + k + "_sug";
      const hint = ta.nextElementSibling; (hint && hint.classList.contains("hint") ? hint : ta).insertAdjacentElement("afterend", box);
      more.placeholder = "Altre voci: tocca per scegliere"; more.autocomplete = "off"; more.setAttribute("autocapitalize", "sentences");
      box.append(chips, more);
      let left = [];
      autoComplete(more, () => left.map(text => ({ text })));
      const have = () => new Set(lines(ta.value).map(x => x.toLowerCase()));
      const add = x => { if (!have().has(x.toLowerCase())) ta.value = (ta.value.trim() ? ta.value.replace(/\s+$/, "") + "\n" : "") + x; draw(); };
      const draw = () => {
        const h = have(); left = all.filter(x => !h.has(x.toLowerCase()));
        chips.innerHTML = left.length ? `<div class="hint">Dai verbali precedenti: tocca per aggiungere.</div><div class="chips">${left.slice(0, 8).map((x, i) => `<button type="button" data-i="${i}">＋ ${esc(x)}</button>`).join("")}</div>` : "";
        chips.querySelectorAll("button").forEach(b => b.onclick = () => add(left[+b.dataset.i]));
        more.parentNode.classList.toggle("hidden", left.length <= 8);
      };
      more.addEventListener("change", () => { const x = more.value.trim(); more.value = ""; if (x) add(x); });
      ta.addEventListener("input", draw); draw();
    };
    suggest("materiale"); suggest("accessori");

    /* foto */
    const thumbs = () => {
      $("v_thumbs").innerHTML = S.photos.map((p, i) => `<div><img src="${p.data}" alt=""><button class="x" data-i="${i}" type="button">×</button></div>`).join("");
      $("v_thumbs").querySelectorAll(".x").forEach(b => b.onclick = () => { S.photos.splice(+b.dataset.i, 1); thumbs(); });
    };
    $("v_btnFoto").onclick = () => $("v_foto").click();
    $("v_foto").onchange = async e => {
      for (const f of e.target.files) S.photos.push({ nome: `verbale_${Date.now()}_${S.photos.length + 1}.jpg`, data: await shrink(f, 1600) });
      e.target.value = ""; thumbs();
    };
    thumbs();

    /* censimento: elenco apparecchi */
    if (t === "censimento") {
      const colonna = () => g("col_codice") || "IMEI";
      const draw = () => {
        $("v_nApp").textContent = `(totale: ${S.apparati.length})`;
        $("v_aList").innerHTML = S.apparati.map((r, i) => `<div class="vr" data-i="${i}"><div><b>${i + 1}.</b> ${esc([r.marca, r.modello].filter(Boolean).join(" ") || "—")}<br><small>${esc(r.codice || "(manca il codice)")}</small></div><button type="button" class="x" data-del="${i}">×</button></div>`).join("");
        $("v_aList").querySelectorAll(".vr").forEach(row => row.onclick = e => {
          const i = +row.dataset.i;
          if (e.target.dataset.del !== undefined) { S.apparati.splice(i, 1); return draw(); }
          const r = S.apparati.splice(i, 1)[0]; marcaAuto = ""; $("v_aMarca").value = r.marca; $("v_aModello").value = r.modello; $("v_aCodice").value = r.codice; draw(); $("v_aCodice").focus();
        });
        const av = controllaApparati(S.apparati, colonna());
        $("v_aWarn").innerHTML = av.length ? `<div class="msg err">${esc(av.slice(0, 4).join("\n") + (av.length > 4 ? `\n…e altri ${av.length - 4} da controllare` : ""))}</div>` : "";
        $("v_aCodLab").textContent = colonna();
        $("v_aCodice").setAttribute("inputmode", colonna().toUpperCase() === "IMEI" ? "numeric" : "text");
      };
      let marcaAuto = "";      // marca rimasta scritta dal giro prima (non digitata da te)
      autoComplete($("v_aMarca"), histItems("ap_marca")); autoComplete($("v_aModello"), histItems("ap_modello"));
      $("v_aModello").addEventListener("change", () => {
        const marca = H.modelli[$("v_aModello").value.trim().toLowerCase()], cur = $("v_aMarca").value.trim();
        if (marca && (!cur || cur === marcaAuto)) { $("v_aMarca").value = marca; marcaAuto = marca; }
      });
      const add = () => {
        const marca = $("v_aMarca").value.trim(), modello = $("v_aModello").value.trim(), codice = $("v_aCodice").value.trim();
        if (!(marca || modello || codice)) return;
        S.apparati.push({ marca, modello, codice }); marcaAuto = marca; $("v_aCodice").value = ""; draw(); $("v_aCodice").focus();
      };
      S.flush = () => { if ($("v_aCodice").value.trim()) add(); };
      $("v_aAdd").onclick = add;
      $("v_aCodice").addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); add(); } });
      $("v_col_codice").addEventListener("change", draw); $("v_col_codice").addEventListener("input", draw);
      $("v_aPasteBtn").onclick = () => { $("v_aPasteBox").classList.toggle("hidden"); $("v_aPaste").focus(); };
      $("v_aPasteOk").onclick = () => { S.apparati.push(...parseApparati($("v_aPaste").value, $("v_aMarca").value.trim(), $("v_aModello").value.trim())); $("v_aPaste").value = ""; $("v_aPasteBox").classList.add("hidden"); draw(); };
      if (S.apparati.length) { const l = S.apparati[S.apparati.length - 1]; $("v_aMarca").value = l.marca; $("v_aModello").value = l.modello; marcaAuto = l.marca; }
      draw();
    }
    /* assegnazione: dotazione */
    if (t === "assegnazione") {
      autoComplete($("v_dTipo"), listItems(TIPI_DOT)); autoComplete($("v_dModello"), histItems("dot_modello"));
      const isSim = () => $("v_dTipo").value.toLowerCase().includes("sim");
      const draw = () => {
        $("v_dList").innerHTML = S.dotazione.map((r, i) => `<div class="vr" data-i="${i}"><div><b>${esc([r.tipo, r.modello].filter(Boolean).join(" · "))}</b><br><small>${esc(righeDotazione(r).join(" · "))}</small></div><button type="button" class="x" data-del="${i}">×</button></div>`).join("") || `<div class="hint">Nessun oggetto ancora.</div>`;
        $("v_dList").querySelectorAll(".vr").forEach(row => row.onclick = e => {
          const i = +row.dataset.i;
          if (e.target.dataset.del !== undefined) { S.dotazione.splice(i, 1); return draw(); }
          const r = S.dotazione.splice(i, 1)[0]; $("v_dTipo").value = r.tipo; $("v_dModello").value = r.modello; $("v_dCodice").value = r.codice; $("v_dIccid").value = r.iccid || ""; $("v_dNumero").value = r.numero || ""; sim(); draw();
        });
      };
      const sim = () => { $("v_dSim").classList.toggle("hidden", !isSim()); $("v_dCodLab").textContent = isSim() ? "Codice (facoltativo per la SIM)" : /smartphone|cellulare|telefono/i.test($("v_dTipo").value) ? "IMEI" : "S/N (numero di serie)"; };
      const add = () => {
        const r = { tipo: $("v_dTipo").value.trim(), modello: $("v_dModello").value.trim(), codice: $("v_dCodice").value.trim(), iccid: $("v_dIccid").value.trim(), numero: $("v_dNumero").value.trim() };
        if (!r.tipo) return;
        if (!isSim()) { r.iccid = ""; r.numero = ""; }
        S.dotazione.push(r); ["v_dModello", "v_dCodice", "v_dIccid", "v_dNumero"].forEach(i => $(i).value = ""); draw();
      };
      /* il seriale proposto da me (dalla scheda utente) non è un oggetto da aggiungere in silenzio */
      S.flush = () => { if (["v_dModello", "v_dCodice", "v_dIccid", "v_dNumero"].some(i => $(i).value.trim() && !(i === "v_dCodice" && $(i).value.trim() === dotAuto))) add(); };
      $("v_dAdd").onclick = add;
      $("v_dTipo").addEventListener("input", () => { sim(); dotHint(); }); $("v_dTipo").addEventListener("change", () => { sim(); dotHint(); });
      sim(); draw();
    }
    $("v_cancel").onclick = () => { const uid = S.uid; closeModal(); if (uid) showVer(uid); };
    $("v_save").onclick = save;
    if (!S.uid && uk && curUser()) picked();      // «Nuovo verbale» dalla scheda di un utente: parto già compilato
  }

  function collect() {
    const t = S.tipo, g = k => { const e = $("v_" + k); return e ? e.value.trim() : ""; };
    const lines = k => g(k) ? $("v_" + k).value.split(/\r?\n/).map(x => x.trim()).filter(Boolean) : [];
    const motivi = () => [...$("modal").querySelectorAll("input[data-m]")].filter(c => c.checked).map(c => MOTIVI[t][+c.dataset.m]);
    const d = { data: isoToIt(g("data")) };
    if (t === "trasferimento") Object.assign(d, { tipologia: g("tipologia"), marca_modello: g("marca_modello"), seriale: g("seriale"), provenienza: g("provenienza"), destinazione: g("destinazione"), assegnatario: g("assegnatario"), contatti: g("contatti"), altro: g("altro"), note: g("note"), data_sped: isoToIt(g("data_sped")), motivi: motivi() });
    if (t === "presa_in_carico") Object.assign(d, { tipologia: g("tipologia"), marca_modello: g("marca_modello"), seriale: g("seriale"), utente: g("utente"), inc: g("inc"), ric: g("ric"), altro: g("altro"), materiale: lines("materiale"), motivi: motivi() });
    if (t === "assegnazione") Object.assign(d, { titolo: g("titolo"), utente: g("utente"), matricola: g("matricola"), dotazione: S.dotazione.map(r => ({ ...r })), accessori: lines("accessori") });
    if (t === "censimento") Object.assign(d, { oggetto: g("oggetto"), descrizione: g("descrizione"), sede_citta: g("sede_citta"), indirizzo: g("indirizzo"), cap: g("cap"), tecnico: g("tecnico"), collega: g("collega"), col_codice: g("col_codice") || "IMEI", apparati: S.apparati.map(r => ({ ...r })) });
    return d;
  }
  function validate(d) {
    if (!$("v_data").value) return "Scrivi la data del verbale.";
    const t = S.tipo;
    if (t === "trasferimento" && !d.tipologia) return "Scrivi la tipologia dell'apparato (esempio: Notebook).";
    if (t === "presa_in_carico") { if (!d.tipologia) return "Scrivi la tipologia del dispositivo (esempio: Notebook)."; if (!d.utente) return "Scrivi l'utente assegnatario."; }
    if (t === "assegnazione") { if (!d.utente) return "Scrivi il nome dell'utente a cui assegni il materiale."; if (!d.dotazione.length) return "Aggiungi almeno un oggetto alla dotazione assegnata."; }
    if (t === "censimento") {
      if (!d.oggetto) return "Scrivi l'oggetto del verbale (esempio: Dispositivi Microsoft Lumia 950).";
      if (!d.tecnico) return "Scrivi il nome del tecnico che redige il verbale.";
      if (!d.apparati.length) return "Aggiungi almeno un apparecchio all'elenco.";
    }
    if (TIPI[t].foto && !S.photos.length) return `Per il ${TIPI[t].titolo.toLowerCase()} la foto è obbligatoria: allega almeno una foto.`;
    return "";
  }
  async function save() {
    if (S.flush) S.flush();
    const d = collect(), err = validate(d);
    if (err) { show($("v_msg"), err, true); $("v_msg").scrollIntoView({ block: "center" }); return; }
    if (S.tipo === "censimento") {
      const av = controllaApparati(d.apparati, d.col_codice);
      if (av.length && !confirm("Ci sono punti da controllare:\n\n" + av.slice(0, 8).join("\n") + (av.length > 8 ? "\n…" : "") + "\n\nSalvo comunque il verbale?")) return;
    }
    const old = S.uid ? await getKey("verbali", S.uid) : null, now = nowIso();
    let firme = old ? { ...(old.firme || {}) } : {};
    if (old && Object.keys(firme).length && (!sameDati(old.dati, d) || (old.foto || []).length !== S.photos.length) &&
        confirm("Questo verbale è già firmato e hai cambiato dei dati.\n\nOK = tolgo le firme (da rifare sul testo nuovo)\nAnnulla = tengo le firme")) firme = {};
    const rec = { uid: S.uid || uuid(), tipo: S.tipo, dati: d, foto: S.photos, firme, creato: old ? old.creato : now, modificato: now };
    await put("verbali", rec);
    if (S.tipo === "censimento") { SET_KEYS.forEach(k => { if (d[k]) settings[k] = d[k]; }); await saveSettings(); }
    closeModal(); await render(); showVer(rec.uid, true);
  }

  /* ---------- pagina del verbale (anteprima + firme) ---------- */
  async function showVer(uid, justSaved) {
    const v = await getKey("verbali", uid); if (!v) return;
    await loadSettings();
    const m = $("modal"), top = m.scrollTop, last = await getMeta("lastSent");
    m.innerHTML = `<div class="card"><h3>${esc(TIPI[v.tipo].titolo)}</h3>
      <div class="meta">${isSigned(v) ? "✍ Firmato" : "Da firmare"} · ${isNewV(v, last) ? "da inviare al PC" : "già nel PC"}</div>
      ${justSaved ? `<div class="msg">Verbale salvato. Fai firmare qui sotto.</div>` : ""}
      <div class="dc">${docHtml(v)}</div>
      ${(v.foto || []).length ? `<label>Foto (${v.foto.length})</label><div class="thumbs">${v.foto.map(f => `<div><img src="${f.data}" alt=""></div>`).join("")}</div>` : ""}
      <button class="btn" type="button" id="vdPdf">📄 Crea PDF (copia per il telefono)</button>
      <button class="btn navy" type="button" id="vdEdit">Modifica i dati</button>
      <button class="btn red" type="button" id="vdDel">Elimina dal telefono</button>
      <button class="btn" type="button" id="vdClose">Chiudi</button><div id="vdMsg"></div>
      <div class="hint">Il Word con le firme dentro si crea sul PC: manda il verbale da Scambio › Prepara file per il PC. L'eliminazione vale solo qui.</div></div>`;
    m.classList.remove("hidden"); m.scrollTop = justSaved ? 0 : top;
    $("vdClose").onclick = () => { closeModal(); render(); };
    $("vdEdit").onclick = () => openForm(v.tipo, v);
    $("vdDel").onclick = async () => { if (confirm("Eliminare questo verbale dal telefono?\nSe è già stato importato nel PC, nel PC resta.")) { await del("verbali", uid); closeModal(); render(); } };
    $("vdPdf").onclick = async () => {
      show($("vdMsg"), "Creo il PDF…");
      try {
        const bytes = await pdfOf(v), file = new File([bytes], fileStem(v) + ".pdf", { type: "application/pdf" });
        if (await deliver(file, $("vdMsg"))) show($("vdMsg"), "PDF creato: " + file.name);
      } catch (err) { show($("vdMsg"), "Non sono riuscito a creare il PDF: " + err.message, true); }
    };
    const store = async fn => { const cur = await getKey("verbali", uid); fn(cur); cur.modificato = nowIso(); delete cur.daPC; await put("verbali", cur); await render(); showVer(uid); };
    const slotLabel = k => slotsOf(v.tipo).find(s => s[0] === k)[1];
    m.querySelectorAll("[data-sign]").forEach(b => b.onclick = async () => {
      const k = b.dataset.sign, r = await signPad(k, slotLabel(k));
      if (!r) return;
      if (r.keep) { myFirma = r.data; await setMeta("firmaTecnico", myFirma); }
      store(cur => { cur.firme = { ...(cur.firme || {}), [k]: r.data }; });
    });
    m.querySelectorAll("[data-mine]").forEach(b => b.onclick = () => store(cur => { cur.firme = { ...(cur.firme || {}), tecnico: myFirma }; }));
    m.querySelectorAll("[data-unsign]").forEach(b => b.onclick = () => store(cur => { const f = { ...(cur.firme || {}) }; delete f[b.dataset.unsign]; cur.firme = f; }));
  }

  /* ---------- scambio con il PC ---------- */
  async function exportItems(mode, last) {
    const all = await getAll("verbali");
    return (mode === "all" ? all : all.filter(v => isNewV(v, last))).map(v => ({ uid: v.uid, tipo: v.tipo, dati: v.dati, foto: (v.foto || []).map(f => ({ nome: f.nome, data: f.data })), firme: v.firme || {}, creato: v.creato, modificato: v.modificato }));
  }
  async function countNew(last) { return (await getAll("verbali")).filter(v => isNewV(v, last)).length; }
  async function merge(data, fromPc) {
    let n = 0, a = 0, p = 0, s = 0;
    for (const it of data.verbali || []) {
      if (!it || !it.uid || !TIPI[it.tipo] || !it.dati || typeof it.dati !== "object") { s++; continue; }
      const uid = String(it.uid), cur = await getKey("verbali", uid), mod = str(it.modificato);
      if (cur && !(mod > (cur.modificato || ""))) { p++; continue; }
      const foto = (it.foto || []).filter(f => f && typeof f.data === "string" && f.data.startsWith("data:")).map(f => ({ nome: String(f.nome || "foto.jpg"), data: f.data }));
      const firme = {};
      for (const k of ["tecnico", "utente"]) { const x = it.firme && it.firme[k]; if (typeof x === "string" && x.startsWith("data:image/")) firme[k] = x; }
      const rec = { uid, tipo: it.tipo, dati: cleanDati(it.dati), foto: foto.length ? foto : (cur ? cur.foto || [] : []), firme: { ...(cur ? cur.firme || {} : {}), ...firme }, creato: it.creato || mod || nowIso(), modificato: mod || nowIso() };
      if (fromPc) rec.daPC = rec.modificato;
      await put("verbali", rec); cur ? a++ : n++;
    }
    const imp = data.impostazioni_verbali;
    if (imp && typeof imp === "object") {
      await loadSettings(); let ch = false;
      for (const k of SET_KEYS) if (!settings[k] && str(imp[k])) { settings[k] = str(imp[k]); ch = true; }
      if (ch) await saveSettings();
    }
    return { n, a, p, s };
  }

  window.VER = { render, chooseTipo, exportItems, countNew, merge, tipi: TIPI, cleanDati, imeiProblem, controllaApparati, parseApparati, docOf };
})();
