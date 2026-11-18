'use strict';
/* Audit kupaca – aplikacija za komercijalistu na terenu.
   Podaci: SharePoint lista „Lične karte kupaca” preko Microsoft Graph-a.
   Rad bez mreže: kupci se čuvaju u telefonu, izmene i fotografije čekaju u redu dok se veza ne vrati. */

const C = window.APP_CONFIG || {};
const DEMO = !C.clientId || /[?&]demo/.test(location.search);
const SCOPES = ['Sites.ReadWrite.All', 'User.Read'];

/* ---------- šifarnici (moraju da se poklapaju sa izborima u listi) ---------- */
const STATUS = ['aktivan', 'neaktivan', 'potencijalni', 'zatvoren'];
const SEG = ['ZDRAVA HRANA KOMPLET - BIOLINE', 'ZDRAVA HRANA NAMENSKA', 'RINFUZNA ROBA / PIJAČNA PRODAVNICA', 'PILJARA', 'APOTEKA / BILJNA APOTEKA', 'MEŠOVITA ROBA / TT - MALA PRODAVNICA', 'MEŠOVITA ROBA / TT - VELIKA PRODAVNICA', 'CANDY SHOP', 'KAFANA / RESTORAN / KAFIĆ SA HRANOM / PICERIJA / BRZA HRANA', 'KETERING / ISHRANA RADNIKA / PROSLAVE', 'POSLASTIČARNICA', 'PEKARA', 'SMOOTHIE SHOP / KAFETERIJA / ČAJDŽINICA / KAFIĆ BEZ HRANE', 'PROIZVODNJA KOLAČA', 'PROIZVODNJA HLEBA I PECIVA', 'PROIZVODNJA KONDITORSKIH PROIZVODA', 'PRERADA MESA', 'PRERADA VOĆA I POVRĆA', 'PAKERAJI', 'OSTALI NESVRSTANI'];
const TER = ['1.01 KRUŠEVAC', '1.02 TRSTENIK', '1.03 ALEKSANDROVAC', '2.01 NEGOTIN', '2.02 ZAJEČAR', '2.03 BOR', '3.01 PROKUPLJE', '3.02 VRANJE', '3.03 LESKOVAC', '4.01 NIŠ', '4.011 NIŠ CENTAR', '4.012 NIŠ SEVER', '4.013 NIŠ JUG', '4.02 ALEKSINAC', '4.03 PIROT', '5.01 PARAĆIN', '5.02 SVILAJNAC', '5.03 JAGODINA', '5.04 KRAGUJEVAC', '5.05 NOVI SAD', '5.06 BEOGRAD', '6.01 KRALJEVO', '6.02 NOVI PAZAR', '6.03 UŽICE', '9 IZVOZ'];
const NRAZ = ['VOZI SE PO RASPOREDU', 'VOZI SE PO POTREBI', 'DOLAZI PO ROBU', 'ŠALJE SE KURIRSKOM SLUŽBOM O NAŠEM TROŠKU', 'ŠALJE SE KURIRSKOM SLUŽBOM O NJIHOVOM TROŠKU'];
const NPOZ = ['KUPAC SE ZOVE PREMA RASPOREDU', 'KUPAC ZOVE SAM PO POTREBI', 'KUPAC NARUČUJE PREKO SAJTA', 'KUPAC DOLAZI SAM'];
const PLAC = ['odmah po prijemu robe', 'na valutu'];
const OBIL = ['801 OBILAZAK NEDELJNO', '802 OBILAZAK DVONEDELJNO', '803 OBILAZAK MESEČNO', '804 OBILAZAK KVARTALNO', '805 OBILAZAK POLUGODIŠNJE', '806 OBILAZAK GODIŠNJE', '809 NE OBILAZI SE'];
const ISH = ['obrađen', 'vratiti se', 'zatvoren / ne postoji', 'odbio saradnju'];
const VEL = ['mali', 'srednji', 'veliki'];
const POT = ['nizak', 'srednji', 'visok'];
const ASORT = ['zdrava hrana', 'čajevi RF', 'Bioline', 'bombone - slatkiši', 'kuhinja - torte', 'grickalice', 'jezgra', 'suvo voće'];
const DANI = ['PONEDELJAK', 'UTORAK', 'SREDA', 'ČETVRTAK', 'PETAK', 'SUBOTA'];
const DANI_K = ['Pon', 'Uto', 'Sre', 'Čet', 'Pet', 'Sub'];

/* ---------- koraci audita ---------- */
const STEPS = [
  { k: 'status', t: 'Status kupca', type: 'one', f: 'StatusNovi', o: STATUS, bs: f => f.field_11 },
  { k: 'seg', t: 'Segment', type: 'one', f: 'AuditSegment', o: SEG, bs: f => f.field_14, find: true },
  { k: 'ter', t: 'Teren', type: 'one', f: 'AuditTeren', o: TER, bs: f => [f.field_8, f.field_9].filter(Boolean).join(' · '), find: true },
  { k: 'nraz', t: 'Način razvoza', type: 'one', f: 'AuditNacinRazvoza', o: NRAZ },
  { k: 'raz', t: 'Dani razvoza', type: 'days', f: 'AuditRazvoz', pre: 'V ', six: true, bs: f => f.field_15 },
  { k: 'npoz', t: 'Način pozivanja', type: 'one', f: 'AuditNacinPozivanja', o: NPOZ },
  { k: 'poz', t: 'Dani poziva za trebovanje', type: 'days', f: 'AuditPozivanje', pre: '', six: false, bs: f => f.field_16 },
  { k: 'kon', t: 'Kontakti', type: 'contacts', bs: f => [f.field_7, f.field_5].filter(Boolean).join(' · ') },
  { k: 'rv', t: 'Radno vreme', type: 'text', f: 'RadnoVreme', ph: 'npr. 8–20, sub 8–14, ned ne radi', quick: ['7–15', '8–16', '8–20', '7–21', 'sub 8–14', 'ned ne radi'] },
  { k: 'plac', t: 'Način plaćanja', type: 'one', f: 'AuditNacinPlacanja', o: PLAC, hint: 'Pitaj kupca: plaća odmah po prijemu robe ili na valutu.' },
  { k: 'vp', t: 'Veličina i potencijal', type: 'pair' },
  { k: 'as', t: 'Naš asortiman', type: 'asort', hint: 'Za svaku grupu: da li je drži ili ne.' },
  { k: 'nab', t: 'Od koga još nabavlja', type: 'area', f: 'NabavljaOd', ph: 'Dobavljači, šta kupuje kod njih, po kojim uslovima…' },
  { k: 'obil', t: 'Plan obilaska', type: 'one', f: 'AuditPlanObilaska', o: OBIL, extra: 'monitoring' },
  { k: 'gps', t: 'Lokacija objekta', type: 'gps' },
  { k: 'foto', t: 'Fotografija objekta', type: 'photo' },
  { k: 'kraj', t: 'Beleška i ishod', type: 'end' }
];

/* ---------- stanje ---------- */
const S = {
  items: [], site: null, list: null, account: null, pca: null,
  view: 'tereni', teren: null, filter: 'open', sort: 'promet', q: '',
  cur: null, step: 0, d: null, photos: [], online: navigator.onLine, pending: 0, syncing: false
};
const $app = document.getElementById('app');

/* ---------- pomoćne ---------- */
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = v => Number(v) || 0;
const money = n => { n = num(n); return n >= 1e6 ? (n / 1e6).toFixed(1).replace('.', ',') + ' mil' : n >= 1000 ? Math.round(n / 1000) + ' hilj.' : (n ? Math.round(n) + ' din' : '—'); };
const cv = v => (v && typeof v === 'object' && 'Value' in v) ? v.Value : v; // vrednost izbora
const arr = v => Array.isArray(v) ? v : (v ? [v] : []);
const split = s => String(s || '').split(',').map(x => x.trim()).filter(Boolean);
const today = () => new Date().toISOString().slice(0, 10);
const daysAgo = d => { if (!d) return null; const t = Date.parse(d); return isNaN(t) ? null : Math.floor((Date.now() - t) / 864e5); };
const fmtDate = d => { if (!d) return '—'; const x = new Date(d); return isNaN(x) ? d : x.toLocaleDateString('sr-RS'); };
const terenOf = it => cv(it.f.AuditTeren) || 'Bez terena';
const phoneOf = it => (split(it.f.field_5)[0] || '').replace(/[^\d+]/g, '');
function toast(msg) { const t = document.getElementById('toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast.h); toast.h = setTimeout(() => t.hidden = true, 2600); }

/* ---------- lokalno skladište ---------- */
const LS_ITEMS = 'ak_items_v1';
function saveItems() { try { localStorage.setItem(LS_ITEMS, JSON.stringify({ at: Date.now(), items: S.items })); } catch (e) { console.warn(e); } }
function loadCached() { try { const x = JSON.parse(localStorage.getItem(LS_ITEMS) || 'null'); return x ? x.items : null; } catch (e) { return null; } }

// Red čekanja u IndexedDB; ako telefon ne dozvoli IndexedDB, red radi u memoriji (dok je aplikacija otvorena).
const DB = {
  db: null, mem: null, seq: 1,
  open() {
    if (this.db) return Promise.resolve(this.db);
    if (this.mem) return Promise.resolve(null);
    return new Promise(res => {
      const done = db => { clearTimeout(tm); if (db) this.db = db; else if (!this.mem) this.mem = []; res(db || null); };
      const tm = setTimeout(() => done(null), 2500);
      try { const r = indexedDB.open('audit-kupaca', 1); r.onupgradeneeded = () => r.result.createObjectStore('q', { keyPath: 'id', autoIncrement: true }); r.onsuccess = () => done(r.result); r.onerror = () => done(null); }
      catch (e) { done(null); }
    });
  },
  async tx(mode, fn) { const db = await this.open(); if (!db) return null; return new Promise((res, rej) => { const t = db.transaction('q', mode); const st = t.objectStore('q'); const out = fn(st); t.oncomplete = () => res(out && out.result !== undefined ? out.result : out); t.onerror = () => rej(t.error); }); },
  async add(op) { await this.open(); if (this.mem) { op.id = this.seq++; this.mem.push(op); return op.id; } return this.tx('readwrite', st => st.add(op)); },
  async all() { await this.open(); return this.mem ? this.mem.slice() : this.tx('readonly', st => st.getAll()); },
  async del(id) { await this.open(); if (this.mem) { this.mem = this.mem.filter(o => o.id !== id); return; } return this.tx('readwrite', st => st.delete(id)); },
  async put(op) { await this.open(); if (this.mem) { const i = this.mem.findIndex(o => o.id === op.id); if (i >= 0) this.mem[i] = op; return; } return this.tx('readwrite', st => st.put(op)); }
};

/* ---------- prijava i Graph ---------- */
async function initAuth() {
  S.pca = new msal.PublicClientApplication({
    auth: { clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + C.tenantId, redirectUri: location.origin + location.pathname.replace(/\/$/, '/index.html') },
    cache: { cacheLocation: 'localStorage' }
  });
  await S.pca.initialize();
  const r = await S.pca.handleRedirectPromise();
  S.account = (r && r.account) || S.pca.getAllAccounts()[0] || null;
}
async function token() {
  try { return (await S.pca.acquireTokenSilent({ scopes: SCOPES, account: S.account })).accessToken; }
  catch (e) { await S.pca.acquireTokenRedirect({ scopes: SCOPES, account: S.account }); throw e; }
}
async function g(path, opt = {}) {
  const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
  const t = await token();
  const r = await fetch(url, { ...opt, headers: { Authorization: 'Bearer ' + t, ...(opt.body && !(opt.body instanceof Blob) ? { 'Content-Type': 'application/json' } : {}), ...(opt.headers || {}) } });
  if (!r.ok) throw new Error(r.status + ' ' + (await r.text()).slice(0, 300));
  return r.status === 204 ? null : r.json();
}
async function resolveList() {
  if (S.site && S.list) return;
  const site = await g('/sites/' + C.spHost + ':' + C.sitePath + '?$select=id');
  S.site = site.id;
  const lists = await g('/sites/' + S.site + '/lists?$select=id,displayName&$top=200');
  const l = lists.value.find(x => x.displayName === C.listName);
  if (!l) throw new Error('Nema liste ' + C.listName);
  S.list = l.id;
}
async function fetchAll() {
  await resolveList();
  let url = '/sites/' + S.site + '/lists/' + S.list + '/items?$expand=fields&$top=999';
  const out = [];
  while (url) { const j = await g(url); j.value.forEach(x => out.push({ id: x.id, f: x.fields })); url = j['@odata.nextLink']; }
  return out;
}

/* ---------- red izmena (rad bez mreže) ---------- */
async function queue(op) { await DB.add(op); await countPending(); flush(); }
async function countPending() { S.pending = (await DB.all()).length; paintSync(); }
async function flush() {
  if (DEMO || S.syncing || !navigator.onLine) return;
  S.syncing = true; paintSync();
  try {
    await resolveList();
    const ops = await DB.all();
    for (const op of ops) {
      if (op.type === 'patch') {
        if (String(op.itemId).startsWith('tmp')) continue;
        await g('/sites/' + S.site + '/lists/' + S.list + '/items/' + op.itemId + '/fields', { method: 'PATCH', body: JSON.stringify(op.fields) });
      } else if (op.type === 'create') {
        const r = await g('/sites/' + S.site + '/lists/' + S.list + '/items', { method: 'POST', body: JSON.stringify({ fields: op.fields }) });
        const it = S.items.find(x => x.id === op.tmp); if (it) it.id = r.id;
        for (const o of await DB.all()) if (o.itemId === op.tmp) { o.itemId = r.id; await DB.put(o); }
        if (S.cur && S.cur.id === op.tmp) S.cur.id = r.id;
        saveItems();
      } else if (op.type === 'photo') {
        const path = encodeURIComponent(C.photoFolder) + '/' + encodeURIComponent(op.name);
        await g('/sites/' + S.site + '/drive/root:/' + path + ':/content', { method: 'PUT', body: op.blob, headers: { 'Content-Type': 'image/jpeg' } });
      }
      await DB.del(op.id);
    }
  } catch (e) { console.warn('sync', e); }
  S.syncing = false; await countPending();
  if ((await DB.all()).some(o => o.type === 'patch' && !String(o.itemId).startsWith('tmp')) && navigator.onLine) setTimeout(flush, 1500);
}
function paintSync() {
  const el = document.getElementById('sync'); if (!el) return;
  const off = !navigator.onLine;
  el.className = 'sync' + (off || S.pending ? ' off' : '');
  el.textContent = DEMO ? 'proba' : off ? (S.pending ? 'bez mreže · ' + S.pending : 'bez mreže') : S.syncing ? 'šaljem…' : S.pending ? 'čeka ' + S.pending : 'sačuvano';
}
addEventListener('online', () => { S.online = true; flush(); paintSync(); });
addEventListener('offline', () => { S.online = false; paintSync(); });

/* ---------- navigacija ---------- */
function go(view, push = true) { S.view = view; if (push) history.pushState({ v: view }, ''); render(); scrollTo(0, 0); }
addEventListener('popstate', () => {
  const back = { kupci: 'tereni', kartica: 'kupci', audit: 'kartica', novi: S.teren !== null ? 'kupci' : 'tereni' }[S.view];
  if (S.view === 'audit' && S.step > 0) { S.step--; history.pushState({ v: 'audit' }, ''); render(); return; }
  if (back) go(back, false);
});

/* ---------- prikazi ---------- */
function bar(title, back) {
  return `<header class="bar">${back ? `<button class="icon" data-a="back" aria-label="Nazad">←</button>` : ''}<h1>${esc(title)}</h1><span id="sync" class="sync"></span>${!back ? `<button class="icon" data-a="refresh" aria-label="Osveži">⟳</button>` : ''}</header>`;
}
function render() {
  const v = { tereni: vTereni, kupci: vKupci, kartica: vKartica, audit: vAudit, novi: vNovi, login: vLogin }[S.view] || vTereni;
  $app.innerHTML = (DEMO ? '<div class="demo">Probni režim – izmišljeni kupci, ništa se ne šalje</div>' : '') + v();
  paintSync();
  const sIn = document.getElementById('q'); if (sIn && S.focusQ) { sIn.focus(); sIn.setSelectionRange(sIn.value.length, sIn.value.length); S.focusQ = false; }
}
function vLogin() {
  return `<div class="login"><h1>Audit kupaca</h1><p class="s" style="margin-bottom:24px">Prijavi se službenim nalogom (@ekoline.rs).</p><button class="btn pri" data-a="login">Prijava</button></div>`;
}
function stats(list) { const tot = list.length, done = list.filter(x => x.f.Obradjeno).length; return { tot, done, pct: tot ? Math.round(done / tot * 100) : 0, prom: list.reduce((a, x) => a + num(x.f.field_19), 0) }; }
function vTereni() {
  const groups = {};
  S.items.forEach(it => { (groups[terenOf(it)] = groups[terenOf(it)] || []).push(it); });
  const keys = Object.keys(groups).sort((a, b) => a === 'Bez terena' ? -1 : b === 'Bez terena' ? 1 : stats(groups[b]).prom - stats(groups[a]).prom);
  const all = stats(S.items);
  const td = S.items.filter(x => (x.f.DatumObilaska || '').slice(0, 10) === today()).length;
  return bar('Tereni') + `<main class="main">
    <div class="sum"><div class="kpi"><span class="s">Obrađeno</span><b>${all.done} / ${all.tot}</b></div><div class="kpi"><span class="s">Danas</span><b>${td}</b></div></div>
    <div class="row" data-a="teren" data-v="*"><div class="grow"><div class="t">Svi kupci</div><div class="s">${all.tot} kupaca · promet 2026: ${money(all.prom)}</div><div class="prog"><i style="width:${all.pct}%"></i></div></div><div class="right s">${all.pct}%</div></div>
    ${keys.map(k => { const s = stats(groups[k]); return `<div class="row" data-a="teren" data-v="${esc(k)}"><div class="grow"><div class="t">${esc(k)}${k === 'Bez terena' ? ' <span class="badge warn">dodeliti</span>' : ''}</div><div class="s">${s.tot} kupaca · obrađeno ${s.done} · ${money(s.prom)}</div><div class="prog"><i style="width:${s.pct}%"></i></div></div><div class="right s">${s.done}/${s.tot}</div></div>`; }).join('')}
  </main><button class="fab" data-a="novi">＋ Nov kupac</button>`;
}
function inTeren() { return S.items.filter(it => S.teren === '*' || terenOf(it) === S.teren); }
function vKupci() {
  const { html, cnt } = kupciList();
  const title = S.teren === '*' ? 'Svi kupci' : S.teren;
  return bar(title, true) + `<main class="main">
    <input id="q" class="search" type="search" placeholder="Naziv, šifra, PIB, adresa…" value="${esc(S.q)}" data-i="q">
    <div class="seg">${[['open', 'Neobrađeni'], ['back', 'Vratiti se'], ['done', 'Obrađeni'], ['all', 'Svi']].map(([k, n]) => `<button class="${S.filter === k ? 'on' : ''}" data-a="filter" data-v="${k}">${n} ${cnt[k]}</button>`).join('')}</div>
    <div class="sorts">Redosled: ${[['promet', 'promet'], ['mesto', 'mesto i adresa'], ['naziv', 'naziv']].map(([k, n]) => `<button class="${S.sort === k ? 'on' : ''}" data-a="sort" data-v="${k}">${n}</button>`).join('')}</div>
    <div id="list">${html}</div>
  </main><button class="fab" data-a="novi">＋ Nov kupac</button>`;
}
function kupciList() {
  const base = inTeren();
  const cnt = { open: 0, back: 0, done: 0, all: base.length };
  base.forEach(it => { const i = cv(it.f.AuditIshod); if (i === 'vratiti se') cnt.back++; else if (it.f.Obradjeno) cnt.done++; else cnt.open++; });
  const q = S.q.trim().toLowerCase();
  let L = base.filter(it => {
    const i = cv(it.f.AuditIshod);
    if (S.filter === 'open' && (it.f.Obradjeno || i === 'vratiti se')) return false;
    if (S.filter === 'back' && i !== 'vratiti se') return false;
    if (S.filter === 'done' && !it.f.Obradjeno) return false;
    if (!q) return true;
    return [it.f.Title, it.f.field_1, it.f.field_2, it.f.field_3, it.f.field_4].some(x => String(x || '').toLowerCase().includes(q));
  });
  const sorters = { promet: (a, b) => num(b.f.field_19) + num(b.f.field_18) / 2 - num(a.f.field_19) - num(a.f.field_18) / 2, naziv: (a, b) => String(a.f.field_1).localeCompare(String(b.f.field_1), 'sr'), mesto: (a, b) => String(a.f.field_4).localeCompare(String(b.f.field_4), 'sr') || String(a.f.field_3).localeCompare(String(b.f.field_3), 'sr') };
  L.sort(sorters[S.sort]);
  return { cnt, html: L.length ? L.slice(0, 300).map(rowKupac).join('') + (L.length > 300 ? `<div class="empty">Prikazano 300 od ${L.length}. Suzi pretragu.</div>` : '') : '<div class="empty">Nema kupaca za ovaj izbor.</div>' };
}
function rowKupac(it) {
  const f = it.f, i = cv(f.AuditIshod), d = daysAgo(f.field_20);
  const badge = i === 'vratiti se' ? '<span class="badge warn">vratiti se</span>' : f.Obradjeno ? '<span class="badge ok">obrađen</span>' : String(f.Title).startsWith('NOVO') ? '<span class="badge brand">nov</span>' : '';
  return `<div class="row" data-a="kupac" data-v="${esc(it.id)}"><div class="grow"><div class="t ell">${esc(f.field_1 || f.Title)}</div><div class="s ell">${esc([f.field_3, f.field_4].filter(Boolean).join(', ') || '—')}</div><div class="s">${esc(cv(f.AuditSegment) || f.field_14 || 'bez segmenta')}</div></div><div class="right"><div class="t">${money(f.field_19)}</div><div class="s">${d === null ? '' : d === 0 ? 'danas' : 'pre ' + d + ' d'}</div>${badge}</div></div>`;
}
function vKartica() {
  const it = S.cur; if (!it) return vTereni(); const f = it.f;
  const tel = phoneOf(it), addr = [f.field_3, f.field_4].filter(Boolean).join(', ');
  const gps = String(f.Lokacija || '').match(/(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)/);
  const nav = gps ? `https://www.google.com/maps/dir/?api=1&destination=${gps[1]},${gps[2]}` : `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(addr + ', Srbija')}`;
  const d = daysAgo(f.field_20), i = cv(f.AuditIshod);
  return bar(f.field_1 || f.Title, true) + `<main class="main">
    <div class="s">Šifra ${esc(f.Title)}${f.field_2 ? ' · PIB ' + esc(f.field_2) : ''}</div>
    <div style="margin:4px 0 2px">${esc(addr || 'Adresa nije upisana')}</div>
    <div>${i ? `<span class="badge ${i === 'vratiti se' ? 'warn' : f.Obradjeno ? 'ok' : ''}">${esc(i)} · ${fmtDate(f.DatumObilaska)}</span>` : '<span class="badge">nije obiđen</span>'} ${cv(f.StatusNovi) ? `<span class="badge brand">${esc(cv(f.StatusNovi))}</span>` : ''}</div>
    <div class="actions">
      <a class="act" ${tel ? `href="tel:${esc(tel)}"` : 'data-a="notel"'}><span>☎</span>Pozovi</a>
      <a class="act" href="${esc(nav)}" target="_blank" rel="noopener"><span>➤</span>Navigacija</a>
      <a class="act" ${f.field_6 ? `href="mailto:${esc(split(f.field_6)[0])}"` : 'data-a="nomail"'}><span>✉</span>Mejl</a>
    </div>
    <div class="kpis">
      <div class="kpi"><span class="s">Promet 2025</span><b>${money(f.field_18)}</b></div>
      <div class="kpi"><span class="s">Promet 2026</span><b>${money(f.field_19)}</b></div>
      <div class="kpi"><span class="s">Poslednja kupovina</span><b>${f.field_20 ? fmtDate(f.field_20) : '—'}</b>${d !== null ? `<span class="s">pre ${d} dana</span>` : ''}</div>
      <div class="kpi"><span class="s">Klasa · operater</span><b>${esc(f.field_12 || '—')}</b><span class="s">${esc(f.field_10 || 'nedodeljen')}</span></div>
    </div>
    <h2>Trenutno u BizniSoft-u</h2>
    <div class="card"><dl class="kv">
      <dt>Status</dt><dd>${esc(f.field_11 || '—')}</dd>
      <dt>Segment</dt><dd>${esc(f.field_14 || '—')}</dd>
      <dt>Teren</dt><dd>${esc([f.field_8, f.field_9].filter(Boolean).join(' · ') || '—')}</dd>
      <dt>Razvoz</dt><dd>${esc(f.field_15 || '—')}</dd>
      <dt>Poziv</dt><dd>${esc(f.field_16 || '—')}</dd>
      <dt>Kontakt</dt><dd>${esc([f.field_7, f.field_5].filter(Boolean).join(' · ') || '—')}</dd>
      ${f.field_17 ? `<dt>NETO</dt><dd>${esc(f.field_17)}</dd>` : ''}
    </dl></div>
    ${f.Beleske ? `<h2>Beleške</h2><div class="card" style="white-space:pre-wrap">${esc(f.Beleske)}</div>` : ''}
  </main><div class="foot"><div class="in"><button class="btn pri" data-a="start">${i ? 'Ponovi audit' : 'Pokreni audit'}</button></div></div>`;
}

/* ---------- audit ---------- */
function startAudit() {
  const f = S.cur.f;
  const bsSeg = split(f.field_14)[0];
  S.d = {
    StatusNovi: cv(f.StatusNovi) || '', AuditSegment: cv(f.AuditSegment) || '', AuditTeren: cv(f.AuditTeren) || '',
    AuditNacinRazvoza: cv(f.AuditNacinRazvoza) || '', AuditRazvoz: arr(f.AuditRazvoz).slice(),
    AuditNacinPozivanja: cv(f.AuditNacinPozivanja) || '', AuditPozivanje: arr(f.AuditPozivanje).slice(),
    NarucujeIme: f.NarucujeIme || f.field_7 || '', NarucujeTel: f.NarucujeTel || phoneOf(S.cur) || '',
    OdlucujeIme: f.OdlucujeIme || '', OdlucujeTel: f.OdlucujeTel || '', PlacaIme: f.PlacaIme || '', PlacaTel: f.PlacaTel || '',
    RadnoVreme: f.RadnoVreme || '', AuditNacinPlacanja: cv(f.AuditNacinPlacanja) || '',
    Velicina: cv(f.Velicina) || '', Potencijal: cv(f.Potencijal) || '',
    DrziAsortiman: arr(f.DrziAsortiman).slice(), NeDrzi: arr(f.NeDrzi).slice(),
    NabavljaOd: f.NabavljaOd || '', AuditPlanObilaska: cv(f.AuditPlanObilaska) || '', AuditMonitoring: !!f.AuditMonitoring,
    Lokacija: f.Lokacija || '', Beleske: '', AuditIshod: '', _bsSeg: bsSeg
  };
  S.photos = []; S.step = 0; go('audit');
}
function vAudit() {
  const st = STEPS[S.step], d = S.d, f = S.cur.f;
  const bsv = st.bs ? st.bs(f) : null;
  const dots = STEPS.map((_, i) => `<i class="${i < S.step ? 'done' : i === S.step ? 'now' : ''}"></i>`).join('');
  let body = '';
  if (st.type === 'one') {
    const was = o => bsv && split(bsv).some(b => { const B = b.toUpperCase(), O = o.toUpperCase(); return B === O || B.endsWith(' ' + O); });
    const opts = st.o.map(o => `<button class="opt ${d[st.f] === o ? 'on' : ''} ${was(o) ? 'was' : ''}" data-a="pick" data-v="${esc(o)}">${esc(o)}</button>`).join('');
    body = (st.find ? `<input class="search" placeholder="Pretraga…" data-i="optq">` : '') + `<div class="opts" id="opts">${opts}</div>`;
    if (st.extra === 'monitoring') body += `<div class="toggle" data-a="mon"><div><div class="t">Monitoring</div><div class="s">Kupac se posebno prati</div></div><div class="sw ${d.AuditMonitoring ? 'on' : ''}"></div></div>`;
  } else if (st.type === 'days') {
    const days = st.six ? 6 : 5, sel = d[st.f];
    body = `<div class="grid"><span></span>${[1, 2, 3, 4].map(w => `<span class="h">${w}. ned</span>`).join('')}
      ${DANI.slice(0, days).map((dn, i) => `<span class="d">${DANI_K[i]}</span>${[1, 2, 3, 4].map(w => { const v = st.pre + dn + ' ' + w; return `<button class="cell ${sel.includes(v) ? 'on' : ''}" data-a="day" data-v="${esc(v)}" aria-label="${esc(v)}"></button>`; }).join('')}`).join('')}</div>
      <div class="quick">${DANI.slice(0, days).map((dn, i) => `<button data-a="dayrow" data-v="${esc(st.pre + dn)}">${DANI_K[i]} svake nedelje</button>`).join('')}<button data-a="dayclear">Obriši</button></div>
      <p class="s">Ciklus od 4 nedelje se ponavlja.</p>`;
  } else if (st.type === 'contacts') {
    const fl = (k, l, t) => `<div class="fld"><label>${l}</label><input data-i="${k}" type="${t}" value="${esc(d[k])}" ${t === 'tel' ? 'inputmode="tel"' : ''}></div>`;
    body = `<div class="card"><div class="t" style="margin-bottom:8px">Ko naručuje</div>${fl('NarucujeIme', 'Ime', 'text')}${fl('NarucujeTel', 'Telefon', 'tel')}</div>
      <div class="card" style="margin-top:8px"><div class="t" style="margin-bottom:8px">Ko odlučuje</div>${fl('OdlucujeIme', 'Ime', 'text')}${fl('OdlucujeTel', 'Telefon', 'tel')}</div>
      <div class="card" style="margin-top:8px"><div class="t" style="margin-bottom:8px">Ko plaća</div>${fl('PlacaIme', 'Ime', 'text')}${fl('PlacaTel', 'Telefon', 'tel')}</div>`;
  } else if (st.type === 'text') {
    body = `<div class="fld"><input data-i="${st.f}" value="${esc(d[st.f])}" placeholder="${esc(st.ph)}"></div><div class="quick">${st.quick.map(q => `<button data-a="quick" data-v="${esc(q)}">${esc(q)}</button>`).join('')}</div>`;
  } else if (st.type === 'area') {
    body = `<div class="fld"><textarea data-i="${st.f}" placeholder="${esc(st.ph)}">${esc(d[st.f])}</textarea></div>`;
  } else if (st.type === 'pair') {
    const grp = (k, o, l) => `<div class="t" style="margin:6px 2px 8px">${l}</div><div class="opts">${o.map(x => `<button class="opt ${d[k] === x ? 'on' : ''}" data-a="pickk" data-k="${k}" data-v="${x}">${x}</button>`).join('')}</div>`;
    body = grp('Velicina', VEL, 'Veličina objekta') + '<div style="height:14px"></div>' + grp('Potencijal', POT, 'Potencijal za nas');
  } else if (st.type === 'asort') {
    body = `<div class="card">${ASORT.map(a => { const y = d.DrziAsortiman.includes(a), n = d.NeDrzi.includes(a); return `<div class="tri"><span class="n">${esc(a)}</span><button class="${y ? 'y' : ''}" data-a="as" data-k="y" data-v="${esc(a)}">drži</button><button class="${n ? 'x' : ''}" data-a="as" data-k="n" data-v="${esc(a)}">ne drži</button></div>`; }).join('')}</div>`;
  } else if (st.type === 'gps') {
    body = `<button class="big-act" data-a="gps"><span>📍</span>${d.Lokacija ? 'Zabeleži ponovo' : 'Zabeleži lokaciju'}</button>
      ${d.Lokacija ? `<p class="s" style="margin-top:10px">Zabeleženo: ${esc(d.Lokacija)}</p>` : '<p class="s" style="margin-top:10px">Stani ispred objekta i pritisni dugme.</p>'}`;
  } else if (st.type === 'photo') {
    body = `<label class="big-act"><span>📷</span>Slikaj objekat<input type="file" accept="image/*" capture="environment" data-i="photo" hidden></label>
      <div class="thumbs">${S.photos.map(p => `<img src="${p.url}" alt="">`).join('')}</div>
      <p class="s">Može i više slika: izlog, police, unutrašnjost.</p>`;
  } else if (st.type === 'end') {
    body = `<div class="fld"><label>Beleška sa obilaska</label><textarea data-i="Beleske" placeholder="Šta je dogovoreno, šta treba uraditi, kada se vratiti…">${esc(d.Beleske)}</textarea></div>
      <div class="t" style="margin:6px 2px 8px">Ishod obilaska</div><div class="opts">${ISH.map(x => `<button class="opt ${d.AuditIshod === x ? 'on' : ''}" data-a="pickk" data-k="AuditIshod" data-v="${esc(x)}">${esc(x)}</button>`).join('')}</div><div class="err" id="err"></div>`;
  }
  const last = S.step === STEPS.length - 1;
  return bar(f.field_1 || f.Title, true) + `<main class="main">
    <div class="dots">${dots}</div><div class="s">Korak ${S.step + 1} od ${STEPS.length}</div>
    <div class="step-h">${esc(st.t)}</div>
    ${bsv !== null && bsv !== undefined ? `<div class="step-bs">U BizniSoft-u: <b>${esc(bsv || 'nije upisano')}</b></div>` : `<div class="step-bs">${esc(st.hint || '')}</div>`}
    ${body}
  </main><div class="foot"><div class="in"><button class="btn" data-a="prev">${S.step ? 'Nazad' : 'Odustani'}</button><button class="btn pri" style="flex:2" data-a="next">${last ? 'Završi i sačuvaj' : 'Dalje'}</button></div></div>`;
}
async function finishAudit() {
  const d = S.d;
  if (!d.AuditIshod) { document.getElementById('err').textContent = 'Izaberi ishod obilaska.'; return; }
  const f = S.cur.f, stamp = new Date().toLocaleString('sr-RS');
  const note = d.Beleske.trim() ? `[${stamp} · ${d.AuditIshod}] ${d.Beleske.trim()}` : `[${stamp} · ${d.AuditIshod}]`;
  const fields = {};
  ['StatusNovi', 'AuditSegment', 'AuditTeren', 'AuditNacinRazvoza', 'AuditNacinPozivanja', 'AuditNacinPlacanja', 'Velicina', 'Potencijal', 'AuditPlanObilaska', 'AuditIshod'].forEach(k => { if (d[k]) fields[k] = d[k]; });
  ['NarucujeIme', 'NarucujeTel', 'OdlucujeIme', 'OdlucujeTel', 'PlacaIme', 'PlacaTel', 'RadnoVreme', 'NabavljaOd', 'Lokacija'].forEach(k => { fields[k] = d[k] || ''; });
  ['AuditRazvoz', 'AuditPozivanje', 'DrziAsortiman', 'NeDrzi'].forEach(k => { fields[k + '@odata.type'] = 'Collection(Edm.String)'; fields[k] = d[k]; });
  fields.AuditMonitoring = !!d.AuditMonitoring;
  fields.Beleske = f.Beleske ? note + '\n' + f.Beleske : note;
  fields.DatumObilaska = today() + 'T12:00:00Z';
  fields.Obradjeno = d.AuditIshod !== 'vratiti se';
  // lokalno odmah
  Object.keys(fields).forEach(k => { if (!k.includes('@')) f[k] = fields[k]; });
  saveItems();
  await queue({ type: 'patch', itemId: S.cur.id, fields });
  for (const p of S.photos) await queue({ type: 'photo', name: p.name, blob: p.blob });
  S.photos.forEach(p => URL.revokeObjectURL(p.url)); S.photos = [];
  toast(navigator.onLine || DEMO ? 'Sačuvano.' : 'Sačuvano u telefonu, šalje se kad bude mreže.');
  S.filter = 'open'; go('kupci');
}

/* ---------- nov kupac ---------- */
function vNovi() {
  return bar('Nov kupac', true) + `<main class="main">
    <p class="s">Radnja koje nema na spisku. Posle upisa odmah kreće audit.</p>
    ${[['n_naziv', 'Naziv radnje', 'text'], ['n_adresa', 'Adresa', 'text'], ['n_mesto', 'Mesto', 'text'], ['n_tel', 'Telefon', 'tel'], ['n_kontakt', 'Kontakt osoba', 'text']].map(([k, l, t]) => `<div class="fld"><label>${l}</label><input id="${k}" type="${t}"></div>`).join('')}
    <div class="err" id="err"></div>
  </main><div class="foot"><div class="in"><button class="btn pri" data-a="novisave">Upiši i pokreni audit</button></div></div>`;
}
async function saveNovi() {
  const v = id => document.getElementById(id).value.trim();
  if (!v('n_naziv')) { document.getElementById('err').textContent = 'Upiši naziv radnje.'; return; }
  const now = new Date(), code = 'NOVO-' + now.toISOString().slice(2, 16).replace(/\D/g, '');
  const fields = { Title: code, field_1: v('n_naziv'), field_3: v('n_adresa'), field_4: v('n_mesto'), field_5: v('n_tel'), field_7: v('n_kontakt'), StatusNovi: 'potencijalni' };
  if (S.teren && S.teren !== '*' && S.teren !== 'Bez terena') fields.AuditTeren = S.teren;
  const it = { id: 'tmp' + now.getTime(), f: { ...fields } };
  S.items.push(it); saveItems();
  await queue({ type: 'create', tmp: it.id, fields });
  S.cur = it; startAudit();
}

/* ---------- događaji ---------- */
document.addEventListener('click', async e => {
  const el = e.target.closest('[data-a]'); if (!el) return;
  const a = el.dataset.a, v = el.dataset.v, d = S.d;
  const rerender = () => { const y = scrollY; render(); scrollTo(0, y); };
  switch (a) {
    case 'back': history.back(); break;
    case 'login': S.pca.loginRedirect({ scopes: SCOPES }); break;
    case 'refresh': await refresh(true); break;
    case 'teren': S.teren = v; S.q = ''; S.filter = 'open'; go('kupci'); break;
    case 'filter': S.filter = v; rerender(); break;
    case 'sort': S.sort = v; rerender(); break;
    case 'kupac': S.cur = S.items.find(x => String(x.id) === v); go('kartica'); break;
    case 'notel': toast('Telefon nije upisan.'); break;
    case 'nomail': toast('Mejl nije upisan.'); break;
    case 'start': startAudit(); break;
    case 'novi': go('novi'); break;
    case 'novisave': await saveNovi(); break;
    case 'pick': d[STEPS[S.step].f] = d[STEPS[S.step].f] === v ? '' : v; rerender(); break;
    case 'pickk': d[el.dataset.k] = d[el.dataset.k] === v ? '' : v; rerender(); break;
    case 'mon': d.AuditMonitoring = !d.AuditMonitoring; rerender(); break;
    case 'day': { const k = STEPS[S.step].f, i = d[k].indexOf(v); i < 0 ? d[k].push(v) : d[k].splice(i, 1); rerender(); break; }
    case 'dayrow': { const k = STEPS[S.step].f, vals = [1, 2, 3, 4].map(w => v + ' ' + w), allOn = vals.every(x => d[k].includes(x)); d[k] = allOn ? d[k].filter(x => !vals.includes(x)) : [...new Set(d[k].concat(vals))]; rerender(); break; }
    case 'dayclear': d[STEPS[S.step].f] = []; rerender(); break;
    case 'quick': { const k = STEPS[S.step].f; d[k] = d[k] ? d[k] + ', ' + v : v; rerender(); break; }
    case 'as': { const y = d.DrziAsortiman, n = d.NeDrzi; [y, n].forEach(l => { const i = l.indexOf(v); if (i >= 0) l.splice(i, 1); }); if (!(el.classList.contains('y') || el.classList.contains('x'))) (el.dataset.k === 'y' ? y : n).push(v); rerender(); break; }
    case 'gps': getGps(); break;
    case 'prev': if (S.step) { S.step--; go('audit', false); } else history.back(); break;
    case 'next': if (S.step < STEPS.length - 1) { S.step++; go('audit', false); } else await finishAudit(); break;
  }
});
document.addEventListener('input', e => {
  const k = e.target.dataset.i; if (!k) return;
  if (k === 'q') { S.q = e.target.value; const l = document.getElementById('list'); if (l) l.innerHTML = kupciList().html; return; }
  if (k === 'optq') { const q = e.target.value.toLowerCase(); document.querySelectorAll('#opts .opt').forEach(b => b.style.display = b.textContent.toLowerCase().includes(q) ? '' : 'none'); return; }
  if (k === 'photo') return;
  if (S.d) S.d[k] = e.target.value;
});
document.addEventListener('change', async e => {
  if (e.target.dataset.i !== 'photo' || !e.target.files.length) return;
  for (const file of e.target.files) {
    const blob = await shrink(file);
    const n = S.photos.length + 1, stamp = new Date().toISOString().slice(0, 16).replace(/\D/g, '');
    S.photos.push({ blob, url: URL.createObjectURL(blob), name: `${S.cur.f.Title}-${stamp}-${n}.jpg` });
  }
  render();
});
function getGps() {
  if (!navigator.geolocation) { toast('Telefon ne daje lokaciju.'); return; }
  toast('Tražim lokaciju…');
  navigator.geolocation.getCurrentPosition(p => {
    S.d.Lokacija = p.coords.latitude.toFixed(6) + ', ' + p.coords.longitude.toFixed(6) + ' (±' + Math.round(p.coords.accuracy) + ' m)';
    render(); toast('Lokacija zabeležena.');
  }, err => toast('Nema lokacije: ' + (err.code === 1 ? 'dozvoli pristup lokaciji' : 'pokušaj ponovo napolju')), { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 });
}
function shrink(file) {
  return new Promise(res => {
    const img = new Image(), url = URL.createObjectURL(file);
    img.onload = () => {
      const m = 1600, k = Math.min(1, m / Math.max(img.width, img.height));
      const c = document.createElement('canvas'); c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
      c.toBlob(b => res(b || file), 'image/jpeg', 0.8);
    };
    img.onerror = () => res(file);
    img.src = url;
  });
}

/* ---------- učitavanje ---------- */
async function refresh(manual) {
  if (DEMO) { if (manual) toast('Probni režim.'); return; }
  if (!navigator.onLine) { if (manual) toast('Nema mreže – radiš sa podacima iz telefona.'); return; }
  try {
    if (manual) toast('Osvežavam…');
    await flush();
    const fresh = await fetchAll();
    const pend = await DB.all();
    // zadrži lokalne nove kupce koji još nisu poslati
    const tmp = S.items.filter(x => String(x.id).startsWith('tmp'));
    S.items = fresh.concat(tmp);
    // ponovo primeni izmene koje još čekaju
    pend.filter(o => o.type === 'patch').forEach(o => { const it = S.items.find(x => x.id === o.itemId); if (it) Object.keys(o.fields).forEach(k => { if (!k.includes('@')) it.f[k] = o.fields[k]; }); });
    if (S.cur) S.cur = S.items.find(x => x.id === S.cur.id) || S.cur;
    saveItems();
    if (S.view !== 'audit' && S.view !== 'novi') render();
    if (manual) toast('Osveženo: ' + S.items.length + ' kupaca.');
  } catch (e) { console.warn(e); toast('Osvežavanje nije uspelo: ' + String(e.message || e).slice(0, 80)); }
}
async function boot() {
  history.replaceState({ v: 'tereni' }, '');
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
  if (DEMO) { S.items = demoItems(); S.view = 'tereni'; render(); return; }
  try { await initAuth(); } catch (e) { $app.innerHTML = `<div class="login"><h1>Greška pri prijavi</h1><p class="s">${esc(e.message)}</p></div>`; return; }
  if (!S.account) { S.view = 'login'; render(); return; }
  const cached = loadCached();
  if (cached) { S.items = cached; render(); }
  await countPending();
  await refresh(false);
  if (!cached && !S.items.length) render();
}

/* ---------- probni podaci (izmišljeni) ---------- */
function demoItems() {
  const raw = [
    ['10001', 'Zdrava kuća', 'Cara Lazara 12', 'Kruševac', '1.01 KRUŠEVAC', 3360000, 2510000, 'A1', 'GORDANA', 'ZDRAVA HRANA NAMENSKA', 'V PONEDELJAK 1, V PONEDELJAK 3', 'PETAK 2', '+381641112233', 'Jelena Petrović'],
    ['10002', 'Biljna apoteka Neven', 'Vidovdanska 4', 'Kruševac', '1.01 KRUŠEVAC', 1840000, 1020000, 'B1', 'GORDANA', 'BILJNA APOTEKA', 'V UTORAK 1', 'PONEDELJAK 1', '+381652223344', 'Marija Ilić'],
    ['10003', 'Pekara Klas', 'Kosovska 8', 'Kruševac', '1.01 KRUŠEVAC', 920000, 610000, 'C1', 'TAMARA', 'PEKARA', '', '', '+381603334455', ''],
    ['10004', 'Piljara Kod Mire', 'Gradska pijaca', 'Kruševac', '1.01 KRUŠEVAC', 210000, 80000, 'C3', '', '', '', '', '', ''],
    ['10005', 'Zrno', 'Vuka Karadžića 3', 'Trstenik', '1.02 TRSTENIK', 1310000, 990000, 'B2', 'PETAR', 'ZDRAVA HRANA NAMENSKA', 'V SREDA 1, V SREDA 3', 'UTORAK 1, UTORAK 3', '+381631234567', 'Dragan'],
    ['10006', 'Slatki kutak', 'Kneza Miloša 21', 'Kragujevac', '5.04 KRAGUJEVAC', 2200000, 1700000, 'A3', 'GORDANA', 'CANDY SHOP', 'V ČETVRTAK 2', 'SREDA 2', '+381649876543', 'Ana'],
    ['10007', 'Bistro Bašta', 'Obrenovićeva 40', 'Niš', '4.01 NIŠ', 640000, 520000, 'C1', 'MILAN', 'KAFANA / RESTORAN / KAFIĆ SA HRANOM / PICERIJA', '', '', '+381611112222', ''],
    ['10008', 'Market Jabuka', 'Zelena pijaca 19', 'Jagodina', '', 400000, 120000, 'C2', '', 'MEŠOVITA ROBA / TT', '', '', '+381621231231', '']
  ];
  return raw.map((r, i) => ({ id: 'd' + i, f: {
    Title: r[0], field_1: r[1], field_3: r[2], field_4: r[3], field_9: r[4], field_8: r[4] ? '1 CENTAR' : '', field_18: r[5], field_19: r[6], field_12: r[7], field_10: r[8],
    field_14: r[9], field_15: r[10], field_16: r[11], field_5: r[12], field_7: r[13], field_11: 'KUPAC AKTIVAN', field_20: new Date(Date.now() - (i * 9 + 2) * 864e5).toISOString().slice(0, 10),
    AuditTeren: r[4] ? { Value: r[4] } : null, AuditSegment: r[9] && SEG.includes(r[9]) ? r[9] : '', AuditRazvoz: split(r[10]), AuditPozivanje: split(r[11]), StatusNovi: 'aktivan', Obradjeno: i === 2
  } }));
}

boot();
