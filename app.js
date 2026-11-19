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
  { k: 'foto', t: 'Fotografija objekta', type: 'photo' },
  { k: 'kraj', t: 'Beleška i ishod', type: 'end' }
];

/* ---------- stanje ---------- */
const S = {
  items: [], site: null, list: null, account: null, pca: null,
  view: 'danas', teren: null, from: 'danas', ktab: 'pregled', log: [], lastV: {}, pravac: null, showAll: false, nk: null, noteNew: false, filter: 'open', sort: 'promet', q: '',
  cur: null, step: 0, d: null, photos: [], online: navigator.onLine, pending: 0, syncing: false,
  logList: null, day: null, visit: null, busy: false
};
const RADIUS = C.visitRadius || 50, MAXACC = C.maxAccuracy || 100, MAPACC = C.mapAccuracy || 10, MAPWAIT = 45000;
const LS_DAY = 'ak_day_v1', LS_VISIT = 'ak_visit_v1';
try { S.day = JSON.parse(localStorage.getItem(LS_DAY) || 'null'); S.visit = JSON.parse(localStorage.getItem(LS_VISIT) || 'null'); } catch (e) { }
const saveDay = () => { try { localStorage.setItem(LS_DAY, JSON.stringify(S.day)); localStorage.setItem(LS_VISIT, JSON.stringify(S.visit)); } catch (e) { } };
const dayOn = () => S.day && S.day.date === new Date().toLocaleDateString('sr-RS') && !S.day.end;
const userName = () => DEMO ? 'proba@ekoline.rs' : (S.account && S.account.username) || '';
const hhmm = t => new Date(t).toLocaleTimeString('sr-RS', { hour: '2-digit', minute: '2-digit' });

/* ---------- lokacija ---------- */
function getPos() {
  if (DEMO) return Promise.resolve({ lat: 43.5800 + Math.random() * 0.0004, lon: 21.3300 + Math.random() * 0.0004, acc: 12 });
  return new Promise((res, rej) => {
    if (!navigator.geolocation) return rej(new Error('Telefon ne daje lokaciju.'));
    navigator.geolocation.getCurrentPosition(
      p => res({ lat: p.coords.latitude, lon: p.coords.longitude, acc: Math.round(p.coords.accuracy) }),
      e => rej(new Error(e.code === 1 ? 'Dozvoli aplikaciji pristup lokaciji (podešavanja pretraživača).' : 'Lokacija nije dostupna. Izađi napolje ili uključi GPS i pokušaj ponovo.')),
      { enableHighAccuracy: true, timeout: 25000, maximumAge: 0 });
  });
}
const parseLoc = s => { const m = String(s || '').match(/(-?\d+\.\d+)\s*,\s*(-?\d+\.\d+)(?:\s*\(±(\d+))?/); return m ? { lat: +m[1], lon: +m[2], acc: m[3] ? +m[3] : MAPACC } : null; };
// precizno mapiranje objekta: prati GPS dok preciznost ne padne na MAPACC ili istekne vreme
function getPrecise(onTick) {
  if (DEMO) return new Promise(res => { let a = 24; const t = setInterval(() => { a -= 7; onTick && onTick(Math.max(a, 4)); if (a <= 4) { clearInterval(t); res({ lat: 43.5800 + Math.random() * 0.00005, lon: 21.3300 + Math.random() * 0.00005, acc: 4 }); } }, 300); });
  return new Promise((res, rej) => {
    if (!navigator.geolocation) return rej(new Error('Telefon ne daje lokaciju.'));
    let best = null, done = false;
    const stop = (fn, v) => { if (done) return; done = true; navigator.geolocation.clearWatch(w); clearTimeout(t); fn(v); };
    const w = navigator.geolocation.watchPosition(p => {
      const c = { lat: p.coords.latitude, lon: p.coords.longitude, acc: Math.round(p.coords.accuracy) };
      if (!best || c.acc < best.acc) best = c;
      onTick && onTick(c.acc);
      if (c.acc <= MAPACC) stop(res, c);
    }, e => { if (e.code === 1) stop(rej, new Error('Dozvoli aplikaciji pristup lokaciji (podešavanja pretraživača).')); }, { enableHighAccuracy: true, maximumAge: 0 });
    const t = setTimeout(() => stop(rej, new Error(best ? `Signal nije dovoljno precizan (najbolje ±${best.acc} m, treba ±${MAPACC} m). Stani ispred ulaza, dalje od zidova i krošnji, pa pokušaj ponovo.` : 'Nema GPS signala. Izađi napolje i pokušaj ponovo.')), MAPWAIT);
  });
}
// provera mesta pre prvog mapiranja: telefon mora biti blizu mesta kupca (OpenStreetMap, keš u telefonu)
const TOWNKM = C.townRadiusKm || 15;
async function townPos(name) {
  const key = String(name || '').trim().toUpperCase(); if (!key) return null;
  let cache = {}; try { cache = JSON.parse(localStorage.getItem('ak_towns_v1') || '{}'); } catch (e) { }
  if (cache[key] !== undefined) return cache[key];
  if (DEMO) { const demo = { 'KRUŠEVAC': { lat: 43.58, lon: 21.33 }, 'TRSTENIK': { lat: 43.617, lon: 21.0 }, 'KRAGUJEVAC': { lat: 44.012, lon: 20.911 }, 'NIŠ': { lat: 43.32, lon: 21.896 }, 'JAGODINA': { lat: 43.977, lon: 21.261 } }; return demo[key] || null; }
  try {
    const ctl = new AbortController(); setTimeout(() => ctl.abort(), 8000);
    const r = await fetch('https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=rs,ba,me&q=' + encodeURIComponent(name), { signal: ctl.signal, headers: { 'Accept-Language': 'sr' } });
    const j = await r.json(); const v = j && j[0] ? { lat: +j[0].lat, lon: +j[0].lon } : null;
    cache[key] = v; try { localStorage.setItem('ak_towns_v1', JSON.stringify(cache)); } catch (e) { }
    return v;
  } catch (e) { return undefined; } // nema mreže – ne može da se proveri
}
const fmtLoc = p => p.lat.toFixed(6) + ', ' + p.lon.toFixed(6) + ' (±' + p.acc + ' m)';
function dist(a, b) { const R = 6371000, r = x => x * Math.PI / 180, dl = r(b.lat - a.lat), dn = r(b.lon - a.lon); const h = Math.sin(dl / 2) ** 2 + Math.cos(r(a.lat)) * Math.cos(r(b.lat)) * Math.sin(dn / 2) ** 2; return Math.round(2 * R * Math.asin(Math.sqrt(h))); }
async function logEvent(type, pos, it, extra = {}) {
  const f = { Title: type, Korisnik: userName(), Vreme: new Date().toISOString() };
  if (pos) { f.Lat = pos.lat; f.Lon = pos.lon; f.Tacnost = pos.acc; }
  if (it) { f.KupacSifra = String(it.f.Title || ''); f.KupacNaziv = String(it.f.field_1 || ''); f.KupacId = String(it.id); }
  Object.assign(f, extra);
  S.log.push({ t: f.Title, v: f.Vreme, s: f.KupacSifra || '', n: f.KupacNaziv || '', k: f.Korisnik, u: f.Udaljenost, m: f.Napomena || '' }); saveLog(); buildLogIndex();
  await queue({ type: 'log', fields: f });
}
if (DEMO) window.__t = { S, get DB() { return DB; } };
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
/* ---------- plan obilaska ---------- */
const OBIL_DAYS = { '801': 7, '802': 14, '803': 30, '804': 91, '805': 182, '806': 365, '809': 0 };
const FAR_TER = ['2.01', '2.02', '2.03']; // daleki tereni (Negotin, Zaječar, Bor): najviše kvartalno
const klasaOf = it => { const k = String(it.f.field_13 || it.f.field_12 || '').trim().toUpperCase(); return /^[A-E]/.test(k) ? k[0] : ''; };
function planOf(it) {
  const a = cv(it.f.AuditPlanObilaska);
  if (a) { const c = String(a).slice(0, 3); return { code: c, days: OBIL_DAYS[c] ?? 0, label: String(a).slice(4).toLowerCase(), set: true }; }
  const t = terenOf(it);
  if (/BEOGRAD|NOVI SAD/.test(t)) return { code: '809', days: 0, label: 'ne obilazi se (predlog)', set: false };
  let c = { A: '802', B: '803', C: '804', D: '805', E: '806' }[klasaOf(it)] || '804';
  if (FAR_TER.some(p => t.startsWith(p)) && OBIL_DAYS[c] < 91) c = '804';
  return { code: c, days: OBIL_DAYS[c], label: OBIL.find(o => o.startsWith(c)).slice(4).toLowerCase() + ' (predlog)', set: false };
}
function lastVisit(it) { const a = S.lastV[String(it.f.Title)] || 0, b = it.f.DatumObilaska ? Date.parse(it.f.DatumObilaska) || 0 : 0; return Math.max(a, b) || null; }
// null = ne ide na plan; over = dana kašnjenja (negativno = za toliko dana dolazi na red)
function dueOf(it) {
  if (cv(it.f.StatusNovi) === 'zatvoren') return null;
  const p = planOf(it), ob = !!it.f.Obici;
  if (!p.days && !ob) return null;
  const lv = lastVisit(it);
  if (ob) return { top: true, ob, lv, p, over: lv ? Math.floor((Date.now() - lv) / 864e5) : null };
  if (!lv) return { top: true, never: true, p, over: null };
  return { top: false, lv, p, over: Math.floor((Date.now() - lv - p.days * 864e5) / 864e5) };
}
const LS_LOG = 'ak_log_v1', LS_PRAVAC = 'ak_pravac_v1';
try { S.log = JSON.parse(localStorage.getItem(LS_LOG) || '[]'); } catch (e) { S.log = []; }
try { const pv = JSON.parse(localStorage.getItem(LS_PRAVAC) || 'null'); S.pravac = pv && pv.date === new Date().toLocaleDateString('sr-RS') ? pv : null; } catch (e) { }
function saveLog() { try { localStorage.setItem(LS_LOG, JSON.stringify(S.log.slice(-5000))); } catch (e) { } }
function buildLogIndex() {
  S.lastV = {};
  S.log.forEach(e => { if (e.s && /POSETE/.test(e.t)) { const t = Date.parse(e.v) || 0; if (t > (S.lastV[e.s] || 0)) S.lastV[e.s] = t; } });
}
buildLogIndex();
/* ---------- filteri ekrana Danas ---------- */
const norm = s => String(s || '').toUpperCase().replace(/Đ/g, 'DJ').normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
const daysOf = v => { const t = norm(arr(v).join(',')), out = DANI.filter(d => t.includes(norm(d))); return out.length ? out : ['nije upisano']; };
const FDEF = {
  pravac: { l: 'po pravcu', ic: '🧭', t: 'Pravac', get: it => [terenOf(it)] },
  kat: { l: 'po kategoriji', ic: '🏷', t: 'Kategorija', get: it => [klasaOf(it) || 'bez kategorije'] },
  tip: { l: 'po tipu kupca', ic: '🏪', t: 'Tip kupca', get: it => [cv(it.f.AuditSegment) || 'bez segmenta'] },
  poz: { l: 'po danu poziva', ic: '📞', t: 'Dan poziva', get: it => daysOf(it.f.AuditPozivanje) },
  raz: { l: 'po danu razvoza', ic: '🚚', t: 'Dan razvoza', get: it => daysOf(it.f.AuditRazvoz) },
  st: { t: 'Status', get: it => [cv(it.f.StatusNovi) || 'nije upisan'] },
  plac: { t: 'Način plaćanja', get: it => [cv(it.f.AuditNacinPlacanja) || 'nije upisan'] },
  op: { t: 'Operater', get: it => [it.f.field_10 || 'nedodeljen'] },
  plan: { t: 'Plan obilaska', get: it => [planOf(it).label] },
  vel: { t: 'Veličina', get: it => [cv(it.f.Velicina) || 'nije upisana'] },
  pot: { t: 'Potencijal', get: it => [cv(it.f.Potencijal) || 'nije upisan'] },
  ob: { t: 'Oznaka „obići”', get: it => [it.f.Obici ? 'označen' : 'nije označen'] },
  lok: { t: 'Lokacija objekta', get: it => [parseLoc(it.f.Lokacija) ? 'zabeležena' : 'nije zabeležena'] },
  fin: { t: 'Plaćanje (BizniSoft)', get: it => { const f = finOf(it), b = bsKupac(it), o = []; if (b && String(b.blok) === '1') o.push('blokiran'); if (f && f.d > 0 && f.m > 30) o.push('kasni više od 30 dana'); if (f && f.d > 0) o.push('ima dospeli dug'); if (f && f.s > 0) o.push('ima saldo'); if (!o.length) o.push('bez duga'); return o; } }
};
const MAINF = ['pravac', 'kat', 'tip', 'poz', 'raz'], MOREF = ['fin', 'st', 'plac', 'op', 'plan', 'vel', 'pot', 'ob', 'lok'];
const LS_FLT = 'ak_flt_v1';
S.flt = {}; S.dq = ''; S.onlyDue = true; S.obOpen = true;
try { const x = JSON.parse(localStorage.getItem(LS_FLT) || 'null'); if (x && x.date === new Date().toLocaleDateString('sr-RS')) S.flt = x.f || {}; } catch (e) { }
if (!S.flt.pravac && S.pravac && S.pravac.t) S.flt.pravac = S.pravac.t;
const saveFlt = () => { try { localStorage.setItem(LS_FLT, JSON.stringify({ date: new Date().toLocaleDateString('sr-RS'), f: S.flt })); } catch (e) { } };
const selOf = k => S.flt[k] || [];
const fMatch = (it, skip) => Object.keys(S.flt).every(k => k === skip || !selOf(k).length || FDEF[k].get(it).some(v => selOf(k).includes(v)));
function fOptions(k) {
  const base = k === 'pravac' ? S.items : S.items.filter(it => !selOf('pravac').length || selOf('pravac').includes(terenOf(it)));
  const c = {}; base.forEach(it => FDEF[k].get(it).forEach(v => c[v] = (c[v] || 0) + 1));
  const ord = { kat: ['A', 'B', 'C', 'D', 'E'], poz: DANI, raz: DANI }[k];
  return Object.keys(c).sort((a, b) => ord ? ((ord.indexOf(a) + 1 || 99) - (ord.indexOf(b) + 1 || 99)) : /^nije|^bez|^ned/.test(a) ? 1 : /^nije|^bez|^ned/.test(b) ? -1 : a.localeCompare(b, 'sr', { numeric: true })).map(v => [v, c[v]]);
}
// prozori: dugme Nazad (Android) zatvara prozor, a ne ekran
function modalOpen(m, close) { m._close = close; history.pushState({ modal: 1 }, ''); document.body.appendChild(m); }
function modalDone(m, cb) { m.remove(); if (history.state && history.state.modal) { S.skipPop = true; S.afterPop = cb || null; history.back(); setTimeout(() => { if (S.afterPop) { const c = S.afterPop; S.afterPop = null; S.skipPop = false; c(); } }, 400); } else if (cb) cb(); }
function pickSheet(title, groups) { // groups: [[key, title, [[val,count]], selected[]]]
  return new Promise(res => {
    const sel = {}; groups.forEach(([k, , , cur]) => sel[k] = new Set(cur));
    const m = document.createElement('div'); m.className = 'modal sheet';
    const paint = () => m.innerHTML = `<div class="box"><div class="mh"><h2>${esc(title)}</h2><button class="x" data-m="x" aria-label="Zatvori">✕</button></div><div class="sbody">${groups.map(([k, t, opts]) => `${groups.length > 1 ? `<div class="sh">${esc(t)}</div>` : ''}${opts.map(([v, n]) => `<button class="ck ${sel[k].has(v) ? 'on' : ''}" data-k="${k}" data-v="${esc(v)}"><span class="box2">${sel[k].has(v) ? '✓' : ''}</span><span class="grow">${esc(v)}</span><span class="s">${n}</span></button>`).join('') || '<div class="s">Nema podataka.</div>'}`).join('')}</div><div class="mbtn"><button class="btn" data-m="clear">Poništi</button><button class="btn pri" data-m="ok">Primeni</button></div></div>`;
    paint();
    m.addEventListener('click', e => {
      e.stopPropagation();
      const c = e.target.closest('.ck'); if (c) { const st = sel[c.dataset.k]; st.has(c.dataset.v) ? st.delete(c.dataset.v) : st.add(c.dataset.v); const sc = m.querySelector('.sbody').scrollTop; paint(); m.querySelector('.sbody').scrollTop = sc; return; }
      const b = e.target.closest('[data-m]');
      if (b && b.dataset.m === 'clear') { Object.values(sel).forEach(x => x.clear()); paint(); return; }
      if (b && b.dataset.m === 'ok') { const out = {}; Object.keys(sel).forEach(k => out[k] = [...sel[k]]); modalDone(m, () => res(out)); return; }
      if ((b && b.dataset.m === 'x') || e.target === m) { modalDone(m, () => res(null)); }
    });
    modalOpen(m, () => { m.remove(); res(null); });
  });
}
async function openFilter(k) {
  const r = k === 'more'
    ? await pickSheet('Još filtera', MOREF.map(x => [x, FDEF[x].t, fOptions(x), selOf(x)]))
    : await pickSheet(FDEF[k].t, [[k, FDEF[k].t, fOptions(k), selOf(k)]]);
  if (!r) return;
  Object.keys(r).forEach(x => { if (r[x].length) S.flt[x] = r[x]; else delete S.flt[x]; });
  saveFlt(); S.showAll = false; render();
}
// teren za nov objekat: po mestu, pa po najbližem kupcu sa lokacijom, pa po izabranom pravcu
function terenFor(mesto, pos) {
  const m = norm(mesto);
  if (m) {
    const c = {}; S.items.forEach(it => { const t = cv(it.f.AuditTeren); if (t && norm(it.f.field_4) === m) c[t] = (c[t] || 0) + 1; });
    const best = Object.entries(c).sort((a, b) => b[1] - a[1])[0]; if (best) return best[0];
    const t = TER.find(x => norm(x.slice(x.indexOf(' ') + 1)) === m); if (t) return t;
  }
  if (pos) { let bd = 15000, bt = null; S.items.forEach(it => { const l = parseLoc(it.f.Lokacija), t = cv(it.f.AuditTeren); if (l && t) { const d = dist(pos, l); if (d < bd) { bd = d; bt = t; } } }); if (bt) return bt; }
  const sel = selOf('pravac'); return sel.length === 1 && sel[0] !== 'Bez terena' ? sel[0] : null;
}
// lokacija i vreme iz slike (EXIF), ako ih telefon nije sklonio
async function exifInfo(file) {
  try {
    const b = new DataView(await file.slice(0, 262144).arrayBuffer());
    if (b.getUint16(0) !== 0xFFD8) return null;
    let o = 2;
    while (o < b.byteLength - 10) {
      const mk = b.getUint16(o), len = b.getUint16(o + 2);
      if (mk === 0xFFE1 && b.getUint32(o + 4) === 0x45786966) {
        const t = o + 10, le = b.getUint16(t) === 0x4949, u16 = p => b.getUint16(p, le), u32 = p => b.getUint32(p, le);
        const ifd = p => { const n = u16(p), out = {}; for (let i = 0; i < n; i++) { const e = p + 2 + i * 12; out[u16(e)] = { cnt: u32(e + 4), off: e + 8 }; } return out; };
        const asc = e => { const p = e.cnt > 4 ? t + u32(e.off) : e.off; let x = ''; for (let i = 0; i < e.cnt - 1; i++) x += String.fromCharCode(b.getUint8(p + i)); return x; };
        const rat = (e, i) => { const p = t + u32(e.off) + i * 8; return u32(p) / (u32(p + 4) || 1); };
        const i0 = ifd(t + u32(t + 4)), res = {};
        if (i0[0x8769]) { const ex = ifd(t + u32(i0[0x8769].off)); if (ex[0x9003]) res.time = asc(ex[0x9003]); }
        if (i0[0x8825]) {
          const gg = ifd(t + u32(i0[0x8825].off));
          if (gg[2] && gg[4]) { const dms = e => rat(e, 0) + rat(e, 1) / 60 + rat(e, 2) / 3600; let lat = dms(gg[2]), lon = dms(gg[4]); if (gg[1] && asc(gg[1]) === 'S') lat = -lat; if (gg[3] && asc(gg[3]) === 'W') lon = -lon; if (lat && lon) { res.lat = +lat.toFixed(6); res.lon = +lon.toFixed(6); } }
        }
        return res;
      }
      if ((mk & 0xFF00) !== 0xFF00) break;
      o += 2 + len;
    }
  } catch (e) { }
  return null;
}
async function fetchLog() {
  if (!S.logList) return;
  let url = '/sites/' + S.site + '/lists/' + S.logList + '/items?$expand=fields($select=Title,Vreme,KupacSifra,KupacNaziv,Korisnik,Udaljenost,Napomena)&$top=999';
  const out = [];
  while (url) { const j = await g(url); j.value.forEach(x => { const f = x.fields; out.push({ t: f.Title, v: f.Vreme, s: f.KupacSifra || '', n: f.KupacNaziv || '', k: f.Korisnik || '', u: f.Udaljenost, m: f.Napomena || '' }); }); url = j['@odata.nextLink']; }
  const pend = (await DB.all()).filter(o => o.type === 'log').map(o => { const f = o.fields; return { t: f.Title, v: f.Vreme, s: f.KupacSifra || '', n: f.KupacNaziv || '', k: f.Korisnik || '', u: f.Udaljenost, m: f.Napomena || '' }; });
  S.log = out.concat(pend).sort((a, b) => String(a.v).localeCompare(String(b.v)));
  saveLog(); buildLogIndex();
}
const phoneOf = it => (split(it.f.field_5)[0] || '').replace(/[^\d+]/g, '');
// potvrda „Da li ste sigurni?” (sopstveni prozor, radi i u aplikaciji na početnom ekranu)
function askSure(title, body, ok = 'Da, siguran sam') {
  return new Promise(res => {
    const m = document.createElement('div'); m.className = 'modal';
    m.innerHTML = `<div class="box"><h2>${esc(title)}</h2><div class="mb">${body}</div><div class="mbtn"><button class="btn" data-m="0">Odustani</button><button class="btn pri" data-m="1">${esc(ok)}</button></div></div>`;
    m.addEventListener('click', e => { const b = e.target.closest('[data-m]'); if (!b && e.target !== m) return; e.stopPropagation(); modalDone(m, () => res(!!b && b.dataset.m === '1')); });
    modalOpen(m, () => { m.remove(); res(false); });
  });
}
function askChoice(title, opts) {
  return new Promise(res => {
    const m = document.createElement('div'); m.className = 'modal';
    m.innerHTML = `<div class="box"><h2>${esc(title)}</h2><div class="opts">${opts.map(([k, t, d]) => `<button class="opt" data-m="${k}"><span><b>${esc(t)}</b><br><span class="s">${esc(d)}</span></span></button>`).join('')}</div><div class="mbtn"><button class="btn" data-m="">Odustani</button></div></div>`;
    m.addEventListener('click', e => { const b = e.target.closest('[data-m]'); if (!b && e.target !== m) return; e.stopPropagation(); modalDone(m, () => res(b ? b.dataset.m : '')); });
    modalOpen(m, () => { m.remove(); res(''); });
  });
}
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
  const ll = lists.value.find(x => x.displayName === (C.logList || 'Dnevnik rada'));
  S.logList = ll ? ll.id : null;
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
        if (S.visit && S.visit.id === op.tmp) S.visit.id = r.id;
        if (S.day && S.day.visits) S.day.visits.forEach(v => { if (v.id === op.tmp) v.id = r.id; });
        saveDay();
        saveItems();
      } else if (op.type === 'x') {
        if (!(await flushX(op))) continue;
      } else if (op.type === 'log') {
        if (!S.logList) throw new Error('Nema liste Dnevnik rada');
        const f = { ...op.fields }; if (String(f.KupacId || '').startsWith('tmp')) { const it = S.items.find(x => x.f.Title === f.KupacSifra); if (it && !String(it.id).startsWith('tmp')) f.KupacId = String(it.id); }
        await g('/sites/' + S.site + '/lists/' + S.logList + '/items', { method: 'POST', body: JSON.stringify({ fields: f }) });
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
  if (S.skipPop) { S.skipPop = false; const c = S.afterPop; S.afterPop = null; if (c) c(); return; }
  const md = [...document.querySelectorAll('.modal')].pop(); if (md && md._close) { md._close(); return; }
  const back = { kupci: 'tereni', kartica: S.from === 'kupci' && S.teren === null ? 'tereni' : (S.from || 'danas'), audit: 'kartica', kraj: 'vise', krajposete: 'kartica', novi: S.from || 'danas', beleska: S.noteNew ? (S.from || 'danas') : 'kartica', thread: S.from2 || 'poruke', mapa: 'danas', pin: 'kartica' }[S.view];
  if (back) go(back, false);
});

/* ---------- prikazi ---------- */
function bar(title, back) {
  return `<header class="bar">${back ? `<button class="icon" data-a="back" aria-label="Nazad">←</button>` : ''}<h1>${esc(title)}</h1><span id="sync" class="sync"></span>${!back ? `<button class="icon" data-a="refresh" aria-label="Osveži">⟳</button>` : ''}</header>`;
}
function saveDraft() { if (S.visit && S.d && S.view === 'audit' && S.cur && S.cur.id === S.visit.id) { S.visit.draft = { d: S.d, step: S.step }; saveDay(); } }
function render() {
  saveDraft(); if (S.visit) visitItem();
  let v = { narudzba: vNarudzba, tiketi: vTiketi, tiket: vTiket, poruke: vPoruke, thread: vThread, nedelja: vNedelja, mapa: vMapa, pin: vPin, aktivnost: vAktivnost, nadzor: vNadzor, admin: vAdmin, danas: vDanas, vise: vVise, beleska: vBeleska, tereni: vTereni, kupci: vKupci, kartica: vKartica, audit: vAudit, novi: vNovi, login: vLogin, kraj: vKraj, krajposete: vKrajPosete }[S.view] || vTereni;
  if (S.view !== 'login' && isKom()) { if (dayOpenOld() && !(S.visit && sameDay(S.visit.start))) v = vJuce; else if (!dayOn() && !(S.visit && sameDay(S.visit.start))) v = vDan; }
  $app.innerHTML = (DEMO ? '<div class="demo">Probni režim – izmišljeni kupci, ništa se ne šalje</div>' : '') + v();
  document.body.classList.toggle('tabs-on', !!document.querySelector('.tabs'));
  document.body.classList.toggle('desk', !isKom());
  if (v === vNedelja) wireWeek(); else if (v === vMapa) showMap(S.pday || ymd()); else if (v === vPin) wirePin(); else if (S.lmap && !document.getElementById('map')) { S.lmap.remove(); S.lmap = null; }
  paintSync();
  const sIn = document.getElementById('q'); if (sIn && S.focusQ) { sIn.focus(); sIn.setSelectionRange(sIn.value.length, sIn.value.length); S.focusQ = false; }
}
function vLogin() {
  return `<div class="login"><h1>Audit kupaca</h1><p class="s" style="margin-bottom:24px">Prijavi se službenim nalogom (@ekoline.rs).</p><button class="btn pri" data-a="login">Prijava</button></div>`;
}
function vDan() {
  const ended = S.day && S.day.date === new Date().toLocaleDateString('sr-RS') && S.day.end;
  return `<div class="login"><h1>${ended ? 'Radni dan je završen' : 'Dobar dan'}</h1>
    <p class="s">${ended ? `Počeo u ${hhmm(S.day.start)}, završio u ${hhmm(S.day.end)}. Ako nastavljaš rad, počni ponovo.` : 'Pre rada na terenu počni radni dan. Beleže se vreme i lokacija.'}</p>
    <p class="s" style="margin-bottom:24px">${esc(userName())}</p>
    <button class="btn pri" data-a="daystart" ${S.busy ? 'disabled' : ''}>${S.busy ? 'Tražim lokaciju…' : (ended ? 'Počni ponovo' : 'Počni radni dan')}</button>
    <div class="err" id="err"></div></div>`;
}
function vKraj() {
  const t = new Date().toLocaleDateString('sr-RS');
  const visits = (S.day && S.day.visits) || [];
  return bar('Kraj radnog dana', true) + `<main class="main">
    <div class="sum"><div class="kpi"><span class="s">Početak</span><b>${S.day ? hhmm(S.day.start) : '—'}</b></div><div class="kpi"><span class="s">Posete danas</span><b>${visits.length}</b></div></div>
    ${visits.map(v => `<div class="row" style="cursor:default"><div class="grow"><div class="t ell">${esc(v.n)}</div><div class="s">${hhmm(v.s)}–${v.e ? hhmm(v.e) : '…'}${v.e ? ' · ' + Math.max(1, Math.round((Date.parse(v.e) - Date.parse(v.s)) / 60000)) + ' min' : ''}</div></div><div class="right s">${esc(v.i || '')}</div></div>`).join('') || '<div class="empty">Danas nije bilo poseta.</div>'}
    ${krajPlanHtml()}
    <div class="err" id="err"></div>
  </main><div class="foot"><div class="in"><button class="btn pri" data-a="dayend" ${S.busy ? 'disabled' : ''}>${S.busy ? 'Završavam…' : 'Završi radni dan'}</button></div></div>`;
}
// za probu: kad GPS ne uspe, nudi se nastavak bez GPS-a (beleži se u dnevnik)
function gpsFail(msg, what, far) {
  S.gpsErr = msg;
  const el = document.getElementById('err'); if (!el) { toast(msg); return; }
  el.innerHTML = esc(msg) + (far ? `<button class="btn" style="width:100%;margin-top:12px" data-a="remap">Lokacija je pogrešna, mapiraj ponovo</button>` : '') + (C.allowNoGps ? `<button class="btn" style="width:100%;margin-top:12px" data-a="nogps" data-v="${what}">Nastavi bez GPS-a</button>` : '');
}
async function noGps(what) {
  const why = 'BEZ GPS-a: ' + (S.gpsErr || 'preskočeno');
  if (what === 'day') {
    await autoCloseOldVisit();
    S.day = { date: new Date().toLocaleDateString('sr-RS'), start: new Date().toISOString(), visits: [], noGps: true };
    saveDay(); await logEvent('POČETAK DANA', null, null, { Napomena: why });
    S.view = 'danas'; render(); toast('Radni dan je počeo bez GPS-a.');
  } else if (what === 'visit') {
    if (S.visit) return;
    beginVisit(S.cur); S.visit.noGps = true; saveDay();
    await logEvent('POČETAK POSETE', null, S.cur, { Napomena: why });
    render(); toast('Poseta je počela bez GPS-a.');
  } else if (what === 'novi') await saveNovi(true);
}
async function dayStart() {
  await autoCloseOldVisit();
  S.busy = true; render();
  try {
    const p = await getPos();
    S.day = { date: new Date().toLocaleDateString('sr-RS'), start: new Date().toISOString(), visits: [] };
    saveDay(); await logEvent('POČETAK DANA', p, null);
    S.busy = false; S.view = 'danas'; render(); toast('Radni dan je počeo u ' + hhmm(S.day.start) + '.');
  } catch (e) { S.busy = false; render(); gpsFail(e.message, 'day'); }
}
async function dayEnd() {
  if (S.visit) { toast('Prvo završi posetu u toku.'); return; }
  const chk = krajPlanCheck(); if (chk) { document.getElementById('err').textContent = chk; return; }
  S.busy = true; render();
  await krajPlanApply();
  let p = null; try { p = await getPos(); } catch (e) { }
  S.day.end = new Date().toISOString(); saveDay();
  await logEvent('KRAJ DANA', p, null, p ? {} : { Napomena: 'lokacija nije dostupna' });
  const mr = await sendReport(ymd());
  S.busy = false; S.view = 'danas'; render(); toast('Radni dan je završen.' + (mr ? ' Izveštaj: ' + mr : ' Izveštaj je poslat.'));
}
const sameDay = t => t && new Date(t).toLocaleDateString('sr-RS') === new Date().toLocaleDateString('sr-RS');
function visitItem() {
  if (!S.visit) return null;
  const it = S.items.find(x => x.id === S.visit.id) || S.items.find(x => S.visit.sifra && x.f.Title === S.visit.sifra) || S.items.find(x => (x.f.field_1 || x.f.Title) === S.visit.n) || null;
  if (it && it.id !== S.visit.id) { const old = S.visit.id; S.visit.id = it.id; S.visit.sifra = it.f.Title; if (S.day && S.day.visits) S.day.visits.forEach(v => { if (v.id === old) v.id = it.id; }); saveDay(); }
  return it;
}
async function autoCloseOldVisit() {
  if (!S.visit || sameDay(S.visit.start)) return;
  await logEvent('KRAJ POSETE', null, visitItem(), { Napomena: `zatvorena automatski (bila otvorena od ${fmtDate(S.visit.start)} ${hhmm(S.visit.start)})` });
  S.visit = null; saveDay();
}
function beginVisit(it, p) {
  try { markPlanVisited(it); } catch (e) { }
  S.visit = { id: it.id, sifra: it.f.Title, n: it.f.field_1 || it.f.Title, start: new Date().toISOString(), audit: false };
  S.day.visits = S.day.visits || []; S.day.visits.push({ id: it.id, n: S.visit.n, s: S.visit.start }); saveDay();
}
// mapiranje objekta na licu mesta: provera mesta, precizan GPS, potvrda
async function mapHere(it, title) {
  const town = it.f.field_4, tp = await townPos(town), here = await getPos();
  if (tp && dist(here, tp) > TOWNKM * 1000 + here.acc) throw new Error(`Kupac je u mestu ${town}, a ti si oko ${Math.round(dist(here, tp) / 1000)} km od tog mesta. Lokacija objekta se beleži samo ispred objekta.`);
  S.townUnchecked = tp === undefined;
  S.busyMsg = 'Stani ispred ulaza, napolju · tražim signal…'; render();
  const p = await getPrecise(a => { S.busyMsg = `Stani ispred ulaza, napolju · ±${a} m${a <= MAPACC ? ' ✓' : '…'}`; const b = document.getElementById('busy'); if (b) b.textContent = S.busyMsg; });
  const ok = await askSure(title, `<p>${esc(it.f.field_1 || it.f.Title)}<br><span class="s">${esc([it.f.field_3, town].filter(Boolean).join(', '))}</span></p><p>Lokacija objekta biće zabeležena na mestu gde stojiš (preciznost ±${p.acc} m). Po njoj se proveravaju sve sledeće posete.</p><p><b>Da li stojiš ispred ulaza ovog objekta?</b></p>`, 'Da, zabeleži');
  if (!ok) { const e = new Error('Lokacija nije zabeležena.'); e.cancel = true; throw e; }
  return p;
}
async function visitStart() {
  const it = S.cur;
  if (S.visit) { toast('Već je u toku poseta kod: ' + S.visit.n); return; }
  S.busy = true; render();
  try {
    let home = parseLoc(it.f.Lokacija), d = null, p, mapped = false;
    if (home) { const tp0 = await townPos(it.f.field_4); if (tp0 && dist(home, tp0) > TOWNKM * 1000) home = null; } // tačka daleko od mesta kupca = pogrešno mapirano, mapira se ponovo
    if (home) {
      p = await getPos();
      if (p.acc > MAXACC) throw new Error(`GPS je neprecizan (±${p.acc} m). Sačekaj malo ili izađi napolje, pa pokušaj ponovo.`);
      d = dist(p, home);
      if (d > RADIUS + p.acc + home.acc) { const e = new Error(`Nisi na lokaciji kupca: udaljen si oko ${d >= 1000 ? (d / 1000).toFixed(1) + ' km' : d + ' m'}. Poseta može da počne samo kod objekta.`); e.far = true; throw e; }
      if (home.acc > MAPACC && p.acc <= MAPACC) mapped = true; // stara, slabija tačka – zameni boljom
    } else {
      p = await mapHere(it, 'Zabeleži lokaciju objekta');
      mapped = true;
    }
    if (mapped) {
      it.f.Lokacija = fmtLoc(p); saveItems();
      await queue({ type: 'patch', itemId: it.id, fields: { Lokacija: it.f.Lokacija } });
    }
    S.busyMsg = '';
    beginVisit(it, p);
    await logEvent('POČETAK POSETE', p, it, Object.assign(d !== null ? { Udaljenost: d } : {}, mapped ? { Napomena: home ? 'lokacija objekta preciznije zabeležena' : 'prva poseta: lokacija objekta zabeležena' + (S.townUnchecked ? ' (mesto nije provereno – nema mreže)' : '') } : {})); S.townUnchecked = false;
    S.busy = false; render(); toast(home ? 'Poseta je počela.' : 'Poseta je počela. Lokacija objekta je zabeležena.');
  } catch (e) { S.busy = false; S.busyMsg = ''; render(); if (e.cancel) toast(e.message); else gpsFail(e.message, 'visit', e.far); }
}
// ispravka pogrešne lokacije: komercijalista sam, uz trag u dnevniku
async function remap() {
  const it = S.cur; if (S.visit) return;
  const old = parseLoc(it.f.Lokacija);
  if (!await askSure('Ispravka lokacije', `<p>Zabeležena lokacija objekta <b>${esc(it.f.field_1 || it.f.Title)}</b> biće zamenjena novom, na mestu gde stojiš.</p><p>Ispravka se beleži u dnevnik rada, sa starom i novom tačkom.</p>`, 'Da, ispravi')) return;
  S.busy = true; render();
  try {
    const p = await mapHere(it, 'Potvrdi novu lokaciju');
    const moved = old ? dist(p, old) : null;
    it.f.Lokacija = fmtLoc(p); saveItems();
    await queue({ type: 'patch', itemId: it.id, fields: { Lokacija: it.f.Lokacija } });
    await logEvent('ISPRAVKA LOKACIJE', p, it, { Udaljenost: moved, Napomena: `stara: ${old ? fmtLoc(old) : '—'}; nova: ${it.f.Lokacija}${moved !== null ? '; pomereno ' + moved + ' m' : ''}${S.townUnchecked ? ' (mesto nije provereno – nema mreže)' : ''}` });
    S.busyMsg = ''; beginVisit(it, p);
    await logEvent('POČETAK POSETE', p, it, { Udaljenost: 0, Napomena: 'posle ispravke lokacije' });
    S.busy = false; render(); toast('Lokacija je ispravljena. Poseta je počela.');
  } catch (e) { S.busy = false; S.busyMsg = ''; render(); if (e.cancel) toast(e.message); else gpsFail(e.message, 'visit'); }
}
const VISH = ['obavljen razgovor', 'urađen audit', 'primljena porudžbina', 'kupac nije bio tu', 'objekat zatvoren', 'ostalo'];
function vKrajPosete() {
  const it = S.items.find(x => x.id === S.visit.id) || S.cur, k = S.vk || (S.vk = { i: S.visit.audit ? 'urađen audit' : '', b: '' });
  return bar('Kraj posete', true) + `<main class="main">
    <div class="t">${esc(S.visit.n)}</div><div class="s">Poseta od ${hhmm(S.visit.start)} · ${Math.max(1, Math.round((Date.now() - Date.parse(S.visit.start)) / 60000))} min${S.visit.audit ? ' · audit urađen' : S.visit.draft ? ' · audit nije završen – nedovršeni podaci se ne čuvaju' : ' · bez audita'}</div>
    <div class="t" style="margin:16px 2px 8px">Ishod posete</div>
    <div class="opts">${VISH.map(x => `<button class="opt ${k.i === x ? 'on' : ''}" data-a="vish" data-v="${esc(x)}">${esc(x)}</button>`).join('')}</div>
    <div class="fld" style="margin-top:14px"><label>Beleška o poseti</label><textarea id="vkb" placeholder="Šta je dogovoreno, šta treba uraditi…">${esc(k.b)}</textarea></div>
    <div class="err" id="err"></div>
  </main><div class="foot"><div class="in"><button class="btn pri" data-a="endvisitsave" ${S.busy ? 'disabled' : ''}>${S.busy ? 'Tražim lokaciju…' : 'Završi posetu'}</button></div></div>`;
}
async function visitEnd() {
  const k = S.vk || {}; k.b = (document.getElementById('vkb') || {}).value || k.b || '';
  if (!k.i) { document.getElementById('err').textContent = 'Izaberi ishod posete.'; return; }
  const it = visitItem();
  S.busy = true; render();
  let p = null; try { p = await getPos(); } catch (e) { }
  const home = it && parseLoc(it.f.Lokacija), extra = { Napomena: k.i + (k.b.trim() ? ': ' + k.b.trim() : '') + (S.visit.audit ? '' : S.visit.draft ? ' (audit započet, nije završen)' : ' (bez audita)') + (p ? '' : ' [lokacija nije dostupna]') };
  if (p && home) extra.Udaljenost = dist(p, home);
  await logEvent('KRAJ POSETE', p, it, extra);
  if (k.b.trim() && it) {
    const note = `[${new Date().toLocaleString('sr-RS')} · poseta: ${k.i}] ${k.b.trim()}`;
    it.f.Beleske = it.f.Beleske ? note + '\n' + it.f.Beleske : note; saveItems();
    await queue({ type: 'patch', itemId: it.id, fields: { Beleske: it.f.Beleske } });
  }
  if (it && it.f.Obici) { it.f.Obici = false; saveItems(); await queue({ type: 'patch', itemId: it.id, fields: { Obici: false } }); }
  const v = (S.day.visits || []).slice().reverse().find(x => x.id === S.visit.id && !x.e); if (v) { v.e = new Date().toISOString(); v.i = k.i; }
  S.visit = null; S.vk = null; S.busy = false; saveDay();
  toast('Poseta je završena.'); history.back();
}
function tabs(on) {
  const n = unreadCount(), P = modOn('por') ? [['poruke', '💬', 'Poruke', n]] : [], TK = modOn('tik') ? [['tiketi', '🎫', 'Tiketi', tikNew()]] : [];
  const T = isKom() ? [['danas', '☀', 'Danas'], ['tereni', '☰', 'Kupci'], ...P, ['vise', '⋯', 'Više']]
    : isAdmin() ? [['aktivnost', '⚡', 'Aktivnost'], ['nadzor', '📊', 'Nadzor'], ['nedelja', '📅', 'Plan'], ...P, ...TK, ['tereni', '☰', 'Kupci'], ['admin', '⚙', 'Admin'], ['vise', '⋯', 'Više']]
    : [['aktivnost', '⚡', 'Aktivnost'], ['nedelja', '📅', 'Plan'], ...P, ...TK, ['tereni', '☰', 'Kupci'], ['vise', '⋯', 'Više']];
  return `<nav class="tabs">${T.map(([k, i, nm, c]) => `<button class="${on === k ? 'on' : ''}" data-a="tab" data-v="${k}"><span>${i}</span>${nm}${c !== undefined ? `<b class="tb">${c || ''}</b>` : ''}</button>`).join('')}</nav>`;
}
function dayStats() {
  const vs = (S.day && S.day.visits) || [], done = vs.filter(v => v.e);
  const avg = done.length ? Math.max(1, Math.round(done.reduce((a, v) => a + (Date.parse(v.e) - Date.parse(v.s)) / 60000, 0) / done.length)) : 0;
  return { n: vs.length, avg, a: (S.day && S.day.audits) || 0, b: (S.day && S.day.notes) || 0 };
}
function danasData() {
  const q = norm(S.dq), anyF = Object.keys(S.flt).some(k => selOf(k).length);
  if (!q && !anyF) return null;
  let base = S.items.filter(it => q ? [it.f.Title, it.f.field_1, it.f.field_2, it.f.field_3, it.f.field_4].some(x => norm(x).includes(q)) && fMatch(it, 'pravac') : fMatch(it));
  const ob = base.filter(it => it.f.Obici && cv(it.f.StatusNovi) !== 'zatvoren').sort((x, y) => String(y.f.Title).localeCompare(String(x.f.Title)));
  const kl = it => 'ABCDE'.indexOf(klasaOf(it)) < 0 ? 9 : 'ABCDE'.indexOf(klasaOf(it));
  const prom = it => num(it.f.field_19) + num(it.f.field_18) / 2;
  let rows = base.filter(it => !it.f.Obici).map(it => ({ it, d: dueOf(it) }));
  if (S.onlyDue && !q) rows = rows.filter(r => r.d && (r.d.top || r.d.over >= -7));
  rows.sort((x, y) => { const a = x.d, b = y.d; if (!a || !b) return (a ? -1 : b ? 1 : 0) || prom(y.it) - prom(x.it); return (a.top === b.top ? 0 : a.top ? -1 : 1) || (a.top ? 0 : b.over - a.over) || kl(x.it) - kl(y.it) || prom(y.it) - prom(x.it); });
  return { ob, rows, q };
}
function danasList() {
  const D = danasData();
  if (!D) return '<div class="empty">Izaberi pravac (jedan ili više terena) ili upiši naziv kupca u pretragu.</div>';
  const lim = S.showAll ? D.rows.length : 40;
  return (D.ob.length ? `<button class="sec" data-a="obtoggle">${S.obOpen ? '▾' : '▸'} Za obići · ${D.ob.length}</button>${S.obOpen ? D.ob.map(it => rowPlan({ it, d: dueOf(it) })).join('') : ''}` : '')
    + `<div class="sech"><span>${D.q ? 'Rezultati pretrage' : S.onlyDue ? 'Na redu' : 'Svi kupci'} · ${D.rows.length}</span>${D.q ? '' : `<button class="lnk" data-a="onlydue">${S.onlyDue ? 'prikaži sve' : 'samo na redu'}</button>`}</div>`
    + (D.rows.slice(0, lim).map(rowPlan).join('') || `<div class="empty">${D.q ? 'Nema kupaca za tu pretragu.' : 'Niko nije na redu za izabrani pravac i filtere.'}</div>`)
    + (D.rows.length > lim ? `<button class="btn" style="width:100%" data-a="showall">Prikaži sve (${D.rows.length})</button>` : '');
}
function rowPlan({ it, d }) {
  const f = it.f, k = klasaOf(it);
  const when = !d ? 'plan: ' + esc(planOf(it).label) : d.ob ? (d.lv ? 'obići · poslednja poseta ' + fmtDate(d.lv) : 'obići · još nije obiđen') : d.never ? 'nikad obiđen' : `poslednja poseta ${fmtDate(d.lv)} · ` + (d.over > 0 ? `<b class="late">kasni ${d.over} d</b>` : d.over === 0 ? '<b class="late">danas na redu</b>' : `na redu za ${-d.over} d`);
  const badges = finFlag(it) + (cv(f.StatusNovi) === 'potencijalni' ? '<span class="badge brand">potencijalni</span> ' : '') + (S.visit && S.visit.id === it.id ? '<span class="badge ok">poseta u toku</span>' : '');
  const sm = S.selMode, on = sm && S.sel && S.sel.has(String(it.id)), inP = sm && planOfDay(S.pday || ymd()).some(p => p.f.Title === String(f.Title));
  return `<div class="row ${on ? 'sel' : ''}" data-a="${sm ? 'selrow' : 'kupac'}" data-v="${esc(it.id)}">${sm ? `<span class="ckb">${on ? '✓' : inP ? '•' : ''}</span>` : ''}<div class="grow"><div class="t ell">${esc(f.field_1 || f.Title)}${inP ? ' <span class="badge">u planu</span>' : ''}</div><div class="s ell">${esc([f.field_3, f.field_4].filter(Boolean).join(', ') || '—')}</div><div class="s">${when}</div>${badges ? `<div>${badges}</div>` : ''}</div><div class="right">${k ? `<div class="kl">${k}</div>` : ''}<div class="s">${money(f.field_19)}</div></div></div>`;
}
function fBtn(k) {
  const d = FDEF[k], sel = selOf(k);
  const txt = sel.length ? sel.slice(0, 2).join(', ') + (sel.length > 2 ? ` +${sel.length - 2}` : '') : d.l;
  return `<button class="fsel ${sel.length ? 'on' : ''}" data-a="fsel" data-v="${k}"><span class="ic">${d.ic}</span><span class="fl ell">${esc(txt)}</span>${sel.length ? `<span class="fx" data-a="fclr" data-v="${k}">✕</span>` : '<span class="cv">▾</span>'}</button>`;
}
function vDanas() {
  const st = dayStats(), more = MOREF.filter(k => selOf(k).length).length, t = ymd(), pd = S.pday || t, today = pd === t;
  const filt = `<div class="fsel srch"><span class="ic">🔍</span><input type="search" data-i="dq" placeholder="Pretraži kupce po nazivu, šifri, adresi" value="${esc(S.dq)}"></div>
    ${MAINF.map(fBtn).join('')}
    <button class="fsel ${more ? 'on' : ''}" data-a="fsel" data-v="more"><span class="ic">⚙</span><span class="fl">Još filtera${more ? ` (${more})` : ''}</span><span class="cv">›</span></button>
    ${Object.keys(S.flt).length ? '<button class="lnk" style="margin:2px 2px 10px" data-a="fclrall">Poništi sve filtere</button>' : ''}`;
  return bar('Danas') + `<main class="main">
    ${visitBanner()}
    ${reminderDue() ? '<div class="row alarm" data-a="kraj"><div class="grow"><div class="t">⚠ Završi radni dan</div><div class="s">Pre odlaska zatvori dan i neobiđene kupce iz plana.</div></div><div class="right">›</div></div>' : ''}
    <div class="sum4"><div class="kpi"><span class="s">Posete</span><b>${st.n}</b></div><div class="kpi"><span class="s">Prosečno</span><b>${st.avg ? st.avg + ' min' : '—'}</b></div><div class="kpi"><span class="s">Auditi</span><b>${st.a}</b></div><div class="kpi"><span class="s">Beleške</span><b>${st.b}</b></div></div>
    <div class="s" style="margin:-4px 2px 10px">${S.day ? 'Radni dan od ' + hhmm(S.day.start) + (S.day.noGps ? ' (bez GPS-a)' : '') : ''}</div>
    ${modOn('plan') ? `<div class="seg daysel"><button class="${today ? 'on' : ''}" data-a="pday" data-v="${t}">Danas</button><button class="${pd === addDays(t, 1) ? 'on' : ''}" data-a="pday" data-v="${addDays(t, 1)}">Sutra</button><button class="${!today && pd !== addDays(t, 1) ? 'on' : ''}" data-a="pday" data-v="cal">📅 ${!today && pd !== addDays(t, 1) ? dLabel(pd) : 'Datum'}</button></div>` : ''}
    ${planBlock(pd)}
    ${today ? tasksBlock() : ''}
    <div class="sech" style="margin-top:6px"><span>${today ? 'Kupci' : 'Dodaj kupce u plan za ' + dLabel(pd)}</span></div>
    ${filt}
    <div id="dlist">${danasList()}</div>
  </main>${S.selMode && S.sel && S.sel.size ? `<div class="selbar"><span>Izabrano: ${S.sel.size}</span><button class="btn pri" data-a="seladd">📅 Dodaj u plan${today ? '' : ' za ' + dLabel(pd)}</button></div>` : '<button class="fab" data-a="novi">＋ Nov kupac</button>'}` + tabs('danas');
}
function vVise() {
  return bar('Više') + `<main class="main">
    <div class="card"><dl class="kv"><dt>Prijavljen</dt><dd>${esc(userName() || '—')}</dd><dt>Kupaca u telefonu</dt><dd>${S.items.length}</dd><dt>Čeka na slanje</dt><dd>${S.pending}${S.pending && !navigator.onLine ? ' (nema mreže)' : ''}</dd><dt>Radni dan</dt><dd>${S.day ? 'od ' + hhmm(S.day.start) + (S.day.end ? ' do ' + hhmm(S.day.end) : '') : '—'}</dd></dl></div>
    <button class="btn" style="width:100%;margin-top:12px" data-a="refresh">⟳ Osveži podatke</button>
    ${isKom() ? '<button class="btn" style="width:100%;margin-top:10px" data-a="kraj">Kraj radnog dana</button>' : ''}
    ${isKom() && modOn('tik') ? `<button class="btn" style="width:100%;margin-top:10px" data-a="mojitik">🎫 Moji tiketi za BizniSoft · ${S.X.tik.filter(t => (t.f.Autor || '').toLowerCase() === me() && t.f.Status !== 'resen').length} otvorenih</button>` : ''}
    ${modOn('ods') && isKom() ? `<button class="btn" style="width:100%;margin-top:10px" data-a="newods">🏖 Prijavi odsustvo</button>${S.X.ods.filter(o => (o.f.Korisnik || '').toLowerCase() === me() && o.f.Do >= ymd()).map(o => `<div class="s" style="margin:6px 2px">${esc(o.f.Vrsta)} ${fmtDate(o.f.Od)}–${fmtDate(o.f.Do)} · ${esc(o.f.Status)}</div>`).join('')}` : ''}
    <div class="s" style="margin:12px 2px">Uloga: ${esc(ROLE_N[role()] || role())}</div>
    <button class="btn" style="width:100%;margin-top:10px" data-a="logout">Odjava</button>
    <p class="s" style="text-align:center;margin-top:18px">Audit kupaca · verzija 19.11.2026 · narudžbine</p>
  </main>` + tabs('vise');
}
function vBeleska() {
  const nw = S.noteNew, it = nw ? null : S.cur, k = S.nk;
  const fld = (id, l, t = 'text', ph = '') => `<div class="fld"><label>${l}</label><input data-i="${id}" type="${t}" placeholder="${esc(ph)}" value="${esc(k[id] || '')}"></div>`;
  return bar(nw ? 'Brza beleška · nov objekat' : 'Brza beleška', true) + `<main class="main">
    ${it ? `<div class="t">${esc(it.f.field_1 || it.f.Title)}</div><div class="s" style="margin-bottom:12px">${esc([it.f.field_3, it.f.field_4].filter(Boolean).join(', '))}</div>` : '<p class="s">Objekat koji si video u prolazu. Upisuje se kao potencijalni kupac, a poseta i lična karta se rade kasnije.</p>'}
    ${nw ? fld('nk_naziv', 'Naziv objekta', 'text', 'ako ne znaš, ostavi prazno') + fld('nk_adresa', 'Adresa') + fld('nk_mesto', 'Mesto') + fld('nk_tel', 'Telefon', 'tel') : ''}
    <div class="fld"><label>Napomena</label><textarea data-i="nk_text" placeholder="Šta si video, zašto vredi doći, kada je otvoreno…">${esc(k.nk_text || '')}</textarea></div>
    <div class="pbtns"><label class="big-act"><span>📷</span>Slikaj<input type="file" accept="image/*" capture="environment" data-i="nphoto" hidden></label><label class="big-act"><span>🖼</span>Iz galerije<input type="file" accept="image/*" multiple data-i="nphoto" hidden></label></div>
    <div class="thumbs">${k.photos.map(p => `<img src="${p.url}" alt="">`).join('')}</div>
    ${k.photoLoc ? `<p class="s ok">📍 Lokacija pročitana sa slike${k.photoLoc.time ? ' (' + esc(k.photoLoc.time.slice(0, 16).replace(/:/g, (m, i) => i < 10 ? '.' : ':')) + ')' : ''}: koristi se za objekat.</p>` : ''}
    <button class="opt ${k.obici ? 'on' : ''}" style="width:100%" data-a="nkobici">${k.obici ? '☑' : '☐'} Obići sledeći put (stavlja ga u plan dana)</button>
    <p class="s">${k.photoLoc ? 'Tačna lokacija objekta se beleži pri prvoj poseti.' : 'Lokacija se beleži okvirno, u trenutku čuvanja (ili sa slike iz galerije, ako je slika ima). Tačna lokacija objekta se beleži pri prvoj poseti.'}</p>
    <div class="err" id="err"></div>
  </main><div class="foot"><div class="in"><button class="btn pri" data-a="notesave" ${S.busy ? 'disabled' : ''}>${S.busy ? 'Čuvam…' : 'Sačuvaj belešku'}</button></div></div>`;
}
function openNote(nw) {
  S.noteNew = nw; S.from = nw ? S.view : S.from;
  S.nk = { obici: nw ? true : !!(S.cur && S.cur.f.Obici), photos: [] };
  go('beleska');
}
async function saveNote() {
  const nw = S.noteNew, k = S.nk, text = String(k.nk_text || '').trim(), er = document.getElementById('err');
  if (nw && !k.nk_naziv && !text && !k.photos.length) { er.textContent = 'Upiši naziv, napomenu ili dodaj sliku.'; return; }
  if (!nw && !text && !k.photos.length && k.obici === !!S.cur.f.Obici) { er.textContent = 'Upiši napomenu ili dodaj sliku.'; return; }
  S.busy = true; render();
  let p = null; try { p = await getPos(); } catch (e) { }
  const now = new Date(), stamp = now.toLocaleString('sr-RS'), line = `[${stamp} · beleška] ${text}`.trim();
  let it;
  if (nw) {
    const code = 'NOVO-' + now.toISOString().slice(2, 16).replace(/\D/g, '');
    const fields = { Title: code, field_1: k.nk_naziv || 'Nepoznat objekat', field_3: k.nk_adresa || '', field_4: k.nk_mesto || '', field_5: k.nk_tel || '', StatusNovi: 'potencijalni', Obici: !!k.obici, Beleske: line };
    const loc = k.photoLoc || (p && p.acc <= MAXACC ? p : null);
    if (loc) fields.Lokacija = fmtLoc(loc);
    const ter = terenFor(fields.field_4, loc); if (ter) fields.AuditTeren = ter;
    it = { id: 'tmp' + now.getTime(), f: { ...fields } };
    S.items.push(it); saveItems();
    await queue({ type: 'create', tmp: it.id, fields });
  } else {
    it = S.cur; const fields = {};
    if (text) { it.f.Beleske = it.f.Beleske ? line + '\n' + it.f.Beleske : line; fields.Beleske = it.f.Beleske; }
    if (k.obici !== !!it.f.Obici) { it.f.Obici = !!k.obici; fields.Obici = !!k.obici; }
    if (k.photoLoc && !parseLoc(it.f.Lokacija)) { it.f.Lokacija = fmtLoc(k.photoLoc); fields.Lokacija = it.f.Lokacija; }
    saveItems();
    if (Object.keys(fields).length) await queue({ type: 'patch', itemId: it.id, fields });
  }
  const ps = now.toISOString().slice(0, 16).replace(/\D/g, '');
  for (let i = 0; i < k.photos.length; i++) await queue({ type: 'photo', name: `${it.f.Title}-${ps}-b${i + 1}.jpg`, blob: k.photos[i].blob });
  k.photos.forEach(x => URL.revokeObjectURL(x.url));
  await logEvent('BELEŠKA', k.photoLoc || p, it, { Napomena: (nw ? 'nov objekat: ' : '') + (text || '(bez teksta)') + (k.photos.length ? ` · slika: ${k.photos.length}` : '') + (k.obici ? ' · obići' : '') + (k.photoLoc ? ' · lokacija sa slike' : p ? '' : ' [bez lokacije]') });
  if (S.day) { S.day.notes = (S.day.notes || 0) + 1; saveDay(); }
  S.busy = false; S.nk = null; S.cur = it; S.ktab = 'posete';
  S.view = 'kartica'; history.replaceState({ v: 'kartica' }, ''); render(); scrollTo(0, 0);
  toast('Beleška je sačuvana.');
}
function visitHistory(it) {
  const code = String(it.f.Title), ev = S.log.filter(e => e.s === code).sort((a, b) => String(a.v).localeCompare(String(b.v)));
  const out = []; const open = {};
  ev.forEach(e => {
    if (e.t === 'POČETAK POSETE') { const r = { kind: 'visit', s: e.v, k: e.k, m: e.m && /BEZ GPS/.test(e.m) ? 'bez GPS-a' : '' }; out.push(r); open[e.k] = r; }
    else if (e.t === 'KRAJ POSETE') { const r = open[e.k]; if (r) { r.e = e.v; r.res = e.m; delete open[e.k]; } else out.push({ kind: 'visit', e: e.v, k: e.k, res: e.m }); }
    else if (e.t === 'AUDIT') { const r = open[e.k]; if (r) r.audit = e.m; else out.push({ kind: 'other', t: 'Audit', v: e.v, k: e.k, m: e.m }); }
    else out.push({ kind: 'other', t: e.t === 'BELEŠKA' ? 'Beleška' : e.t === 'ISPRAVKA LOKACIJE' ? 'Ispravka lokacije' : e.t, v: e.v, k: e.k, m: e.m });
  });
  return out.reverse();
}
function vPosete(it) {
  const h = visitHistory(it), who = k => esc(String(k || '').split('@')[0]);
  const old = it.f.DatumObilaska && !h.some(r => r.kind === 'visit') ? `<div class="hrow"><div class="t">Obilazak ${fmtDate(it.f.DatumObilaska)}</div><div class="s">${esc(cv(it.f.AuditIshod) || '')}</div></div>` : '';
  if (!h.length && !old) return '<div class="empty">Nema zabeleženih poseta ni beleški.</div>';
  return h.map(r => r.kind === 'visit'
    ? `<div class="hrow"><div class="t">Poseta ${fmtDate(r.s || r.e)} · ${r.s ? hhmm(r.s) : '?'}–${r.e ? hhmm(r.e) : 'u toku'}${r.s && r.e ? ' · ' + Math.max(1, Math.round((Date.parse(r.e) - Date.parse(r.s)) / 60000)) + ' min' : ''}</div><div class="s">${who(r.k)}${r.audit ? ' · audit: ' + esc(r.audit) : ''}${r.m ? ' · ' + r.m : ''}</div>${r.res ? `<div class="s2">${esc(r.res)}</div>` : ''}</div>`
    : `<div class="hrow"><div class="t">${esc(r.t)} · ${fmtDate(r.v)} ${r.v ? hhmm(r.v) : ''}</div><div class="s">${who(r.k)}</div>${r.m ? `<div class="s2">${esc(r.m)}</div>` : ''}</div>`).join('') + old;
}
function stats(list) { const tot = list.length, done = list.filter(x => x.f.Obradjeno).length; return { tot, done, pct: tot ? Math.round(done / tot * 100) : 0, prom: list.reduce((a, x) => a + num(x.f.field_19), 0) }; }
function vTereni() {
  const groups = {};
  S.items.forEach(it => { (groups[terenOf(it)] = groups[terenOf(it)] || []).push(it); });
  const keys = Object.keys(groups).sort((a, b) => a === 'Bez terena' ? -1 : b === 'Bez terena' ? 1 : stats(groups[b]).prom - stats(groups[a]).prom);
  const all = stats(S.items);
  const td = S.items.filter(x => (x.f.DatumObilaska || '').slice(0, 10) === today()).length;
  return bar('Kupci') + `<main class="main">
    ${visitBanner()}
    <div class="sum"><div class="kpi"><span class="s">Obrađeno</span><b>${all.done} / ${all.tot}</b></div><div class="kpi"><span class="s">Danas</span><b>${td}</b></div></div>
    <div class="row" data-a="teren" data-v="*"><div class="grow"><div class="t">Svi kupci</div><div class="s">${all.tot} kupaca · promet 2026: ${money(all.prom)}</div><div class="prog"><i style="width:${all.pct}%"></i></div></div><div class="right s">${all.pct}%</div></div>
    ${keys.map(k => { const s = stats(groups[k]); return `<div class="row" data-a="teren" data-v="${esc(k)}"><div class="grow"><div class="t">${esc(k)}${k === 'Bez terena' ? ' <span class="badge warn">dodeliti</span>' : ''}</div><div class="s">${s.tot} kupaca · obrađeno ${s.done} · ${money(s.prom)}</div><div class="prog"><i style="width:${s.pct}%"></i></div></div><div class="right s">${s.done}/${s.tot}</div></div>`; }).join('')}
  </main><button class="fab" data-a="novi">＋ Nov kupac</button>` + tabs('tereni');
}
function inTeren() { return S.items.filter(it => S.teren === '*' || terenOf(it) === S.teren); }
function visitBanner() {
  return S.visit ? `<div class="row" data-a="govisit" style="border-color:var(--brand);background:var(--brand-bg,transparent)"><div class="grow"><div class="t">Poseta u toku: ${esc(S.visit.n)}</div><div class="s">od ${hhmm(S.visit.start)}${S.visit.draft ? ' · audit nedovršen, korak ' + (S.visit.draft.step + 1) + '/' + STEPS.length : S.visit.audit ? ' · audit urađen' : ''}</div></div><div class="right"><b>Nastavi ›</b></div></div>` : '';
}
function vKupci() {
  const { html, cnt } = kupciList();
  const title = S.teren === '*' ? 'Svi kupci' : S.teren;
  return bar(title, true) + `<main class="main">
    ${visitBanner()}
    <input id="q" class="search" type="search" placeholder="Naziv, šifra, PIB, adresa…" value="${esc(S.q)}" data-i="q">
    <div class="seg">${[['open', 'Neobrađeni'], ['pot', 'Potencijalni / obići'], ['back', 'Vratiti se'], ['done', 'Obrađeni'], ['all', 'Svi']].map(([k, n]) => `<button class="${S.filter === k ? 'on' : ''}" data-a="filter" data-v="${k}">${n} ${cnt[k]}</button>`).join('')}</div>
    <div class="sorts">Redosled: ${[['promet', 'promet'], ['mesto', 'mesto i adresa'], ['naziv', 'naziv']].map(([k, n]) => `<button class="${S.sort === k ? 'on' : ''}" data-a="sort" data-v="${k}">${n}</button>`).join('')}</div>
    <div id="list">${html}</div>
  </main><button class="fab" data-a="novi">＋ Nov kupac</button>` + tabs('tereni');
}
function kupciList() {
  const base = inTeren();
  const cnt = { open: 0, back: 0, done: 0, pot: 0, all: base.length };
  const isPot = it => it.f.Obici || cv(it.f.StatusNovi) === 'potencijalni';
  base.forEach(it => { const i = cv(it.f.AuditIshod); if (isPot(it)) cnt.pot++; if (i === 'vratiti se') cnt.back++; else if (it.f.Obradjeno) cnt.done++; else cnt.open++; });
  const q = S.q.trim().toLowerCase();
  let L = base.filter(it => {
    const i = cv(it.f.AuditIshod);
    if (S.filter === 'open' && (it.f.Obradjeno || i === 'vratiti se')) return false;
    if (S.filter === 'back' && i !== 'vratiti se') return false;
    if (S.filter === 'done' && !it.f.Obradjeno) return false;
    if (S.filter === 'pot' && !isPot(it)) return false;
    if (!q) return true;
    return [it.f.Title, it.f.field_1, it.f.field_2, it.f.field_3, it.f.field_4].some(x => String(x || '').toLowerCase().includes(q));
  });
  const sorters = { promet: (a, b) => num(b.f.field_19) + num(b.f.field_18) / 2 - num(a.f.field_19) - num(a.f.field_18) / 2, naziv: (a, b) => String(a.f.field_1).localeCompare(String(b.f.field_1), 'sr'), mesto: (a, b) => String(a.f.field_4).localeCompare(String(b.f.field_4), 'sr') || String(a.f.field_3).localeCompare(String(b.f.field_3), 'sr') };
  L.sort(S.filter === 'pot' ? (a, b) => String(b.f.Title).localeCompare(String(a.f.Title)) : sorters[S.sort]);
  return { cnt, html: L.length ? L.slice(0, 300).map(rowKupac).join('') + (L.length > 300 ? `<div class="empty">Prikazano 300 od ${L.length}. Suzi pretragu.</div>` : '') : '<div class="empty">Nema kupaca za ovaj izbor.</div>' };
}
function rowKupac(it) {
  const f = it.f, i = cv(f.AuditIshod), d = daysAgo(f.field_20);
  const badge = i === 'vratiti se' ? '<span class="badge warn">vratiti se</span>' : f.Obradjeno ? '<span class="badge ok">obrađen</span>' : String(f.Title).startsWith('NOVO') ? '<span class="badge brand">nov</span>' : ''
  const badge2 = badge + ' ' + finFlag(it);
  return `<div class="row" data-a="kupac" data-v="${esc(it.id)}"><div class="grow"><div class="t ell">${esc(f.field_1 || f.Title)}</div><div class="s ell">${esc([f.field_3, f.field_4].filter(Boolean).join(', ') || '—')}</div><div class="s">${esc(cv(f.AuditSegment) || f.field_14 || 'bez segmenta')}</div></div><div class="right"><div class="t">${money(f.field_19)}</div><div class="s">${d === null ? '' : d === 0 ? 'danas' : 'pre ' + d + ' d'}</div>${badge2}</div></div>`;
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
    ${pjBlock(it)}
    <div class="kact">
      ${modOn('nar') && modOn('tik') ? `<button data-a="kord"><span>🛒</span>Narudžbina${ordLoad()[it.id] && Object.keys(ordLoad()[it.id].q || {}).length ? ' <b class="tb">●</b>' : ''}</button>` : ''}
      ${modOn('plan') ? `<button data-a="kplan"><span>📅</span>U plan</button>` : ''}
      ${modOn('zad') ? `<button data-a="ktask"><span>✅</span>Zadatak</button>` : ''}
      ${modOn('por') ? `<button data-a="kmsg"><span>💬</span>Poruka${(() => { const n = S.X.por.filter(m => m.f.KupacSifra === String(f.Title) && unread(m)).length; return n ? ` <b class="tb">${n}</b>` : ''; })()}</button>` : ''}
      ${modOn('konk') ? `<button data-a="kkonk"><span>🏁</span>Konkurencija</button>` : ''}
      ${modOn('mapa') ? `<button data-a="kpin"><span>📍</span>Na mapi</button>` : ''}
    </div>
    ${(() => { const P = S.X.plan.filter(p => p.f.Title === String(f.Title) && p.f.Status === 'planiran' && p.f.Datum >= ymd()).sort((a, b) => a.f.Datum.localeCompare(b.f.Datum)); return P.length ? `<div class="s" style="margin:6px 2px">📅 U planu: ${P.map(p => dLabel(p.f.Datum) + (p.f.Korisnik !== me() ? ' (' + esc(nameOf(p.f.Korisnik)) + ')' : '')).join(', ')}</div>` : ''; })()}
    ${modOn('pot') ? `<div class="potrow"><span class="s">Potencijal</span>${[1, 2, 3, 4, 5].map(v => `<button class="${num(f.PotencijalOcena) === v ? 'on' : ''}" data-a="pot" data-v="${v}">${v}</button>`).join('')}</div>` : ''}
    ${modOn('zad') ? (() => { const Z = custTasks(f.Title).sort((a, b) => (a.f.Status === 'urađen') - (b.f.Status === 'urađen') || num(a.f.Prioritet) - num(b.f.Prioritet)).slice(0, 8); return Z.length ? `<div class="sech"><span>Zadaci kod ovog kupca</span></div>${Z.map(z => taskHtml(z)).join('')}` : ''; })() : ''}
    ${modOn('konk') ? (() => { const K = S.X.konk.filter(k => k.f.KupacSifra === String(f.Title)).sort((a, b) => String(b.f.Vreme).localeCompare(String(a.f.Vreme)))[0]; return K ? `<div class="card" style="margin-top:8px"><b>Konkurencija</b> (${fmtDate(K.f.Vreme)}): ${esc(K.f.Konkurenti)}${K.f.Grupe ? ' · ' + esc(K.f.Grupe) : ''}${K.f.Cene ? ' · cene: ' + esc(K.f.Cene) : ''}${K.f.Uslovi ? ' · uslovi: ' + esc(K.f.Uslovi) : ''}${K.f.Napomena ? '<div class="s">' + esc(K.f.Napomena) + '</div>' : ''}</div>` : ''; })() : ''}
    ${(() => { const d = dueOf(it), p = planOf(it); return `<div class="card plan"><b>Plan obilaska:</b> ${esc(p.label)}${d && !d.top ? ` · ${d.over > 0 ? '<b class="late">kasni ' + d.over + ' d</b>' : d.over === 0 ? '<b class="late">danas na redu</b>' : 'na redu za ' + (-d.over) + ' d'}` : d && d.never ? ' · nikad obiđen' : ''}${it.f.Obici ? ' · <span class="badge warn">obići</span>' : ''}</div>`; })()}
    <div class="seg ktabs"><button class="${S.ktab !== 'posete' && S.ktab !== 'dok' ? 'on' : ''}" data-a="ktab" data-v="pregled">Pregled</button><button class="${S.ktab === 'posete' ? 'on' : ''}" data-a="ktab" data-v="posete">Posete · ${visitHistory(it).length}</button>${modOn('fin') && S.SY.dokumenta ? `<button class="${S.ktab === 'dok' ? 'on' : ''}" data-a="ktab" data-v="dok">Dokumenta</button>` : ''}</div>
    ${S.ktab === 'posete' ? vPosete(it) : S.ktab === 'dok' ? vDokumenta(it) : `
    ${modOn('fin') ? finBlock(it) + prometBlock(it) + stavkeBlock(it) : ''}
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
    ${f.Beleske ? `<h2>Beleške</h2><div class="card" style="white-space:pre-wrap">${esc(f.Beleske)}</div>` : ''}`}
    <div class="err" id="err"></div>
  </main><div class="foot"><div class="in">${visitFoot(it, i)}</div></div>`;
}
function visitFoot(it, i) {
  if (S.busy) return `<button class="btn pri" id="busy" disabled>${esc(S.busyMsg || 'Tražim lokaciju…')}</button>`;
  if (!S.visit) return `<button class="btn" data-a="note">Brza beleška</button><button class="btn pri" style="flex:2" data-a="visit">Započni posetu</button>`;
  if (S.visit.id !== it.id) return `<button class="btn" data-a="govisit">Poseta u toku: ${esc(S.visit.n)} ›</button>`;
  if (S.visit.draft) return `<button class="btn" data-a="endvisit">Završi posetu</button><button class="btn pri" style="flex:2" data-a="start">Nastavi audit · korak ${S.visit.draft.step + 1}/${STEPS.length}</button>`;
  return `<button class="btn" data-a="endvisit">Završi posetu</button><button class="btn pri" style="flex:2" data-a="start">${S.visit.audit ? 'Audit je urađen · ponovi' : i ? 'Ponovi audit' : 'Pokreni audit'}</button>`;
}

/* ---------- audit ---------- */
function startAudit(fresh) {
  if (!S.visit || S.visit.id !== S.cur.id) { toast('Audit može samo tokom posete kod ovog kupca.'); return; }
  if (S.visit.draft && !fresh) { S.d = S.visit.draft.d; S.step = S.visit.draft.step; go('audit'); return; }
  if (fresh) { S.visit.draft = null; saveDay(); }
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
  S.d._d0 = JSON.parse(JSON.stringify(S.d));
  S.photos = []; S.step = 0; S.visit.draft = { d: S.d, step: 0 }; saveDay(); go('audit');
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
    body = `<div class="pbtns"><label class="big-act"><span>📷</span>Slikaj<input type="file" accept="image/*" capture="environment" data-i="photo" hidden></label><label class="big-act"><span>🖼</span>Iz galerije<input type="file" accept="image/*" multiple data-i="photo" hidden></label></div>
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
  const KEY = [['StatusNovi', 'Status'], ['AuditSegment', 'Segment'], ['AuditTeren', 'Teren'], ['AuditNacinPlacanja', 'Način plaćanja'], ['AuditPlanObilaska', 'Plan obilaska'], ['AuditNacinRazvoza', 'Način razvoza'], ['AuditNacinPozivanja', 'Način pozivanja']];
  const ch = KEY.filter(([k]) => cv(f[k]) && d[k] && cv(f[k]) !== d[k]);
  const closing = ['zatvoren', 'neaktivan'].includes(d.StatusNovi) && cv(f.StatusNovi) !== d.StatusNovi;
  if ((ch.length || closing) && !await askSure('Menjaš važne podatke', `<p>Kod kupca <b>${esc(f.field_1 || f.Title)}</b> menjaš:</p><ul>${ch.map(([k, n]) => `<li><b>${n}:</b> ${esc(cv(f[k]))} → <b>${esc(d[k])}</b></li>`).join('')}${closing && !ch.some(([k]) => k === 'StatusNovi') ? `<li><b>Status:</b> → <b>${esc(d.StatusNovi)}</b></li>` : ''}</ul><p>Da li su izmene tačne?</p>`, 'Da, sačuvaj')) return;
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
  if (S.visit && S.visit.id === S.cur.id) { S.visit.audit = true; S.visit.draft = null; }
  if (S.day) S.day.audits = (S.day.audits || 0) + 1; saveDay();
  S.d = null;
  try { tikFromAudit(S.cur, d); } catch (e) { console.warn('tiket', e); }
  await logEvent('AUDIT', null, S.cur, { Napomena: d.AuditIshod });
  toast((navigator.onLine || DEMO ? 'Audit je sačuvan.' : 'Audit je sačuvan u telefonu, šalje se kad bude mreže.') + ' Poseta i dalje traje.');
  S.step = 0; history.back();
}

/* ---------- nov kupac ---------- */
function vNovi() {
  return bar('Nov kupac', true) + `<main class="main">
    <p class="s">Radnja koje nema na spisku. Posle upisa odmah kreće audit.</p>
    ${[['n_naziv', 'Naziv radnje', 'text'], ['n_adresa', 'Adresa', 'text'], ['n_mesto', 'Mesto', 'text'], ['n_tel', 'Telefon', 'tel'], ['n_kontakt', 'Kontakt osoba', 'text']].map(([k, l, t]) => `<div class="fld"><label>${l}</label><input id="${k}" type="${t}"></div>`).join('')}
    <div class="err" id="err"></div>
  </main><div class="foot"><div class="in"><button class="btn pri" data-a="novisave">Upiši i pokreni audit</button></div></div>`;
}
async function saveNovi(skipGps) {
  const v = id => document.getElementById(id).value.trim();
  if (!v('n_naziv')) { document.getElementById('err').textContent = 'Upiši naziv radnje.'; return; }
  const now = new Date(), code = 'NOVO-' + now.toISOString().slice(2, 16).replace(/\D/g, '');
  const fields = { Title: code, field_1: v('n_naziv'), field_3: v('n_adresa'), field_4: v('n_mesto'), field_5: v('n_tel'), field_7: v('n_kontakt'), StatusNovi: 'potencijalni' };
  if (S.visit) { document.getElementById('err').textContent = 'Prvo završi posetu kod: ' + S.visit.n; return; }
  const btn = document.querySelector('[data-a=novisave]'), er = document.getElementById('err'); let p = null;
  if (!skipGps) {
  btn.disabled = true; btn.textContent = 'Stani ispred ulaza, napolju · tražim signal…'; er.textContent = '';
  try { p = await getPrecise(a => { btn.textContent = `Stani ispred ulaza, napolju · ±${a} m${a <= MAPACC ? ' ✓' : '…'}`; }); }
  catch (e) { btn.disabled = false; btn.textContent = 'Upiši i pokreni audit'; gpsFail(e.message, 'novi'); return; }
  if (!await askSure('Zabeleži lokaciju objekta', `<p>${esc(fields.field_1)}<br><span class="s">${esc([fields.field_3, fields.field_4].filter(Boolean).join(', '))}</span></p><p>Lokacija novog objekta biće zabeležena na mestu gde stojiš (preciznost ±${p.acc} m).</p><p><b>Da li stojiš ispred ulaza ovog objekta?</b></p>`, 'Da, zabeleži')) { btn.disabled = false; btn.textContent = 'Upiši i pokreni audit'; return; }
  fields.Lokacija = fmtLoc(p);
  }
  { const ter = terenFor(fields.field_4, p); if (ter) fields.AuditTeren = ter; else if (S.teren && S.teren !== '*' && S.teren !== 'Bez terena') fields.AuditTeren = S.teren; }
  const it = { id: 'tmp' + now.getTime(), f: { ...fields } };
  S.items.push(it); saveItems();
  await queue({ type: 'create', tmp: it.id, fields });
  S.cur = it; beginVisit(it, p);
  await logEvent('POČETAK POSETE', p, it, { Napomena: p ? 'nov kupac: lokacija objekta zabeležena' : 'nov kupac, BEZ GPS-a: ' + (S.gpsErr || '') });
  history.replaceState({ v: 'kartica' }, ''); S.view = 'kartica'; startAudit();
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
    case 'kupac': S.cur = S.items.find(x => String(x.id) === v); S.from = S.view; S.ktab = 'pregled'; go('kartica'); break;
    case 'tab': S.view = v; S.showAll = false; history.replaceState({ v }, ''); render(); scrollTo(0, 0); break;
    case 'fsel': await openFilter(v); break;
    case 'fintoggle': S.finOpen = !S.finOpen; rerender(); break;
    case 'pjtoggle': S.pjOpen = !S.pjOpen; rerender(); break;
    case 'pjimport': await pjImport(); break;
    case 'sttoggle': S.stOpen = !S.stOpen; rerender(); break;
    case 'pday': { let d = v; if (v === 'cal') d = await askDate('Izaberi dan za plan', S.pday && S.pday !== ymd() ? S.pday : addDays(ymd(), 2)); if (d) { S.pday = d; S.selMode = false; S.sel = new Set(); render(); } break; }
    case 'pup': planMove(v, -1); break;
    case 'pdn': planMove(v, 1); break;
    case 'pmenu': await planMenu(v); break;
    case 'predlog': await predlogPlana(); break;
    case 'selmode': S.selMode = !S.selMode; S.sel = new Set(); rerender(); break;
    case 'selrow': { S.sel = S.sel || new Set(); S.sel.has(v) ? S.sel.delete(v) : S.sel.add(v); rerender(); break; }
    case 'seladd': { const d = S.pday || ymd(); const its = S.items.filter(x => S.sel.has(String(x.id))); const r = addToPlan(its, d); S.selMode = false; S.sel = new Set(); render(); toast(`Dodato u plan za ${dLabel(d)}: ${r.added}${r.skip ? ' (već u planu: ' + r.skip + ')' : ''}`); break; }
    case 'mapa': go('mapa'); break;
    case 'mapday': S.mapDay = v; render(); break;
    case 'task': await openTask(v); break;
    case 'newtask': await newTask(null); break;
    case 'tktoggle': S.tkOpen = !S.tkOpen; rerender(); break;
    case 'thread': openThread(v); break;
    case 'tikf': S.tikF = v; rerender(); break;
    case 'kord': openOrder(S.cur); break;
    case 'otip': S.ord.tip = v; ordSave(); render(); scrollTo(0, 0); break;
    case 'otipch': S.ord.tip = ''; ordSave(); render(); break;
    case 'otab': S.ord.tab = v; if (v === 'obicno') { S.ord.s = ''; S.ord.kat = ''; } ordSave(); render(); scrollTo(0, 0); break;
    case 'oq': { const a = S.SY.artikli.art[v], cur = num(S.ord.q[v]); const st = el.dataset.k === 'tp' ? (a.tp || 1) : ordStep(a) * num(el.dataset.k); ordSetQ(v, cur + st); ordRefresh(); break; }
    case 'oclear': if (await askSure('Obriši narudžbinu', '<p>Brišu se sve količine iz ove narudžbine.</p>', 'Obriši')) { S.ord.q = {}; ordSave(); render(); } break;
    case 'osend': await ordSend(); break;
    case 'tiket': S.tikId = v; go('tiket'); break;
    case 'mojitik': S.tikF = 'svi'; go('tiketi'); break;
    case 'tikpre': await tikAct('pre'); break;
    case 'tikvrati': await tikAct('vrati'); break;
    case 'tikres': await tikAct('res'); break;
    case 'tikaut': { const t = S.X.tik.find(x => x.id === S.tikId); if (t) { S.toWho = isKom() ? 'op' : 'kom'; openThread(t.f.KupacSifra || '', `Tiket „${t.f.Title}”: `); } break; }
    case 'towho': S.toWho = v; S.msgTxt = (document.querySelector('[data-i=msgtxt]') || {}).value || ''; rerender(); break;
    case 'sendtpl': { const inp = document.querySelector('[data-i=msgtxt]'); const cur = inp ? inp.value.trim() : ''; sendMsg(cur && /…$/.test(v) ? v.replace(/…$/, ' ') + cur : v); break; }
    case 'sendmsg': sendMsg((document.querySelector('[data-i=msgtxt]') || {}).value); break;
    case 'msg2task': msgToTask(v); break;
    case 'kplan': { const d = await pickDate('Dodaj u plan: ' + (S.cur.f.field_1 || '')); if (d) { const r = addToPlan([S.cur], d); toast(r.added ? 'Dodato u plan za ' + dLabel(d) + '.' : 'Već je u planu za ' + dLabel(d) + '.'); render(); } break; }
    case 'ktask': await newTask(S.cur); break;
    case 'kmsg': openThread(String(S.cur.f.Title)); break;
    case 'kkonk': await addKonk(S.cur); break;
    case 'kpin': go('pin'); break;
    case 'pinsave': await pinSave(); break;
    case 'pot': await setPot(S.cur, num(v)); break;
    case 'newods': await newAbsence(); break;
    case 'odsok': case 'odsno': { const o = S.X.ods.find(x => x.id === v); if (o) xSet('ods', o, { Status: a === 'odsok' ? 'odobreno' : 'odbijeno', Odobrio: me() }); render(); break; }
    case 'wprev': S.week = addDays(S.week || weekStart(ymd()), -7); render(); break;
    case 'wnext': S.week = addDays(S.week || weekStart(ymd()), 7); render(); break;
    case 'wtoday': S.week = weekStart(ymd()); render(); break;
    case 'nzprev': S.nzDay = addDays(S.nzDay || ymd(), -1); render(); break;
    case 'nznext': S.nzDay = addDays(S.nzDay || ymd(), 1); render(); break;
    case 'nztoday': S.nzDay = ymd(); render(); break;
    case 'nzmail': { const w = window.open('', '_blank'); if (w) { w.document.write(reportHtml(S.nzDay || ymd(), S.planUser || komUsers()[0])); w.document.close(); } break; }
    case 'adminsave': adminSave(); break;
    case 'jr': S.jr = v; rerender(); break;
    case 'juceok': await closeYesterday(); break;
    case 'kdr': S.kd[el.dataset.k].r = v; rerender(); break;
    case 'kda': { const k = S.kd[el.dataset.k]; if (v === 'date') { const d = await askDate('Prebaci na dan', k.d || nextWorkday(ymd())); if (!d) break; k.d = d; } k.a = v; rerender(); break; }
    case 'fclr': delete S.flt[v]; saveFlt(); render(); break;
    case 'fclrall': S.flt = {}; saveFlt(); render(); break;
    case 'obtoggle': S.obOpen = !S.obOpen; rerender(); break;
    case 'onlydue': S.onlyDue = !S.onlyDue; S.showAll = false; rerender(); break;
    case 'showall': S.showAll = true; rerender(); break;
    case 'ktab': S.ktab = v; rerender(); break;
    case 'note': openNote(false); break;
    case 'nkobici': S.nk.obici = !S.nk.obici; rerender(); break;
    case 'notesave': await saveNote(); break;
    case 'logout': if (DEMO) toast('Probni režim.'); else if (await askSure('Odjava', '<p>Odjaviti se sa aplikacije? Podaci koji čekaju na slanje ostaju u telefonu.</p>', 'Odjavi me')) S.pca.logoutRedirect({ account: S.account }); break;
    case 'notel': toast('Telefon nije upisan.'); break;
    case 'nomail': toast('Mejl nije upisan.'); break;
    case 'start': startAudit(); break;
    case 'novi': { const c = await askChoice('Nov objekat', [['visit', 'Poseta novom kupcu', 'Ušao sam, razgovaram i radim ličnu kartu (audit).'], ['note', 'Brza beleška', 'Prolazio sam: slika, napomena i okvirna lokacija, pa dolazim sledeći put.']]); if (c === 'visit') { S.from = S.view; go('novi'); } else if (c === 'note') openNote(true); break; }
    case 'novisave': await saveNovi(); break;
    case 'pick': d[STEPS[S.step].f] = d[STEPS[S.step].f] === v ? '' : v; rerender(); break;
    case 'pickk': d[el.dataset.k] = d[el.dataset.k] === v ? '' : v; rerender(); break;
    case 'mon': d.AuditMonitoring = !d.AuditMonitoring; rerender(); break;
    case 'day': { const k = STEPS[S.step].f, i = d[k].indexOf(v); i < 0 ? d[k].push(v) : d[k].splice(i, 1); rerender(); break; }
    case 'dayrow': { const k = STEPS[S.step].f, vals = [1, 2, 3, 4].map(w => v + ' ' + w), allOn = vals.every(x => d[k].includes(x)); d[k] = allOn ? d[k].filter(x => !vals.includes(x)) : [...new Set(d[k].concat(vals))]; rerender(); break; }
    case 'dayclear': d[STEPS[S.step].f] = []; rerender(); break;
    case 'quick': { const k = STEPS[S.step].f; d[k] = d[k] ? d[k] + ', ' + v : v; rerender(); break; }
    case 'as': { const y = d.DrziAsortiman, n = d.NeDrzi; [y, n].forEach(l => { const i = l.indexOf(v); if (i >= 0) l.splice(i, 1); }); if (!(el.classList.contains('y') || el.classList.contains('x'))) (el.dataset.k === 'y' ? y : n).push(v); rerender(); break; }
    case 'daystart': await dayStart(); break;
    case 'dayend': await dayEnd(); break;
    case 'kraj': go('kraj'); break;
    case 'visit': await visitStart(); break;
    case 'govisit': { const it = visitItem(); if (it) { S.cur = it; S.from = S.view; S.ktab = 'pregled'; go('kartica'); } else { S.vk = null; go('krajposete'); } break; }
    case 'endvisit': S.vk = null; go('krajposete'); break;
    case 'nogps': await noGps(v); break;
    case 'remap': await remap(); break;
    case 'startnew': startAudit(true); break;
    case 'vish': S.vk.b = document.getElementById('vkb').value; S.vk.i = S.vk.i === v ? '' : v; rerender(); break;
    case 'endvisitsave': await visitEnd(); break;
    case 'prev': if (S.step) { S.step--; go('audit', false); } else history.back(); break;
    case 'next': if (S.step < STEPS.length - 1) { S.step++; go('audit', false); } else await finishAudit(); break;
  }
});
document.addEventListener('input', e => {
  const k = e.target.dataset.i; if (!k) return;
  if (k === 'pq') { S.pq = e.target.value; const y = scrollY; render(); const i = document.querySelector('[data-i=pq]'); if (i) { i.focus(); i.setSelectionRange(i.value.length, i.value.length); } scrollTo(0, y); return; }
  if (k === 'msgtxt') { S.msgTxt = e.target.value; return; }
  if (k === 'ordq' && S.ord) { S.ord.s = e.target.value; if (S.ord.tab === 'korpa') S.ord.tab = 'sve'; ordSave(); ordRefresh(); return; }
  if (k === 'ordkat' && S.ord) { S.ord.kat = e.target.value; if (S.ord.tab !== 'sve') S.ord.tab = 'sve'; ordSave(); ordRefresh(); return; }
  if (k.startsWith('oq_') && S.ord) { ordSetQ(k.slice(3), e.target.value); const t = ordTotal(S.cur), f = document.querySelector('.foot .grow'); if (f) f.innerHTML = `<b>${din(t.sum)}</b><div class="s">${t.n} artikala · bez PDV</div>`; const b = document.querySelector('[data-a=osend]'); if (b) b.disabled = !t.n; return; }
  if (k === 'planuser') { S.planUser = e.target.value; render(); return; }
  if (k.startsWith('kdt_')) { const x = S.kd[k.slice(4)]; if (x) x.t = e.target.value; return; }
  if (k === 'dq') { S.dq = e.target.value; S.showAll = false; const l = document.getElementById('dlist'); if (l) l.innerHTML = danasList(); else { render(); const i = document.querySelector('[data-i=dq]'); if (i) { i.focus(); i.setSelectionRange(i.value.length, i.value.length); } } return; }
  if (k === 'q') { S.q = e.target.value; const l = document.getElementById('list'); if (l) l.innerHTML = kupciList().html; return; }
  if (k === 'optq') { const q = e.target.value.toLowerCase(); document.querySelectorAll('#opts .opt').forEach(b => b.style.display = b.textContent.toLowerCase().includes(q) ? '' : 'none'); return; }
  if (k === 'photo' || k === 'nphoto') return;
  if (k.startsWith('nk_')) { if (S.nk) S.nk[k] = e.target.value; return; }
  if (S.d) { S.d[k] = e.target.value; saveDraft(); }
});
document.addEventListener('change', async e => {
  if (e.target.dataset.i === 'nphoto' && e.target.files.length && S.nk) {
    for (const file of e.target.files) {
      if (!S.nk.photoLoc) { const x = await exifInfo(file); if (x && x.lat) S.nk.photoLoc = { lat: x.lat, lon: x.lon, acc: 30, time: x.time }; }
      const blob = await shrink(file); S.nk.photos.push({ blob, url: URL.createObjectURL(blob) });
    }
    e.target.value = '';
    render(); return;
  }
  if (e.target.dataset.i !== 'photo' || !e.target.files.length) return;
  for (const file of e.target.files) {
    const blob = await shrink(file);
    const n = S.photos.length + 1, stamp = new Date().toISOString().slice(0, 16).replace(/\D/g, '');
    S.photos.push({ blob, url: URL.createObjectURL(blob), name: `${S.cur.f.Title}-${stamp}-${n}.jpg` });
  }
  render();
});
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
/* ================= PLANIRANJE, ZADACI, PORUKE, KONKURENCIJA, ODSUSTVA, NADZOR, ADMIN ================= */
const ROLES = Object.fromEntries(Object.entries(C.roles || {}).map(([k, v]) => [k.toLowerCase(), v]));
const NAMES = Object.fromEntries(Object.entries(C.names || {}).map(([k, v]) => [k.toLowerCase(), v]));
const role = () => DEMO ? (new URLSearchParams(location.search).get('role') || 'kom') : (ROLES[userName().toLowerCase()] || C.defaultRole || 'op');
const me = () => userName().toLowerCase();
const isKom = () => role() === 'kom';
const isAdmin = () => role() === 'admin';
const nameOf = e => { e = String(e || '').toLowerCase(); return NAMES[e] || ({ kom: 'Komercijalista', op: 'Call centar', mag: 'Magacin', admin: 'Admin' })[e] || e.split('@')[0] || '—'; };
const ROLE_N = { kom: 'Petar (komercijalista)', op: 'Call centar (Goca, Tamara)', mag: 'Magacin', admin: 'Admin' };
const komUsers = () => Object.keys(ROLES).filter(k => ROLES[k] === 'kom');
const ymd = d => { const x = d ? new Date(d) : new Date(); return x.getFullYear() + '-' + String(x.getMonth() + 1).padStart(2, '0') + '-' + String(x.getDate()).padStart(2, '0'); };
const addDays = (s, n) => { const x = new Date(s + 'T12:00:00'); x.setDate(x.getDate() + n); return ymd(x); };
const dLabel = s => { const t = ymd(); if (s === t) return 'danas'; if (s === addDays(t, 1)) return 'sutra'; if (s === addDays(t, 2)) return 'prekosutra'; if (s === addDays(t, -1)) return 'juče'; return new Date(s + 'T12:00:00').toLocaleDateString('sr-Latn-RS', { weekday: 'short', day: 'numeric', month: 'numeric' }); };
const DOW = ['nedelja', 'ponedeljak', 'utorak', 'sreda', 'četvrtak', 'petak', 'subota'];

/* ---------- podaci iz BizniSoft-a (folder Sinhronizacija, puni ga program na serveru) ---------- */
const SYNC_FILES = ['finansije', 'promet', 'dokumenta', 'stavke', 'kupci', 'stanje', 'artikli', 'zalihe'];
S.SY = {}; S.SYtag = {};
try { S.SYtag = JSON.parse(localStorage.getItem('ak_sytag_v1') || '{}'); } catch (e) { }
async function syncLoadCache() {
  if (!('caches' in window)) return;
  try { const c = await caches.open('ak-sync'); for (const n of SYNC_FILES) { const r = await c.match('/sync/' + n + '.json'); if (r) S.SY[n] = await r.json(); } } catch (e) { console.warn('sync cache', e); }
}
async function syncFetch() {
  if (DEMO || !navigator.onLine) return;
  await resolveList();
  const ch = await g('/sites/' + S.site + '/drive/root:/Sinhronizacija:/children?$top=50');
  const c = ('caches' in window) ? await caches.open('ak-sync') : null;
  for (const f of ch.value) {
    const n = f.name.replace(/\.json$/, ''); if (!SYNC_FILES.includes(n)) continue;
    if (S.SYtag[n] === f.eTag && S.SY[n]) continue;
    const r = f['@microsoft.graph.downloadUrl'] ? await fetch(f['@microsoft.graph.downloadUrl']) : await fetch('https://graph.microsoft.com/v1.0/sites/' + S.site + '/drive/root:/Sinhronizacija/' + f.name + ':/content', { headers: { Authorization: 'Bearer ' + await token() } }); if (!r.ok) continue;
    const txt = await r.text();
    try { S.SY[n] = JSON.parse(txt); S.SYtag[n] = f.eTag; if (c) await c.put('/sync/' + n + '.json', new Response(txt, { headers: { 'Content-Type': 'application/json' } })); } catch (e) { console.warn('sync', n, e); }
  }
  try { localStorage.setItem('ak_sytag_v1', JSON.stringify(S.SYtag)); } catch (e) { }
}
const bsKey = it => { const v = it && (it.f.Centrala || it.f.Title); const n = Number(v); return isNaN(n) ? '' : String(n); };
const finOf = it => (S.SY.finansije && S.SY.finansije.k[bsKey(it)]) || null;
const bsKupac = it => { const k = bsKey(it); if (!k || !S.SY.kupci) return null; if (!S.SY._kmap) { S.SY._kmap = {}; S.SY.kupci.kupci.forEach(x => S.SY._kmap[String(x.s)] = x); } return S.SY._kmap[k] || null; };
function finFlag(it) {
  const f = finOf(it), b = bsKupac(it);
  if (b && String(b.blok) === '1') return '<span class="badge bad">blokiran</span>';
  if (f && f.d > 0 && f.m > 30) return `<span class="badge bad">kasni ${f.m} d</span>`;
  if (f && f.d > 0) return `<span class="badge warn">dospelo ${money(f.d)}</span>`;
  return '';
}
const ago = s => s ? Math.floor((Date.now() - Date.parse(s)) / 864e5) : null;
function finBlock(it) {
  if (!modOn('fin')) return '';
  const f = finOf(it), b = bsKupac(it), gen = S.SY.finansije && S.SY.finansije.gen;
  if (!S.SY.finansije) return '';
  const lim = b && Number(b.lim) > 0 ? Number(b.lim) : 0;
  if (!f) return `<div class="card fin ok"><b>Finansije:</b> nema duga${b && String(b.blok) === '1' ? ' · <b class="late">BLOKIRAN</b>' : ''}<div class="s">stanje iz BizniSoft-a ${fmtDate(gen)} ${hhmm(gen)}</div></div>`;
  const late = f.d > 0;
  return `<div class="card fin ${late ? 'bad' : ''}">
    <div class="fgrid"><div><span class="s">Saldo</span><b>${money(f.s)}</b></div><div><span class="s">Dospelo</span><b class="${late ? 'late' : ''}">${money(f.d)}</b></div><div><span class="s">Kasni najviše</span><b class="${f.m > 30 ? 'late' : ''}">${f.m ? f.m + ' d' : '—'}</b></div><div><span class="s">Poslednja uplata</span><b>${f.u ? fmtDate(f.u) : '—'}</b>${f.u ? `<span class="s">${money(f.ui)} · pre ${ago(f.u)} d</span>` : ''}</div></div>
    ${lim ? `<div class="s">Kreditni limit ${money(lim)}${f.s > lim ? ' · <b class="late">prekoračen</b>' : ''}</div>` : ''}${b && String(b.blok) === '1' ? '<div><b class="late">Kupac je BLOKIRAN u BizniSoft-u</b></div>' : ''}
    ${f.n ? `<button class="sec" data-a="fintoggle">${S.finOpen ? '▾' : '▸'} Neplaćene fakture · ${f.n}</button>${S.finOpen ? `<table class="tbl">${f.o.map(o => { const d = ago(o[2]); return `<tr><td>${esc(String(o[0]).replace(/^VFNL\w+-VFNL\w+-/, 'KA-'))}<div class="s">${fmtDate(o[1])}</div></td><td>${money(o[4])}${o[4] !== o[3] ? `<div class="s">od ${money(o[3])}</div>` : ''}</td><td>${o[2] ? fmtDate(o[2]) : ''}${d > 0 ? `<div class="late s">kasni ${d} d</div>` : ''}</td></tr>`; }).join('')}</table>` : ''}` : ''}
    <div class="s">stanje iz BizniSoft-a ${fmtDate(gen)} ${hhmm(gen)}</div></div>`;
}
function prometBlock(it) {
  const P = S.SY.promet && S.SY.promet.k[bsKey(it)]; if (!P) return '';
  const now = new Date(), ms = []; for (let i = 23; i >= 0; i--) { const d = new Date(now.getFullYear(), now.getMonth() - i, 1); ms.push(d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0')); }
  const v = ms.map(m => (P.m[m] || [0])[0]), last12 = v.slice(12).reduce((a, b) => a + b, 0), prev12 = v.slice(0, 12).reduce((a, b) => a + b, 0), mx = Math.max(1, ...v.slice(12));
  const tr = prev12 ? Math.round((last12 - prev12) / prev12 * 100) : null;
  return `<div class="card"><div class="fgrid"><div><span class="s">Promet 12 mes.${isPJ(it) ? ' (cela firma)' : ''}</span><b>${money(last12)}</b></div><div><span class="s">Prethodnih 12</span><b>${money(prev12)}</b></div><div><span class="s">Trend</span><b class="${tr !== null && tr < -10 ? 'late' : ''}">${tr === null ? '—' : (tr > 0 ? '+' : '') + tr + '%'}</b></div><div><span class="s">Poslednja kupovina</span><b>${P.z ? fmtDate(P.z) : '—'}</b></div></div>
    <div class="bars">${ms.slice(12).map((m, i) => `<div title="${m}: ${money(v[12 + i])}"><i style="height:${Math.round(v[12 + i] / mx * 100)}%"></i><span>${m.slice(5)}</span></div>`).join('')}</div></div>`;
}
function stavkeBlock(it) {
  const st = S.SY.stavke; const L = st && st.k[bsKey(it)]; if (!L) return '';
  const A = st.artikli, now = Date.now();
  const top = L.slice(0, 10), stopped = L.filter(x => x[3] && (now - Date.parse(x[3])) / 864e5 > 90 && x[4] >= 2).slice(0, 10);
  const row = x => `<tr><td>${esc((A[x[0]] || [x[0]])[0])}<div class="s">${esc((A[x[0]] || [])[1] || x[0])}</div></td><td>${money(x[2])}</td><td class="s">${fmtDate(x[3])}</td></tr>`;
  return `<button class="sec" data-a="sttoggle">${S.stOpen ? '▾' : '▸'} Šta kupuje (12 meseci)</button>${S.stOpen ? `<table class="tbl">${top.map(row).join('')}</table>${stopped.length ? `<div class="sh" style="margin-top:10px">⚠ Prestao da kupuje (ranije redovno, nije kupio 90+ dana)</div><table class="tbl">${stopped.map(row).join('')}</table>` : ''}` : ''}`;
}
function vDokumenta(it) {
  const D = S.SY.dokumenta && S.SY.dokumenta.k[bsKey(it)];
  if (!D) return '<div class="empty">Nema dokumenata u poslednjih 12 meseci (ili podaci još nisu stigli).</div>';
  const N = { FA: 'Faktura', PF: 'Profaktura', NK: 'Narudžbina', PK: 'Povraćaj', OK: 'Odobrenje' };
  return `<table class="tbl">${D.map(d => d && d.value && !Array.isArray(d) ? d.value : d).map(d => `<tr><td><b>${N[d[0]] || d[0]}</b><div class="s">${esc(d[1])}</div></td><td>${fmtDate(d[2])}</td><td>${money(d[3])}<div class="s">${d[6]} st.</div></td></tr>`).join('')}</table><p class="s">stanje iz BizniSoft-a ${fmtDate(S.SY.dokumenta.gen)} ${hhmm(S.SY.dokumenta.gen)}</p>`;
}
/* ---------- poslovne jedinice ---------- */
const isPJ = it => !!(it && it.f.Centrala);
const pjOf = c => { const k = String(Number(c.f.Title)); return S.items.filter(x => x.f.Centrala === k).sort((a, b) => num(a.f.PJ) - num(b.f.PJ)); };
const centralaOf = it => S.items.find(x => !x.f.Centrala && String(Number(x.f.Title)) === String(it.f.Centrala));
function pjMissing() {
  const K = S.SY.kupci; if (!K) return [];
  const have = new Set(S.items.map(x => String(x.f.Title)));
  const byK = {}; S.items.forEach(x => { if (!x.f.Centrala) byK[String(Number(x.f.Title))] = x; });
  return K.pj.filter(p => byK[String(p.k)] && !have.has(p.k + '/' + p.s)).map(p => ({ p, c: byK[String(p.k)] }));
}
async function pjImport() {
  const M = pjMissing(); if (!M.length) { toast('Nema novih poslovnih jedinica.'); return; }
  if (!await askSure('Poslovne jedinice', `<p>Dodati <b>${M.length}</b> poslovnih jedinica iz BizniSoft-a u listu kupaca? Svaka dobija svoju ličnu kartu, plan i posete; finansije ostaju na centrali.</p>`, 'Dodaj')) return;
  for (const { p, c } of M) {
    const f = c.f, fields = { Title: p.k + '/' + p.s, Centrala: String(p.k), PJ: num(p.s), field_1: `${f.field_1 || ''} · ${p.n || 'PJ ' + p.s}`, field_2: f.field_2 || '', field_3: p.adr || '', field_4: p.mesto || '', field_5: p.tel || '', field_7: p.kon || '' };
    ['field_8', 'field_10', 'field_11', 'field_12', 'field_13', 'field_14'].forEach(k => { if (f[k]) fields[k] = f[k]; });
    if (cv(f.StatusNovi)) fields.StatusNovi = cv(f.StatusNovi);
    if (cv(f.AuditSegment)) fields.AuditSegment = cv(f.AuditSegment);
    const t = terenFor(p.mesto, null) || cv(f.AuditTeren); if (t) fields.AuditTeren = t;
    const it = { id: 'tmp' + Date.now() + Math.random().toString(36).slice(2, 6), f: { ...fields } };
    S.items.push(it); await queue({ type: 'create', tmp: it.id, fields });
  }
  saveItems(); toast(`Dodato ${M.length} poslovnih jedinica. Šalje se na SharePoint…`); render();
}
function pjBlock(it) {
  if (isPJ(it)) { const c = centralaOf(it); return `<div class="card pjc">🏢 Poslovna jedinica ${esc(String(it.f.PJ || ''))} firme ${c ? `<button class="lnk" data-a="kupac" data-v="${esc(c.id)}">${esc(c.f.field_1 || c.f.Title)} ›</button>` : esc(it.f.Centrala)}<div class="s">Finansije, dokumenta i promet su na nivou cele firme.</div></div>`; }
  const P = pjOf(it); if (!P.length) return '';
  return `<button class="sec" data-a="pjtoggle">${S.pjOpen ? '▾' : '▸'} Poslovne jedinice · ${P.length}</button>${S.pjOpen ? P.map(rowKupac).join('') : ''}`;
}
/* ---------- liste ---------- */
const XL = { plan: 'Plan poseta', zad: 'Zadaci', por: 'Poruke', ods: 'Odsustva', konk: 'Konkurencija', pod: 'Podešavanja', tik: 'Tiketi' };
const LS_X = 'ak_x_v1';
S.X = { plan: [], zad: [], por: [], ods: [], konk: [], pod: [], tik: [] }; S.XL = {};
try { const x = JSON.parse(localStorage.getItem(LS_X) || 'null'); if (x) Object.assign(S.X, x); } catch (e) { }
if (!S.X.tik) S.X.tik = [];
const saveX = () => { try { localStorage.setItem(LS_X, JSON.stringify(S.X)); } catch (e) { } };
async function resolveX() {
  if (Object.keys(S.XL).length) return;
  await resolveList();
  const lists = await g('/sites/' + S.site + '/lists?$select=id,displayName&$top=200');
  Object.entries(XL).forEach(([k, n]) => { const l = lists.value.find(x => x.displayName === n); if (l) S.XL[k] = l.id; });
}
async function fetchX(k) {
  if (!S.XL[k]) return;
  let url = '/sites/' + S.site + '/lists/' + S.XL[k] + '/items?$expand=fields&$top=999';
  const out = [];
  while (url) { const j = await g(url); j.value.forEach(x => { const f = { ...x.fields }; ['@odata.etag', 'ContentType', 'Attachments', 'AuthorLookupId', 'EditorLookupId', '_UIVersionString', 'Edit', 'LinkTitleNoMenu', 'LinkTitle', 'ItemChildCount', 'FolderChildCount', '_ComplianceFlags', '_ComplianceTag', '_ComplianceTagWrittenTime', '_ComplianceTagUserId', 'AppAuthorLookupId', 'AppEditorLookupId'].forEach(z => delete f[z]); out.push({ id: String(x.id), f }); }); url = j['@odata.nextLink']; }
  const pend = await DB.all();
  const tmp = S.X[k].filter(x => String(x.id).startsWith('tmp'));
  pend.filter(o => o.type === 'x' && o.k === k && o.act === 'patch').forEach(o => { const it = out.find(x => x.id === String(o.itemId)); if (it) Object.assign(it.f, o.fields); });
  S.X[k] = out.concat(tmp);
}
async function fetchAllX(keys) {
  if (DEMO) return;
  await resolveX();
  for (const k of keys || Object.keys(XL)) { try { await fetchX(k); } catch (e) { console.warn('x', k, e); } }
  saveX();
}
function xAdd(k, fields) {
  const it = { id: 'tmp' + Date.now() + Math.random().toString(36).slice(2, 6), f: { ...fields, Created: new Date().toISOString() } };
  S.X[k].push(it); saveX();
  queue({ type: 'x', k, act: 'create', tmp: it.id, fields });
  return it;
}
function xSet(k, it, fields) {
  Object.assign(it.f, fields); saveX();
  queue({ type: 'x', k, act: 'patch', itemId: it.id, fields });
}
async function flushX(op) {
  await resolveX();
  if (!S.XL[op.k]) throw new Error('Nema liste ' + XL[op.k]);
  if (op.act === 'create') {
    const r = await g('/sites/' + S.site + '/lists/' + S.XL[op.k] + '/items', { method: 'POST', body: JSON.stringify({ fields: op.fields }) });
    const it = S.X[op.k].find(x => x.id === op.tmp); if (it) it.id = String(r.id);
    try {
      const f = op.fields;
      if (op.k === 'por') notifyTeams(f.Za, `💬 ${nameOf(f.Od)}${f.KupacNaziv ? ' · ' + f.KupacNaziv : ''}: ${f.Tekst || f.Title}`);
      if (op.k === 'zad' && f.Za && (f.Za || '').toLowerCase() !== (f.Zadao || '').toLowerCase()) notifyTeams(f.Za, `✅ Nov zadatak od ${nameOf(f.Zadao)} (${(PRI[num(f.Prioritet) || 2] || [''])[0]})${f.KupacNaziv ? ' · ' + f.KupacNaziv : ''}: ${f.Title}${f.Rok ? ' · rok ' + f.Rok : ''}`);
    } catch (e) { console.warn('notify', e); }
    for (const o of await DB.all()) if (o.type === 'x' && String(o.itemId) === op.tmp) { o.itemId = String(r.id); await DB.put(o); }
    saveX();
  } else {
    if (String(op.itemId).startsWith('tmp')) return false;
    await g('/sites/' + S.site + '/lists/' + S.XL[op.k] + '/items/' + op.itemId + '/fields', { method: 'PATCH', body: JSON.stringify(op.fields) });
  }
  return true;
}

/* ---------- Teams obaveštenja (šalje ih aplikacija pošiljaoca) ---------- */
const APP_URL = 'https://ssashas-hub.github.io/audit-kupaca/index.html';
function recipientsOf(za) {
  za = String(za || '').toLowerCase();
  if (za.includes('@')) return [za];
  return Object.keys(ROLES).filter(e => ROLES[e] === za && e !== 'powerapps@ekolineks.onmicrosoft.com');
}
async function notifyTeams(za, text) {
  if (DEMO) return;
  const to = recipientsOf(za).filter(e => e !== me()); if (!to.length) return;
  let tk;
  try { tk = (await S.pca.acquireTokenSilent({ scopes: ['Chat.Create', 'ChatMessage.Send', 'Mail.Send'], account: S.account })).accessToken; }
  catch (e) { try { tk = (await S.pca.acquireTokenSilent({ scopes: ['Mail.Send'], account: S.account })).accessToken; } catch (e2) { console.warn('notify token', e2); return; } }
  const H = { Authorization: 'Bearer ' + tk, 'Content-Type': 'application/json' };
  const html = `${esc(text)}<br><a href="${APP_URL}">Otvori aplikaciju Audit kupaca</a>`;
  let chats = {}; try { chats = JSON.parse(localStorage.getItem('ak_chats_v1') || '{}'); } catch (e) { }
  for (const r of to) {
    try {
      if (ROLES[r] === 'mag') throw new Error('mejl');
      let id = chats[r];
      if (!id) {
        const c = await fetch('https://graph.microsoft.com/v1.0/chats', { method: 'POST', headers: H, body: JSON.stringify({ chatType: 'oneOnOne', members: [me(), r].map(u => ({ '@odata.type': '#microsoft.graph.aadUserConversationMember', roles: ['owner'], 'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${u}')` })) }) });
        if (!c.ok) throw new Error('chat ' + c.status);
        id = chats[r] = (await c.json()).id; localStorage.setItem('ak_chats_v1', JSON.stringify(chats));
      }
      const m = await fetch(`https://graph.microsoft.com/v1.0/chats/${id}/messages`, { method: 'POST', headers: H, body: JSON.stringify({ body: { contentType: 'html', content: html } }) });
      if (!m.ok) { delete chats[r]; localStorage.setItem('ak_chats_v1', JSON.stringify(chats)); throw new Error('msg ' + m.status); }
    } catch (e) {
      // rezerva: mejl
      try { await fetch('https://graph.microsoft.com/v1.0/me/sendMail', { method: 'POST', headers: H, body: JSON.stringify({ message: { subject: 'Audit kupaca: ' + text.slice(0, 80), body: { contentType: 'HTML', content: ROLES[r] === 'mag' ? `${esc(text)}<br><br>Odgovorite na ovaj mejl, odgovor stiže direktno pošiljaocu.` : html }, toRecipients: [{ emailAddress: { address: r } }] }, saveToSentItems: false }) }); } catch (e2) { console.warn('notify', r, e, e2); }
    }
  }
}
/* ---------- podešavanja (admin) ---------- */
const DEF_SET = {
  moduli: { '*': { fin: 1, plan: 1, zad: 1, por: 1, konk: 1, pot: 1, ods: 1, mapa: 1, tik: 1, nar: 1 } },
  podsetnik: '17:30',
  konkurenti: ['Špajz', 'Viva', 'Macrobiotic', 'NS COM', 'Mrvica'],
  praznici: ['2026-01-01', '2026-01-02', '2026-01-07', '2026-02-16', '2026-02-17', '2026-04-10', '2026-04-11', '2026-04-12', '2026-04-13', '2026-05-01', '2026-05-02', '2026-11-11', '2027-01-01', '2027-01-02', '2027-01-07', '2027-02-15', '2027-02-16', '2027-04-30', '2027-05-01', '2027-05-02', '2027-05-03', '2027-05-04', '2027-11-11'],
  sabloni: {
    op2kom: ['Pitaj za neplaćenu fakturu', 'Ponudi aktuelnu akciju', 'Kupac se žali na dostavu, proveri', 'Kupac nije naručio duže vreme, obiđi ga', 'Donesi katalog / cenovnik', 'Proveri policu i izlaganje', 'Kupac traži da ga posetiš'],
    kom2op: ['Kupac traži poziv', 'Kupac želi da naruči', 'Promenio kontakt osobu', 'Zatvorio objekat', 'Nezadovoljan cenom', 'Nezadovoljan dostavom', 'Traži cenovnik / katalog', 'Pošalji mu ponudu'],
    kom2mag: ['Ima li na stanju…', 'Kada stiže roba…', 'Kupac prijavio oštećenu robu', 'Kupac prijavio pogrešnu robu', 'Fali roba u isporuci'],
    odgovori: ['Urađeno', 'Kupac nije bio tu', 'Plaća do petka', 'Odbija', 'Javiće se', 'Treba mi još informacija', 'Na stanju je', 'Nema na stanju', 'Stiže sledeće nedelje']
  },
  razlozi: ['nije stigao, nestalo vremena', 'kupac zatvoren ili odsutan', 'promena plana, hitno drugo', 'kupac otkazao ili tražio drugi termin', 'put ili vreme', 'ostalo'],
  vrsteOdsustva: ['godišnji odmor', 'bolovanje', 'slobodan dan', 'službeno drugde', 'ostalo'],
  predlogBroj: 12,
  sinhronizacija: { kupci: 60, promet: 60, dokumenta: 30, finansije: 30, stavke: 1440, artikli: 60, zalihe: 30, od: 7, do: 19, nocu: 2 },
  kpi: { posetaDnevno: null, postovanjePlana: null, primarniZadaci: null },
  pragoviPotencijala: { 1: '', 2: '', 3: '', 4: '', 5: '' }
};
function getSet(k) {
  const it = S.X.pod.find(x => x.f.Title === k);
  if (it) { try { return JSON.parse(it.f.Vrednost); } catch (e) { } }
  return JSON.parse(JSON.stringify(DEF_SET[k]));
}
function putSet(k, v) {
  const it = S.X.pod.find(x => x.f.Title === k), val = JSON.stringify(v);
  if (it) xSet('pod', it, { Vrednost: val, Izmenio: me() }); else xAdd('pod', { Title: k, Vrednost: val, Izmenio: me() });
}
const MODS = [['fin', 'Finansije i dokumenta'], ['plan', 'Planiranje'], ['zad', 'Zadaci'], ['por', 'Poruke'], ['konk', 'Konkurencija'], ['pot', 'Potencijal 1–5'], ['ods', 'Odsustva'], ['mapa', 'Mapa i ruta'], ['tik', 'Tiketi za BizniSoft'], ['nar', 'Narudžbine']];
function modOn(m, user) {
  const s = getSet('moduli'), u = (user || me()).toLowerCase();
  const v = s[u] && m in s[u] ? s[u][m] : s['*'] && m in s['*'] ? s['*'][m] : 1;
  return !!v;
}
const isHoliday = s => getSet('praznici').includes(s);
function absentOn(s, user) { return S.X.ods.some(o => (o.f.Korisnik || '').toLowerCase() === (user || me()) && o.f.Status !== 'odbijeno' && o.f.Od <= s && s <= o.f.Do); }
function nextWorkday(s, user) { let d = addDays(s, 1); for (let i = 0; i < 30; i++) { const w = new Date(d + 'T12:00:00').getDay(); if (w !== 0 && !isHoliday(d) && !absentOn(d, user)) return d; d = addDays(d, 1); } return d; }

/* ---------- plan ---------- */
const planOfDay = (s, user) => S.X.plan.filter(p => p.f.Datum === s && (p.f.Korisnik || '').toLowerCase() === (user || me()).toLowerCase() && !['izbačen', 'prebačen'].includes(p.f.Status)).sort((a, b) => num(a.f.Redosled) - num(b.f.Redosled));
const itemBySifra = s => S.items.find(x => String(x.f.Title) === String(s));
function visitedOn(sifra, s, user) {
  return S.log.some(e => e.s === String(sifra) && /POSETE/.test(e.t) && e.v && ymd(e.v) === s && (!user || (e.k || '').toLowerCase() === user.toLowerCase()));
}
function addToPlan(its, s, user) {
  user = (user || me()).toLowerCase();
  const cur = planOfDay(s, user); let n = cur.length ? num(cur[cur.length - 1].f.Redosled) : 0, added = 0, skip = 0;
  its.forEach(it => {
    if (cur.some(p => p.f.Title === String(it.f.Title))) { skip++; return; }
    xAdd('plan', { Title: String(it.f.Title), KupacId: String(it.id), KupacNaziv: it.f.field_1 || '', Datum: s, Redosled: ++n, Korisnik: user, Status: 'planiran', Dodao: me() });
    added++;
  });
  return { added, skip };
}
async function pickDate(title, sug) {
  const t = ymd();
  const opts = [['d:' + addDays(t, 1), 'Sutra', dLabel(addDays(t, 1)) + ' · ' + DOW[new Date(addDays(t, 1) + 'T12:00:00').getDay()]], ['d:' + addDays(t, 2), 'Prekosutra', DOW[new Date(addDays(t, 2) + 'T12:00:00').getDay()]], ['d:' + t, 'Danas', ''], ['cal', '📅 Izaberi datum', '']];
  if (sug) opts.unshift(['d:' + sug, 'Sledeći radni dan', dLabel(sug)]);
  const c = await askChoice(title, opts);
  if (!c) return null;
  if (c.startsWith('d:')) return c.slice(2);
  return await askDate(title);
}
function askDate(title, val) {
  return new Promise(res => {
    const m = document.createElement('div'); m.className = 'modal';
    m.innerHTML = `<div class="box"><div class="mh"><h2>${esc(title)}</h2><button class="x" data-m="x">✕</button></div><input type="date" class="dinp" value="${val || addDays(ymd(), 1)}" min="${ymd()}"><div class="mbtn"><button class="btn" data-m="x">Odustani</button><button class="btn pri" data-m="ok">Izaberi</button></div></div>`;
    m.addEventListener('click', e => { const b = e.target.closest('[data-m]'); if (!b && e.target !== m) return; e.stopPropagation(); const v = m.querySelector('input').value; modalDone(m, () => res(b && b.dataset.m === 'ok' && v ? v : null)); });
    modalOpen(m, () => { m.remove(); res(null); });
  });
}
function askText(title, ph, val) {
  return new Promise(res => {
    const m = document.createElement('div'); m.className = 'modal';
    m.innerHTML = `<div class="box"><div class="mh"><h2>${esc(title)}</h2><button class="x" data-m="x">✕</button></div><textarea class="dinp" rows="3" placeholder="${esc(ph || '')}">${esc(val || '')}</textarea><div class="mbtn"><button class="btn" data-m="x">Odustani</button><button class="btn pri" data-m="ok">Sačuvaj</button></div></div>`;
    m.addEventListener('click', e => { const b = e.target.closest('[data-m]'); if (!b && e.target !== m) return; e.stopPropagation(); const v = m.querySelector('textarea').value; modalDone(m, () => res(b && b.dataset.m === 'ok' ? v : null)); });
    modalOpen(m, () => { m.remove(); res(null); });
  });
}
function planRowHtml(p, i, s, user) {
  const it = itemBySifra(p.f.Title), f = it ? it.f : { field_1: p.f.KupacNaziv, Title: p.f.Title };
  const done = p.f.Status === 'obiđen' || visitedOn(p.f.Title, s, user), act = S.visit && it && S.visit.id === it.id;
  return `<div class="prow ${done ? 'done' : ''}">
    <div class="pn">${done ? '✓' : i + 1}</div>
    <div class="grow" data-a="${it ? 'kupac' : ''}" data-v="${it ? esc(it.id) : ''}"><div class="t ell">${esc(f.field_1 || f.Title)}</div><div class="s ell">${esc([f.field_3, f.field_4].filter(Boolean).join(', ') || '')}</div>${p.f.Napomena ? `<div class="s2">📝 ${esc(p.f.Napomena)}</div>` : ''}${act ? '<span class="badge ok">poseta u toku</span>' : ''}${(p.f.Dodao || '').toLowerCase() !== (p.f.Korisnik || '').toLowerCase() ? `<div class="s">dodao: ${esc(nameOf(p.f.Dodao))}</div>` : ''}</div>
    <div class="pact"><button data-a="pup" data-v="${p.id}" aria-label="Gore">▲</button><button data-a="pdn" data-v="${p.id}" aria-label="Dole">▼</button><button data-a="pmenu" data-v="${p.id}" aria-label="Više">⋯</button></div>
  </div>`;
}
function planBlock(s) {
  if (!modOn('plan')) return '';
  const user = me(), P = planOfDay(s, user), done = P.filter(p => p.f.Status === 'obiđen' || visitedOn(p.f.Title, s, user)).length;
  const next = s === ymd() ? P.find(p => !(p.f.Status === 'obiđen' || visitedOn(p.f.Title, s, user))) : null, nit = next && itemBySifra(next.f.Title);
  const nl = nit && parseLoc(nit.f.Lokacija), nav = nit ? (nl ? `https://www.google.com/maps/dir/?api=1&destination=${nl.lat},${nl.lon}` : `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent([nit.f.field_3, nit.f.field_4].filter(Boolean).join(', ') + ', Srbija')}`) : '';
  return `<div class="sech"><span>Plan · ${dLabel(s)} · ${P.length ? `obiđeno ${done} od ${P.length}` : 'prazan'}</span><span>${P.length && modOn('mapa') ? '<button class="lnk" data-a="mapa">🗺 Mapa</button>' : ''}</span></div>
    ${P.length ? `<div class="pbar"><i style="width:${Math.round(done / P.length * 100)}%"></i></div>` : ''}
    ${nit ? `<div class="card next"><div class="grow"><div class="s">Sledeći kupac</div><div class="t">${esc(nit.f.field_1 || nit.f.Title)}</div></div><a class="btn pri" href="${esc(nav)}" target="_blank" rel="noopener">➤ Vozi</a></div>` : ''}
    ${P.map((p, i) => planRowHtml(p, i, s, user)).join('')}
    ${P.length && done === P.length && s === ymd() ? '<div class="card ok">🎉 Plan za danas je ispunjen. Svaka čast!</div>' : ''}
    <div class="pbtns2"><button class="btn" data-a="predlog">✨ Predlog plana</button><button class="btn" data-a="selmode">${S.selMode ? '✕ Prekini biranje' : '☑ Biraj kupce'}</button></div>`;
}
async function planMenu(id) {
  const p = S.X.plan.find(x => x.id === id); if (!p) return;
  const c = await askChoice(p.f.KupacNaziv || p.f.Title, [['nap', '📝 Napomena', p.f.Napomena || 'npr. doneti katalog'], ['mov', '📅 Prebaci na drugi dan', ''], ['del', '✕ Izbaci iz plana', '']]);
  if (c === 'nap') { const t = await askText('Napomena', 'npr. doneti katalog, pitati za fakturu', p.f.Napomena); if (t !== null) xSet('plan', p, { Napomena: t, Izmenio: me() }); }
  if (c === 'mov') { const d = await pickDate('Prebaci na dan', nextWorkday(p.f.Datum)); if (d) movePlan(p, d, 'premešteno'); }
  if (c === 'del') { if (await askSure('Izbaci iz plana', `<p>Izbaciti <b>${esc(p.f.KupacNaziv)}</b> iz plana za ${dLabel(p.f.Datum)}?</p>`, 'Izbaci')) xSet('plan', p, { Status: 'izbačen', Izmenio: me() }); }
  render();
}
function movePlan(p, d, razlog) {
  xSet('plan', p, { Status: 'prebačen', Razlog: razlog || '', Izmenio: me() });
  const it = itemBySifra(p.f.Title) || { id: p.f.KupacId, f: { Title: p.f.Title, field_1: p.f.KupacNaziv } };
  addToPlan([it], d, p.f.Korisnik);
  const np = S.X.plan[S.X.plan.length - 1]; if (np && p.f.Napomena) np.f.Napomena = p.f.Napomena;
}
function planMove(id, dir) {
  const p = S.X.plan.find(x => x.id === id); if (!p) return;
  const P = planOfDay(p.f.Datum, p.f.Korisnik), i = P.indexOf(p), j = i + dir; if (j < 0 || j >= P.length) return;
  [P[i], P[j]] = [P[j], P[i]];
  P.forEach((x, n) => { if (num(x.f.Redosled) !== n + 1) xSet('plan', x, { Redosled: n + 1, Izmenio: me() }); });
  render();
}
function markPlanVisited(it) {
  const p = planOfDay(ymd()).find(x => x.f.Title === String(it.f.Title));
  if (p && p.f.Status !== 'obiđen') xSet('plan', p, { Status: 'obiđen' });
}
async function predlogPlana() {
  const s = S.pday || ymd(), D = danasData();
  const pool = D ? D.rows.filter(r => r.d && (r.d.top || r.d.over >= -7)).map(r => r.it).concat(D.ob) : [];
  if (!pool.length) { toast('Prvo izaberi pravac (i filtere) za koji praviš plan.'); return; }
  const already = planOfDay(s).map(p => p.f.Title), cand = pool.filter(it => !already.includes(String(it.f.Title))).slice(0, getSet('predlogBroj') || 12);
  if (!cand.length) { toast('Svi kupci na redu su već u planu.'); return; }
  const r = await pickSheet(`Predlog plana · ${dLabel(s)}`, [['k', 'Kupci', cand.map(it => [it.f.field_1 + ' · ' + (it.f.field_4 || ''), klasaOf(it) || '']), cand.map(it => it.f.field_1 + ' · ' + (it.f.field_4 || ''))]]);
  if (!r) return;
  const pick = cand.filter(it => r.k.includes(it.f.field_1 + ' · ' + (it.f.field_4 || '')));
  const x = addToPlan(pick, s); toast(`Dodato u plan: ${x.added}`); render();
}

/* ---------- kraj dana: neobiđeni iz plana ---------- */
function unvisitedPlan(s) { return planOfDay(s).filter(p => p.f.Status !== 'obiđen' && !visitedOn(p.f.Title, s, me())); }
S.kd = {}; // id plana -> {r: razlog, a: 'next'|'date'|'del', d: datum, t: tekst}
function krajPlanHtml() {
  if (!modOn('plan')) return '';
  const U = unvisitedPlan(ymd()); if (!U.length) return '';
  const R = getSet('razlozi');
  return `<h2>Neobiđeni iz plana · ${U.length}</h2><p class="s">Za svakog izaberi razlog i šta dalje.</p>` + U.map(p => {
    const k = S.kd[p.id] || (S.kd[p.id] = { a: 'next', d: nextWorkday(ymd()) });
    return `<div class="card kd"><div class="t">${esc(p.f.KupacNaziv || p.f.Title)}</div>
      <div class="chipsel">${R.map(r => `<button class="${k.r === r ? 'on' : ''}" data-a="kdr" data-k="${p.id}" data-v="${esc(r)}">${esc(r)}</button>`).join('')}</div>
      ${k.r === 'ostalo' ? `<input class="dinp" data-i="kdt_${p.id}" placeholder="Upiši razlog" value="${esc(k.t || '')}">` : ''}
      <div class="seg" style="margin-top:8px"><button class="${k.a === 'next' ? 'on' : ''}" data-a="kda" data-k="${p.id}" data-v="next">${dLabel(nextWorkday(ymd()))}</button><button class="${k.a === 'date' ? 'on' : ''}" data-a="kda" data-k="${p.id}" data-v="date">${k.a === 'date' && k.d ? dLabel(k.d) : '📅 drugi dan'}</button><button class="${k.a === 'del' ? 'on' : ''}" data-a="kda" data-k="${p.id}" data-v="del">izbaci</button></div></div>`;
  }).join('');
}
function krajPlanCheck() {
  const U = unvisitedPlan(ymd());
  for (const p of U) { const k = S.kd[p.id]; if (!k || !k.r || (k.r === 'ostalo' && !String(k.t || '').trim())) return 'Za svakog neobiđenog kupca iz plana izaberi razlog.'; }
  return '';
}
async function krajPlanApply() {
  for (const p of unvisitedPlan(ymd())) {
    const k = S.kd[p.id], razlog = k.r === 'ostalo' ? 'ostalo: ' + k.t.trim() : k.r, it = itemBySifra(p.f.Title);
    if (k.a === 'del') xSet('plan', p, { Status: 'izbačen', Razlog: razlog, Izmenio: me() });
    else movePlan(p, k.a === 'date' && k.d ? k.d : nextWorkday(ymd()), razlog);
    await logEvent('PLAN NIJE OBIĐEN', null, it || { id: p.f.KupacId, f: { Title: p.f.Title, field_1: p.f.KupacNaziv } }, { Napomena: razlog + (k.a === 'del' ? ' · izbačen iz plana' : ' · prebačen na ' + (k.a === 'date' && k.d ? k.d : nextWorkday(ymd()))) });
  }
  S.kd = {};
}

/* ---------- dnevni izveštaj mejlom ---------- */
function reportHtml(s, user) {
  user = (user || me()).toLowerCase();
  const ev = S.log.filter(e => e.v && ymd(e.v) === s && (e.k || '').toLowerCase() === user);
  const st = ev.find(e => e.t === 'POČETAK DANA'), en = ev.filter(e => e.t === 'KRAJ DANA').pop();
  const vis = []; const open = {};
  ev.forEach(e => { if (e.t === 'POČETAK POSETE') { open[e.s] = { n: e.n, s: e.v, m: e.m }; vis.push(open[e.s]); } if (e.t === 'KRAJ POSETE' && open[e.s]) { open[e.s].e = e.v; open[e.s].res = e.m; } });
  const P = S.X.plan.filter(p => p.f.Datum === s && (p.f.Korisnik || '').toLowerCase() === user && p.f.Status !== 'izbačen' || (p.f.Datum === s && p.f.Status === 'prebačen' && (p.f.Korisnik || '').toLowerCase() === user));
  const nis = ev.filter(e => e.t === 'PLAN NIJE OBIĐEN'), bel = ev.filter(e => e.t === 'BELEŠKA');
  const zd = S.X.zad.filter(z => z.f.Uradjeno && ymd(z.f.Uradjeno) === s && (z.f.Uradio || '').toLowerCase() === user);
  const td = (a) => `<td style="padding:4px 8px;border-bottom:1px solid #ddd">${a}</td>`;
  return `<div style="font-family:Arial,sans-serif;font-size:14px"><h2 style="margin:0 0 6px">Dnevni izveštaj · ${esc(nameOf(user))} · ${new Date(s + 'T12:00:00').toLocaleDateString('sr-RS')}</h2>
    <p>Radni dan: <b>${st ? hhmm(st.v) : '—'}</b> – <b>${en ? hhmm(en.v) : '—'}</b> · Posete: <b>${vis.length}</b> · Plan: <b>${P.length}</b> · Neobiđeno: <b>${nis.length}</b> · Beleške: <b>${bel.length}</b> · Urađeni zadaci: <b>${zd.length}</b>${(() => { const n = ev.filter(e => e.t === 'NARUDŽBINA'); return n.length ? ` · Narudžbine: <b>${n.length}</b>` : ''; })()}</p>${(() => { const n = ev.filter(e => e.t === 'NARUDŽBINA'); return n.length ? `<h3>Narudžbine</h3><table style="border-collapse:collapse">${n.map(e => `<tr>${td(esc(e.n || ''))}${td(hhmm(e.v))}${td(esc(e.m || ''))}</tr>`).join('')}</table>` : ''; })()}
    <h3>Posete</h3><table style="border-collapse:collapse">${vis.map(v => `<tr>${td(esc(v.n))}${td(hhmm(v.s) + '–' + (v.e ? hhmm(v.e) : '?'))}${td(v.e ? Math.max(1, Math.round((Date.parse(v.e) - Date.parse(v.s)) / 60000)) + ' min' : '')}${td(esc(v.res || ''))}${td(/BEZ GPS/.test(v.m || '') ? 'bez GPS-a' : '')}</tr>`).join('') || '<tr><td>—</td></tr>'}</table>
    ${nis.length ? `<h3>Neobiđeni iz plana</h3><table style="border-collapse:collapse">${nis.map(e => `<tr>${td(esc(e.n))}${td(esc(e.m))}</tr>`).join('')}</table>` : ''}
    ${bel.length ? `<h3>Beleške</h3><table style="border-collapse:collapse">${bel.map(e => `<tr>${td(esc(e.n))}${td(esc(e.m))}</tr>`).join('')}</table>` : ''}
    ${zd.length ? `<h3>Urađeni zadaci</h3><table style="border-collapse:collapse">${zd.map(z => `<tr>${td(esc(z.f.Title))}${td(esc(z.f.KupacNaziv || ''))}${td(esc(z.f.Izvrsenje || ''))}</tr>`).join('')}</table>` : ''}
    <p style="color:#888;font-size:12px">Automatski izveštaj aplikacije Audit kupaca.</p></div>`;
}
async function sendReport(s) {
  const to = C.reportTo; if (!to || DEMO) return 'probni režim';
  try {
    const t = (await S.pca.acquireTokenSilent({ scopes: ['Mail.Send'], account: S.account })).accessToken;
    const r = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', { method: 'POST', headers: { Authorization: 'Bearer ' + t, 'Content-Type': 'application/json' }, body: JSON.stringify({ message: { subject: `Dnevni izveštaj · ${nameOf(me())} · ${new Date(s + 'T12:00:00').toLocaleDateString('sr-RS')}`, body: { contentType: 'HTML', content: reportHtml(s) }, toRecipients: [{ emailAddress: { address: to } }] }, saveToSentItems: true }) });
    if (!r.ok) throw new Error(r.status);
    return '';
  } catch (e) { console.warn('mail', e); return 'mejl nije poslat (' + String(e.message || e).slice(0, 40) + ')'; }
}

/* ---------- zadaci ---------- */
const PRI = { 1: ['primarni', 'p1'], 2: ['sekundarni', 'p2'], 3: ['tercijarni', 'p3'] };
const myTasks = () => S.X.zad.filter(z => z.f.Status !== 'urađen' && z.f.Status !== 'otkazan' && ((z.f.Za || '').toLowerCase() === me() || z.f.Za === role()));
const custTasks = sifra => S.X.zad.filter(z => z.f.KupacSifra === String(sifra) && z.f.Status !== 'otkazan');
function taskHtml(z, withCust) {
  const p = PRI[num(z.f.Prioritet) || 2], late = z.f.Rok && z.f.Status !== 'urađen' && z.f.Rok < ymd();
  return `<div class="task ${p[1]} ${z.f.Status === 'urađen' ? 'done' : ''}" data-a="task" data-v="${z.id}"><div class="pr">${p[0]}</div><div class="grow"><div class="t">${z.f.Status === 'urađen' ? '✓ ' : ''}${esc(z.f.Title)}</div>${z.f.Opis ? `<div class="s">${esc(z.f.Opis)}</div>` : ''}<div class="s">${withCust && z.f.KupacNaziv ? '🏪 ' + esc(z.f.KupacNaziv) + ' · ' : ''}od: ${esc(nameOf(z.f.Zadao))}${z.f.Za ? ' → ' + esc(nameOf(z.f.Za)) : ''}${z.f.Rok ? ` · rok ${late ? '<b class="late">' : ''}${dLabel(z.f.Rok)}${late ? '</b>' : ''}` : ''}</div>${z.f.Status === 'urađen' ? `<div class="s2">✓ ${esc(z.f.Izvrsenje || '')} · ${esc(nameOf(z.f.Uradio))}</div>` : ''}</div></div>`;
}
function tasksBlock() {
  if (!modOn('zad')) return '';
  const T = myTasks().filter(z => !z.f.KupacSifra).sort((a, b) => num(a.f.Prioritet) - num(b.f.Prioritet) || String(a.f.Rok || '9').localeCompare(String(b.f.Rok || '9')));
  const TK = myTasks().filter(z => z.f.KupacSifra);
  return `<div class="sech"><span>Moji zadaci · ${T.length + TK.length}</span><button class="lnk" data-a="newtask">＋ Zadatak</button></div>${T.map(z => taskHtml(z)).join('')}${TK.length ? `<button class="sec" data-a="tktoggle">${S.tkOpen ? '▾' : '▸'} Zadaci kod kupaca · ${TK.length}</button>${S.tkOpen ? TK.map(z => taskHtml(z, true)).join('') : ''}` : ''}${!T.length && !TK.length ? '<div class="s" style="margin:0 2px 8px">Nema otvorenih zadataka.</div>' : ''}`;
}
async function newTask(it) {
  const forOpts = isKom() ? [['me', 'Za mene', ''], ['op', 'Za call centar', 'Goca, Tamara'], ['mag', 'Za magacin', '']] : [['kom', 'Za Petra (komercijalista)', ''], ['op', 'Za call centar', ''], ['mag', 'Za magacin', '']];
  const za = await askChoice('Nov zadatak' + (it ? ' · ' + (it.f.field_1 || '') : ''), forOpts); if (!za) return;
  const S2 = getSet('sabloni'), tpl = za === 'mag' ? S2.kom2mag : (za === 'op' ? S2.kom2op : S2.op2kom);
  const r = await taskForm(tpl); if (!r) return;
  xAdd('zad', { Title: r.title, Opis: r.opis || '', Prioritet: r.pri, KupacSifra: it ? String(it.f.Title) : '', KupacNaziv: it ? it.f.field_1 || '' : '', Rok: r.rok || '', Zadao: me(), Za: za === 'me' ? me() : za, Status: 'otvoren' });
  toast('Zadatak je zadat.'); render();
}
function taskForm(tpl) {
  return new Promise(res => {
    const st = { title: '', pri: 2, rok: '', opis: '' };
    const m = document.createElement('div'); m.className = 'modal sheet';
    const paint = () => m.innerHTML = `<div class="box"><div class="mh"><h2>Zadatak</h2><button class="x" data-m="x">✕</button></div><div class="sbody">
      <div class="sh">Šta (izaberi ili upiši)</div><div class="chipsel">${tpl.map(t => `<button class="${st.title === t ? 'on' : ''}" data-t="${esc(t)}">${esc(t)}</button>`).join('')}</div>
      <input class="dinp" data-f="title" placeholder="ili upiši zadatak" value="${esc(st.title)}">
      <div class="sh">Prioritet</div><div class="seg">${[1, 2, 3].map(p => `<button class="${st.pri === p ? 'on' : ''}" data-p="${p}">${PRI[p][0]}</button>`).join('')}</div>
      <div class="sh">Rok (nije obavezan)</div><input type="date" class="dinp" data-f="rok" value="${st.rok}">
      <div class="sh">Napomena</div><input class="dinp" data-f="opis" value="${esc(st.opis)}"></div>
      <div class="mbtn"><button class="btn" data-m="x">Odustani</button><button class="btn pri" data-m="ok">Zadaj</button></div></div>`;
    paint();
    const sync = () => m.querySelectorAll('[data-f]').forEach(i => st[i.dataset.f] = i.value);
    m.addEventListener('click', e => {
      e.stopPropagation(); const t = e.target.closest('[data-t]'), p = e.target.closest('[data-p]'), b = e.target.closest('[data-m]');
      if (t) { sync(); st.title = t.dataset.t; paint(); return; }
      if (p) { sync(); st.pri = +p.dataset.p; paint(); return; }
      if (b && b.dataset.m === 'ok') { sync(); if (!st.title.trim()) { m.querySelector('[data-f=title]').focus(); return; } modalDone(m, () => res({ ...st, title: st.title.trim() })); return; }
      if (b || e.target === m) { modalDone(m, () => res(null)); }
    });
    modalOpen(m, () => { m.remove(); res(null); });
  });
}
async function openTask(id) {
  const z = S.X.zad.find(x => x.id === id); if (!z) return;
  if (z.f.Status === 'urađen') { if (await askSure('Zadatak je urađen', `<p><b>${esc(z.f.Title)}</b></p><p>${esc(z.f.Izvrsenje || '')}</p><p class="s">${esc(nameOf(z.f.Uradio))} · ${fmtDate(z.f.Uradjeno)}</p>`, 'Vrati kao otvoren')) { xSet('zad', z, { Status: 'otvoren' }); render(); } return; }
  const ans = getSet('sabloni').odgovori;
  const opts = ans.map(a => ['a:' + a, a, '']).concat([['txt', '✍ Upiši šta je urađeno', ''], ['msg', '💬 Napiši poruku o ovome', ''], ...(z.f.Zadao === me() || isAdmin() ? [['cancel', '✕ Otkaži zadatak', '']] : [])]);
  const c = await askChoice('✓ ' + z.f.Title, opts); if (!c) return;
  if (c === 'cancel') { xSet('zad', z, { Status: 'otkazan' }); render(); return; }
  if (c === 'msg') { openThread(z.f.KupacSifra || '', z.f.Title); return; }
  let txt = c.startsWith('a:') ? c.slice(2) : await askText('Šta je urađeno', '', ''); if (txt === null) return;
  xSet('zad', z, { Status: 'urađen', Izvrsenje: txt, Uradio: me(), Uradjeno: new Date().toISOString() });
  if (z.f.KupacSifra) { const it = itemBySifra(z.f.KupacSifra); await logEvent('ZADATAK', null, it || null, { Napomena: z.f.Title + ' → ' + txt }); }
  toast('Zadatak je označen kao urađen.'); render();
}

/* ---------- poruke ---------- */
const forMe = m => { const z = (m.f.Za || '').toLowerCase(); return z === me() || z === role() || (z === 'svi'); };
const unread = m => (m.f.Od || '').toLowerCase() !== me() && forMe(m) && !String(m.f.Procitali || '').toLowerCase().split(',').includes(me());
const unreadCount = () => modOn('por') ? S.X.por.filter(unread).length + (isKom() ? 0 : 0) : 0;
function threads() {
  const T = {};
  S.X.por.forEach(m => { const k = m.f.KupacSifra || ''; (T[k] = T[k] || { k, n: m.f.KupacNaziv || 'Opšte poruke', msgs: [] }).msgs.push(m); });
  return Object.values(T).map(t => { t.msgs.sort((a, b) => String(a.f.Vreme).localeCompare(String(b.f.Vreme))); t.last = t.msgs[t.msgs.length - 1]; t.un = t.msgs.filter(unread).length; return t; }).sort((a, b) => String(b.last.f.Vreme).localeCompare(String(a.last.f.Vreme)));
}
function vPoruke() {
  const T = threads(), q = norm(S.pq || '');
  const L = T.filter(t => !q || norm(t.n).includes(q));
  return bar('Poruke') + `<main class="main">
    <div class="fsel srch"><span class="ic">🔍</span><input type="search" data-i="pq" placeholder="Pretraži razgovore po kupcu" value="${esc(S.pq || '')}"></div>
    <button class="btn" style="width:100%;margin-bottom:10px" data-a="thread" data-v="">💬 Opšte poruke${(T.find(t => t.k === '') || {}).un ? ` <span class="bdg">${T.find(t => t.k === '').un}</span>` : ''}</button>
    ${L.filter(t => t.k).map(t => `<div class="row" data-a="thread" data-v="${esc(t.k)}"><div class="grow"><div class="t ell">${esc(t.n)}</div><div class="s ell">${esc(nameOf(t.last.f.Od))}: ${esc(t.last.f.Title)}</div></div><div class="right"><div class="s">${dLabel(ymd(t.last.f.Vreme))}</div>${t.un ? `<span class="bdg">${t.un}</span>` : ''}</div></div>`).join('') || '<div class="empty">Još nema razgovora o kupcima. Poruku o kupcu šalješ sa njegove kartice (💬).</div>'}
  </main>` + tabs('poruke');
}
function openThread(k, prefill) { S.thr = k; S.thrPre = prefill || ''; S.from2 = S.view; go('thread'); }
function vThread() {
  const k = S.thr || '', it = k ? itemBySifra(k) : null, T = S.X.por.filter(m => (m.f.KupacSifra || '') === k).sort((a, b) => String(a.f.Vreme).localeCompare(String(b.f.Vreme)));
  T.filter(unread).forEach(m => xSet('por', m, { Procitali: [String(m.f.Procitali || ''), me()].filter(Boolean).join(',') }));
  const S2 = getSet('sabloni'), tpl = isKom() ? (S.toWho === 'mag' ? S2.kom2mag : S2.kom2op) : S2.op2kom;
  const who = S.toWho || (isKom() ? 'op' : 'kom');
  const tos = isKom() ? [['op', 'Call centar'], ['mag', 'Magacin'], ['admin', 'Admin']] : [['kom', 'Petar'], ['op', 'Call centar'], ['mag', 'Magacin']];
  return bar(it ? it.f.field_1 || k : (k ? 'Razgovor · ' + k : 'Opšte poruke'), true) + `<main class="main chat">
    ${it ? `<button class="lnk" data-a="kupac" data-v="${esc(it.id)}">› Otvori karticu kupca</button>` : ''}
    ${T.map(m => { const mine = (m.f.Od || '').toLowerCase() === me(); return `<div class="msg ${mine ? 'me' : ''}"><div class="mt">${esc(m.f.Title)}</div>${m.f.Tekst && m.f.Tekst !== m.f.Title ? `<div class="mx">${esc(m.f.Tekst)}</div>` : ''}<div class="mm">${esc(nameOf(m.f.Od))} → ${esc(nameOf(m.f.Za))} · ${dLabel(ymd(m.f.Vreme))} ${hhmm(m.f.Vreme)}${m.f.ZadatakId ? ' · ✓ postao zadatak' : ''}${!mine && !m.f.ZadatakId ? ` · <button class="lnk" data-a="msg2task" data-v="${m.id}">→ zadatak za mene</button>` : ''}</div></div>`; }).join('') || '<div class="empty">Još nema poruka. Izaberi gotovu poruku ispod ili upiši svoju.</div>'}
  </main><div class="foot chatfoot"><div class="in" style="flex-direction:column;align-items:stretch">
    <div class="seg">${tos.map(([v, n]) => `<button class="${who === v ? 'on' : ''}" data-a="towho" data-v="${v}">${n}</button>`).join('')}</div>
    <div class="chipsel tpl">${tpl.concat(S2.odgovori).map(t => `<button data-a="sendtpl" data-v="${esc(t)}">${esc(t)}</button>`).join('')}</div>
    <div style="display:flex;gap:8px"><input class="dinp" style="margin:0;flex:1" data-i="msgtxt" placeholder="Upiši poruku…" value="${esc(S.msgTxt || S.thrPre || '')}"><button class="btn pri" data-a="sendmsg">Pošalji</button></div>
  </div></div>`;
}
function sendMsg(text) {
  text = String(text || '').trim(); if (!text) return;
  const k = S.thr || '', it = k ? itemBySifra(k) : null, who = S.toWho || (isKom() ? 'op' : 'kom');
  xAdd('por', { Title: text.slice(0, 250), Tekst: text, Od: me(), Za: who, KupacSifra: k, KupacNaziv: it ? it.f.field_1 || '' : (k ? k : ''), Vreme: new Date().toISOString(), Procitali: me() });
  S.msgTxt = ''; S.thrPre = ''; render(); scrollTo(0, document.body.scrollHeight);
}
function msgToTask(id) {
  const m = S.X.por.find(x => x.id === id); if (!m) return;
  const z = xAdd('zad', { Title: m.f.Title, Opis: m.f.Tekst !== m.f.Title ? m.f.Tekst || '' : '', Prioritet: 2, KupacSifra: m.f.KupacSifra || '', KupacNaziv: m.f.KupacNaziv || '', Zadao: m.f.Od, Za: me(), Status: 'otvoren' });
  xSet('por', m, { ZadatakId: 1 }); toast('Poruka je postala tvoj zadatak.'); render();
}

/* ---------- tiketi za BizniSoft ---------- */
const TIK_ST = { novo: ['Novo', 'warn'], preuzet: ['Preuzet', 'brand'], resen: ['Rešen', 'ok'] };
const tikRows = t => { try { const x = JSON.parse(t.f.Izmene || '[]'); return Array.isArray(x) ? x : []; } catch (e) { return []; } };
const tikNew = () => modOn('tik') ? S.X.tik.filter(t => t.f.Status === 'novo').length : 0;
const daySet = s => new Set((Array.isArray(s) ? s : String(s || '').split(',')).map(x => norm(x)).filter(Boolean));
const sameSet = (a, b) => a.size === b.size && [...a].every(x => b.has(x));
const telN = s => String(s || '').replace(/\D/g, '').slice(-8);
function tikDiff(f, d) {
  const r = [], N = norm, d0 = d._d0 || null;
  if (d.StatusNovi && !N(f.field_11).split(/\s+/).includes(N(d.StatusNovi))) r.push(['Status', f.field_11 || '', d.StatusNovi]);
  if (d.AuditSegment && !(f.field_14 && (N(d.AuditSegment).includes(N(f.field_14)) || N(f.field_14).includes(N(d.AuditSegment))))) r.push(['Segment', f.field_14 || '', d.AuditSegment]);
  if (d.AuditTeren && N(d.AuditTeren) !== N(f.field_9)) r.push(['Pravac', f.field_9 || '', d.AuditTeren]);
  if (d.AuditRazvoz && d.AuditRazvoz.length && !sameSet(daySet(d.AuditRazvoz), daySet(f.field_15))) r.push(['Dani razvoza', f.field_15 || '', d.AuditRazvoz.join(', ')]);
  if (d.AuditPozivanje && d.AuditPozivanje.length && !sameSet(daySet(d.AuditPozivanje), daySet(f.field_16))) r.push(['Dani poziva', f.field_16 || '', d.AuditPozivanje.join(', ')]);
  if (d.NarucujeIme && N(d.NarucujeIme) !== N(f.field_7)) r.push(['Kontakt osoba', f.field_7 || '', d.NarucujeIme]);
  if (d.NarucujeTel && telN(d.NarucujeTel) !== telN(f.field_5)) r.push(['Telefon', f.field_5 || '', d.NarucujeTel]);
  if (!d0) return r;
  const K = { 'Status': 'StatusNovi', 'Segment': 'AuditSegment', 'Pravac': 'AuditTeren', 'Dani razvoza': 'AuditRazvoz', 'Dani poziva': 'AuditPozivanje', 'Kontakt osoba': 'NarucujeIme', 'Telefon': 'NarucujeTel' };
  const same = (a, b) => JSON.stringify(Array.isArray(a) ? a.slice().sort() : a || '') === JSON.stringify(Array.isArray(b) ? b.slice().sort() : b || '');
  // BizniSoft nema vrednost, a Petar je nije menjao (aplikacija ju je sama predložila): nije izmena za BizniSoft
  return r.filter(x => String(x[1] || '').trim() || !same(d[K[x[0]]], d0[K[x[0]]]));
}
function tikFromAudit(it, d) {
  if (!modOn('tik')) return;
  const f = it.f, k = String(f.Title), old = S.X.tik.filter(t => t.f.KupacSifra === k);
  if (/^NOVO-/.test(k)) {
    if (old.some(t => t.f.Vrsta === 'novi')) return;
    const rows = [['Naziv', '', f.field_1], ['Adresa', '', f.field_3], ['Mesto', '', f.field_4], ['Telefon', '', d.NarucujeTel || f.field_5], ['Kontakt osoba', '', d.NarucujeIme || f.field_7], ['Segment', '', d.AuditSegment], ['Pravac', '', d.AuditTeren], ['Dani razvoza', '', (d.AuditRazvoz || []).join(', ')], ['Dani poziva', '', (d.AuditPozivanje || []).join(', ')], ['Način plaćanja', '', d.AuditNacinPlacanja], ['Lokacija', '', d.Lokacija || f.Lokacija]].filter(x => x[2]);
    xAdd('tik', { Title: 'Nov kupac: ' + (f.field_1 || k), Vrsta: 'novi', KupacSifra: k, KupacNaziv: f.field_1 || '', Izmene: JSON.stringify(rows), Autor: me(), Vreme: new Date().toISOString(), Status: 'novo' });
    return;
  }
  const done = new Set(old.flatMap(t => tikRows(t).map(x => x[0] + '|' + norm(x[2]))));
  const rows = tikDiff(f, d).filter(x => !done.has(x[0] + '|' + norm(x[2])));
  if (!rows.length) return;
  xAdd('tik', { Title: 'Izmena: ' + rows.map(x => x[0]).join(', '), Vrsta: 'izmena', KupacSifra: k, KupacNaziv: f.field_1 || '', Izmene: JSON.stringify(rows), Autor: me(), Vreme: new Date().toISOString(), Status: 'novo' });
}
function tikCheck(t) {
  if (t.f.Vrsta === 'narudzba') return t.f.Status === 'resen' ? ordCheck(t) : '';
  if (t.f.Status !== 'resen' || !S.SY.kupci) return '';
  if (t.f.Vrsta === 'novi') { const s = t.f.BSSifra; if (!s) return ''; const ok = S.SY.kupci.kupci.some(x => String(x.s) === String(Number(s))); return ok ? '<span class="badge ok">✓ vidi se u BS</span>' : '<span class="badge bad">⚠ nije vidljivo u BS</span>'; }
  const it = itemBySifra(t.f.KupacSifra); if (!it) return '';
  let b = null;
  if (it.f.Centrala) { const p = (S.SY.kupci.pj || []).find(x => String(x.k) === String(it.f.Centrala) && String(x.s) === String(it.f.PJ)); if (p) b = { kon: p.kon, tel: p.tel }; } else b = bsKupac(it);
  if (!b) return '';
  const chk = tikRows(t).filter(x => x[0] === 'Kontakt osoba' || x[0] === 'Telefon');
  if (!chk.length) return '';
  const ok = chk.every(x => x[0] === 'Telefon' ? telN(b.tel) === telN(x[2]) : norm(b.kon) === norm(x[2]));
  return ok ? '<span class="badge ok">✓ vidi se u BS</span>' : '<span class="badge bad">⚠ nije vidljivo u BS</span>';
}
function vTiketi() {
  const fl = S.tikF || 'otvoreni', mine = isKom();
  let L = S.X.tik.slice().sort((a, b) => String(b.f.Vreme || '').localeCompare(String(a.f.Vreme || '')));
  if (mine) L = L.filter(t => (t.f.Autor || '').toLowerCase() === me());
  const cnt = s => L.filter(t => s === 'otvoreni' ? t.f.Status !== 'resen' : t.f.Status === s).length;
  const V = L.filter(t => fl === 'svi' ? true : fl === 'otvoreni' ? t.f.Status !== 'resen' : t.f.Status === fl);
  return bar(mine ? 'Moji tiketi' : 'Tiketi za BizniSoft', mine) + `<main class="main">
    <div class="seg">${[['otvoreni', 'Otvoreni'], ['novo', 'Novi'], ['preuzet', 'Preuzeti'], ['resen', 'Rešeni'], ['svi', 'Svi']].map(([k, n]) => `<button class="${fl === k ? 'on' : ''}" data-a="tikf" data-v="${k}">${n}${k !== 'svi' ? ' · ' + cnt(k) : ''}</button>`).join('')}</div>
    ${V.map(t => { const s = TIK_ST[t.f.Status] || ['?', '']; return `<div class="row" data-a="tiket" data-v="${esc(t.id)}"><div class="grow"><div class="t ell">${esc(t.f.KupacNaziv || t.f.KupacSifra)}</div><div class="s ell">${esc(t.f.Title)}</div><div class="s">${esc(nameOf(t.f.Autor))} · ${dLabel(ymd(t.f.Vreme))} ${hhmm(t.f.Vreme)}${t.f.Preuzeo && t.f.Status === 'preuzet' ? ' · preuzeo ' + esc(nameOf(t.f.Preuzeo)) : ''}</div></div><div class="right"><span class="badge ${s[1]}">${s[0]}</span>${tikCheck(t)}</div></div>`; }).join('') || '<div class="empty">Nema tiketa u ovom prikazu.</div>'}
  </main>` + (mine ? '' : tabs('tiketi'));
}
function vTiket() {
  const t = S.X.tik.find(x => x.id === S.tikId); if (!t) return bar('Tiket', true) + '<main class="main"><div class="empty">Tiket nije pronađen.</div></main>';
  const f = t.f, s = TIK_ST[f.Status] || ['?', ''], it = itemBySifra(f.KupacSifra), mine = (f.Preuzeo || '').toLowerCase() === me(), can = !isKom() || isAdmin();
  const wait = (a, b) => { if (!a || !b) return ''; const m = Math.round((Date.parse(b) - Date.parse(a)) / 60000); return m < 60 ? m + ' min' : m < 1440 ? Math.round(m / 60) + ' h' : Math.round(m / 1440) + ' d'; };
  return bar(f.Vrsta === 'novi' ? 'Tiket · nov kupac' : f.Vrsta === 'narudzba' ? 'Tiket · narudžbina' : 'Tiket · izmena', true) + `<main class="main">
    <div class="t">${esc(f.KupacNaziv || f.KupacSifra)}</div><div class="s">šifra u aplikaciji: ${esc(f.KupacSifra)}${it && it.f.Centrala ? ` · poslovna jedinica ${esc(it.f.PJ)} centrale ${esc(it.f.Centrala)}` : ''}</div>
    ${it ? `<button class="lnk" data-a="kupac" data-v="${esc(it.id)}">› Otvori karticu kupca</button>` : ''}
    ${f.Vrsta === 'narudzba' ? ordTikHtml(t) : `<div class="card"><table class="tbl"><tr><th>Polje</th><th>U BizniSoft-u</th><th>Treba da bude</th></tr>${tikRows(t).map(r => `<tr><td><b>${esc(r[0])}</b></td><td class="s">${esc(r[1] || '—')}</td><td>${esc(r[2] || '')}</td></tr>`).join('')}</table></div>`}
    <div class="card"><dl class="kv"><dt>Status</dt><dd><span class="badge ${s[1]}">${s[0]}</span> ${tikCheck(t)}</dd>
      <dt>Napravio</dt><dd>${esc(nameOf(f.Autor))} · ${fmtDate(f.Vreme)} ${hhmm(f.Vreme)}</dd>
      ${f.Preuzeo ? `<dt>Preuzeo</dt><dd>${esc(nameOf(f.Preuzeo))} · ${fmtDate(f.Preuzeto)} ${hhmm(f.Preuzeto)} <span class="s">(čekao ${wait(f.Vreme, f.Preuzeto)})</span></dd>` : ''}
      ${f.Resio ? `<dt>Rešio</dt><dd>${esc(nameOf(f.Resio))} · ${fmtDate(f.Reseno)} ${hhmm(f.Reseno)} <span class="s">(rešavanje ${wait(f.Preuzeto || f.Vreme, f.Reseno)})</span></dd>` : ''}
      ${f.BSSifra ? `<dt>${f.Vrsta === 'narudzba' ? 'Broj NK' : 'Šifra u BS'}</dt><dd>${esc(f.BSSifra)}</dd>` : ''}
      ${f.Komentar ? `<dt>Komentar</dt><dd>${esc(f.Komentar)}</dd>` : ''}</dl></div>
    <button class="btn" style="width:100%" data-a="tikaut">💬 Piši autoru (${esc(nameOf(f.Autor))})</button>
  </main>${can && f.Status !== 'resen' ? `<div class="foot"><div class="in">${f.Status === 'novo' ? '<button class="btn pri" data-a="tikpre">Preuzmi</button>' : `${mine ? '<button class="btn" data-a="tikvrati">Vrati</button>' : `<span class="s grow">Preuzeo ${esc(nameOf(f.Preuzeo))}</span>`}<button class="btn pri" style="flex:2" data-a="tikres">Rešeno</button>`}</div></div>` : ''}`;
}
async function tikAct(a) {
  const t = S.X.tik.find(x => x.id === S.tikId); if (!t) return;
  const now = new Date().toISOString();
  if (a === 'pre') { if (t.f.Status !== 'novo') return; xSet('tik', t, { Status: 'preuzet', Preuzeo: me(), Preuzeto: now }); toast('Tiket je tvoj.'); }
  if (a === 'vrati') { if (!await askSure('Vrati tiket', '<p>Tiket se vraća u „Novo“ i neko drugi može da ga preuzme.</p>', 'Vrati')) return; xSet('tik', t, { Status: 'novo', Preuzeo: '', Preuzeto: null }); }
  if (a === 'res') {
    const fx = { Status: 'resen', Resio: me(), Reseno: now };
    if (t.f.Status === 'novo') { fx.Preuzeo = me(); fx.Preuzeto = now; }
    if (t.f.Vrsta === 'narudzba') {
      const nk = String(await askText('Broj narudžbine (NK) u BizniSoft-u', 'npr. NK400009600', t.f.BSSifra || '') || '').trim().toUpperCase();
      if (!nk) return;
      fx.BSSifra = /^\d+$/.test(nk) ? 'NK400' + nk.padStart(6, '0') : nk;
    }
    if (t.f.Vrsta === 'novi') {
      const s = String(await askText('Šifra kupca u BizniSoft-u', 'npr. 3721', t.f.BSSifra || '') || '').trim();
      if (!s) return;
      if (!/^\d+$/.test(s)) { toast('Šifra treba da bude broj.'); return; }
      fx.BSSifra = s;
    }
    const c = await askText('Komentar (nije obavezno)', 'npr. uneto, promenjen i drugi telefon', t.f.Komentar || '');
    if (c === null) return;
    if (String(c).trim()) fx.Komentar = String(c).trim();
    if (fx.BSSifra && t.f.Vrsta === 'novi') {
      const s = fx.BSSifra, it = itemBySifra(t.f.KupacSifra), dup = S.items.find(x => String(x.f.Title) === s);
      if (it && !dup) { it.f.Title = s; saveItems(); await queue({ type: 'patch', itemId: it.id, fields: { Title: s } }); fx.KupacSifra = s; }
      else if (dup) toast('Kupac ' + s + ' već postoji u aplikaciji, pa NOVO zapis nije preimenovan.');
    }
    xSet('tik', t, fx); toast('Tiket je rešen.');
  }
  render();
}

/* ---------- narudžbine ---------- */
const din = n => (Math.round(num(n) * 100) / 100).toLocaleString('sr-RS', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const qfmt = n => (Math.round(num(n) * 1000) / 1000).toLocaleString('sr-RS');
const LS_ORD = 'ak_ord_v1';
function artIdx() {
  const A = S.SY.artikli; if (!A) return null;
  if (S.SY._ai && S.SY._ai.gen === A.gen) return S.SY._ai;
  const cen = {}, aat = {};
  (A.cen || []).forEach(r => { ((cen[r.c] = cen[r.c] || {})[r.a] = cen[r.c][r.a] || []).push(r); });
  Object.entries(A.artatr || {}).forEach(([t, L]) => (L || []).forEach(s => { (aat[s] = aat[s] || new Set()).add(Number(t)); }));
  const km = {}; Object.values(A.art).forEach(a => { const k = (a.kat || '').trim(); if (k && !/^izba/i.test(k) && !km[k.toUpperCase()]) km[k.toUpperCase()] = k; });
  const kats = Object.values(km).sort((a, b) => a.localeCompare(b, 'sr'));
  return S.SY._ai = { gen: A.gen, cen, aat, kats };
}
const dateOk = (r, d) => (!r.od || String(r.od).slice(0, 10) <= d) && (!r.dd || String(r.dd).slice(0, 10) >= d);
function priceFor(kup, pj, sifra, tip) {
  const A = S.SY.artikli, X = artIdx(); if (!A || !X) return null;
  const d = ymd(), k = Number(kup), s = Number(sifra), ki = (A.kup || {})[String(k)] || {};
  const pick = n => { const R = ((X.cen[n] || {})[s] || []).filter(r => dateOk(r, d) && (!r.p || r.p === k) && (!r.pj || r.pj === Number(pj))); R.sort((a, b) => ((b.p ? 2 : 0) + (b.pj ? 1 : 0)) - ((a.p ? 2 : 0) + (a.pj ? 1 : 0))); return R.length ? R[0].v : 0; };
  const base = pick(ki.c || 'CENOVNIK') || pick('CENOVNIK');
  const katr = new Set((A.kupatr || {})[String(k)] || []), aatr = X.aat[s] || new Set();
  const ak = (A.akc || []).filter(x => dateOk(x, d) && (x.a === s || (!x.a && x.aa && aatr.has(x.aa))) && (!x.k || x.k === k) && (!x.ka || katr.has(x.ka)));
  const fx = ak.filter(x => x.c > 0).sort((a, b) => (!!b.k - !!a.k) || (!!b.a - !!a.a) || (a.pr - b.pr));
  if (fx.length) return { base, cena: fx[0].c, rab: 0, neto: fx[0].c, akc: fx[0].n || 'akcija' };
  const ar = ak.filter(x => x.r > 0 && !(x.c > 0)).sort((a, b) => (!!b.k - !!a.k) || (!!b.a - !!a.a));
  let rab = 0, izvor = '';
  if (ar.length) { rab = ar[0].r; izvor = ar[0].n || 'akcija'; }
  else {
    const U = (A.usl || []).filter(u => dateOk(u, d) && (u.k || u.ka) && (!u.k || u.k === k) && (!u.ka || katr.has(u.ka)) && (!u.a || u.a === s) && (!u.aa || aatr.has(u.aa)) && (!u.d || u.d === tip));
    U.sort((a, b) => ((b.k ? 8 : 0) + (b.a ? 4 : b.aa ? 2 : 0) + (b.d ? 1 : 0)) - ((a.k ? 8 : 0) + (a.a ? 4 : a.aa ? 2 : 0) + (a.d ? 1 : 0)));
    if (U.length) { rab = U[0].r; izvor = U[0].k ? 'lični rabat' : 'rabat kategorije'; }
    else if (ki.r) { rab = ki.r; izvor = 'rabat kupca'; }
  }
  return { base, cena: base, rab, neto: base * (1 - rab / 100), akc: ar.length ? izvor : '', izvor };
}
function ordKup(it) { return { k: Number(bsKey(it)), pj: it.f.Centrala ? num(it.f.PJ) : 0 }; }
function ordLoad() { try { return JSON.parse(localStorage.getItem(LS_ORD) || '{}'); } catch (e) { return {}; } }
function ordSave() { const all = ordLoad(); if (S.ord) { if (Object.keys(S.ord.q).length || S.ord.tip) all[S.ord.id] = S.ord; else delete all[S.ord.id]; } try { localStorage.setItem(LS_ORD, JSON.stringify(all)); } catch (e) { } }
function openOrder(it) {
  if (!S.SY.artikli) { toast('Artikli još nisu stigli iz BizniSoft-a. Osveži podatke kad bude mreže.'); return; }
  if (!bsKey(it)) { toast('Kupac još nema BizniSoft šifru. Narudžbina se pravi kada bude unet u BizniSoft.'); return; }
  S.ord = ordLoad()[it.id] || { id: it.id, tip: '', q: {}, s: '', kat: '', tab: 'obicno' };
  S.cur = it; go('narudzba');
}
function ordUsual(it) {
  const st = S.SY.stavke, L = st && st.k[bsKey(it)]; if (!L) return [];
  return L.map(x => ({ s: String(x[0]), avg: x[4] ? x[1] / x[4] : x[1], z: x[3] })).filter(x => S.SY.artikli.art[x.s]);
}
function ordStep(a) { return a && a.kp > 0 ? a.kp : 1; }
function ordRow(it, s, hint) {
  const A = S.SY.artikli, a = A.art[s]; if (!a) return '';
  const o = S.ord, k = ordKup(it), p = priceFor(k.k, k.pj, s, o.tip) || {}, q = num(o.q[s]);
  const z = S.SY.zalihe && S.SY.zalihe.k[s], fr = z ? num(z.f) : null;
  const pak = [a.kp > 0 ? `kom. pak. ${qfmt(a.kp)}` : '', a.tp > 0 ? `trans. pak. ${qfmt(a.tp)}` : ''].filter(Boolean).join(' · ');
  const odd = q && a.kp > 1 && Math.abs(q / a.kp - Math.round(q / a.kp)) > 1e-6;
  return `<div class="orow${q ? ' on' : ''}"><div class="grow">
    <div class="t">${esc(a.n)}</div>
    <div class="s">${esc(s)}${a.kat ? ' · ' + esc(a.kat) : ''} · ${esc(a.jm || '')}${pak ? ' · ' + pak : ''}</div>
    <div class="s">${p.neto > 0 ? (p.akc && p.rab === 0 ? `<span class="badge brand">${esc(p.akc)}</span> <b>${din(p.neto)}</b>${p.base > p.neto ? ` <s>${din(p.base)}</s>` : ''}` : `${din(p.cena)}${p.rab ? ` −${qfmt(p.rab)}% = <b>${din(p.neto)}</b>` : ''}${p.akc ? ` <span class="badge brand">${esc(p.akc)}</span>` : ''}`) : '<span class="late">nema cene</span>'}
      · ${fr === null ? 'stanje —' : `<span class="${fr <= 0 || (q && q > fr) ? 'late' : ''}">slobodno ${qfmt(fr)}</span>`}${hint ? ' · ' + hint : ''}</div>
    ${odd ? `<div class="s late">nije ceo broj kom. pakovanja (${qfmt(a.kp)})</div>` : ''}
  </div><div class="qty"><button data-a="oq" data-v="${esc(s)}" data-k="-1">−</button><input type="number" inputmode="decimal" data-i="oq_${esc(s)}" value="${q || ''}" placeholder="0"><button data-a="oq" data-v="${esc(s)}" data-k="1">+</button>${a.tp > 0 ? `<button class="tp" data-a="oq" data-v="${esc(s)}" data-k="tp">+TP</button>` : ''}</div></div>`;
}
function ordTotal(it) {
  const k = ordKup(it); let n = 0, sum = 0;
  Object.entries(S.ord.q).forEach(([s, q]) => { if (num(q) > 0) { n++; const p = priceFor(k.k, k.pj, s, S.ord.tip); if (p) sum += p.neto * num(q); } });
  return { n, sum };
}
function ordWarn(it) {
  const f = finOf(it), b = bsKupac(it), w = [];
  if (b && String(b.blok) === '1') w.push('Kupac je <b>blokiran</b> u BizniSoft-u.');
  if (f && f.d > 0) w.push(`Kupac ima dospeli dug <b>${money(f.d)}</b>${f.m ? `, kasni najviše ${f.m} dana` : ''}.`);
  return w;
}
function vNarudzba() {
  const it = S.cur, o = S.ord; if (!it || !o) return vTereni();
  const A = S.SY.artikli, X = artIdx(), k = ordKup(it), w = ordWarn(it);
  if (!o.tip) {
    const r = t => { const p = priceFor(k.k, k.pj, 0, t); return p ? p.rab : 0; };
    return bar('Narudžbina', true) + `<main class="main">
      <div class="t">${esc(it.f.field_1 || it.f.Title)}</div>${it.f.Centrala ? `<div class="s">poslovna jedinica · cene i rabat centrale ${esc(it.f.Centrala)}</div>` : ''}
      ${w.length ? `<div class="card warnc">${w.join('<br>')}</div>` : ''}
      <div class="sh" style="margin-top:12px">Kako kupac plaća ovu narudžbinu?</div>
      <button class="big-act" style="width:100%;margin-bottom:10px" data-a="otip" data-v="OA"><span>📄</span>Na valutu (otpremnica) · rabat ${qfmt(r('OA'))}%</button>
      <button class="big-act" style="width:100%" data-a="otip" data-v="PF"><span>💵</span>Plaća odmah (profaktura) · rabat ${qfmt(r('PF'))}%</button>
      <p class="s">Rabat je iz uslova u BizniSoft-u. Na artiklima sa akcijskom cenom nema rabata.</p>
    </main>`;
  }
  const us = ordUsual(it), qn = norm(o.s || '');
  let list;
  if (o.tab === 'korpa') list = Object.keys(o.q).filter(s => num(o.q[s]) > 0).map(s => ordRow(it, s));
  else if (o.tab === 'obicno' && !qn && !o.kat) list = us.slice(0, 80).map(x => ordRow(it, x.s, `obično ${qfmt(Math.round(x.avg * 10) / 10)}`));
  else {
    const L = Object.entries(A.art).filter(([s, a]) => !/^izba/i.test((a.kat || '').trim()) && (!o.kat || (a.kat || '').trim().toUpperCase() === o.kat.toUpperCase()) && (!qn || norm(a.n + ' ' + s + ' ' + (a.kat || '')).includes(qn)));
    list = L.slice(0, 60).map(([s]) => ordRow(it, s)); if (L.length > 60) list.push(`<div class="s" style="margin:8px 2px">Prikazano 60 od ${L.length}. Suzi pretragu.</div>`);
  }
  const t = ordTotal(it), kc = Object.keys(o.q).filter(s => num(o.q[s]) > 0).length;
  const zg = S.SY.zalihe ? S.SY.zalihe.gen : null;
  return bar('Narudžbina · ' + (o.tip === 'PF' ? 'plaća odmah' : 'na valutu'), true) + `<main class="main">
    <div class="t ell">${esc(it.f.field_1 || it.f.Title)}</div>
    ${w.length ? `<div class="card warnc">${w.join('<br>')}</div>` : ''}
    <div class="seg">${[['obicno', `Obično · ${us.length}`], ['sve', 'Svi artikli'], ['korpa', `Korpa · ${kc}`]].map(([v, n]) => `<button class="${o.tab === v ? 'on' : ''}" data-a="otab" data-v="${v}">${n}</button>`).join('')}</div>
    ${o.tab !== 'korpa' ? `<div class="fsel srch"><span class="ic">🔍</span><input type="search" data-i="ordq" placeholder="Naziv, šifra ili katalog" value="${esc(o.s || '')}"></div>
    <select class="dinp" data-i="ordkat"><option value="">Svi katalozi</option>${X.kats.map(c => `<option ${o.kat === c ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>` : ''}
    <div id="olist">${list.join('') || '<div class="empty">Nema artikala za prikaz.</div>'}</div>
    <p class="s">Cene su bez PDV-a. Konačnu cenu računa BizniSoft pri unosu.${zg ? ` Stanje magacina 400 od ${fmtDate(zg)} ${hhmm(zg)}.` : ''} <button class="lnk" data-a="otipch">Promeni način plaćanja</button></p>
  </main><div class="foot"><div class="in"><div class="grow"><b>${din(t.sum)}</b><div class="s">${t.n} artikala · bez PDV</div></div><button class="btn" data-a="oclear">Obriši</button><button class="btn pri" data-a="osend" ${t.n ? '' : 'disabled'}>Pošalji</button></div></div>`;
}
function ordSetQ(s, v) {
  v = Math.max(0, Math.round(num(v) * 1000) / 1000);
  if (v) S.ord.q[s] = v; else delete S.ord.q[s];
  ordSave();
}
function ordRefresh() {
  const y = scrollY, f = document.activeElement && document.activeElement.dataset && document.activeElement.dataset.i;
  render(); scrollTo(0, y);
  if (f) { const i = document.querySelector(`[data-i="${f}"]`); if (i) { i.focus(); try { i.setSelectionRange(i.value.length, i.value.length); } catch (e) { } } }
}
async function ordSend() {
  const it = S.cur, o = S.ord, k = ordKup(it), A = S.SY.artikli;
  const L = Object.entries(o.q).filter(([, q]) => num(q) > 0).map(([s, q]) => { const p = priceFor(k.k, k.pj, s, o.tip) || {}; return [s, (A.art[s] || {}).n || s, num(q), Math.round((p.cena || 0) * 100) / 100, p.rab || 0, Math.round((p.neto || 0) * num(q) * 100) / 100, p.akc || '']; });
  if (!L.length) return;
  const sum = L.reduce((a, x) => a + x[5], 0), w = ordWarn(it);
  const low = L.filter(x => { const z = S.SY.zalihe && S.SY.zalihe.k[x[0]]; return z && x[2] > num(z.f); });
  if (!await askSure('Pošalji narudžbinu', `<p><b>${esc(it.f.field_1 || it.f.Title)}</b><br>${o.tip === 'PF' ? 'plaća odmah (PF)' : 'na valutu (OA)'} · ${L.length} artikala · <b>${din(sum)} din</b> bez PDV</p>${w.length ? `<p class="late">${w.join('<br>')}</p>` : ''}${low.length ? `<p class="late">Nema dovoljno na stanju: ${low.map(x => esc(x[1])).join(', ')}</p>` : ''}<p>Narudžbina ide operateru kao tiket i unosi se u BizniSoft.</p>`, 'Pošalji')) return;
  const nap = await askText('Napomena uz narudžbinu (nije obavezno)', 'npr. isporuka u petak, pozvati pre dostave', '');
  if (nap === null) return;
  xAdd('tik', { Title: `Narudžbina · ${L.length} art · ${din(sum)} din`, Vrsta: 'narudzba', KupacSifra: String(it.f.Title), KupacNaziv: it.f.field_1 || '', Izmene: JSON.stringify({ tip: o.tip, kupac: k.k, pj: k.pj, l: L, iznos: Math.round(sum * 100) / 100, nap: String(nap || '').trim() }), Autor: me(), Vreme: new Date().toISOString(), Status: 'novo' });
  await logEvent('NARUDŽBINA', null, it, { Napomena: `${o.tip} · ${L.length} art · ${din(sum)} din${nap ? ' · ' + String(nap).trim() : ''}` });
  S.ord.q = {}; S.ord.tip = ''; ordSave(); S.ord = null;
  toast('Narudžbina je poslata.'); history.back();
}
const ordOf = t => { try { const x = JSON.parse(t.f.Izmene || '{}'); return x && x.l ? x : null; } catch (e) { return null; } };
function ordTikHtml(t) {
  const x = ordOf(t); if (!x) return '';
  return `<div class="card"><div class="s">${x.tip === 'PF' ? 'Plaća odmah · profaktura (PF)' : 'Na valutu · otpremnica (OA)'} · šifra kupca u BS: <b>${esc(x.kupac)}</b>${x.pj ? ` · poslovna jedinica <b>${esc(x.pj)}</b>` : ''}</div>
    <table class="tbl"><tr><th>Artikal</th><th>Kol.</th><th>Cena</th><th>Rab.</th><th>Iznos</th></tr>${x.l.map(r => `<tr><td>${esc(r[1])}<div class="s">${esc(r[0])}${r[6] ? ' · ' + esc(r[6]) : ''}</div></td><td>${qfmt(r[2])}</td><td>${din(r[3])}</td><td>${r[4] ? qfmt(r[4]) + '%' : ''}</td><td>${din(r[5])}</td></tr>`).join('')}
    <tr><td colspan="4"><b>Ukupno bez PDV</b></td><td><b>${din(x.iznos)}</b></td></tr></table>${x.nap ? `<p><b>Napomena:</b> ${esc(x.nap)}</p>` : ''}</div>`;
}
function ordCheck(t) {
  const x = ordOf(t), nk = String(t.f.BSSifra || '').trim().toUpperCase(); if (!x || !nk || !S.SY.dokumenta) return '';
  const D = (S.SY.dokumenta.k[String(x.kupac)] || []).map(d => d && d.value && !Array.isArray(d) ? d.value : d);
  const d = D.find(d => String(d[1]).toUpperCase() === nk || String(d[1]).toUpperCase().endsWith(nk));
  return d ? `<span class="badge ok">✓ ${esc(d[1])} u BS · ${din(d[3])}</span>` : '<span class="badge bad">⚠ NK nije vidljiva u BS</span>';
}

/* ---------- konkurencija ---------- */
function konkForm(it) {
  return new Promise(res => {
    const st = { k: new Set(), g: new Set(), c: '', u: '', n: '', o: '' };
    const KO = getSet('konkurenti');
    const m = document.createElement('div'); m.className = 'modal sheet';
    const paint = () => m.innerHTML = `<div class="box"><div class="mh"><h2>Konkurencija · ${esc(it.f.field_1 || '')}</h2><button class="x" data-m="x">✕</button></div><div class="sbody">
      <div class="sh">Čiju robu drži</div><div class="chipsel">${KO.map(k => `<button class="${st.k.has(k) ? 'on' : ''}" data-k="${esc(k)}">${esc(k)}</button>`).join('')}</div>
      <input class="dinp" data-f="o" placeholder="ostali (upiši naziv)" value="${esc(st.o)}">
      <div class="sh">Za koje grupe proizvoda</div><div class="chipsel">${ASORT.map(a => `<button class="${st.g.has(a) ? 'on' : ''}" data-g="${esc(a)}">${esc(a)}</button>`).join('')}</div>
      <div class="sh">Cene u odnosu na nas</div><div class="seg">${['jeftiniji', 'slično', 'skuplji'].map(x => `<button class="${st.c === x ? 'on' : ''}" data-c="${x}">${x}</button>`).join('')}</div>
      <div class="sh">Uslovi (rabat, valuta)</div><div class="seg">${['bolji', 'slično', 'lošiji'].map(x => `<button class="${st.u === x ? 'on' : ''}" data-u="${x}">${x}</button>`).join('')}</div>
      <div class="sh">Napomena</div><input class="dinp" data-f="n" value="${esc(st.n)}"></div>
      <div class="mbtn"><button class="btn" data-m="x">Odustani</button><button class="btn pri" data-m="ok">Sačuvaj</button></div></div>`;
    paint();
    const sync = () => m.querySelectorAll('[data-f]').forEach(i => st[i.dataset.f] = i.value);
    m.addEventListener('click', e => {
      e.stopPropagation(); const b = e.target.closest('button'); if (!b && e.target !== m) return;
      if (b && b.dataset.k) { sync(); st.k.has(b.dataset.k) ? st.k.delete(b.dataset.k) : st.k.add(b.dataset.k); paint(); return; }
      if (b && b.dataset.g) { sync(); st.g.has(b.dataset.g) ? st.g.delete(b.dataset.g) : st.g.add(b.dataset.g); paint(); return; }
      if (b && b.dataset.c) { sync(); st.c = st.c === b.dataset.c ? '' : b.dataset.c; paint(); return; }
      if (b && b.dataset.u) { sync(); st.u = st.u === b.dataset.u ? '' : b.dataset.u; paint(); return; }
      if (b && b.dataset.m === 'ok') { sync(); const ks = [...st.k].concat(st.o.split(',').map(x => x.trim()).filter(Boolean).map(x => 'ostali: ' + x)); if (!ks.length) { m.querySelector('[data-f=o]').focus(); return; } modalDone(m, () => res({ ks, gs: [...st.g], c: st.c, u: st.u, n: st.n })); return; }
      if (!b || b.dataset.m) { modalDone(m, () => res(null)); }
    });
    modalOpen(m, () => { m.remove(); res(null); });
  });
}
async function addKonk(it) {
  const r = await konkForm(it); if (!r) return;
  xAdd('konk', { Title: String(it.f.Title), KupacSifra: String(it.f.Title), KupacNaziv: it.f.field_1 || '', Konkurenti: r.ks.join(', '), Grupe: r.gs.join(', '), Cene: r.c, Uslovi: r.u, Napomena: r.n, Korisnik: me(), Vreme: new Date().toISOString() });
  await logEvent('KONKURENCIJA', null, it, { Napomena: r.ks.join(', ') + (r.gs.length ? ' · ' + r.gs.join(', ') : '') });
  toast('Konkurencija je zabeležena.'); render();
}

/* ---------- potencijal 1–5 ---------- */
async function setPot(it, v) {
  const old = num(it.f.PotencijalOcena); if (old === v) return;
  it.f.PotencijalOcena = v; saveItems();
  await queue({ type: 'patch', itemId: it.id, fields: { PotencijalOcena: v } });
  await logEvent('POTENCIJAL', null, it, { Napomena: (old ? old + ' → ' : '') + v });
  render();
}

/* ---------- odsustva ---------- */
async function newAbsence() {
  const V = getSet('vrsteOdsustva');
  const c = await askChoice('Prijavi odsustvo', V.map(v => [v, v, ''])); if (!c) return;
  const od = await askDate('Od kog dana', ymd()); if (!od) return;
  const doo = await askDate('Do kog dana (uključujući)', od); if (!doo) return;
  xAdd('ods', { Title: c, Korisnik: me(), Od: od, Do: doo < od ? od : doo, Vrsta: c, Status: 'prijavljeno' });
  toast('Odsustvo je prijavljeno.'); render();
}

/* ---------- desktop: nedeljni plan ---------- */
function weekStart(s) { const x = new Date(s + 'T12:00:00'), w = (x.getDay() + 6) % 7; x.setDate(x.getDate() - w); return ymd(x); }
function vNedelja() {
  const ws = S.week || weekStart(ymd()), user = S.planUser || komUsers()[0] || me();
  const days = [0, 1, 2, 3, 4, 5].map(i => addDays(ws, i));
  const D = danasData(), pool = D ? D.rows.map(r => r.it).concat(D.ob) : [];
  const col = s => { const P = planOfDay(s, user), hol = isHoliday(s), ab = absentOn(s, user); return `<div class="wcol ${s === ymd() ? 'today' : ''}" data-day="${s}"><div class="wh">${DOW[new Date(s + 'T12:00:00').getDay()]} ${new Date(s + 'T12:00:00').toLocaleDateString('sr-Latn-RS', { day: 'numeric', month: 'numeric' })}${hol ? ' · praznik' : ''}${ab ? ' · odsutan' : ''}<span class="s"> · ${P.length}</span></div>${P.map((p, i) => { const done = p.f.Status === 'obiđen' || visitedOn(p.f.Title, s, user); return `<div class="wit ${done ? 'done' : ''}" draggable="true" data-pid="${p.id}"><b>${i + 1}.</b> ${esc(p.f.KupacNaziv || p.f.Title)}${p.f.Napomena ? `<div class="s">📝 ${esc(p.f.Napomena)}</div>` : ''}<button class="wx" data-a="pmenu" data-v="${p.id}">⋯</button></div>`; }).join('')}<div class="wdrop">prevuci kupca ovde</div></div>`; };
  return bar('Plan po nedeljama') + `<main class="main wide">
    <div class="wtop"><button class="btn" data-a="wprev">‹</button><b>${new Date(ws + 'T12:00:00').toLocaleDateString('sr-RS')} – ${new Date(addDays(ws, 5) + 'T12:00:00').toLocaleDateString('sr-RS')}</b><button class="btn" data-a="wnext">›</button><button class="btn" data-a="wtoday">Ova nedelja</button>
      ${komUsers().length > 1 ? `<select data-i="planuser">${komUsers().map(u => `<option value="${u}" ${u === user ? 'selected' : ''}>${esc(nameOf(u))}</option>`).join('')}</select>` : `<span class="s">Komercijalista: ${esc(nameOf(user))}</span>`}</div>
    <div class="wgrid">
      <div class="wside"><div class="fsel srch"><span class="ic">🔍</span><input type="search" data-i="dq" placeholder="Pretraži kupce" value="${esc(S.dq)}"></div>${MAINF.map(fBtn).join('')}
        <div class="s" style="margin:6px 2px">${pool.length ? `Prevuci kupca u dan · ${pool.length}` : 'Izaberi pravac ili pretraži kupce.'}</div>
        <div class="wpool">${pool.slice(0, 150).map(it => `<div class="wit" draggable="true" data-kid="${esc(it.id)}">${esc(it.f.field_1 || it.f.Title)}<div class="s">${esc(it.f.field_4 || '')} · ${klasaOf(it) || '–'}${(() => { const d = dueOf(it); return d ? d.never ? ' · nikad' : d.over > 0 ? ' · kasni ' + d.over + ' d' : '' : ''; })()}</div></div>`).join('')}</div></div>
      <div class="wdays">${days.map(col).join('')}</div>
      ${modOn('mapa') ? `<div class="wmap"><div class="s">Mapa dana: ${days.map(s => `<button class="lnk ${S.mapDay === s ? 'on' : ''}" data-a="mapday" data-v="${s}">${DOW[new Date(s + 'T12:00:00').getDay()].slice(0, 3)}</button>`).join(' ')}</div><div id="map" class="mapbox"></div></div>` : ''}
    </div></main>` + tabs('nedelja');
}
function wireWeek() {
  const user = S.planUser || komUsers()[0] || me();
  document.querySelectorAll('.wit[draggable]').forEach(el => el.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', JSON.stringify({ pid: el.dataset.pid, kid: el.dataset.kid })); }));
  document.querySelectorAll('.wcol').forEach(col => {
    col.addEventListener('dragover', e => { e.preventDefault(); col.classList.add('over'); });
    col.addEventListener('dragleave', () => col.classList.remove('over'));
    col.addEventListener('drop', e => {
      e.preventDefault(); col.classList.remove('over');
      let d; try { d = JSON.parse(e.dataTransfer.getData('text/plain')); } catch (x) { return; }
      const day = col.dataset.day, items = [...col.querySelectorAll('.wit')], before = items.find(x => e.clientY < x.getBoundingClientRect().top + x.offsetHeight / 2);
      if (d.kid) { const it = S.items.find(x => String(x.id) === d.kid); if (it) addToPlan([it], day, user); }
      else if (d.pid) {
        const p = S.X.plan.find(x => x.id === d.pid); if (!p) return;
        if (p.f.Datum !== day) { xSet('plan', p, { Datum: day, Redosled: 999, Izmenio: me() }); }
      }
      // redosled po mestu spuštanja
      const P = planOfDay(day, user); const moved = d.pid ? P.find(x => x.id === d.pid) : P[P.length - 1];
      if (moved && before) { const rest = P.filter(x => x !== moved), bi = rest.findIndex(x => x.id === before.dataset.pid); rest.splice(bi < 0 ? rest.length : bi, 0, moved); rest.forEach((x, n) => { if (num(x.f.Redosled) !== n + 1) xSet('plan', x, { Redosled: n + 1, Izmenio: me() }); }); }
      else P.forEach((x, n) => { if (num(x.f.Redosled) !== n + 1) xSet('plan', x, { Redosled: n + 1, Izmenio: me() }); });
      render();
    });
  });
  if (document.getElementById('map')) showMap(S.mapDay || ymd(), user);
}

/* ---------- mapa ---------- */
function loadLeaflet() {
  if (window.L) return Promise.resolve();
  return new Promise((res, rej) => {
    const c = document.createElement('link'); c.rel = 'stylesheet'; c.href = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css'; document.head.appendChild(c);
    const s = document.createElement('script'); s.src = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'; s.onload = res; s.onerror = () => rej(new Error('Mapa nije dostupna (nema mreže).')); document.head.appendChild(s);
  });
}
function routeUrl(pts) {
  if (!pts.length) return '';
  const f = p => p.lat + ',' + p.lon, last = pts[pts.length - 1], mid = pts.slice(0, -1).slice(0, 9);
  return `https://www.google.com/maps/dir/?api=1&destination=${f(last)}${mid.length ? '&waypoints=' + encodeURIComponent(mid.map(f).join('|')) : ''}&travelmode=driving`;
}
async function showMap(s, user) {
  const el = document.getElementById('map'); if (!el) return;
  try { await loadLeaflet(); } catch (e) { el.innerHTML = `<div class="empty">${esc(e.message)}</div>`; return; }
  const P = planOfDay(s, user || me()), pts = [];
  P.forEach((p, i) => { const it = itemBySifra(p.f.Title), l = it && parseLoc(it.f.Lokacija); if (l) pts.push({ ...l, n: i + 1, name: p.f.KupacNaziv, done: p.f.Status === 'obiđen' || visitedOn(p.f.Title, s, user) }); });
  if (S.lmap) { S.lmap.remove(); S.lmap = null; }
  const map = S.lmap = L.map(el).setView(pts.length ? [pts[0].lat, pts[0].lon] : [43.58, 21.33], pts.length ? 11 : 8);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '© OpenStreetMap' }).addTo(map);
  pts.forEach(p => L.marker([p.lat, p.lon], { icon: L.divIcon({ className: 'mnum' + (p.done ? ' done' : ''), html: String(p.n), iconSize: [26, 26] }) }).addTo(map).bindPopup(esc(p.n + '. ' + p.name)));
  if (pts.length > 1) { L.polyline(pts.map(p => [p.lat, p.lon]), { weight: 3, opacity: .7 }).addTo(map); map.fitBounds(pts.map(p => [p.lat, p.lon]), { padding: [30, 30] }); }
  const miss = P.length - pts.length, info = document.getElementById('mapinfo');
  if (info) info.innerHTML = `${pts.length ? `<a class="btn pri" href="${esc(routeUrl(pts.filter(p => !p.done)))}" target="_blank" rel="noopener">➤ Otvori rutu u Google mapama</a>` : ''}${miss ? `<p class="s">${miss} kupaca iz plana nema zabeleženu lokaciju i nisu na mapi.</p>` : ''}`;
}
function vMapa() {
  return bar('Mapa · plan ' + dLabel(S.pday || ymd()), true) + `<main class="main"><div id="map" class="mapbox tall"></div><div id="mapinfo" style="margin-top:10px"></div></main>`;
}
// postavljanje lokacije kupca čiodom
function vPin() {
  const it = S.cur;
  return bar('Lokacija na mapi', true) + `<main class="main"><p class="s">${esc(it.f.field_1 || '')} · ${esc([it.f.field_3, it.f.field_4].filter(Boolean).join(', '))}</p><p class="s">Pomeri čiodu na ulaz objekta. Lokacija sa mape važi dok je prva poseta ne potvrdi ili ispravi.</p><div id="map" class="mapbox tall"></div></main><div class="foot"><div class="in"><button class="btn pri" data-a="pinsave">Sačuvaj lokaciju</button></div></div>`;
}
async function wirePin() {
  const el = document.getElementById('map'); if (!el) return;
  try { await loadLeaflet(); } catch (e) { el.innerHTML = `<div class="empty">${esc(e.message)}</div>`; return; }
  const it = S.cur, l = parseLoc(it.f.Lokacija);
  let c = l ? [l.lat, l.lon] : null;
  if (!c) { const t = await townPos(it.f.field_4); c = t ? [t.lat, t.lon] : [43.58, 21.33]; }
  if (S.lmap) { S.lmap.remove(); S.lmap = null; }
  const map = S.lmap = L.map(el).setView(c, l ? 17 : 14);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '© OpenStreetMap' }).addTo(map);
  S.pin = L.marker(c, { draggable: true }).addTo(map);
  map.on('click', e => S.pin.setLatLng(e.latlng));
}
async function pinSave() {
  const it = S.cur, ll = S.pin.getLatLng(), old = it.f.Lokacija;
  if (!await askSure('Sačuvaj lokaciju', `<p>Lokacija objekta <b>${esc(it.f.field_1 || '')}</b> biće postavljena na mesto čiode.</p>`, 'Da, sačuvaj')) return;
  it.f.Lokacija = `${ll.lat.toFixed(6)}, ${ll.lng.toFixed(6)} (±50 m) mapa`; saveItems();
  await queue({ type: 'patch', itemId: it.id, fields: { Lokacija: it.f.Lokacija } });
  await logEvent('LOKACIJA SA MAPE', { lat: ll.lat, lon: ll.lng, acc: 50 }, it, { Napomena: 'stara: ' + (old || '—') });
  toast('Lokacija je sačuvana.'); history.back();
}

/* ---------- aktivnost (call centar, admin) ---------- */
function vAktivnost() {
  const ev = S.log.slice().reverse().filter(e => e.t !== 'POTENCIJAL').slice(0, 150);
  const op = S.X.zad.filter(z => z.f.Status !== 'urađen' && z.f.Status !== 'otkazan' && ((z.f.Za || '').toLowerCase() === me() || z.f.Za === role()));
  let lastD = '';
  return bar('Aktivnost') + `<main class="main wide">
    ${op.length ? `<div class="sech"><span>Zadaci za mene · ${op.length}</span></div>${op.map(z => taskHtml(z, true)).join('')}` : ''}
    <div class="sech"><span>Šta se dešava na terenu</span><button class="lnk" data-a="newtask">＋ Zadatak za Petra</button></div>
    ${ev.map(e => { const d = e.v ? ymd(e.v) : ''; const h = d !== lastD ? `<div class="dsep">${dLabel(d)}</div>` : ''; lastD = d; const it = e.s && itemBySifra(e.s); return h + `<div class="hrow" ${it ? `data-a="kupac" data-v="${esc(it.id)}"` : ''}><div class="t">${esc(e.t.toLowerCase())}${e.n ? ' · ' + esc(e.n) : ''}</div><div class="s">${hhmm(e.v)} · ${esc(nameOf(e.k))}${e.u ? ' · ' + e.u + ' m' : ''}</div>${e.m ? `<div class="s2">${esc(e.m)}</div>` : ''}</div>`; }).join('') || '<div class="empty">Još nema aktivnosti.</div>'}
  </main>` + tabs('aktivnost');
}

/* ---------- nadzor (admin) ---------- */
function vNadzor() {
  const s = S.nzDay || ymd(), user = S.planUser || komUsers()[0] || '';
  const ev = S.log.filter(e => e.v && ymd(e.v) === s && (e.k || '').toLowerCase() === user);
  const st = ev.find(e => e.t === 'POČETAK DANA'), en = ev.filter(e => e.t === 'KRAJ DANA').pop();
  const P = S.X.plan.filter(p => p.f.Datum === s && (p.f.Korisnik || '').toLowerCase() === user && p.f.Status !== 'izbačen');
  const pd = P.filter(p => p.f.Status === 'obiđen' || visitedOn(p.f.Title, s, user)).length, pAll = P.filter(p => p.f.Status !== 'prebačen').length + P.filter(p => p.f.Status === 'prebačen').length;
  const vis = ev.filter(e => e.t === 'POČETAK POSETE').length, nog = ev.filter(e => /BEZ GPS/.test(e.m || '')).length;
  const zOpen = S.X.zad.filter(z => z.f.Status === 'otvoren' && ((z.f.Za || '').toLowerCase() === user || z.f.Za === 'kom'));
  const zLate = zOpen.filter(z => z.f.Rok && z.f.Rok < ymd()), zP1 = zOpen.filter(z => num(z.f.Prioritet) === 1);
  // 7 dana
  const wk = [...Array(7)].map((_, i) => addDays(s, -6 + i)).map(d => { const pp = S.X.plan.filter(p => p.f.Datum === d && (p.f.Korisnik || '').toLowerCase() === user && p.f.Status !== 'izbačen'); const ok = pp.filter(p => p.f.Status === 'obiđen' || visitedOn(p.f.Title, d, user)).length; const vv = S.log.filter(e => e.v && ymd(e.v) === d && (e.k || '').toLowerCase() === user && e.t === 'POČETAK POSETE').length; return { d, pp: pp.length, ok, vv }; });
  const ods = S.X.ods.filter(o => o.f.Status === 'prijavljeno');
  return bar('Nadzor') + `<main class="main wide">
    <div class="wtop"><button class="btn" data-a="nzprev">‹</button><b>${new Date(s + 'T12:00:00').toLocaleDateString('sr-Latn-RS', { weekday: 'long', day: 'numeric', month: 'numeric', year: 'numeric' })}</b><button class="btn" data-a="nznext">›</button><button class="btn" data-a="nztoday">Danas</button><span class="s">${esc(nameOf(user))}</span><button class="btn" data-a="nzmail">✉ Izveštaj</button></div>
    ${ods.length ? `<div class="card warnc"><b>Prijavljena odsustva za odobrenje:</b>${ods.map(o => `<div style="display:flex;gap:8px;align-items:center;margin-top:6px"><span class="grow">${esc(nameOf(o.f.Korisnik))} · ${esc(o.f.Vrsta)} · ${fmtDate(o.f.Od)}–${fmtDate(o.f.Do)}</span><button class="btn" data-a="odsok" data-v="${o.id}">Odobri</button><button class="btn" data-a="odsno" data-v="${o.id}">Odbij</button></div>`).join('')}</div>` : ''}
    <div class="sum4"><div class="kpi"><span class="s">Radni dan</span><b>${st ? hhmm(st.v) : '—'}–${en ? hhmm(en.v) : (st ? '…' : '—')}</b></div><div class="kpi"><span class="s">Posete</span><b>${vis}</b></div><div class="kpi"><span class="s">Plan</span><b>${pd}/${pAll}</b></div><div class="kpi"><span class="s">Bez GPS-a</span><b>${nog}</b></div></div>
    <div class="sum4"><div class="kpi"><span class="s">Otvoreni zadaci</span><b>${zOpen.length}</b></div><div class="kpi"><span class="s">Primarni</span><b>${zP1.length}</b></div><div class="kpi"><span class="s">Kasne</span><b class="${zLate.length ? 'late' : ''}">${zLate.length}</b></div><div class="kpi"><span class="s">Beleške</span><b>${ev.filter(e => e.t === 'BELEŠKA').length}</b></div></div>
    ${st && !en && s < ymd() ? '<div class="card warnc">⚠ Dan nije završen.</div>' : ''}
    <h2>Poslednjih 7 dana</h2><div class="card"><table class="tbl"><tr><th>Dan</th><th>Posete</th><th>Plan</th><th>Poštovanje plana</th></tr>${wk.map(w => `<tr><td>${dLabel(w.d)}</td><td>${w.vv}</td><td>${w.ok}/${w.pp}</td><td>${w.pp ? Math.round(w.ok / w.pp * 100) + '%' : '—'}</td></tr>`).join('')}</table></div>
    <h2>Izveštaj dana</h2><div class="card rep">${reportHtml(s, user)}</div>
  </main>` + tabs('nadzor');
}

/* ---------- admin ---------- */
function vAdmin() {
  const mods = getSet('moduli'), users = ['*', ...Object.keys(ROLES)], sab = getSet('sabloni');
  const ta = (k, v, rows = 5) => `<textarea class="dinp" rows="${rows}" data-set="${k}">${esc(v)}</textarea>`;
  const counts = {}; S.X.konk.forEach(k => String(k.f.Konkurenti || '').split(',').map(x => x.trim()).filter(x => x.startsWith('ostali:')).forEach(x => counts[x.slice(7).trim()] = (counts[x.slice(7).trim()] || 0) + 1));
  return bar('Admin') + `<main class="main wide">
    <h2>Moduli po korisniku</h2><p class="s">Kvačica = ekran je uključen. „Svi” važi za one koji nemaju svoje podešavanje.</p>
    <div class="card" style="overflow:auto"><table class="tbl"><tr><th>Korisnik</th>${MODS.map(([, n]) => `<th>${n}</th>`).join('')}</tr>${users.map(u => `<tr><td>${u === '*' ? '<b>Svi</b>' : esc(nameOf(u)) + `<div class="s">${esc(ROLES[u] || '')}</div>`}</td>${MODS.map(([m]) => { const on = u === '*' ? !!(mods['*'] || {})[m] : modOn(m, u); return `<td><input type="checkbox" data-mod="${esc(u)}|${m}" ${on ? 'checked' : ''}></td>`; }).join('')}</tr>`).join('')}</table></div>
    <h2>Sinhronizacija sa BizniSoft-om</h2>
    ${(() => { const sy = getSet('sinhronizacija'), st = S.SY.stanje; const N = { kupci: 'Kupci i jedinice', promet: 'Promet', dokumenta: 'Dokumenta', finansije: 'Finansije', stavke: 'Stavke (šta kupuje)', artikli: 'Artikli, cene, akcije, rabati', zalihe: 'Stanje magacina 400' };
      return `<div class="card" style="overflow:auto"><table class="tbl"><tr><th>Skup</th><th>Na koliko minuta</th><th>Poslednje osvežavanje</th><th>Rezultat</th></tr>${Object.keys(N).map(k => { const x = st && st.skupovi && st.skupovi[k]; return `<tr><td>${N[k]}</td><td><input class="dinp" style="width:90px;margin:0" type="number" min="0" data-sync="${k}" value="${esc(sy[k])}"></td><td>${x ? fmtDate(x.last) + ' ' + hhmm(x.last) : '—'}</td><td class="${x && !x.ok ? 'late' : ''}">${x ? esc(x.poruka || '') : '—'}</td></tr>`; }).join('')}</table>
      <p class="s">0 = isključeno · 1440 = jednom noću. Radno vreme za češće osvežavanje: od <input class="dinp" style="width:60px;display:inline-block;margin:0" type="number" data-sync="od" value="${esc(sy.od)}"> do <input class="dinp" style="width:60px;display:inline-block;margin:0" type="number" data-sync="do" value="${esc(sy.do)}"> h, noćno osvežavanje u <input class="dinp" style="width:60px;display:inline-block;margin:0" type="number" data-sync="nocu" value="${esc(sy.nocu)}"> h. Program na serveru čita ova podešavanja na svakih 5 minuta.${st ? ' Server: ' + esc(st.server || '') + ', poslednji rad ' + fmtDate(st.gen) + ' ' + hhmm(st.gen) : ''}</p></div>`; })()}
    <h2>Poslovne jedinice</h2>${(() => { const K = S.SY.kupci, n = S.items.filter(isPJ).length, m = pjMissing().length; return `<div class="card">U BizniSoft-u: <b>${K ? K.pj.length : '—'}</b> jedinica · u aplikaciji: <b>${n}</b>${m ? ` · nove za dodavanje: <b>${m}</b> <button class="btn" data-a="pjimport">Dodaj nove</button>` : ''}<p class="s">Dodaju se samo jedinice kupaca koji su u aplikaciji. Svaka jedinica ima svoju ličnu kartu, plan i posete; finansije ostaju na centrali.</p></div>`; })()}
    <h2>Šabloni poruka i zadataka</h2><p class="s">Jedan šablon po redu.</p>
    <div class="agrid"><div><div class="sh">Call centar → Petar</div>${ta('sab.op2kom', sab.op2kom.join('\n'))}</div><div><div class="sh">Petar → call centar</div>${ta('sab.kom2op', sab.kom2op.join('\n'))}</div><div><div class="sh">Petar → magacin</div>${ta('sab.kom2mag', sab.kom2mag.join('\n'))}</div><div><div class="sh">Odgovori dodirom</div>${ta('sab.odgovori', sab.odgovori.join('\n'))}</div></div>
    <div class="agrid"><div><div class="sh">Konkurenti</div>${ta('konkurenti', getSet('konkurenti').join('\n'))}${Object.keys(counts).length ? `<div class="s">Ručno upisani: ${Object.entries(counts).sort((a, b) => b[1] - a[1]).map(([k, n]) => esc(k) + ' (' + n + ')').join(', ')}</div>` : ''}</div><div><div class="sh">Razlozi za neobiđene</div>${ta('razlozi', getSet('razlozi').join('\n'))}</div><div><div class="sh">Praznici (GGGG-MM-DD)</div>${ta('praznici', getSet('praznici').join('\n'), 8)}</div><div><div class="sh">Vrste odsustva</div>${ta('vrsteOdsustva', getSet('vrsteOdsustva').join('\n'))}</div></div>
    <div class="agrid"><div><div class="sh">Podsetnik za kraj dana (sat)</div><input class="dinp" data-set="podsetnik" value="${esc(getSet('podsetnik'))}"></div><div><div class="sh">Predlog plana: broj kupaca</div><input class="dinp" type="number" data-set="predlogBroj" value="${esc(getSet('predlogBroj'))}"></div><div><div class="sh">KPI ciljevi (za kasnije)</div>${ta('kpi', JSON.stringify(getSet('kpi'), null, 1), 4)}</div><div><div class="sh">Pragovi potencijala 1–5 (za kasnije)</div>${ta('pragoviPotencijala', JSON.stringify(getSet('pragoviPotencijala'), null, 1), 4)}</div></div>
    <button class="btn pri" style="width:100%;margin-top:12px" data-a="adminsave">Sačuvaj podešavanja</button>
  </main>` + tabs('admin');
}
function adminSave() {
  const mods = getSet('moduli');
  document.querySelectorAll('[data-mod]').forEach(c => { const [u, m] = c.dataset.mod.split('|'); mods[u] = mods[u] || {}; mods[u][m] = c.checked ? 1 : 0; });
  putSet('moduli', mods);
  const sy = getSet('sinhronizacija'); document.querySelectorAll('[data-sync]').forEach(i => sy[i.dataset.sync] = Math.max(0, num(i.value))); putSet('sinhronizacija', sy);
  const lines = k => document.querySelector(`[data-set="${k}"]`).value.split('\n').map(x => x.trim()).filter(Boolean);
  const sab = { op2kom: lines('sab.op2kom'), kom2op: lines('sab.kom2op'), kom2mag: lines('sab.kom2mag'), odgovori: lines('sab.odgovori') };
  putSet('sabloni', sab); putSet('konkurenti', lines('konkurenti')); putSet('razlozi', lines('razlozi')); putSet('praznici', lines('praznici')); putSet('vrsteOdsustva', lines('vrsteOdsustva'));
  putSet('podsetnik', document.querySelector('[data-set="podsetnik"]').value.trim() || '17:30');
  putSet('predlogBroj', num(document.querySelector('[data-set="predlogBroj"]').value) || 12);
  try { putSet('kpi', JSON.parse(document.querySelector('[data-set="kpi"]').value)); putSet('pragoviPotencijala', JSON.parse(document.querySelector('[data-set="pragoviPotencijala"]').value)); } catch (e) { toast('KPI ili pragovi nisu ispravno napisani; ostalo je sačuvano.'); render(); return; }
  toast('Podešavanja su sačuvana.'); render();
}

/* ---------- podsetnik i jučerašnji dan ---------- */
function reminderDue() {
  if (!isKom() || !dayOn()) return false;
  const [h, m] = String(getSet('podsetnik') || '17:30').split(':').map(Number), n = new Date();
  return n.getHours() * 60 + n.getMinutes() >= h * 60 + (m || 0);
}
const dayOpenOld = () => isKom() && S.day && S.day.date !== new Date().toLocaleDateString('sr-RS') && !S.day.end && S.day.start;
function vJuce() {
  const R = ['zaboravio sam', 'telefon se ugasio', 'nije bilo mreže', 'ostalo'];
  return `<div class="login"><h1>Juče nije završen radni dan</h1><p class="s">Dan je počeo ${fmtDate(S.day.start)} u ${hhmm(S.day.start)}. Pre novog dana zatvori jučerašnji: izaberi razlog.</p>
    <div class="opts">${R.map(r => `<button class="opt ${S.jr === r ? 'on' : ''}" data-a="jr" data-v="${esc(r)}">${esc(r)}</button>`).join('')}</div>
    <button class="btn pri" style="margin-top:14px" data-a="juceok" ${S.jr ? '' : 'disabled'}>Zatvori jučerašnji dan</button></div>`;
}
async function closeYesterday() {
  const old = ymd(S.day.start), when = new Date().toISOString();
  await logEvent('KRAJ DANA', null, null, { Napomena: 'dan zatvoren naknadno (' + S.jr + '), počeo ' + fmtDate(S.day.start) });
  // neobiđeni iz jučerašnjeg plana prelaze na danas
  planOfDay(old).filter(p => p.f.Status !== 'obiđen' && !visitedOn(p.f.Title, old, me())).forEach(p => movePlan(p, ymd(), 'dan nije završen'));
  S.day.end = when; S.jr = ''; saveDay(); render();
}

/* ---------- osvežavanje u pozadini (poruke, zadaci, plan) ---------- */
async function pollX() {
  if (DEMO || !navigator.onLine || document.hidden || !S.account) return;
  const before = unreadCount();
  try { await flush(); await fetchAllX(['por', 'zad', 'plan', 'tik']); } catch (e) { return; }
  const after = unreadCount();
  if (after > before) toast(`📩 Nova poruka (${after})`);
  if (!['audit', 'novi', 'beleska', 'krajposete', 'thread', 'admin', 'pin'].includes(S.view)) { const y = scrollY; render(); scrollTo(0, y); }
  else paintBadge();
}
function paintBadge() { const b = document.querySelector('[data-v=poruke] .tb'); const n = unreadCount(); if (b) b.textContent = n || ''; }
setInterval(pollX, 60000);

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
    try { await fetchLog(); } catch (e) { console.warn('log', e); }
    try { await fetchAllX(); } catch (e) { console.warn('x', e); }
    try { await syncFetch(); } catch (e) { console.warn('sync', e); }
    // ponovo primeni izmene koje još čekaju
    pend.filter(o => o.type === 'patch').forEach(o => { const it = S.items.find(x => x.id === o.itemId); if (it) Object.keys(o.fields).forEach(k => { if (!k.includes('@')) it.f[k] = o.fields[k]; }); });
    if (S.cur) S.cur = S.items.find(x => x.id === S.cur.id) || S.cur;
    saveItems();
    if (!['audit', 'novi', 'beleska', 'krajposete', 'thread', 'admin', 'pin'].includes(S.view)) render();
    if (manual) toast('Osveženo: ' + S.items.length + ' kupaca.');
  } catch (e) { console.warn(e); toast('Osvežavanje nije uspelo: ' + String(e.message || e).slice(0, 80)); }
}
// ako je aplikacija zatvorena tokom posete, vrati se pravo na tog kupca
function resumeVisit() {
  if (!S.visit) return false;
  if (!sameDay(S.visit.start)) return false; // stara poseta: zatvara se pri početku dana
  const it = visitItem(); if (!it) return false;
  S.cur = it; go('kartica'); toast('Nastavljaš posetu: ' + S.visit.n); return true;
}
async function boot() {
  history.replaceState({ v: 'danas' }, '');
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    // nova verzija aplikacije: odmah je aktiviraj i jednom osveži stranicu (osim usred unosa)
    let reloaded = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => { if (reloaded || ['audit', 'novi', 'beleska', 'krajposete'].includes(S.view)) return; reloaded = true; location.reload(); });
    navigator.serviceWorker.register('sw.js').then(r => { const skip = () => r.waiting && r.waiting.postMessage('skip'); skip(); r.addEventListener('updatefound', () => { const n = r.installing; n && n.addEventListener('statechange', () => { if (n.state === 'installed') skip(); }); }); r.update().catch(() => {}); }).catch(() => {});
  }
  if (DEMO) { S.items = demoItems(); S.view = isKom() ? 'danas' : 'aktivnost'; resumeVisit() || render(); return; }
  try { await initAuth(); } catch (e) { $app.innerHTML = `<div class="login"><h1>Greška pri prijavi</h1><p class="s">${esc(e.message)}</p></div>`; return; }
  if (!S.account) { S.view = 'login'; render(); return; }
  if (!isKom() && S.view === 'danas') { S.view = 'aktivnost'; history.replaceState({ v: 'aktivnost' }, ''); }
  await syncLoadCache();
  const cached = loadCached();
  if (cached) { S.items = cached; resumeVisit() || render(); }
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
