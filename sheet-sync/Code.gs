/**
 * Core Pilates — Google Sheet "Core-Daily 2026"  <->  Studio desk (two-way)
 * Every 5 minutes:
 *   • new rows typed in the monthly "Daily …" tab  → become sales / expenses in the desk
 *   • sales / expenses made in the desk            → are added as rows in the sheet
 * Column A (#) gets a small ID so a row is never copied twice ("·" = before the sync started).
 * Setup: Run ▸ setupSheetSync once and approve.
 */
const SS_ID = '10_ZSd2xLr3kDEHwbT_cXqdPTOU3Ktcx3hb2kh8IbLMc';
const PROJECT = 'core-pilates---system';
const BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents';
const TZ = 'Europe/Tirane';
const MONTHS = [['january','janar','jan'],['february','shkurt','feb'],['march','mars','mar'],['april','prill','apr'],['may','maj'],['june','qershor','jun'],
  ['july','korrik','jul'],['august','gusht','aug'],['september','septembre','shtator','sep'],['october','tetor','oct'],['november','novembre','nentor','nov'],['december','decembre','dhjetor','dec']];
const SYNC_RE = /^(·|[SD][0-9a-zA-Z]{6,})$/; // our marks; plain row numbers in # are left alone
const MON_EN = ['January','February','March','April','May','June','July','August','September','October','November','December'];

function setupSheetSync() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sheetSync').timeBased().everyMinutes(5).create();
  const P = PropertiesService.getScriptProperties();
  P.setProperty('setupAt', String(Date.now()));
  P.setProperty('pend', '{}');
  // mark every row that already exists, so the past is not copied again
  const ss = SpreadsheetApp.openById(SS_ID), now = new Date();
  [0, -1].forEach(k => {
    const d = new Date(now.getFullYear(), now.getMonth() + k, 1), sh = monthTab_(ss, d.getFullYear(), d.getMonth(), false);
    if (!sh) return;
    const H = header_(sh); if (!H) return;
    const marks = [];
    const notes = [];
    rows_(sh, H).forEach(r => { if (!r.id && r.name && r.date) (r.a0 ? notes : marks).push('A' + r.row); });
    for (let i = 0; i < marks.length; i += 200) sh.getRangeList(marks.slice(i, i + 200)).setValue('·');
    notes.forEach(a => sh.getRange(a).setNote('·'));
  });
  patch_('settings/sheetsync', { connected: true, sheet: ss.getName(), at: Date.now() }, ['connected', 'sheet', 'at']);
}

function sheetSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try { const ss = SpreadsheetApp.openById(SS_ID); sheetToDesk_(ss); deskToSheet_(ss); }
  finally { lock.releaseLock(); }
}

/* ---------- sheet → desk ---------- */
function sheetToDesk_(ss) {
  const P = PropertiesService.getScriptProperties(), pend = JSON.parse(P.getProperty('pend') || '{}'), seen = {}, now = new Date();
  const tabs = [monthTab_(ss, now.getFullYear(), now.getMonth(), false)];
  if (now.getDate() <= 5) { const d = new Date(now.getFullYear(), now.getMonth() - 1, 1); tabs.push(monthTab_(ss, d.getFullYear(), d.getMonth(), false)); }
  tabs.filter(Boolean).forEach(sh => {
    const H = header_(sh); if (!H) return;
    rows_(sh, H).forEach(r => {
      if (r.id || !r.name || !r.date) return;
      const money = r.cash + r.poss + r.debit + r.expense;
      if (!money && !r.typeOfPayment) return;                 // still being typed
      const h = sh.getSheetId() + ':' + r.hash; seen[h] = 1;
      if (!pend[h]) { pend[h] = Date.now(); return; }          // wait one round so half-typed rows are not taken
      const id = 'S' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
      write_('sheetin/' + id, { status: 'new', at: Date.now(), tab: sh.getName(), row: r.row, date: r.date, name: r.name, typeOfPayment: r.typeOfPayment,
        category: r.category, type: r.type, n: r.n, start: r.start, instructor: r.instructor, cash: r.cash, poss: r.poss, debit: r.debit, expense: r.expense, operator: r.operator, note: r.note });
      if (r.a0) sh.getRange(r.row, 1).setNote(id); else sh.getRange(r.row, 1).setValue(id);
      delete pend[h];
    });
  });
  Object.keys(pend).forEach(k => { if (!seen[k]) delete pend[k]; });
  P.setProperty('pend', JSON.stringify(pend));
}

/* ---------- desk → sheet ---------- */
function deskToSheet_(ss) {
  const since = +PropertiesService.getScriptProperties().getProperty('setupAt') || Date.now();
  const pays = query_('payments', since), exps = query_('expenses', since);
  if (!pays.length && !exps.length) return;
  const staff = {}; list_('staff').forEach(d => staff[d.id] = String(d.name || '').split(/\s+/)[0]);
  const cache = {};
  const client = id => { if (!id) return null; if (!(id in cache)) cache[id] = get_('clients/' + id); return cache[id]; };
  const items = [];
  pays.forEach(p => { if (p.sheetId || p.src === 'sheet' || !p.date) return; items.push(['payments', p]); });
  exps.forEach(e => { if (e.sheetId || e.src === 'sheet' || !e.date) return; items.push(['expenses', e]); });
  items.sort((a, b) => (a[1].ts || 0) - (b[1].ts || 0));
  items.forEach(([col, x]) => {
    const [y, m] = x.date.split('-').map(Number), sh = monthTab_(ss, y, m - 1, true); if (!sh) return;
    const H = header_(sh); if (!H) return;
    const sid = 'D' + x.id.slice(-10), v = {};
    const pay = methodSplit_(x);
    v.date = dmy_(x.date); v.month = MON_EN[m - 1].slice(0, 3) + '.'; v.operator = cap_(x.by || ''); v.note = x.note || '';
    if (col === 'expenses') {
      Object.assign(v, { name: x.note || x.cat || 'Expense', typeOfPayment: 'For Studio', category: 'Expense', type: 'Studio', n: 1, start: v.date, paid: v.date, instructor: 'Studio', expense: +x.amount || 0, note: x.cat && x.note ? x.cat : '' });
    } else if (x.kind === 'product') {
      const mm = /^(.*?)\s*×\s*(\d+)$/.exec(x.item || ''), c = client(x.client);
      Object.assign(v, { name: mm ? mm[1] : (x.item || 'Product'), typeOfPayment: 'For Studio', category: 'Sale', type: 'Studio', n: mm ? +mm[2] : 1, start: v.date, paid: v.date, instructor: 'Studio', cash: pay.cash, poss: pay.card });
      if (c && c.name) v.note = [c.name, v.note].filter(Boolean).join(' · ');
    } else {
      const c = client(x.client) || {}, it = String(x.item || ''), parts = it.split(' · '), t = parts[0], n = +(parts[1] || '').replace(/\D/g, '') || '';
      const typ = { 'Private': 'Private', 'Semi private': 'Group', 'Duet': 'Duo', 'Trio': 'Trio', 'Clinic': 'Clinic', 'Yoga': 'Group', 'Mat Pilates': 'Group', 'Functional Private': 'Functional Private' }[t] || t;
      Object.assign(v, { name: c.name || x.name || 'Client', typeOfPayment: /^debt/i.test(it) ? 'Debt Payment' : n === 1 ? 'Single Session' : 'Package Renewal', category: t === 'Yoga' ? 'Yoga' : 'Pilates',
        type: typ, n: n, start: dmy_(c.packageStart && c.plan === it ? c.packageStart : x.date), paid: v.date, instructor: (c.trainers || []).map(t => staff[t]).filter(Boolean).join(', '), cash: pay.cash, poss: pay.card });
    }
    const row = new Array(H.width).fill('');
    row[0] = sid;
    Object.keys(H.col).forEach(k => { if (k in v && v[k] !== undefined && v[k] !== '' && v[k] !== 0) row[H.col[k]] = v[k]; });
    sh.appendRow(row);
    patch_(col + '/' + x.id, { sheetId: sid }, ['sheetId']);
  });
}
function methodSplit_(x) {
  const a = +x.amount || 0;
  if (x.split) return { cash: +x.split.Cash || 0, card: +x.split.Card || 0 };
  if (x.method === 'Card') return { cash: 0, card: a };
  return { cash: a, card: 0 };
}

/* ---------- sheet helpers ---------- */
function monthTab_(ss, y, m, create) {
  const words = MONTHS[m], tabs = ss.getSheets();
  const match = tabs.filter(s => { const n = s.getName().toLowerCase(); if (n.indexOf('daily') < 0) return false;
    if (!words.some(w => new RegExp('\\b' + w + '\\b').test(n))) return false;
    const yr = /(20\d\d)/.exec(n); return yr ? +yr[1] === y : y === 2026; });
  if (match.length) return match[0];
  if (!create) return null;
  // new month: copy the previous month's tab and empty it
  const pd = new Date(y, m - 1, 1), prev = monthTab_(ss, pd.getFullYear(), pd.getMonth(), false); if (!prev) return null;
  const H = header_(prev); if (!H) return null;
  const sh = prev.copyTo(ss).setName('Daily ' + MON_EN[m] + ' ' + y);
  ss.setActiveSheet(sh); ss.moveActiveSheet(prev.getIndex() + 1);
  const last = sh.getLastRow(); if (last > H.row) sh.getRange(H.row + 1, 1, last - H.row, sh.getLastColumn()).clearContent();
  return sh;
}
function header_(sh) {
  const top = sh.getRange(1, 1, Math.min(6, Math.max(1, sh.getLastRow())), Math.max(1, sh.getLastColumn())).getDisplayValues();
  for (let r = 0; r < top.length; r++) {
    const h = top[r].map(x => String(x).trim().toLowerCase().replace(/\s+/g, ' '));
    const fn = h.indexOf('full name'); if (fn < 0) continue;
    const col = {}, f = (k, ...names) => { const i = h.findIndex(x => names.indexOf(x) >= 0); if (i >= 0) col[k] = i; };
    f('month', 'month'); f('name', 'full name'); f('typeOfPayment', 'type of payment'); f('category', 'category'); f('type', 'type');
    f('n', '#of session /product', '# of session /product', '#of session / product'); f('start', 'start of packages'); f('paid', 'package payment');
    f('instructor', 'instructor'); f('cash', 'cash'); f('poss', 'poss', 'pos'); f('debit', 'debit'); f('expense', 'expense'); f('total', 'total'); f('operator', 'operator'); f('note', 'note');
    col.date = h.indexOf('date') >= 0 ? h.indexOf('date') : (col.month || 2) - 1;
    if (col.n === undefined && col.type !== undefined) col.n = col.type + 1;
    if (col.start === undefined && col.n !== undefined) col.start = col.n + 1;
    if (col.paid === undefined && col.start !== undefined) col.paid = col.start + 1;
    return { row: r + 1, col: col, width: h.length };
  }
  return null;
}
function rows_(sh, H) {
  const last = sh.getLastRow(); if (last <= H.row) return [];
  const n = last - H.row, vals = sh.getRange(H.row + 1, 1, n, H.width).getValues(), disp = sh.getRange(H.row + 1, 1, n, H.width).getDisplayValues(), C = H.col;
  const notes = sh.getRange(H.row + 1, 1, n, 1).getNotes();
  return vals.map((v, i) => {
    const g = k => C[k] === undefined ? '' : v[C[k]], s = k => String(g(k) === null ? '' : g(k)).trim();
    const a0 = String(v[0] || '').trim(), nt = String(notes[i][0] || '').trim();
    return { row: H.row + 1 + i, id: SYNC_RE.test(a0) ? a0 : SYNC_RE.test(nt) ? nt : '', a0: a0, date: iso_(g('date')), name: s('name').replace(/\s+/g, ' '), typeOfPayment: s('typeOfPayment'), category: s('category'),
      type: s('type'), n: num_(g('n')) || 1, start: iso_(g('start')), instructor: s('instructor'), cash: num_(g('cash')), poss: num_(g('poss')), debit: num_(g('debit')),
      expense: num_(g('expense')), operator: s('operator'), note: s('note'), hash: Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, disp[i].slice(1).join('|'))).slice(0, 16) };
  });
}
function num_(v) { if (typeof v === 'number') return v; const s = String(v || '').replace(/[^\d.,-]/g, '').replace(/,/g, ''); return +s || 0; }
function iso_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = /(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/.exec(String(v || '')); if (!m) return '';
  return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}
function dmy_(iso) { const p = String(iso || '').split('-'); return p.length === 3 ? p[2] + '.' + p[1] + '.' + p[0] : ''; }
function cap_(s) { s = String(s || ''); return s ? s.charAt(0).toUpperCase() + s.slice(1) : ''; }

/* ---------- Firestore helpers ---------- */
function query_(coll, since) {
  const q = UrlFetchApp.fetch(BASE + ':runQuery', { method: 'post', contentType: 'application/json', headers: auth_(), muteHttpExceptions: true,
    payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: coll }], where: { fieldFilter: { field: { fieldPath: 'ts' }, op: 'GREATER_THAN_OR_EQUAL', value: { integerValue: String(since) } } }, limit: 500 } }) });
  return JSON.parse(q.getContentText() || '[]').filter(r => r.document).map(r => Object.assign(decode_(r.document.fields), { id: r.document.name.split('/').pop() }));
}
function list_(coll) {
  const r = UrlFetchApp.fetch(BASE + '/' + coll + '?pageSize=300', { headers: auth_(), muteHttpExceptions: true }), j = JSON.parse(r.getContentText() || '{}');
  return (j.documents || []).map(d => Object.assign(decode_(d.fields), { id: d.name.split('/').pop() }));
}
function auth_() { return { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }; }
function enc_(v) {
  if (v === null || v === undefined) return { nullValue: null };
  if (typeof v === 'boolean') return { booleanValue: v };
  if (typeof v === 'number') return Number.isInteger(v) ? { integerValue: String(v) } : { doubleValue: v };
  if (Array.isArray(v)) return { arrayValue: { values: v.map(enc_) } };
  if (typeof v === 'object') { const f = {}; Object.keys(v).forEach(k => f[k] = enc_(v[k])); return { mapValue: { fields: f } }; }
  return { stringValue: String(v) };
}
function dec_(v) {
  if ('stringValue' in v) return v.stringValue; if ('integerValue' in v) return +v.integerValue; if ('doubleValue' in v) return v.doubleValue;
  if ('booleanValue' in v) return v.booleanValue; if ('mapValue' in v) return decode_(v.mapValue.fields || {});
  if ('arrayValue' in v) return (v.arrayValue.values || []).map(dec_); return null;
}
function decode_(f) { const o = {}; Object.keys(f || {}).forEach(k => o[k] = dec_(f[k])); return o; }
function fields_(o) { const f = {}; Object.keys(o).forEach(k => f[k] = enc_(o[k])); return f; }
function get_(path) { const r = UrlFetchApp.fetch(BASE + '/' + path, { headers: auth_(), muteHttpExceptions: true }); return r.getResponseCode() === 200 ? decode_(JSON.parse(r.getContentText()).fields) : null; }
function write_(path, obj) { return patch_(path, obj, null); }
function patch_(path, obj, mask) {
  const qs = mask ? '?' + mask.map(m => 'updateMask.fieldPaths=' + encodeURIComponent(m)).join('&') : '';
  const r = UrlFetchApp.fetch(BASE + '/' + path + qs, { method: 'patch', contentType: 'application/json', headers: auth_(), muteHttpExceptions: true, payload: JSON.stringify({ fields: fields_(obj) }) });
  if (r.getResponseCode() >= 300) throw new Error('Firestore ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 300));
  return r;
}
