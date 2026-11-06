/**
 * Core Pilates — Gmail → Studio desk Inbox
 * Runs every 2 minutes: copies new emails into the desk Inbox, and sends replies typed in the desk.
 * Setup: run `setup` once (Run ▸ setup) and approve the permissions.
 */
const PROJECT = 'core-pilates---system';
const BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents';
const SEARCH = 'in:inbox -category:promotions -category:social -category:updates -category:forums';

function setup() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(1).create();
  PropertiesService.getScriptProperties().setProperty('since', String(Math.floor(Date.now() / 1000) - 3 * 86400));
  sync();
  patch_('settings/inbox', { email: { connected: true, account: Session.getActiveUser().getEmail(), at: Date.now() } }, ['email']);
}

function sync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try { pullMail_(); sendOutbox_(); } finally { lock.releaseLock(); }
}

/* ---------- incoming ---------- */
function pullMail_() {
  const props = PropertiesService.getScriptProperties();
  const since = +(props.getProperty('since') || Math.floor(Date.now() / 1000) - 86400);
  const me = Session.getActiveUser().getEmail().toLowerCase();
  const threads = GmailApp.search(SEARCH + ' after:' + since, 0, 50);
  let newest = since;
  threads.forEach(th => {
    const tid = 'em_' + th.getId();
    th.getMessages().forEach(m => {
      const t = Math.floor(m.getDate().getTime() / 1000);
      if (t <= since) return;
      newest = Math.max(newest, t);
      const from = m.getFrom(), addr = (from.match(/<([^>]+)>/) || [0, from])[1].trim().toLowerCase();
      const out = addr === me;
      const name = out ? null : (from.replace(/<[^>]+>/, '').replace(/"/g, '').trim() || addr);
      const text = clean_(m.getPlainBody());
      const mid = 'em_' + m.getId();
      if (out) { const p = PropertiesService.getScriptProperties(), k = 'sent_' + tid, s0 = p.getProperty(k); if (s0 && text.indexOf(s0) === 0) { p.deleteProperty(k); return; } }
      if (exists_('messages/' + mid)) return;
      write_('messages/' + mid, { thread: tid, dir: out ? 'out' : 'in', text: text, subject: m.getSubject(), at: m.getDate().getTime(), by: out ? 'email' : '', status: 'sent' });
      const cur = get_('threads/' + tid) || {};
      const upd = { channel: 'email', last: (out ? 'You: ' : '') + text.slice(0, 140), lastAt: m.getDate().getTime(), subject: m.getSubject() };
      if (!out) { upd.name = name; upd.handle = addr; upd.unread = (+cur.unread || 0) + 1; }
      else if (!cur.handle) { const to = (m.getTo().match(/<([^>]+)>/) || [0, m.getTo()])[1].trim().toLowerCase(); upd.handle = to; upd.name = to; }
      patch_('threads/' + tid, upd, Object.keys(upd));
    });
  });
  props.setProperty('since', String(newest));
}

/* ---------- outgoing (replies typed in the desk) ---------- */
function sendOutbox_() {
  const q = UrlFetchApp.fetch(BASE + ':runQuery', {
    method: 'post', contentType: 'application/json', headers: auth_(), muteHttpExceptions: true,
    payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: 'outbox' }], where: { compositeFilter: { op: 'AND', filters: [
      { fieldFilter: { field: { fieldPath: 'channel' }, op: 'EQUAL', value: { stringValue: 'email' } } },
      { fieldFilter: { field: { fieldPath: 'status' }, op: 'EQUAL', value: { stringValue: 'queued' } } }] } }, limit: 20 } })
  });
  const rows = JSON.parse(q.getContentText() || '[]').filter(r => r.document);
  rows.forEach(r => {
    const id = r.document.name.split('/').pop(), d = decode_(r.document.fields);
    let status = 'sent', err = '';
    try {
      const gid = String(d.thread || '').replace(/^em_/, '');
      const th = gid ? GmailApp.getThreadById(gid) : null;
      if (th) { th.reply(d.text); PropertiesService.getScriptProperties().setProperty('sent_' + d.thread, String(d.text).slice(0, 60)); } else if (d.to) GmailApp.sendEmail(d.to, 'Core Pilates Tirana', d.text); else throw new Error('No email address');
    } catch (e) { status = 'failed'; err = String(e.message || e); }
    patch_('outbox/' + id, { status: status, error: err, sentAt: Date.now() }, ['status', 'error', 'sentAt']);
    patch_('messages/' + id, { status: status }, ['status']);
  });
}

/* ---------- helpers ---------- */
function clean_(t) {
  t = String(t || '').replace(/\r/g, '');
  const cut = t.search(/\n(On .+wrote:|-----Original Message-----|From: .+\nSent: )/);
  if (cut > 0) t = t.slice(0, cut);
  return t.replace(/\n{3,}/g, '\n\n').trim().slice(0, 4000);
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
function exists_(path) { return UrlFetchApp.fetch(BASE + '/' + path, { headers: auth_(), muteHttpExceptions: true }).getResponseCode() === 200; }
function write_(path, obj) { return patch_(path, obj, null); }
function patch_(path, obj, mask) {
  const qs = mask ? '?' + mask.map(m => 'updateMask.fieldPaths=' + encodeURIComponent(m)).join('&') : '';
  const r = UrlFetchApp.fetch(BASE + '/' + path + qs, { method: 'patch', contentType: 'application/json', headers: auth_(), muteHttpExceptions: true, payload: JSON.stringify({ fields: fields_(obj) }) });
  if (r.getResponseCode() >= 300) throw new Error('Firestore ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 300));
  return r;
}
