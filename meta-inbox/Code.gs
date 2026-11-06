/**
 * Core Pilates — Instagram + Facebook Messenger → Studio desk Inbox
 * Every minute: copies new DMs into the desk Inbox and sends replies typed in the desk.
 * Setup: Project Settings ▸ Script properties: PAGE_ID, PAGE_TOKEN (long-lived Page token), optional IG_ID.
 * Then run `setup` once and approve.
 */
const PROJECT = 'core-pilates---system';
const BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents';
const GRAPH = 'https://graph.facebook.com/v23.0/';

function setup() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(1).create();
  const P = props_();
  const page = gget_(P.PAGE_ID + '?fields=name,instagram_business_account{id,username}');
  const ig = page.instagram_business_account || {};
  if (ig.id) PropertiesService.getScriptProperties().setProperty('IG_ID', ig.id);
  sync();
  patch_('settings/inbox', { facebook: { connected: true, page: page.name, at: Date.now() }, instagram: ig.id ? { connected: true, account: '@' + ig.username, at: Date.now() } : { connected: false, note: 'No Instagram business account linked to the Page' } }, ['facebook', 'instagram']);
}

function sync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try { ['messenger', 'instagram'].forEach(pull_); sendOutbox_(); } finally { lock.releaseLock(); }
}

function props_() { return PropertiesService.getScriptProperties().getProperties(); }
function gget_(path) {
  const P = props_(), sep = path.indexOf('?') < 0 ? '?' : '&';
  const r = UrlFetchApp.fetch(GRAPH + path + sep + 'access_token=' + encodeURIComponent(P.PAGE_TOKEN), { muteHttpExceptions: true });
  const j = JSON.parse(r.getContentText() || '{}');
  if (j.error) throw new Error('Meta: ' + j.error.message);
  return j;
}

/* ---------- incoming ---------- */
function pull_(platform) {
  const P = props_(), ch = platform === 'instagram' ? 'instagram' : 'facebook';
  if (ch === 'instagram' && !P.IG_ID) return;
  const me = [P.PAGE_ID, P.IG_ID].filter(Boolean);
  const res = gget_(P.PAGE_ID + '/conversations?platform=' + platform + '&limit=25&fields=id,updated_time,participants,messages.limit(15){id,message,from,created_time}');
  (res.data || []).forEach(cv => {
    const tid = (ch === 'instagram' ? 'ig_' : 'fb_') + cv.id;
    const key = 'seen_' + tid, seen = +(PropertiesService.getScriptProperties().getProperty(key) || 0);
    const upd = Date.parse(cv.updated_time) || 0;
    if (upd && upd <= seen) return;
    const other = ((cv.participants || {}).data || []).find(p => me.indexOf(p.id) < 0) || {};
    const msgs = ((cv.messages || {}).data || []).slice().reverse();
    const cur = get_('threads/' + tid) || {};
    let unread = +cur.unread || 0, last = null;
    msgs.forEach(m => {
      const at = Date.parse(m.created_time) || Date.now();
      if (at <= seen) return;
      const mid = 'mm_' + String(m.id).replace(/[^A-Za-z0-9_-]/g, '').slice(-120);
      if (exists_('messages/' + mid)) return;
      const out = me.indexOf((m.from || {}).id) >= 0, text = String(m.message || '(attachment)').slice(0, 4000);
      if (out) { const p = PropertiesService.getScriptProperties(), k = 'sent_' + tid, s0 = p.getProperty(k); if (s0 && text.indexOf(s0) === 0) { p.deleteProperty(k); last = { text: text, at: at, out: true }; return; } }
      write_('messages/' + mid, { thread: tid, dir: out ? 'out' : 'in', text: text, at: at, by: out ? ch : '', status: 'sent' });
      if (!out) unread++;
      last = { text: text, at: at, out: out };
    });
    const u = { channel: ch, uid: other.id || cur.uid || '', name: other.name || other.username || cur.name || 'Unknown', handle: other.username ? '@' + other.username : (cur.handle || ''), unread: unread };
    if (last) { u.last = (last.out ? 'You: ' : '') + last.text.slice(0, 140); u.lastAt = last.at; }
    patch_('threads/' + tid, u, Object.keys(u));
    PropertiesService.getScriptProperties().setProperty(key, String(upd || Date.now()));
  });
}

/* ---------- outgoing ---------- */
function sendOutbox_() {
  const P = props_();
  ['facebook', 'instagram'].forEach(ch => {
    const q = UrlFetchApp.fetch(BASE + ':runQuery', { method: 'post', contentType: 'application/json', headers: auth_(), muteHttpExceptions: true,
      payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: 'outbox' }], where: { compositeFilter: { op: 'AND', filters: [
        { fieldFilter: { field: { fieldPath: 'channel' }, op: 'EQUAL', value: { stringValue: ch } } },
        { fieldFilter: { field: { fieldPath: 'status' }, op: 'EQUAL', value: { stringValue: 'queued' } } }] } }, limit: 20 } }) });
    JSON.parse(q.getContentText() || '[]').filter(r => r.document).forEach(r => {
      const id = r.document.name.split('/').pop(), d = decode_(r.document.fields);
      let status = 'sent', err = '';
      try {
        const th = get_('threads/' + d.thread) || {};
        if (!th.uid) throw new Error('No recipient id on this chat yet');
        const res = UrlFetchApp.fetch(GRAPH + P.PAGE_ID + '/messages?access_token=' + encodeURIComponent(P.PAGE_TOKEN), { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          payload: JSON.stringify({ recipient: { id: th.uid }, messaging_type: 'RESPONSE', message: { text: d.text } }) });
        const j = JSON.parse(res.getContentText() || '{}');
        if (j.error) throw new Error(j.error.message);
        PropertiesService.getScriptProperties().setProperty('sent_' + d.thread, String(d.text).slice(0, 60));
      } catch (e) { status = 'failed'; err = String(e.message || e); }
      patch_('outbox/' + id, { status: status, error: err, sentAt: Date.now() }, ['status', 'error', 'sentAt']);
      patch_('messages/' + id, { status: status === 'sent' ? 'sent' : 'failed: ' + err }, ['status']);
    });
  });
}

/* ---------- helpers ---------- */
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
