/**
 * Core Pilates — Google Calendar (coretirana@gmail.com) → Studio desk bookings
 * Every 5 minutes: copies bookings from the calendar (yesterday → next 60 days) into the desk.
 * Setup: share coretirana@gmail.com calendar (See all event details) with the account running this script, then Run ▸ setup.
 */
const PROJECT = 'core-pilates---system';
const BASE = 'https://firestore.googleapis.com/v1/projects/' + PROJECT + '/databases/(default)/documents';
const TZ = 'Europe/Tirane';
const DAYS_AHEAD = 60;
const CAL_ID = 'coretirana@gmail.com'; // calendar shared with the account running this script

function setup() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sync').timeBased().everyMinutes(5).create();
  sync();
  patch_('settings/calendar', { connected: true, account: CAL_ID, at: Date.now() }, ['connected', 'account', 'at']);
}

function sync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const from = new Date(Date.now() - 86400000), to = new Date(Date.now() + DAYS_AHEAD * 86400000);
    const fromD = Utilities.formatDate(from, TZ, 'yyyy-MM-dd');
    const events = (CalendarApp.getCalendarById(CAL_ID) || CalendarApp.getDefaultCalendar()).getEvents(from, to);
    const have = existing_(fromD), seen = {};
    events.forEach(ev => {
      if (ev.isAllDayEvent()) return;
      const s = ev.getStartTime(), e = ev.getEndTime();
      const id = 'cal_' + ev.getId().replace(/[^A-Za-z0-9]/g, '').slice(0, 60) + '_' + s.getTime();
      const guests = ev.getGuestList().map(g => g.getEmail().toLowerCase()).filter(x => x && x !== 'coretirana@gmail.com');
      const doc = {
        date: Utilities.formatDate(s, TZ, 'yyyy-MM-dd'), time: Utilities.formatDate(s, TZ, 'HH:mm'), end: Utilities.formatDate(e, TZ, 'HH:mm'),
        title: String(ev.getTitle() || 'Booking').slice(0, 200), note: String(ev.getDescription() || '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 500),
        guests: guests.join(', '), src: 'gcal'
      };
      seen[id] = 1;
      const old = have[id];
      if (old && ['date', 'time', 'end', 'title', 'note', 'guests'].every(k => String(old[k] || '') === String(doc[k]))) return;
      write_('bookings/' + id, doc);
    });
    Object.keys(have).forEach(id => { if (!seen[id] && have[id].src === 'gcal') del_('bookings/' + id); });
  } finally { lock.releaseLock(); }
}

function existing_(fromD) {
  const q = UrlFetchApp.fetch(BASE + ':runQuery', {
    method: 'post', contentType: 'application/json', headers: auth_(), muteHttpExceptions: true,
    payload: JSON.stringify({ structuredQuery: { from: [{ collectionId: 'bookings' }], where: { fieldFilter: { field: { fieldPath: 'date' }, op: 'GREATER_THAN_OR_EQUAL', value: { stringValue: fromD } } }, limit: 2000 } })
  });
  const out = {};
  JSON.parse(q.getContentText() || '[]').filter(r => r.document).forEach(r => { out[r.document.name.split('/').pop()] = decode_(r.document.fields); });
  return out;
}

/* ---------- Firestore helpers ---------- */
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
function write_(path, obj) { return patch_(path, obj, null); }
function del_(path) { UrlFetchApp.fetch(BASE + '/' + path, { method: 'delete', headers: auth_(), muteHttpExceptions: true }); }
function patch_(path, obj, mask) {
  const qs = mask ? '?' + mask.map(m => 'updateMask.fieldPaths=' + encodeURIComponent(m)).join('&') : '';
  const r = UrlFetchApp.fetch(BASE + '/' + path + qs, { method: 'patch', contentType: 'application/json', headers: auth_(), muteHttpExceptions: true, payload: JSON.stringify({ fields: fields_(obj) }) });
  if (r.getResponseCode() >= 300) throw new Error('Firestore ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 300));
  return r;
}
