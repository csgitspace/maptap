/*************************************************************
 * THE COUNCIL OF CHRIS — MapTap Competition
 * Google Apps Script backend  ·  Google Sheet as the database
 *
 * Deploy:  New deployment → Web app → Execute as: Me →
 *          Who has access: Anyone (or Anyone with Google account)
 * On first run the script auto-creates a Sheet named
 * "Council of Chris — MapTap Data" in your Drive. No manual setup.
 *************************************************************/

const MEMBERS_SHEET = 'Members';
const SCORES_SHEET  = 'Scores';
const PALETTE = ["#D9A94A","#3E8E9C","#C6493B","#69B98C","#B07FD1","#E08A3C","#5E9FE0","#D06B9C","#7FC4A8","#C9A227"];
const EMOJIS  = ["🧭","🌍","🗺️","⛰️","🏔️","🌋","🏝️","🧗","🚩","⚓"];

/* ---------- web app entry ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Council of Chris · MapTap Competition')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- spreadsheet bootstrap ---------- */
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { /* recreate below */ } }
  const ss = SpreadsheetApp.create('Council of Chris — MapTap Data');
  setupSheets_(ss);
  props.setProperty('SS_ID', ss.getId());
  return ss;
}

function setupSheets_(ss) {
  const ms = ss.getSheetByName(MEMBERS_SHEET) || ss.insertSheet(MEMBERS_SHEET);
  ms.clear();
  ms.getRange(1, 1, 1, 4).setValues([['id', 'name', 'color', 'emoji']]);
  ms.setFrozenRows(1);

  const sc = ss.getSheetByName(SCORES_SHEET) || ss.insertSheet(SCORES_SHEET);
  sc.clear();
  sc.getRange(1, 1, 1, 9).setValues([['date', 'memberId', 'total', 'updated', 'r1', 'r2', 'r3', 'r4', 'r5']]);
  sc.getRange('A:A').setNumberFormat('@'); // keep the date column as plain text
  sc.setFrozenRows(1);

  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
}

/* ---------- helpers ---------- */
function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v).trim().slice(0, 10);
}

function readMembers_(ss) {
  const vals = ss.getSheetByName(MEMBERS_SHEET).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    out.push({
      id: String(vals[i][0]),
      name: String(vals[i][1]),
      color: String(vals[i][2] || PALETTE[0]),
      emoji: String(vals[i][3] || '🧭')
    });
  }
  return out;
}

/* ---------- league settings (scoring mode) ---------- */
const SETTINGS_SHEET = 'Settings';
const DEFAULT_SETTINGS = { mode: 'total', bestN: 5 };

function ensureSettingsSheet_(ss) {
  let sh = ss.getSheetByName(SETTINGS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SETTINGS_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    sh.appendRow(['mode', DEFAULT_SETTINGS.mode]);
    sh.appendRow(['bestN', DEFAULT_SETTINGS.bestN]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readSettings_(ss) {
  const sh = ensureSettingsSheet_(ss);
  const vals = sh.getDataRange().getValues();
  const out = { mode: DEFAULT_SETTINGS.mode, bestN: DEFAULT_SETTINGS.bestN };
  for (let i = 1; i < vals.length; i++) {
    const k = String(vals[i][0]);
    if (k === 'mode') out.mode = String(vals[i][1] || DEFAULT_SETTINGS.mode);
    if (k === 'bestN') out.bestN = Number(vals[i][1]) || DEFAULT_SETTINGS.bestN;
  }
  return out;
}

function saveSettings(s) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sh = ensureSettingsSheet_(ss);
    const vals = sh.getDataRange().getValues();
    const want = { mode: String(s.mode || DEFAULT_SETTINGS.mode), bestN: Number(s.bestN) || DEFAULT_SETTINGS.bestN };
    const seen = {};
    for (let i = 1; i < vals.length; i++) {
      const k = String(vals[i][0]);
      if (want[k] !== undefined) { sh.getRange(i + 1, 2).setValue(want[k]); seen[k] = true; }
    }
    Object.keys(want).forEach(function (k) { if (!seen[k]) sh.appendRow([k, want[k]]); });
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/* older sheets were 4 columns wide — add the round columns in place */
function ensureRoundCols_(sh) {
  if (sh.getLastColumn() < 9) {
    sh.getRange(1, 5, 1, 5).setValues([['r1', 'r2', 'r3', 'r4', 'r5']]);
  }
  return sh;
}

/* ---------- read everything ---------- */
function getData() {
  const ss = getSpreadsheet_();
  const members = readMembers_(ss);
  const entries = {};
  const sh = ensureRoundCols_(ss.getSheetByName(SCORES_SHEET));
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const date = vals[i][0], mid = vals[i][1], total = vals[i][2];
    if (!date || mid === '' || mid === null) continue;
    const d = normDate_(date);
    if (!entries[d]) entries[d] = {};
    const rec = { total: Number(total) };
    const rounds = [];
    for (let c = 4; c <= 8; c++) {
      const v = vals[i][c];
      if (v === '' || v === null || v === undefined) { rounds.length = 0; break; }
      rounds.push(Number(v));
    }
    if (rounds.length === 5) rec.rounds = rounds;
    entries[d][String(mid)] = rec;
  }
  return { members: members, entries: entries, settings: readSettings_(ss), seasons: readSeasons_(ss) };
}

/* ---------- write scores (upsert one day) ---------- */
function saveDay(date, scores) {
  date = normDate_(date);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sh = ensureRoundCols_(ss.getSheetByName(SCORES_SHEET));
    const vals = sh.getDataRange().getValues();
    const idx = {};
    for (let i = 1; i < vals.length; i++) idx[normDate_(vals[i][0]) + '|' + String(vals[i][1])] = i + 1;

    const now = new Date();
    const toAppend = [];
    (scores || []).forEach(function (s) {
      let t = Math.round(Number(s.total));
      if (isNaN(t)) return;
      t = Math.max(0, Math.min(1000, t));
      let r = ['', '', '', '', ''];
      if (s.rounds && s.rounds.length === 5) {
        r = s.rounds.map(function (x) {
          const n = Math.round(Number(x));
          return isNaN(n) ? '' : Math.max(0, Math.min(1000, n));
        });
      }
      const key = date + '|' + String(s.memberId);
      if (idx[key]) {
        sh.getRange(idx[key], 3).setValue(t);
        sh.getRange(idx[key], 4).setValue(now);
        sh.getRange(idx[key], 5, 1, 5).setValues([r]);
      } else {
        toAppend.push([date, String(s.memberId), t, now].concat(r));
      }
    });
    if (toAppend.length) sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, 9).setValues(toAppend);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function deleteScore(date, memberId) {
  date = normDate_(date);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSpreadsheet_().getSheetByName(SCORES_SHEET);
    const vals = sh.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (normDate_(vals[i][0]) === date && String(vals[i][1]) === String(memberId)) sh.deleteRow(i + 1);
    }
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/* ---------- roster ---------- */
function addMember(name) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const members = readMembers_(ss);
    name = String(name || '').trim().slice(0, 24);
    if (!name) throw new Error('Name required');
    if (members.some(function (m) { return m.name.toLowerCase() === name.toLowerCase(); }))
      throw new Error(name + ' is already on the council');
    const i = members.length;
    const id = 'm' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    ss.getSheetByName(MEMBERS_SHEET).appendRow([id, name, PALETTE[i % PALETTE.length], EMOJIS[i % EMOJIS.length]]);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function updateMember(id, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSpreadsheet_().getSheetByName(MEMBERS_SHEET);
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(id)) {
        if (patch.name != null)  sh.getRange(i + 1, 2).setValue(String(patch.name).trim().slice(0, 24));
        if (patch.color != null) sh.getRange(i + 1, 3).setValue(patch.color);
        if (patch.emoji != null) sh.getRange(i + 1, 4).setValue(patch.emoji);
        break;
      }
    }
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function removeMember(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSpreadsheet_().getSheetByName(MEMBERS_SHEET);
    const vals = sh.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === String(id)) sh.deleteRow(i + 1);
    }
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/* ---------- data management ---------- */
function clearAll() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { setupSheets_(getSpreadsheet_()); return getData(); }
  finally { lock.releaseLock(); }
}

function bulkImport(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    setupSheets_(ss);
    const ms = ss.getSheetByName(MEMBERS_SHEET);
    const sc = ss.getSheetByName(SCORES_SHEET);
    (data.members || []).forEach(function (m) {
      ms.appendRow([m.id, m.name, m.color || PALETTE[0], m.emoji || '🧭']);
    });
    const now = new Date();
    const rows = [];
    Object.keys(data.entries || {}).forEach(function (d) {
      Object.keys(data.entries[d]).forEach(function (mid) {
        rows.push([normDate_(d), mid, Number(data.entries[d][mid].total), now]);
      });
    });
    if (rows.length) sc.getRange(sc.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function seedSample() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    setupSheets_(ss);
    const ms = ss.getSheetByName(MEMBERS_SHEET);
    const sc = ss.getSheetByName(SCORES_SHEET);
    const names = ['Chris', 'Dana', 'Miguel', 'Priya', 'Sam'];
    const means = [905, 872, 918, 860, 889], spread = [38, 55, 30, 62, 45];
    const ids = names.map(function (n, i) {
      const id = 'm' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
      ms.appendRow([id, n, PALETTE[i % PALETTE.length], EMOJIS[i % EMOJIS.length]]);
      return id;
    });
    let seed = 20260601;
    function rnd() { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; }
    function gauss() { let u = 0, v = 0; while (!u) u = rnd(); while (!v) v = rnd(); return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v); }
    const tz = Session.getScriptTimeZone(), now = new Date();
    const rows = [];
    for (let i = 0; i < 28; i++) {
      const d = new Date(); d.setDate(now.getDate() - 27 + i);
      const iso = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      ids.forEach(function (id, mi) {
        if (rnd() < 0.12) return;
        let s = Math.round(means[mi] + i * 0.6 + gauss() * spread[mi]);
        s = Math.max(420, Math.min(1000, s));
        rows.push([iso, id, s, now]);
      });
    }
    if (rows.length) sc.getRange(sc.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================= *
 *  Seasons
 * ============================================================= */
const SEASONS_SHEET = 'Seasons';

function ensureSeasonsSheet_(ss) {
  let sh = ss.getSheetByName(SEASONS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SEASONS_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['name', 'start', 'end']]);
    sh.getRange('B:C').setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readSeasons_(ss) {
  const sh = ensureSeasonsSheet_(ss);
  const vals = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    out.push({
      name: String(vals[i][0]),
      start: normDate_(vals[i][1]),
      end: vals[i][2] ? normDate_(vals[i][2]) : ''
    });
  }
  return out;
}

/* Close the running season and open a new one. Scores are never deleted —
   a season is just a named date range, so history stays recomputable. */
function startSeason(name, startDate) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sh = ensureSeasonsSheet_(ss);
    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const start = startDate ? normDate_(startDate) : today;
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] && !vals[i][2]) {
        const prevEnd = new Date(new Date(start + 'T00:00:00').getTime() - 86400000);
        sh.getRange(i + 1, 3).setValue(Utilities.formatDate(prevEnd, tz, 'yyyy-MM-dd'));
      }
    }
    sh.appendRow([String(name || 'New season').slice(0, 40), start, '']);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================= *
 *  Announcements — post a result to Banter exactly once
 * ============================================================= */
function announceIfNew(key, text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_();
    const sh = ensureSettingsSheet_(ss);
    const marker = 'ann_' + String(key);
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === marker) return null; // already announced
    }
    sh.appendRow([marker, '1']);
    const chat = ensureChatSheet_(ss);
    const id = 'c' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    chat.appendRow([id, Date.now(), '__system__', String(text).slice(0, 280), '{}']);
    return getChat();
  } finally {
    lock.releaseLock();
  }
}

/* commissioner convenience: jump to the underlying Sheet */
function getSheetUrl() { return getSpreadsheet_().getUrl(); }

/* The real, shareable web-app URL.
   NOTE: client-side location.href returns the sandboxed iframe address
   (…googleusercontent.com/userCodeAppPanel), which won't load for anyone
   else. This returns the proper /exec deployment URL. */
function getAppUrl() {
  try { return ScriptApp.getService().getUrl() || ''; }
  catch (e) { return ''; }
}

/* ============================================================= *
 *  Banter — the chat / bragging-rights feed
 * ============================================================= */
const CHAT_SHEET = 'Chat';

function ensureChatSheet_(ss) {
  let sh = ss.getSheetByName(CHAT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CHAT_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['id', 'ts', 'memberId', 'text', 'reactions']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getChat() {
  const sh = ensureChatSheet_(getSpreadsheet_());
  const vals = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    let reactions = {};
    try { reactions = JSON.parse(vals[i][4] || '{}'); } catch (e) {}
    out.push({
      id: String(vals[i][0]),
      ts: Number(vals[i][1]) || 0,
      memberId: String(vals[i][2]),
      text: String(vals[i][3]),
      reactions: reactions
    });
  }
  out.sort(function (a, b) { return a.ts - b.ts; });
  return out.slice(-200);
}

function postMessage(memberId, text) {
  text = String(text || '').trim().slice(0, 280);
  if (!text) throw new Error('Empty message');
  if (!memberId) throw new Error('Pick who you are first');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureChatSheet_(getSpreadsheet_());
    const id = 'c' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    sh.appendRow([id, Date.now(), String(memberId), text, '{}']);
    return getChat();
  } finally {
    lock.releaseLock();
  }
}

function reactMessage(id, memberId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureChatSheet_(getSpreadsheet_());
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(id)) {
        let r = {};
        try { r = JSON.parse(vals[i][4] || '{}'); } catch (e) {}
        const arr = r.fire || [];
        const at = arr.indexOf(String(memberId));
        if (at >= 0) arr.splice(at, 1); else arr.push(String(memberId));
        r.fire = arr;
        sh.getRange(i + 1, 5).setValue(JSON.stringify(r));
        break;
      }
    }
    return getChat();
  } finally {
    lock.releaseLock();
  }
}

function deleteMessage(id, memberId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureChatSheet_(getSpreadsheet_());
    const vals = sh.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === String(id) && String(vals[i][2]) === String(memberId)) sh.deleteRow(i + 1);
    }
    return getChat();
  } finally {
    lock.releaseLock();
  }
}
