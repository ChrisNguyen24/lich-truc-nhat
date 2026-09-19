/**
 * Backend for the "Lịch trực nhật" web app, hosted on Google Apps Script.
 *
 * Standalone project: it reaches the spreadsheet by id, so it does not have to
 * be bound to the Sheet. Set SHEET_ID to the id in the spreadsheet URL.
 *
 * Storage layout in the spreadsheet:
 *   _state  A1 holds the raw JSON state written by the app (source of truth).
 *   Lịch    A human-readable schedule regenerated on every save.
 *   Thu chi The shared ledger: one row per income or expense entry.
 *
 * Deploy: Deploy > New deployment > Web app,
 *   Execute as: Me, Who has access: Anyone.
 * Then paste the /exec URL into the web app.
 */

var SHEET_ID = '1EAxhi_FQPCGJymHY_l_uclL1XsR_IL4EkbGAarJYWdE';
var STATE_SHEET = '_state';
var VIEW_SHEET = 'Lịch';
var LEDGER_SHEET = 'Thu chi';
var LEDGER_HEADER = ['id', 'Ngày', 'Loại', 'Số tiền', 'Nội dung', 'Người', 'Ghi chú', 'Đã giải ngân'];
var VIEW_WEEKS = 26;
var DAY = 86400000;
var WEEK = 7 * DAY;

function doGet(e) {
  try {
    var want = e && e.parameter ? e.parameter.resource : '';
    if (want === 'ledger') return json({ ok: true, entries: readLedger() });
    return json({ ok: true, state: readState(), entries: readLedger() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var body = JSON.parse(e.postData.contents);

    if (body.kind === 'ledger.add') {
      var added = appendEntry(body.entry);
      return json({ ok: true, entry: added, entries: readLedger() });
    }
    if (body.kind === 'ledger.import') {
      writeLedger((body.entries || []).map(normalizeEntry));
      return json({ ok: true, entries: readLedger() });
    }
    if (body.kind === 'ledger.delete') {
      deleteEntry(String(body.id));
      return json({ ok: true, entries: readLedger() });
    }
    if (body.kind === 'ledger.settle') {
      setSettled(String(body.id), !!body.settled);
      return json({ ok: true, entries: readLedger() });
    }

    var incoming = normalize(body);
    var current = readState();
    // Last write wins, but never let a stale client overwrite newer data.
    if (current && current.updatedAt > incoming.updatedAt) {
      return json({ ok: true, state: current, note: 'stale' });
    }
    writeState(incoming);
    renderSheet(incoming);
    return json({ ok: true, state: incoming });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readState() {
  var raw = sheet(STATE_SHEET).getRange('A1').getValue();
  if (!raw) return { members: [], start: '', skips: [], swaps: [], updatedAt: 0 };
  return normalize(JSON.parse(raw));
}

function writeState(state) {
  sheet(STATE_SHEET).getRange('A1').setValue(JSON.stringify(state));
}

function normalize(o) {
  o = o || {};
  return {
    members: Array.isArray(o.members) ? o.members.map(String) : [],
    start: /^\d{4}-\d{2}-\d{2}$/.test(o.start || '') ? o.start : '',
    skips: Array.isArray(o.skips) ? o.skips.map(Number).filter(function (n) { return n >= 0; }) : [],
    swaps: Array.isArray(o.swaps) ? o.swaps.filter(function (p) { return Array.isArray(p) && p.length === 2; }) : [],
    updatedAt: Number(o.updatedAt) || 0
  };
}

/** Same rotation rules as the web app: base order, then skips, then swaps. */
function schedule(state, count) {
  var n = state.members.length;
  var out = [];
  if (!n) return out;
  var skips = state.skips.slice().sort(function (a, b) { return a - b; });
  for (var w = 0; w < count; w++) {
    var offset = 0;
    for (var i = 0; i < skips.length; i++) if (skips[i] <= w) offset++;
    out.push({ week: w, name: state.members[(w + offset) % n], skipped: skips.indexOf(w) >= 0 });
  }
  state.swaps.forEach(function (pair) {
    var a = Number(pair[0]), b = Number(pair[1]);
    if (a < count && b < count) {
      var t = out[a].name; out[a].name = out[b].name; out[b].name = t;
      out[a].swapped = true; out[b].swapped = true;
    }
  });
  return out;
}

function renderSheet(state) {
  var sh = sheet(VIEW_SHEET);
  sh.clear();
  var rows = [['Tuần', 'Từ ngày', 'Đến ngày', 'Người trực', 'Ghi chú']];
  if (state.start && state.members.length) {
    var parts = state.start.split('-');
    var start = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    var currentWeek = Math.floor((new Date() - start) / WEEK);
    var from = Math.max(0, currentWeek);
    schedule(state, from + VIEW_WEEKS).slice(from).forEach(function (r) {
      var begin = new Date(start.getTime() + r.week * WEEK);
      var end = new Date(begin.getTime() + 6 * DAY);
      var notes = [];
      if (r.skipped) notes.push('đã bỏ lượt');
      if (r.swapped) notes.push('đã đổi');
      rows.push(['Tuần ' + (r.week + 1), begin, end, r.name, notes.join(', ')]);
    });
  }
  sh.getRange(1, 1, rows.length, 5).setValues(rows);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold');
  if (rows.length > 1) sh.getRange(2, 2, rows.length - 1, 2).setNumberFormat('dd/MM/yyyy');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 5);
}


// ---- Ledger: one row per income ("thu") or expense ("chi") entry ----

function ledgerSheet() {
  var sh = sheet(LEDGER_SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, LEDGER_HEADER.length).setValues([LEDGER_HEADER]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // Keep dates as plain yyyy-MM-dd text so they survive the round trip untouched.
  sh.getRange('B2:B').setNumberFormat('@');
  return sh;
}

/** Sheets may hand back a Date, a yyyy-MM-dd string, or dd/MM/yyyy text. */
function toISODate(v) {
  if (v instanceof Date || Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function normalizeEntry(o) {
  o = o || {};
  var type = o.type === 'thu' ? 'thu' : 'chi';
  return {
    id: String(o.id || (Date.now() + '-' + Math.random().toString(36).slice(2, 8))),
    date: toISODate(o.date),
    type: type,
    amount: Math.round(Number(o.amount) || 0),
    note: String(o.note || ''),
    who: String(o.who || ''),
    memo: String(o.memo || ''),
    settled: o.settled === true || o.settled === 'TRUE' || o.settled === 'true'
  };
}

function entryToRow(en) {
  return [en.id, en.date, en.type, en.amount, en.note, en.who, en.memo, en.settled];
}

function readLedger() {
  var sh = ledgerSheet();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, LEDGER_HEADER.length).getValues();
  return values
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      return normalizeEntry({
        id: r[0],
        date: toISODate(r[1]),
        type: r[2],
        amount: r[3],
        note: r[4],
        who: r[5],
        memo: r[6],
        settled: r[7]
      });
    });
}

function writeLedger(entries) {
  var sh = ledgerSheet();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, LEDGER_HEADER.length).clearContent();
  if (!entries.length) return;
  entries.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  sh.getRange(2, 1, entries.length, LEDGER_HEADER.length).setValues(entries.map(entryToRow));
  sh.getRange(2, 4, entries.length, 1).setNumberFormat('#,##0 "đ"');
  sh.autoResizeColumns(1, LEDGER_HEADER.length);
}

function appendEntry(entry) {
  var en = normalizeEntry(entry);
  var sh = ledgerSheet();
  sh.appendRow(entryToRow(en));
  sh.getRange(sh.getLastRow(), 4).setNumberFormat('#,##0 "đ"');
  return en;
}

function setSettled(id, settled) {
  var sh = ledgerSheet();
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) sh.getRange(i + 2, 8).setValue(settled);
  }
}

function deleteEntry(id) {
  var sh = ledgerSheet();
  var last = sh.getLastRow();
  if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === id) sh.deleteRow(i + 2);
  }
}
