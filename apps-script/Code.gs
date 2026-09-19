/**
 * Backend for the "Lịch trực nhật" web app, hosted on Google Apps Script.
 *
 * Standalone project: it reaches the spreadsheet by id, so it does not have to
 * be bound to the Sheet. Set SHEET_ID to the id in the spreadsheet URL.
 *
 * Storage layout in the spreadsheet:
 *   _state  A1 holds the raw JSON state written by the app (source of truth).
 *   Lịch    A human-readable schedule regenerated on every save.
 *
 * Deploy: Deploy > New deployment > Web app,
 *   Execute as: Me, Who has access: Anyone.
 * Then paste the /exec URL into the web app.
 */

var SHEET_ID = '1EAxhi_FQPCGJymHY_l_uclL1XsR_IL4EkbGAarJYWdE';
var STATE_SHEET = '_state';
var VIEW_SHEET = 'Lịch';
var VIEW_WEEKS = 26;
var DAY = 86400000;
var WEEK = 7 * DAY;

function doGet() {
  try {
    return json({ ok: true, state: readState() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var incoming = normalize(JSON.parse(e.postData.contents));
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
