/**
 * Rebuilt Training — Google Apps Script backend
 *
 * Deployed at:
 *   https://script.google.com/macros/s/AKfycbyRSuUKBjwlPZiY8E7IZZUEQ8u-uwwuTXZnQ0nTPqtfk-bu3zyAa6RKZYw7687arJVs/exec
 *
 * This file is the source of truth for the deployed script. After editing,
 * paste into the Apps Script project and create a new version of the
 * existing deployment so the /exec URL stays the same — do NOT create a
 * new deployment, because every HTML page hard-codes the current URL.
 *
 * Sheets:
 *   - "Tracking"  — knowledge-check + OLT submissions (existing schema)
 *                   columns: email | module_id | module_name | type |
 *                            score_pct | passed | completed_at | role | class
 *   - "PageViews" — page-view tracking (auto-created on first event)
 *                   columns: client_ts | server_received_at | email | role |
 *                            page | title | referrer
 */

var SHEET_ID = '1qBOTsZpMHaZ2C5Fh_zemNONvpUzFuTg6tLvRsiqyAHg';
var TRACKING_SHEET = 'Tracking';
var PAGEVIEWS_SHEET = 'PageViews';
var PASS_THRESHOLD = 0.70;

// ─── Acquisition Skills Practice — attendance tracking ─────────────────────
// Roster lives in a separate workbook owned by ops; the deploying user
// must have at least Viewer access to ROSTER_SHEET_ID for reads to succeed.
var ROSTER_SHEET_ID            = '1_eC6wN-PKDZCHy89uoXZYYQp7W2dbitEwuRRmMriTkA';
var ROSTER_TAB                 = 'ActiveEmployees';
var SKILLS_ATTENDANCE_SHEET    = 'SkillsAttendance';
var SKILLS_SESSIONS_SHEET      = 'SkillsSessions';
// Legacy attendance workbook (the original Tue/Thu Google Sheet) — read once
// during a manual migration via migrateLegacySkillsAttendance_().
var LEGACY_ATTENDANCE_SHEET_ID = '1x9QESI4TGuhw8Ijk1-5ihr7exyZFN4sDLZck9HwHNpc';

// ─── Routing ────────────────────────────────────────────────────────────────

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ success: false, error: 'invalid_json' });
  }

  switch (payload && payload.type) {
    case 'KC':
    case 'OLT':         return handleCompletion_(payload);
    case 'PV':          return handlePageView_(payload);
    case 'ATT_SET':     return handleAttendanceSet_(payload);
    case 'ATT_ADD_DATE':return handleAttendanceAddDate_(payload);
    default:            return jsonOut_({ success: false, error: 'unknown_type' });
  }
}

// JSONP-only GET so the gate page can fetch a user's completions without
// running into Apps Script's CORS redirect.
//
// Modes:
//   default            — pass ?email=<addr> to get that user's completions
//   admin              — pass ?mode=admin to get every completion row (admin.html)
//   skills_attendance  — pass ?mode=skills_attendance to load roster + sessions
//                        + attendance marks (acq_skills_practice.html)
function doGet(e) {
  var mode  = String((e.parameter && e.parameter.mode) || '').toLowerCase().trim();
  var email = String((e.parameter && e.parameter.email) || '').toLowerCase().trim();
  var cb    = String((e.parameter && e.parameter.callback) || '');

  if (mode === 'skills_attendance') {
    var payload;
    try {
      payload = { success: true, data: getSkillsAttendance_() };
    } catch (err) {
      payload = { success: false, error: String(err && err.message || err) };
    }
    return jsonpOrJson_(payload, cb);
  }

  var completions;
  if (mode === 'admin') {
    completions = getAllCompletions_();
  } else if (email) {
    completions = getCompletionsForEmail_(email);
  } else {
    completions = [];
  }

  return jsonpOrJson_({ success: true, completions: completions }, cb);
}

function jsonpOrJson_(obj, cb) {
  var json = JSON.stringify(obj);
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Knowledge-check + OLT submissions ──────────────────────────────────────

function handleCompletion_(p) {
  var sheet = getOrCreateSheet_(TRACKING_SHEET, [
    'email', 'module_id', 'module_name', 'type',
    'score_pct', 'passed', 'completed_at', 'role', 'class'
  ]);

  var scorePct = Number(p.score_pct) || 0;
  sheet.appendRow([
    String(p.email || '').toLowerCase().trim(),
    String(p.module_id || ''),
    String(p.module_name || ''),
    String(p.type || ''),
    scorePct,
    scorePct >= (PASS_THRESHOLD * 100),
    new Date(),
    String(p.role || ''),
    ''  // class — populated manually / by other tooling
  ]);
  return jsonOut_({ success: true });
}

function getCompletionsForEmail_(email) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TRACKING_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var idx = {
    email:       header.indexOf('email'),
    module_id:   header.indexOf('module_id'),
    module_name: header.indexOf('module_name'),
    type:        header.indexOf('type'),
    score_pct:   header.indexOf('score_pct'),
    passed:      header.indexOf('passed')
  };
  if (idx.email < 0 || idx.module_id < 0) return [];

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(row[idx.email] || '').toLowerCase().trim() !== email) continue;
    if (!row[idx.module_id]) continue;  // skip placeholder/empty rows
    out.push({
      module_id:   row[idx.module_id],
      module_name: row[idx.module_name],
      type:        row[idx.type],
      score_pct:   Number(row[idx.score_pct]) || 0,
      passed:      row[idx.passed] === true || String(row[idx.passed]).toUpperCase() === 'TRUE'
    });
  }
  return out;
}

// Returns every completion row, with the extra fields (email, role, class)
// the admin dashboard relies on for grouping and filtering.
function getAllCompletions_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TRACKING_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var idx = {
    email:       header.indexOf('email'),
    module_id:   header.indexOf('module_id'),
    module_name: header.indexOf('module_name'),
    type:        header.indexOf('type'),
    score_pct:   header.indexOf('score_pct'),
    passed:      header.indexOf('passed'),
    role:        header.indexOf('role'),
    cls:         header.indexOf('class')
  };
  if (idx.email < 0 || idx.module_id < 0) return [];

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var email = String(row[idx.email] || '').toLowerCase().trim();
    if (!email) continue;
    if (!row[idx.module_id]) continue;
    out.push({
      email:       email,
      module_id:   row[idx.module_id],
      module_name: row[idx.module_name],
      type:        row[idx.type],
      score_pct:   Number(row[idx.score_pct]) || 0,
      passed:      row[idx.passed] === true || String(row[idx.passed]).toUpperCase() === 'TRUE',
      role:        idx.role >= 0 ? String(row[idx.role] || '') : '',
      class:       idx.cls  >= 0 ? formatClassValue_(row[idx.cls]) : ''
    });
  }
  return out;
}

// Class cells may be Date objects (Sheets auto-parses things like "4/8/2026")
// or plain strings. Normalize Dates to M/D/YYYY so the admin dropdown shows
// a clean date instead of "Wed Apr 08 2026 00:00:00 GMT-0500 (...)".
function formatClassValue_(v) {
  if (v instanceof Date) {
    var tz = Session.getScriptTimeZone() || 'America/Chicago';
    return Utilities.formatDate(v, tz, 'M/d/yyyy');
  }
  return String(v || '').trim();
}

// ─── Page-view tracking ─────────────────────────────────────────────────────

function handlePageView_(p) {
  var sheet = getOrCreateSheet_(PAGEVIEWS_SHEET, [
    'client_ts', 'server_received_at', 'email', 'role',
    'page', 'title', 'referrer'
  ]);

  sheet.appendRow([
    p.ts ? new Date(Number(p.ts)) : '',
    new Date(),
    String(p.email || '').toLowerCase().trim(),
    String(p.role || ''),
    String(p.page || ''),
    String(p.title || ''),
    String(p.referrer || '')
  ]);
  return jsonOut_({ success: true });
}

// ─── Acq Skills Practice — attendance ───────────────────────────────────────
//
// Storage model (long-format, one row per mark):
//   SkillsAttendance: employee_key | employee_name | session_date | present
//                     | recorded_at | recorded_by
//   SkillsSessions:   session_date | created_at | created_by | notes
//
// session_date is normalized to YYYY-MM-DD on the way in. employee_key is the
// roster row's email when present, otherwise a slug of the name.
//
// Reads aggregate the latest mark per (employee_key, session_date) so toggles
// can be append-only (no row mutation) and stay race-free under concurrent
// writes.

function getSkillsAttendance_() {
  return {
    roster:   readAcqRoster_(),
    sessions: readSkillsSessions_(),
    marks:    readLatestSkillsMarks_(),
    fetched_at: new Date().toISOString()
  };
}

function readAcqRoster_() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  } catch (err) {
    throw new Error('Cannot open roster sheet — ensure the deploying user has Viewer access to ' + ROSTER_SHEET_ID + '. Underlying: ' + err);
  }
  var sheet = ss.getSheetByName(ROSTER_TAB);
  if (!sheet) throw new Error('Roster tab "' + ROSTER_TAB + '" not found in roster workbook.');
  if (sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function(h){ return String(h || '').trim().toLowerCase(); });

  function findCol() {
    for (var a = 0; a < arguments.length; a++) {
      var i = header.indexOf(String(arguments[a]).toLowerCase());
      if (i >= 0) return i;
    }
    return -1;
  }
  var iName    = findCol('employee name', 'name', 'full name');
  var iEmail   = findCol('email', 'work email', 'rebuilt email', 'company email');
  var iTeam    = findCol('team', 'department', 'function');
  var iRole    = findCol('role', 'title', 'position');
  var iActive  = findCol('active', 'status', 'employment status');
  var iHire    = findCol('hire date', 'start date', 'date of hire', 'hired', 'hire_date', 'start_date');
  var iManager = findCol('manager', 'reports to', 'direct manager', 'supervisor', 'reports_to');

  if (iName < 0)  throw new Error('Roster missing an "Employee Name" column.');
  if (iTeam < 0)  throw new Error('Roster missing a "Team" column.');

  var seen = {};
  var out  = [];
  for (var r = 1; r < values.length; r++) {
    var row  = values[r];
    var team = String(row[iTeam] || '').trim().toLowerCase();
    if (team !== 'acquisition') continue;

    if (iActive >= 0) {
      var status = String(row[iActive] || '').trim().toLowerCase();
      if (status && status !== 'active' && status !== 'true' && status !== 'yes') continue;
    }

    var name  = String(row[iName] || '').trim();
    if (!name) continue;
    var email = iEmail >= 0 ? String(row[iEmail] || '').toLowerCase().trim() : '';
    var key   = email || slugifyName_(name);
    if (seen[key]) continue;
    seen[key] = true;

    out.push({
      key:       key,
      name:      name,
      email:     email,
      role:      iRole >= 0 ? String(row[iRole] || '').trim() : '',
      hire_date: iHire >= 0 ? normalizeDateStr_(row[iHire]) : '',
      manager:   iManager >= 0 ? String(row[iManager] || '').trim() : ''
    });
  }
  out.sort(function(a, b) { return a.name.localeCompare(b.name); });
  return out;
}

function readSkillsSessions_() {
  var sheet = getOrCreateSheet_(SKILLS_SESSIONS_SHEET, [
    'session_date', 'created_at', 'created_by', 'notes'
  ]);
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var seen = {};
  var out  = [];
  for (var i = 1; i < values.length; i++) {
    var d = normalizeDateStr_(values[i][0]);
    if (!d || seen[d]) continue;
    seen[d] = true;
    out.push({
      session_date: d,
      created_by:   String(values[i][2] || ''),
      notes:        String(values[i][3] || '')
    });
  }
  out.sort(function(a, b) { return a.session_date < b.session_date ? -1 : a.session_date > b.session_date ? 1 : 0; });
  return out;
}

// Returns one record per (employee_key, session_date) — the most recent mark
// wins. Sessions referenced only in attendance (not yet in SkillsSessions) are
// still surfaced so the UI can render them as columns.
function readLatestSkillsMarks_() {
  var sheet = getOrCreateSheet_(SKILLS_ATTENDANCE_SHEET, [
    'employee_key', 'employee_name', 'session_date', 'present',
    'recorded_at', 'recorded_by'
  ]);
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();

  // Bucket rows by (key, date), keeping the row with the latest recorded_at.
  var bucket = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var key = String(row[0] || '').toLowerCase().trim();
    var d   = normalizeDateStr_(row[2]);
    if (!key || !d) continue;
    var ts  = parseTs_(row[4]);
    var bk  = key + '|' + d;
    var prev = bucket[bk];
    if (!prev || ts >= prev.ts) {
      bucket[bk] = {
        key:    key,
        name:   String(row[1] || ''),
        date:   d,
        present:row[3] === true || String(row[3]).toLowerCase() === 'true',
        ts:     ts
      };
    }
  }
  return Object.keys(bucket).map(function(k) {
    var b = bucket[k];
    return {
      key:          b.key,
      employee_name:b.name,
      session_date: b.date,
      present:      b.present
    };
  });
}

function handleAttendanceSet_(p) {
  var key  = String(p.key  || p.email_key || '').toLowerCase().trim();
  var name = String(p.name || p.employee_name || '').trim();
  var date = normalizeDateStr_(p.session_date);
  var present = p.present === true || String(p.present).toLowerCase() === 'true';
  var by   = String(p.recorded_by || '').toLowerCase().trim();

  if (!key || !date) return jsonOut_({ success: false, error: 'missing_key_or_date' });

  var sheet = getOrCreateSheet_(SKILLS_ATTENDANCE_SHEET, [
    'employee_key', 'employee_name', 'session_date', 'present',
    'recorded_at', 'recorded_by'
  ]);
  sheet.appendRow([key, name, date, present, new Date(), by]);

  // Implicit session creation: if marking a new date, also register it as a
  // session so the column persists even after the mark is toggled off.
  ensureSession_(date, by);

  return jsonOut_({ success: true });
}

function handleAttendanceAddDate_(p) {
  var date = normalizeDateStr_(p.session_date);
  var by   = String(p.recorded_by || '').toLowerCase().trim();
  if (!date) return jsonOut_({ success: false, error: 'invalid_date' });
  ensureSession_(date, by, String(p.notes || ''));
  return jsonOut_({ success: true, session_date: date });
}

function ensureSession_(date, by, notes) {
  var sheet = getOrCreateSheet_(SKILLS_SESSIONS_SHEET, [
    'session_date', 'created_at', 'created_by', 'notes'
  ]);
  if (sheet.getLastRow() >= 2) {
    var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      if (normalizeDateStr_(values[i][0]) === date) return;
    }
  }
  sheet.appendRow([date, new Date(), by || '', notes || '']);
}

// ─── Date helpers ───────────────────────────────────────────────────────────
function normalizeDateStr_(v) {
  if (!v) return '';
  if (v instanceof Date && !isNaN(v.getTime())) return formatYmd_(v);
  var s = String(v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return formatYmd_(d);
  return '';
}
function formatYmd_(d) {
  var tz = Session.getScriptTimeZone() || 'America/Chicago';
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}
function parseTs_(v) {
  if (v instanceof Date) return v.getTime();
  var n = Date.parse(String(v));
  return isNaN(n) ? 0 : n;
}
function slugifyName_(name) {
  return String(name).toLowerCase().trim()
    .replace(/[^a-z0-9]+/g, '.')
    .replace(/^\.+|\.+$/g, '');
}

// ─── One-time legacy import ─────────────────────────────────────────────────
// Run manually from the Apps Script editor after the new sheets exist:
//   1. In the function dropdown, pick migrateLegacySkillsAttendance → Run.
//   2. Approve the Drive scope prompt the first time.
// Idempotent: rows from a prior import are tagged recorded_by='legacy_import'
// and the function refuses to re-import if any such row already exists.
//
// If this returns 0, run inspectLegacyAttendance (also in the dropdown) to
// see what the legacy sheet's header row actually looks like — the heuristic
// expects a "Name"-ish first column and date-parseable column headers.
function migrateLegacySkillsAttendance() {
  return migrateLegacySkillsAttendance_();
}
function inspectLegacyAttendance() {
  var src = SpreadsheetApp.openById(LEGACY_ATTENDANCE_SHEET_ID);
  var report = src.getSheets().map(function(sheet) {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow === 0 || lastCol === 0) {
      return { tab: sheet.getName(), empty: true };
    }
    var rows = Math.min(lastRow, 5);
    var cols = Math.min(lastCol, 8);
    var sample = sheet.getRange(1, 1, rows, cols).getValues().map(function(r) {
      return r.map(function(v) {
        if (v instanceof Date) return 'Date(' + Utilities.formatDate(v, Session.getScriptTimeZone() || 'America/Chicago', 'yyyy-MM-dd') + ')';
        return v;
      });
    });
    return { tab: sheet.getName(), rows: lastRow, cols: lastCol, sample: sample };
  });
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function migrateLegacySkillsAttendance_() {
  var dst = getOrCreateSheet_(SKILLS_ATTENDANCE_SHEET, [
    'employee_key', 'employee_name', 'session_date', 'present',
    'recorded_at', 'recorded_by'
  ]);
  if (dst.getLastRow() >= 2) {
    var existing = dst.getRange(2, 6, dst.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]) === 'legacy_import') {
        throw new Error('Legacy import already ran. Clear SkillsAttendance rows where recorded_by=legacy_import and re-run.');
      }
    }
  }

  var src = SpreadsheetApp.openById(LEGACY_ATTENDANCE_SHEET_ID);
  var sheets = src.getSheets();
  var imported = 0;
  var rosterByName = {};
  readAcqRoster_().forEach(function(r) {
    rosterByName[r.name.toLowerCase()] = r;
  });

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 2) continue;

    var values = sheet.getDataRange().getValues();
    var header = values[0];

    // Find the name column (heuristic: leftmost cell whose lowered header
    // looks like "name", "employee", or "associate"; falls back to col 0).
    var nameCol = 0;
    for (var c = 0; c < header.length; c++) {
      var h = String(header[c] || '').toLowerCase().trim();
      if (h === 'name' || h === 'employee' || h === 'employee name' || h === 'associate' || h === 'rep') {
        nameCol = c;
        break;
      }
    }

    // Date columns: any header that parses as a Date.
    var dateCols = [];
    for (var c2 = 0; c2 < header.length; c2++) {
      if (c2 === nameCol) continue;
      var v = header[c2];
      var d = (v instanceof Date) ? v : new Date(v);
      if (v && !isNaN(d.getTime())) {
        dateCols.push({ col: c2, date: formatYmd_(d) });
      }
    }
    if (dateCols.length === 0) continue;

    for (var r = 1; r < values.length; r++) {
      var row  = values[r];
      var name = String(row[nameCol] || '').trim();
      if (!name) continue;

      var roster = rosterByName[name.toLowerCase()];
      var key    = roster ? roster.key : slugifyName_(name);

      for (var k = 0; k < dateCols.length; k++) {
        var cell = row[dateCols[k].col];
        var present = isPresentTruthy_(cell);
        if (present) {
          dst.appendRow([key, roster ? roster.name : name, dateCols[k].date, true, new Date(), 'legacy_import']);
          imported++;
        }
      }
      // Also register every encountered date as a session.
      for (var k2 = 0; k2 < dateCols.length; k2++) {
        ensureSession_(dateCols[k2].date, 'legacy_import');
      }
    }
  }

  return imported;
}
function isPresentTruthy_(v) {
  if (v === true) return true;
  if (typeof v === 'number') return v !== 0;
  var s = String(v || '').trim().toLowerCase();
  if (!s) return false;
  return s === 'x' || s === '✓' || s === 'y' || s === 'yes' || s === 'p' || s === 'present' || s === '1' || s === 'true';
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
