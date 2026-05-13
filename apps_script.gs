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
// People tagged Team=Acquisition in the roster but who shouldn't appear on
// the skills practice roll (e.g. acq leaders who manage the org but don't
// take attendance themselves). Lower-case email or full name.
var ACQ_ROSTER_EXCLUDE = [
  'patrick.solomon@rebuilt.com',
  'patrick solomon'
];
// Emails allowed to remove session dates from the skills practice roll.
// Removal is destructive (deletes the SkillsSessions row plus every mark
// recorded on that date), so it's restricted to acq leadership.
var SESSION_ADMIN_EMAILS = [
  'mike.spalding@rebuilt.com',
  'patrick.solomon@rebuilt.com'
];
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
    case 'ATT_SET':       return handleAttendanceSet_(payload);
    case 'ATT_ADD_DATE':  return handleAttendanceAddDate_(payload);
    case 'ATT_REMOVE_DATE': return handleAttendanceRemoveDate_(payload);
    default:              return jsonOut_({ success: false, error: 'unknown_type' });
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
  // Fuzzy fallback — return the first header that includes `needle` and none
  // of the `exclude` substrings (e.g. avoid grabbing "Manager Email" when
  // we want the manager's display name).
  function findColContaining(needle, exclude) {
    var ex = exclude || [];
    for (var i = 0; i < header.length; i++) {
      var h = header[i];
      if (h.indexOf(needle) < 0) continue;
      var skip = false;
      for (var k = 0; k < ex.length; k++) {
        if (h.indexOf(ex[k]) >= 0) { skip = true; break; }
      }
      if (!skip) return i;
    }
    return -1;
  }
  var iName    = findCol('employee name', 'name', 'full name');
  var iEmail   = findCol('email', 'work email', 'rebuilt email', 'company email');
  var iTeam    = findCol('team', 'department', 'function');
  var iRole    = findCol('role', 'title', 'position');
  var iActive  = findCol('active', 'status', 'employment status');
  var iHire    = findCol(
    'hire date', 'hire_date', 'hiredate', 'hired',
    'start date', 'start_date', 'startdate',
    'date of hire', 'date hired', 'employment start date', 'employment start',
    'original hire date'
  );
  if (iHire < 0) iHire = findColContaining('hire', ['anniversary']);
  var iManager = findCol(
    'manager', 'manager name', 'managers name', "manager's name",
    'reports to', 'reports_to', 'reportsto',
    'direct manager', 'direct supervisor', 'supervisor',
    'reporting manager',
    'employee manager', 'current manager', 'people manager'
  );
  if (iManager < 0) iManager = findColContaining('manager', ['email', ' id', '_id', 'phone']);
  if (iManager < 0) iManager = findColContaining('reports to', []);
  if (iManager < 0) iManager = findColContaining('supervisor', ['email', ' id', '_id']);

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
    if (ACQ_ROSTER_EXCLUDE.indexOf(email) >= 0) continue;
    if (ACQ_ROSTER_EXCLUDE.indexOf(name.toLowerCase()) >= 0) continue;
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

// Hard-deletes the SkillsSessions row for `date` and every SkillsAttendance
// row stamped with that date. Restricted to SESSION_ADMIN_EMAILS — clients
// also gate the UI, but this server-side check is the actual enforcement
// point for a destructive op.
function handleAttendanceRemoveDate_(p) {
  var date = normalizeDateStr_(p.session_date);
  var by   = String(p.recorded_by || '').toLowerCase().trim();
  if (!date) return jsonOut_({ success: false, error: 'invalid_date' });
  if (SESSION_ADMIN_EMAILS.indexOf(by) < 0) {
    return jsonOut_({ success: false, error: 'unauthorized' });
  }

  var removedSessions = 0;
  var removedMarks    = 0;

  var sess = getOrCreateSheet_(SKILLS_SESSIONS_SHEET, [
    'session_date', 'created_at', 'created_by', 'notes'
  ]);
  if (sess.getLastRow() >= 2) {
    var sv = sess.getDataRange().getValues();
    // Walk bottom-up so deleteRow doesn't shift unscanned rows.
    for (var i = sv.length - 1; i >= 1; i--) {
      if (normalizeDateStr_(sv[i][0]) === date) {
        sess.deleteRow(i + 1);
        removedSessions++;
      }
    }
  }

  var att = getOrCreateSheet_(SKILLS_ATTENDANCE_SHEET, [
    'employee_key', 'employee_name', 'session_date', 'present',
    'recorded_at', 'recorded_by'
  ]);
  if (att.getLastRow() >= 2) {
    var av = att.getDataRange().getValues();
    for (var j = av.length - 1; j >= 1; j--) {
      if (normalizeDateStr_(av[j][2]) === date) {
        att.deleteRow(j + 1);
        removedMarks++;
      }
    }
  }

  return jsonOut_({
    success: true,
    session_date: date,
    removed_sessions: removedSessions,
    removed_marks: removedMarks
  });
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
// Clears any rows previously written by migrateLegacySkillsAttendance so the
// import can be re-run from a clean slate. Manual-only — never invoked from
// the web app.
function resetLegacySkillsImport() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SKILLS_ATTENDANCE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('Nothing to reset.');
    return 0;
  }
  var values = sheet.getDataRange().getValues();
  var removed = 0;
  for (var r = values.length - 1; r >= 1; r--) {
    if (String(values[r][5] || '') === 'legacy_import') {
      sheet.deleteRow(r + 1);
      removed++;
    }
  }
  Logger.log('Removed ' + removed + ' legacy_import rows.');
  return removed;
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

// Diagnostic: prints the ActiveEmployees tab's header row plus 3 sample
// Acquisition rows so we can see which column is the manager column. Run
// from the Apps Script editor → Run dropdown → check Execution log.
function inspectRoster() {
  var ss = SpreadsheetApp.openById(ROSTER_SHEET_ID);
  var sheet = ss.getSheetByName(ROSTER_TAB);
  if (!sheet) {
    Logger.log('Tab "' + ROSTER_TAB + '" not found.');
    return null;
  }
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('HEADERS (' + lastCol + '):');
  for (var i = 0; i < header.length; i++) {
    Logger.log('  [' + i + '] ' + JSON.stringify(header[i]));
  }

  var values = sheet.getDataRange().getValues();
  var teamCol = -1;
  for (var c = 0; c < header.length; c++) {
    var h = String(header[c] || '').toLowerCase().trim();
    if (h === 'team' || h === 'department' || h === 'function') { teamCol = c; break; }
  }

  var samples = [];
  for (var r = 1; r < values.length && samples.length < 3; r++) {
    var team = teamCol >= 0 ? String(values[r][teamCol] || '').toLowerCase().trim() : '';
    if (teamCol < 0 || team === 'acquisition') {
      var row = {};
      for (var c2 = 0; c2 < header.length; c2++) {
        var key = String(header[c2] || ('col' + c2));
        row[key] = values[r][c2];
      }
      samples.push(row);
    }
  }
  Logger.log('SAMPLE ACQ ROWS (' + samples.length + '):');
  Logger.log(JSON.stringify(samples, null, 2));

  // Also report which column readAcqRoster_ currently picks as manager.
  try {
    var roster = readAcqRoster_();
    var withMgr = roster.filter(function(r){ return r.manager; }).length;
    Logger.log('readAcqRoster_ returned ' + roster.length + ' reps; ' + withMgr + ' have a manager value populated.');
    if (roster.length > 0) {
      Logger.log('First rep sample: ' + JSON.stringify(roster[0], null, 2));
    }
  } catch (err) {
    Logger.log('readAcqRoster_ failed: ' + err);
  }

  return { headers: header, samples: samples };
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
  Logger.log('LEGACY IMPORT — source workbook: ' + src.getName());

  var sheets = src.getSheets();
  Logger.log('  Tabs found: ' + sheets.map(function(s){return s.getName();}).join(', '));

  var rosterByName = {};   // exact-name match
  var rosterByLast = {};   // fallback: last token
  readAcqRoster_().forEach(function(r) {
    rosterByName[normalizeName_(r.name)] = r;
    var parts = r.name.trim().split(/\s+/);
    if (parts.length > 0) {
      var last = parts[parts.length - 1].toLowerCase();
      if (!rosterByLast[last]) rosterByLast[last] = r;
    }
  });
  Logger.log('  Roster: ' + Object.keys(rosterByName).length + ' Acquisition reps loaded');

  var imported = 0;
  var unmatched = {};

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var tabName = sheet.getName();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    Logger.log('  Tab "' + tabName + '": ' + lastRow + ' rows × ' + lastCol + ' cols');
    if (lastRow < 2 || lastCol < 2) { Logger.log('    skipped (too small)'); continue; }

    var values = sheet.getDataRange().getValues();

    // Auto-detect the header row by scanning the first 8 rows for the one
    // with the most date-parseable cells (≥ 2 dates required).
    var headerInfo = findHeaderRow_(values);
    if (!headerInfo) {
      Logger.log('    no header row with ≥ 2 date columns found — skipping');
      continue;
    }
    Logger.log('    header row: index ' + headerInfo.row + ' (' + headerInfo.dateCols.length + ' date columns)');
    Logger.log('    header preview: ' + JSON.stringify(values[headerInfo.row].slice(0, 12).map(formatCellForLog_)));

    var nameCol  = headerInfo.nameCol;
    var dateCols = headerInfo.dateCols;
    Logger.log('    name column index: ' + nameCol);
    Logger.log('    detected dates: ' + dateCols.map(function(d){return d.date;}).join(', '));

    var tabImported = 0;
    for (var r = headerInfo.row + 1; r < values.length; r++) {
      var row  = values[r];
      var name = String(row[nameCol] || '').trim();
      if (!name) continue;

      var roster = lookupRoster_(name, rosterByName, rosterByLast);
      if (!roster) {
        unmatched[name] = (unmatched[name] || 0) + 1;
      }
      var key = roster ? roster.key : slugifyName_(name);

      for (var k = 0; k < dateCols.length; k++) {
        var cell = row[dateCols[k].col];
        if (isPresentTruthy_(cell)) {
          dst.appendRow([key, roster ? roster.name : name, dateCols[k].date, true, new Date(), 'legacy_import']);
          imported++;
          tabImported++;
        }
      }
    }
    Logger.log('    imported from this tab: ' + tabImported);

    for (var k2 = 0; k2 < dateCols.length; k2++) {
      ensureSession_(dateCols[k2].date, 'legacy_import');
    }
  }

  var unmatchedNames = Object.keys(unmatched);
  if (unmatchedNames.length > 0) {
    Logger.log('Unmatched names (used slug fallback): ' + unmatchedNames.slice(0, 30).join(' | '));
  }
  Logger.log('TOTAL IMPORTED: ' + imported);
  return imported;
}

// Scan the first 8 rows for the row most likely to be the header — the one
// with the most date-parseable cells (Date instances or strings the strict
// parser accepts), requiring ≥ 2 dates so noise rows are rejected.
function findHeaderRow_(values) {
  var maxRows = Math.min(values.length, 8);
  var best = null;
  for (var r = 0; r < maxRows; r++) {
    var row = values[r];
    var dateCols = [];
    var nameCol = -1;
    for (var c = 0; c < row.length; c++) {
      var v = row[c];
      var d = strictParseDate_(v);
      if (d) {
        dateCols.push({ col: c, date: d });
      } else if (nameCol < 0) {
        var h = String(v || '').toLowerCase().trim();
        if (h === 'name' || h === 'employee' || h === 'employee name' ||
            h === 'associate' || h === 'rep' || h === 'team member') {
          nameCol = c;
        }
      }
    }
    if (dateCols.length >= 2) {
      if (nameCol < 0) {
        // No explicit name header; pick the first non-date column as the name
        // column (typically col A holds names in these sheets).
        for (var c2 = 0; c2 < row.length; c2++) {
          var isDateCol = dateCols.some(function(dc){ return dc.col === c2; });
          if (!isDateCol) { nameCol = c2; break; }
        }
        if (nameCol < 0) nameCol = 0;
      }
      if (!best || dateCols.length > best.dateCols.length) {
        best = { row: r, nameCol: nameCol, dateCols: dateCols };
      }
    }
  }
  return best;
}

// Stricter than `new Date(v)` — only accepts genuine Date objects or strings
// that match a date-like pattern. Rejects words like "May" that JS would
// otherwise coerce into a date.
function strictParseDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return formatYmd_(v);
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  // Require M/D, M-D, M/D/YY, M/D/YYYY, or YYYY-MM-DD shapes (optionally with
  // a leading day-of-week prefix like "Tue 5/7").
  var m = s.match(/^(?:(?:mon|tue|wed|thu|fri|sat|sun)\w*\s+)?(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?$/i);
  if (m) {
    var mo = parseInt(m[1], 10);
    var da = parseInt(m[2], 10);
    var yr = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    if (yr < 100) yr += 2000;
    if (mo >= 1 && mo <= 12 && da >= 1 && da <= 31) {
      return formatYmd_(new Date(yr, mo - 1, da));
    }
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return '';
}

function normalizeName_(name) {
  return String(name || '').toLowerCase().trim().replace(/\s+/g, ' ');
}

// Match by exact normalized name; if no hit, try "Last, First" → "First Last";
// finally try last-token match (e.g. legacy "Caceres" matches "Andrew Caceres").
function lookupRoster_(name, rosterByName, rosterByLast) {
  var norm = normalizeName_(name);
  if (rosterByName[norm]) return rosterByName[norm];

  var comma = norm.match(/^([^,]+),\s*(.+)$/);
  if (comma) {
    var flipped = (comma[2] + ' ' + comma[1]).trim();
    if (rosterByName[flipped]) return rosterByName[flipped];
  }

  var parts = norm.split(/\s+/);
  if (parts.length > 0) {
    var last = parts[parts.length - 1];
    if (rosterByLast[last]) return rosterByLast[last];
  }
  return null;
}

function formatCellForLog_(v) {
  if (v instanceof Date) return 'Date(' + formatYmd_(v) + ')';
  return v;
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
