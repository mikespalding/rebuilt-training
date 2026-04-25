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
 * Sheets are auto-created on first event:
 *   - "Completions" — knowledge-check submissions  (type: "KC")
 *   - "PageViews"   — page-view tracking           (type: "PV")
 */

var PASS_THRESHOLD = 0.70;

// ─── Routing ────────────────────────────────────────────────────────────────

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ success: false, error: 'invalid_json' });
  }

  switch (payload && payload.type) {
    case 'KC': return handleKnowledgeCheck_(payload);
    case 'PV': return handlePageView_(payload);
    default:   return jsonOut_({ success: false, error: 'unknown_type' });
  }
}

// JSONP-only GET so the gate page can fetch a user's completions without
// running into Apps Script's CORS redirect.
function doGet(e) {
  var email = String((e.parameter && e.parameter.email) || '').toLowerCase().trim();
  var cb    = String((e.parameter && e.parameter.callback) || '');
  var body  = { success: true, completions: email ? getCompletionsForEmail_(email) : [] };
  var json  = JSON.stringify(body);

  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Knowledge-check submissions ────────────────────────────────────────────

function handleKnowledgeCheck_(p) {
  var sheet = getOrCreateSheet_('Completions', [
    'server_received_at',
    'email',
    'role',
    'module_id',
    'module_name',
    'score_pct',
    'passed'
  ]);

  var scorePct = Number(p.score_pct) || 0;
  sheet.appendRow([
    new Date(),
    String(p.email || '').toLowerCase().trim(),
    String(p.role || ''),
    String(p.module_id || ''),
    String(p.module_name || ''),
    scorePct,
    scorePct >= (PASS_THRESHOLD * 100)
  ]);
  return jsonOut_({ success: true });
}

function getCompletionsForEmail_(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Completions');
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getDataRange().getValues();
  var header = values[0];
  var idx = {
    email:       header.indexOf('email'),
    module_id:   header.indexOf('module_id'),
    module_name: header.indexOf('module_name'),
    score_pct:   header.indexOf('score_pct'),
    passed:      header.indexOf('passed')
  };

  var out = [];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (String(row[idx.email] || '').toLowerCase().trim() !== email) continue;
    out.push({
      module_id:   row[idx.module_id],
      module_name: row[idx.module_name],
      score_pct:   Number(row[idx.score_pct]) || 0,
      passed:      row[idx.passed] === true || row[idx.passed] === 'TRUE'
    });
  }
  return out;
}

// ─── Page-view tracking ─────────────────────────────────────────────────────

function handlePageView_(p) {
  var sheet = getOrCreateSheet_('PageViews', [
    'client_ts',
    'server_received_at',
    'email',
    'role',
    'page',
    'title',
    'referrer'
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

// ─── Helpers ────────────────────────────────────────────────────────────────

function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
