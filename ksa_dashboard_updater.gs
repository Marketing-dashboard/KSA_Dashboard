// KSA Performance Dashboard — sheet → GitHub, daily refresh
// Sheet: https://docs.google.com/spreadsheets/d/12bR6eDJL2TS_nD913YuJRF5-x4-ekpur2I1whFT3gGM ("KSA_Raw_data")
// Repo:  https://github.com/Marketing-dashboard/KSA_Dashboard (index.html, served by GitHub Pages)
//
// Reads:  Google    — Date | Campaign Name | Campaign Type | Impressions | Spends | Clicks | Conversions
//         Meta      — Date | Campaign Name | Leads | Spend | Impressions | Clicks
//         Reference — Campaign Name | Model (brand)
//         Daywise Triggers — Day | Triggers | Model | Source (Google/Meta)
// Writes: the DATA_RAW / TRIGGERS_RAW / BUILD_TS blocks of index.html.
//
// Brand = Reference lookup by campaign name; if missing, the 2nd part of the name
// (KSA-<Brand>-...). The dashboard only counts trigger days inside its own date range.
//
// Every global is prefixed KSAD_ so this file can share the sheet's Apps Script project
// with other scripts.
//
// Setup (one time), in Extensions → Apps Script of the sheet above:
//   1. + → Script, name it "KSA Dashboard", paste this file.
//   2. Save a GitHub token (repo scope on Marketing-dashboard) with a temporary function,
//      run it once, then delete it:
//        function tmp() { KSAD_setToken('<github token>'); }
//   3. Run KSAD_update() once to check it works.
//   4. Run KSAD_createDailyTrigger() once → runs every day ~10-11 AM IST
//      (after the Google Ads script has filled the Google tab).

var KSAD_OWNER     = 'Marketing-dashboard';
var KSAD_REPO      = 'KSA_Dashboard';
var KSAD_FILE      = 'index.html';
var KSAD_TZ        = 'Asia/Kolkata';
var KSAD_TOKEN_KEY = 'KSAD_GITHUB_TOKEN';

function KSAD_setToken(token) {
  PropertiesService.getScriptProperties().setProperty(KSAD_TOKEN_KEY, token);
}

function KSAD_createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'KSAD_update') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('KSAD_update').timeBased().everyDays(1).atHour(10).inTimezone(KSAD_TZ).create();
}

function KSAD_update() {
  var token = PropertiesService.getScriptProperties().getProperty(KSAD_TOKEN_KEY);
  if (!token) throw new Error(KSAD_TOKEN_KEY + ' not set. Run KSAD_setToken(token) first.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var brandOf = KSAD_brandMap_(ss);
  var rows = [];

  // Google: [date, 'G', brand, campaign, spends, leads(conversions), impressions, clicks]
  KSAD_values_(ss, 'Google').forEach(function(r) {
    var d = KSAD_date_(r[0], tz), c = String(r[1] || '').trim();
    if (!d || !c) return;
    rows.push([d, 'G', brandOf(c), c, KSAD_num_(r[4]), KSAD_num_(r[6]), KSAD_num_(r[3]), KSAD_num_(r[5])]);
  });
  // Meta: Date | Campaign Name | Leads | Spend | Impressions | Clicks
  KSAD_values_(ss, 'Meta').forEach(function(r) {
    var d = KSAD_date_(r[0], tz), c = String(r[1] || '').trim();
    if (!d || !c) return;
    rows.push([d, 'M', brandOf(c), c, KSAD_num_(r[3]), KSAD_num_(r[2]), KSAD_num_(r[4]), KSAD_num_(r[5])]);
  });
  rows.sort(function(a, b) {
    return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : a[3] < b[3] ? -1 : a[3] > b[3] ? 1 : 0;
  });
  if (!rows.length) throw new Error('No rows found in the Google / Meta tabs; dashboard not changed.');

  var dataBlock = 'var DATA_RAW = [\n' + rows.map(function(r) {
    return '  [' + JSON.stringify(r[0]) + ',' + JSON.stringify(r[1]) + ',' + JSON.stringify(r[2]) + ',' +
           JSON.stringify(r[3]) + ',' + KSAD_round_(r[4]) + ',' + KSAD_round_(r[5]) + ',' + r[6] + ',' + r[7] + ']';
  }).join(',\n') + '\n];';
  var trigBlock = KSAD_triggersBlock_(ss, tz, rows[rows.length - 1][0].slice(0, 4));

  // Read the current file, swap the data blocks, write it back.
  var api = 'https://api.github.com/repos/' + KSAD_OWNER + '/' + KSAD_REPO + '/contents/' + KSAD_FILE;
  var hdr = {Authorization: 'token ' + token, Accept: 'application/vnd.github.v3+json'};
  var res = UrlFetchApp.fetch(api, {headers: hdr, muteHttpExceptions: true});
  if (res.getResponseCode() !== 200) throw new Error('GitHub read failed: HTTP ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
  var meta = JSON.parse(res.getContentText());
  var html = Utilities.newBlob(Utilities.base64Decode(meta.content.replace(/\n/g, ''))).getDataAsString('UTF-8');
  if (html.indexOf('var DATA_RAW = [') === -1) throw new Error('DATA_RAW block missing in ' + KSAD_FILE + '; not updated.');

  html = KSAD_replace_(html, 'var DATA_RAW = [', '];', dataBlock);
  html = KSAD_replace_(html, '// Triggered leads by day:', '];', trigBlock);
  html = html.replace(/var BUILD_TS = '[^']*';/,
    "var BUILD_TS = '" + Utilities.formatDate(new Date(), KSAD_TZ, 'yyyy-MM-dd HH:mm') + " IST';");

  var put = UrlFetchApp.fetch(api, {
    method: 'put', headers: hdr, contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({
      message: 'Data refresh: ' + Utilities.formatDate(new Date(), KSAD_TZ, 'yyyy-MM-dd'),
      content: Utilities.base64Encode(Utilities.newBlob(html).getBytes()),
      sha: meta.sha
    })
  });
  if (put.getResponseCode() >= 300) throw new Error('GitHub write failed: HTTP ' + put.getResponseCode() + ' ' + put.getContentText().slice(0, 200));
  Logger.log('Pushed ' + rows.length + ' rows (' + rows[0][0] + ' → ' + rows[rows.length - 1][0] + ') + triggers.');
}

// Rows of a tab without the header; [] if the tab doesn't exist.
function KSAD_values_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

// campaign name → brand, from Reference (A: Campaign Name, B: Model), else KSA-<Brand>-...
function KSAD_brandMap_(ss) {
  var map = {};
  KSAD_values_(ss, 'Reference').forEach(function(r) {
    var c = String(r[0] || '').trim().toLowerCase(), b = String(r[1] || '').trim();
    if (c && b) map[c] = b;
  });
  return function(campaign) {
    var b = map[campaign.toLowerCase()];
    if (b) return b;
    var p = campaign.split(/[-_ ]+/);
    var guess = p[0].toUpperCase() === 'KSA' && p[1] ? p[1] : p[0];
    return guess.charAt(0).toUpperCase() + guess.slice(1).toLowerCase();
  };
}

// Daywise Triggers tab (Day | Triggers | Model | Source) → TRIGGERS_RAW block,
// [date, source G/M, brand, triggers]. Rows with a blank / zero count are skipped.
// A day without a year (e.g. "30-Sep" kept as text) takes the year of defaultYear.
function KSAD_triggersBlock_(ss, tz, defaultYear) {
  var rows = [];
  KSAD_values_(ss, 'Daywise Triggers').forEach(function(r) {
    var d = KSAD_date_(r[0], tz, defaultYear), n = KSAD_num_(r[1]), b = String(r[2] || '').trim();
    var s = String(r[3] || '').trim().toLowerCase();
    if (!d || !b || !n) return;
    var src = s.indexOf('meta') === 0 || s.indexOf('facebook') === 0 || s === 'fb' ? 'M' : 'G';
    rows.push([d, src, b, n]);
  });
  rows.sort(function(a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[2] < b[2] ? -1 : a[2] > b[2] ? 1 : 0; });
  return '// Triggered leads by day: [date, source G/M, brand, triggers]\nvar TRIGGERS_RAW = [\n' + rows.map(function(r) {
    return '  [' + JSON.stringify(r[0]) + ',' + JSON.stringify(r[1]) + ',' + JSON.stringify(r[2]) + ',' + r[3] + ']';
  }).join(',\n') + '\n];';
}

function KSAD_replace_(html, startMarker, endMarker, block) {
  var s = html.indexOf(startMarker);
  if (s === -1) return html;
  var e = html.indexOf(endMarker, s + startMarker.length);
  if (e === -1) return html;
  return html.substring(0, s) + block + html.substring(e + endMarker.length);
}

var KSAD_MONTHS = {jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12};

// Date cell → 'yyyy-MM-dd' (Date objects, 13-Oct-2026, 13-Oct, 10/13/2026, 2026-10-13).
function KSAD_date_(v, tz, defaultYear) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim(), m;
  if (!s) return '';
  if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return m[1] + '-' + m[2] + '-' + m[3];
  if ((m = s.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/)) && KSAD_MONTHS[m[2].toLowerCase()])
    return m[3] + '-' + ('0' + KSAD_MONTHS[m[2].toLowerCase()]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  if (defaultYear && (m = s.match(/^(\d{1,2})-([A-Za-z]{3})$/)) && KSAD_MONTHS[m[2].toLowerCase()])
    return defaultYear + '-' + ('0' + KSAD_MONTHS[m[2].toLowerCase()]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  if ((m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/)))
    return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  return '';
}

function KSAD_num_(v) {
  var n = typeof v === 'number' ? v : parseFloat(String(v || '').replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function KSAD_round_(n) {
  return Math.round(n * 100) / 100;
}
