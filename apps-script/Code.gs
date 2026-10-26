/**
 * Padel tournament registration -> Google Sheet.
 *
 * Paste this file into the Apps Script editor of the Google Sheet that should
 * collect registrations (Extensions > Apps Script), then deploy it as a web app.
 * Full steps are in README.md.
 */

// Optional: cap the number of players. 0 = unlimited.
var MAX_PLAYERS = 24;

var SHEET_NAME = 'נרשמים';

var FIELDS = [
  ['timestamp', 'תאריך הרשמה'],
  ['name', 'שם מלא'],
  ['phone', 'טלפון'],
  ['email', 'אימייל'],
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(FIELDS.map(function (f) { return f[1]; }));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, FIELDS.length).setFontWeight('bold');
    sheet.setRightToLeft(true);
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Strips a leading =, +, - or @ so a submitted value can't run as a formula.
function clean_(value) {
  var s = String(value || '').trim().slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

// 050-123 4567, +972501234567, 0501234567 and 501234567 (a number cell that
// lost its leading zero) all become 501234567.
function normPhone_(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (d.indexOf('972') === 0) d = d.slice(3);
  return d.replace(/^0+/, '');
}

function normEmail_(v) {
  return String(v || '').trim().toLowerCase();
}

function findDuplicate_(sheet, p) {
  var n = sheet.getLastRow() - 1;
  if (n < 1) return null;
  var phoneCol = indexOf_('phone'), emailCol = indexOf_('email');
  var rows = sheet.getRange(2, 1, n, FIELDS.length).getValues();
  var phone = normPhone_(p.phone), email = normEmail_(p.email);
  for (var i = 0; i < rows.length; i++) {
    if (phone && normPhone_(rows[i][phoneCol]) === phone) return 'duplicate_phone';
    if (email && normEmail_(rows[i][emailCol]) === email) return 'duplicate_email';
  }
  return null;
}

function indexOf_(key) {
  for (var i = 0; i < FIELDS.length; i++) if (FIELDS[i][0] === key) return i;
  return -1;
}

function register_(p) {
  // Honeypot field: real users never fill it in.
  if (p.website) return { ok: true };

  var required = ['name', 'phone', 'email'];
  for (var i = 0; i < required.length; i++) {
    if (!String(p[required[i]] || '').trim()) {
      return { ok: false, error: 'missing_' + required[i] };
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    if (MAX_PLAYERS > 0 && sheet.getLastRow() - 1 >= MAX_PLAYERS) {
      return { ok: false, error: 'full' };
    }
    var dup = findDuplicate_(sheet, p);
    if (dup) return { ok: false, error: dup };
    var row = FIELDS.map(function (f) {
      if (f[0] === 'timestamp') return new Date();
      var v = clean_(p[f[0]]);
      // keep the phone as text so Sheets doesn't drop the leading zero
      return f[0] === 'phone' && v.charAt(0) !== "'" ? "'" + v : v;
    });
    sheet.appendRow(row);
    return { ok: true, count: sheet.getLastRow() - 1 };
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  return json_(register_((e && e.parameter) || {}));
}

// Returns the registration count, or registers when action=register.
// With ?callback=fn it answers as JSONP, which some browsers allow when they
// block reading a cross-site fetch.
function doGet(e) {
  var p = (e && e.parameter) || {};
  var result = p.action === 'register'
    ? register_(p)
    : { ok: true, count: getSheet_().getLastRow() - 1, max: MAX_PLAYERS };
  if (p.callback && /^[A-Za-z_$][\w$.]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(result);
}
