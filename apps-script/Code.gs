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

function doPost(e) {
  var p = (e && e.parameter) || {};

  // Honeypot field: real users never fill it in.
  if (p.website) return json_({ ok: true });

  var required = ['name', 'phone', 'email'];
  for (var i = 0; i < required.length; i++) {
    if (!String(p[required[i]] || '').trim()) {
      return json_({ ok: false, error: 'missing_' + required[i] });
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getSheet_();
    if (MAX_PLAYERS > 0 && sheet.getLastRow() - 1 >= MAX_PLAYERS) {
      return json_({ ok: false, error: 'full' });
    }
    var row = FIELDS.map(function (f) {
      return f[0] === 'timestamp' ? new Date() : clean_(p[f[0]]);
    });
    sheet.appendRow(row);
    return json_({ ok: true, count: sheet.getLastRow() - 1 });
  } finally {
    lock.releaseLock();
  }
}

// Lets the page show how many players are registered.
function doGet() {
  var sheet = getSheet_();
  return json_({ ok: true, count: sheet.getLastRow() - 1, max: MAX_PLAYERS });
}
