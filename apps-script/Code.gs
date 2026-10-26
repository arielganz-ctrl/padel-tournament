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
  var result;
  if (p.action === 'register') result = register_(p);
  else if (p.action === 'rankinfo') result = rankInfo_(p.t);
  else if (p.action === 'rank') result = rankSubmit_(p.t, p.order);
  else result = { ok: true, count: getSheet_().getLastRow() - 1, max: MAX_PLAYERS };
  if (p.callback && /^[A-Za-z_$][\w$.]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(result);
}

/* ===================== שיבוץ זוגות (דירוג טיר 1 ← טיר 2) ===================== */

var RANK_PAGE_URL = 'https://arielganz-ctrl.github.io/padel-tournament/rank.html';
var RANK_SHEET_NAME = 'דירוגים';
var TIER_HEADER = 'טיר (1/2)';
var LINK_HEADER = 'קישור לדירוג';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎾 שיבוץ זוגות')
    .addItem('1. הוספת עמודת טיר', 'setupTiers')
    .addItem('2. יצירת קישורי דירוג לטיר 1', 'makeRankLinks')
    .addToUi();
}

// Tier and link live right after the registration fields.
function tierCol_() { return FIELDS.length + 1; }
function linkCol_() { return FIELDS.length + 2; }

function setupTiers() {
  var sheet = getSheet_();
  sheet.getRange(1, tierCol_(), 1, 2).setValues([[TIER_HEADER, LINK_HEADER]]).setFontWeight('bold');
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2'], true)
    .setAllowInvalid(false)
    .setHelpText('1 = טיר 1, 2 = טיר 2, ריק = לא בהגרלה')
    .build();
  sheet.getRange(2, tierCol_(), Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  SpreadsheetApp.getUi().alert('נוספה עמודת "' + TIER_HEADER + '". בחרו לכל שחקן 1 או 2 (ריק = לא בהגרלה), ואז הריצו "יצירת קישורי דירוג".');
}

function makeRankLinks() {
  var sheet = getSheet_();
  var players = readPlayers_();
  var made = 0;
  players.forEach(function (pl) {
    if (pl.tier === '1' && !pl.token) {
      var token = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
      sheet.getRange(pl.row, linkCol_()).setValue(RANK_PAGE_URL + '?t=' + token);
      made++;
    }
  });
  var t1 = players.filter(function (pl) { return pl.tier === '1'; }).length;
  var t2 = players.filter(function (pl) { return pl.tier === '2'; }).length;
  SpreadsheetApp.getUi().alert('נוצרו ' + made + ' קישורים חדשים.\nטיר 1: ' + t1 + ' שחקנים, טיר 2: ' + t2 + ' שחקנים.\nשלחו לכל שחקן טיר 1 את הקישור שבשורה שלו.');
}

function readPlayers_() {
  var sheet = getSheet_();
  var n = sheet.getLastRow() - 1;
  if (n < 1) return [];
  var values = sheet.getRange(2, 1, n, linkCol_()).getValues();
  var nameCol = indexOf_('name');
  return values.map(function (r, i) {
    var link = String(r[linkCol_() - 1] || '');
    var m = link.match(/[?&]t=([A-Za-z0-9]+)/);
    return {
      row: i + 2,
      name: String(r[nameCol] || '').trim(),
      tier: String(r[tierCol_() - 1] || '').trim(),
      token: m ? m[1] : '',
    };
  }).filter(function (pl) { return pl.name; });
}

function tier2Names_(players) {
  return players.filter(function (pl) { return pl.tier === '2'; })
    .map(function (pl) { return pl.name; });
}

function findRanker_(players, token) {
  if (!token) return null;
  for (var i = 0; i < players.length; i++) {
    if (players[i].tier === '1' && players[i].token === String(token)) return players[i];
  }
  return null;
}

function getRankSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(RANK_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RANK_SHEET_NAME);
    sheet.setRightToLeft(true);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Columns: עודכן | שחקן טיר 1 | קוד | מקום 1 | מקום 2 | ...
function findRankRow_(sheet, token) {
  var n = sheet.getLastRow() - 1;
  if (n < 1) return null;
  var tokens = sheet.getRange(2, 3, n, 1).getValues();
  for (var i = 0; i < tokens.length; i++) {
    if (String(tokens[i][0]) === token) return i + 2;
  }
  return null;
}

function rankInfo_(token) {
  var players = readPlayers_();
  var me = findRanker_(players, token);
  if (!me) return { ok: false, error: 'bad_token' };
  var names = tier2Names_(players);
  var existing = null;
  var sheet = getRankSheet_();
  var row = findRankRow_(sheet, me.token);
  if (row) {
    var saved = sheet.getRange(row, 4, 1, Math.max(names.length, 1)).getValues()[0]
      .map(String).filter(function (v) { return v; });
    // only reuse a saved ranking if it still covers exactly the current tier 2
    if (sameSet_(saved, names)) existing = saved;
  }
  return { ok: true, name: me.name, players: names, ranking: existing };
}

function rankSubmit_(token, orderJson) {
  var order;
  try { order = JSON.parse(orderJson); } catch (e) { return { ok: false, error: 'bad_order' }; }
  if (!Array.isArray(order)) return { ok: false, error: 'bad_order' };
  order = order.map(function (v) { return String(v).trim(); });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var players = readPlayers_();
    var me = findRanker_(players, token);
    if (!me) return { ok: false, error: 'bad_token' };
    if (!sameSet_(order, tier2Names_(players))) return { ok: false, error: 'players_changed' };

    var sheet = getRankSheet_();
    var header = ['עודכן', 'שחקן טיר 1', 'קוד'];
    for (var i = 1; i <= order.length; i++) header.push('מקום ' + i);
    sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');

    var rowValues = [new Date(), me.name, me.token].concat(order.map(clean_));
    var row = findRankRow_(sheet, me.token) || sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, sheet.getMaxColumns()).clearContent();
    sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function sameSet_(a, b) {
  if (a.length !== b.length) return false;
  var x = a.slice().sort(), y = b.slice().sort();
  for (var i = 0; i < x.length; i++) if (x[i] !== y[i]) return false;
  return true;
}
