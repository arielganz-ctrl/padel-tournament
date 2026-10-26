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
var FIXED_HEADER = 'זוג קבוע עם';
var PAIRS_SHEET_NAME = 'זוגות';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎾 שיבוץ זוגות')
    .addItem('1. הוספת עמודת טיר', 'setupTiers')
    .addItem('2. יצירת קישורי דירוג לטיר 1', 'makeRankLinks')
    .addItem('3. חישוב זוגות', 'computePairs')
    .addToUi();
}

// Tier and link live right after the registration fields.
function tierCol_() { return FIELDS.length + 1; }
function linkCol_() { return FIELDS.length + 2; }
function fixedCol_() { return FIELDS.length + 3; }

function setupTiers() {
  var sheet = getSheet_();
  sheet.getRange(1, tierCol_(), 1, 3).setValues([[TIER_HEADER, LINK_HEADER, FIXED_HEADER]]).setFontWeight('bold');
  var rows = Math.max(sheet.getMaxRows() - 1, 1);
  var names = sheet.getRange(2, indexOf_('name') + 1, rows, 1);
  sheet.getRange(2, fixedCol_(), rows, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(names, true)
    .setHelpText('לזוג שידוע מראש: בחרו כאן את בן הזוג. הם ידרגו כרגיל אבל לא ייכנסו לחישוב.')
    .build());
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['1', '2'], true)
    .setAllowInvalid(false)
    .setHelpText('1 = טיר 1, 2 = טיר 2, ריק = לא בהגרלה')
    .build();
  sheet.getRange(2, tierCol_(), rows, 1).setDataValidation(rule);
  SpreadsheetApp.getUi().alert('נוספו העמודות "' + TIER_HEADER + '" ו-"' + FIXED_HEADER + '".\n' +
    'בחרו לכל שחקן 1 או 2 (ריק = לא בהגרלה).\n' +
    'לזוג שידוע מראש: בשורה של שחקן טיר 1 בחרו את בן הזוג בעמודה "' + FIXED_HEADER + '".\n' +
    'אחר כך הריצו "יצירת קישורי דירוג".');
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
  var values = sheet.getRange(2, 1, n, fixedCol_()).getValues();
  var nameCol = indexOf_('name');
  return values.map(function (r, i) {
    var link = String(r[linkCol_() - 1] || '');
    var m = link.match(/[?&]t=([A-Za-z0-9]+)/);
    return {
      row: i + 2,
      name: String(r[nameCol] || '').trim(),
      tier: String(r[tierCol_() - 1] || '').trim(),
      token: m ? m[1] : '',
      fixed: String(r[fixedCol_() - 1] || '').trim(),
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

/* ===================== חישוב זוגות ===================== */

// Everyone ranks honestly; we pick the assignment with the lowest total rank
// (= happiest on average). Ties go to the assignment with fewer bad matches.
function computePairs() {
  var ui = SpreadsheetApp.getUi();
  var players = readPlayers_();
  var byName = {};
  players.forEach(function (pl) { byName[pl.name] = pl; });

  // fixed pairs, whichever side the partner was filled in on
  var fixedPairs = [], taken = {};
  players.forEach(function (pl) {
    if (!pl.fixed || taken[pl.name]) return;
    var other = byName[pl.fixed];
    if (!other) return;
    var t1 = pl.tier === '1' ? pl : other.tier === '1' ? other : null;
    var t2 = t1 === pl ? other : pl;
    if (!t1 || t2.tier !== '2') return;
    fixedPairs.push([t1.name, t2.name]);
    taken[t1.name] = taken[t2.name] = true;
  });

  var t1 = players.filter(function (pl) { return pl.tier === '1' && !taken[pl.name]; });
  var t2 = players.filter(function (pl) { return pl.tier === '2' && !taken[pl.name]; });
  if (!t1.length || !t2.length) {
    ui.alert('אין מספיק שחקנים לחישוב. ודאו שסימנתם טיר 1 וטיר 2.');
    return;
  }

  // rankings by token
  var rankSheet = getRankSheet_();
  var rankings = {};
  var allT2 = tier2Names_(players);
  var nr = rankSheet.getLastRow() - 1;
  if (nr > 0) {
    rankSheet.getRange(2, 1, nr, 3 + allT2.length).getValues().forEach(function (r) {
      rankings[String(r[2])] = r.slice(3).map(String).filter(function (v) { return v; });
    });
  }

  var missing = [];
  var n = Math.max(t1.length, t2.length);
  var BIG = 100000;
  var cost = [], rankOf = [];
  for (var i = 0; i < n; i++) {
    cost.push([]); rankOf.push([]);
    var list = i < t1.length ? rankings[t1[i].token] : null;
    if (i < t1.length && !list) missing.push(t1[i].name);
    for (var j = 0; j < n; j++) {
      if (i >= t1.length || j >= t2.length) { cost[i].push(0); rankOf[i].push(null); continue; }
      var r = list ? list.indexOf(t2[j].name) + 1 : 0;
      if (!r) r = allT2.length; // unranked counts as last place
      rankOf[i].push(list ? r : null);
      cost[i].push(r * BIG + r * r);
    }
  }

  var assign = hungarian_(cost);

  var out = [['זוג', 'שחקן טיר 1', 'שחקן טיר 2', 'מקום בדירוג של טיר 1', 'סוג']];
  var num = 1, sum = 0, cnt = 0, firsts = 0;
  for (var a = 0; a < t1.length; a++) {
    var b = assign[a];
    if (b < t2.length) {
      var rk = rankOf[a][b];
      out.push([num++, t1[a].name, t2[b].name, rk || '-', 'הגרלה']);
      if (rk) { sum += rk; cnt++; if (rk === 1) firsts++; }
    } else {
      out.push(['', t1[a].name, '(אין בן זוג)', '', 'הגרלה']);
    }
  }
  for (var c = 0; c < t2.length; c++) {
    if (assign.indexOf(c) === -1 || assign.indexOf(c) >= t1.length) out.push(['', '(אין בן זוג)', t2[c].name, '', 'הגרלה']);
  }
  fixedPairs.forEach(function (fp) {
    var list = rankings[(byName[fp[0]] || {}).token];
    var rk = list ? list.indexOf(fp[1]) + 1 : 0;
    out.push([num++, fp[0], fp[1], rk || '-', 'זוג קבוע']);
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PAIRS_SHEET_NAME) || ss.insertSheet(PAIRS_SHEET_NAME);
  sheet.clear();
  sheet.setRightToLeft(true);
  sheet.getRange(1, 1, out.length, 5).setValues(out);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 5);
  ss.setActiveSheet(sheet);

  var msg = 'החישוב הסתיים. הזוגות בלשונית "' + PAIRS_SHEET_NAME + '".\n\n' +
    'זוגות בהגרלה: ' + Math.min(t1.length, t2.length) + ', זוגות קבועים: ' + fixedPairs.length + '\n' +
    (cnt ? 'מקום ממוצע שכל שחקן טיר 1 קיבל: ' + (sum / cnt).toFixed(2) + '\n' +
      'קיבלו את הבחירה הראשונה שלהם: ' + firsts + '\n' : '');
  if (missing.length) msg += '\n⚠️ עוד לא דירגו (חושבו כאילו אין להם העדפה): ' + missing.join(', ');
  if (t1.length !== t2.length) msg += '\n⚠️ מספר שחקני טיר 1 וטיר 2 בהגרלה לא שווה (' + t1.length + ' מול ' + t2.length + ').';
  ui.alert(msg);
}

// Hungarian algorithm, square cost matrix -> assign[row] = col (min total).
function hungarian_(a) {
  var n = a.length, INF = 1e18;
  var u = [], v = [], p = [], way = [];
  for (var i = 0; i <= n; i++) { u.push(0); v.push(0); p.push(0); way.push(0); }
  for (i = 1; i <= n; i++) {
    p[0] = i;
    var j0 = 0, minv = [], used = [];
    for (var j = 0; j <= n; j++) { minv.push(INF); used.push(false); }
    do {
      used[j0] = true;
      var i0 = p[j0], delta = INF, j1 = 0;
      for (j = 1; j <= n; j++) {
        if (used[j]) continue;
        var cur = a[i0 - 1][j - 1] - u[i0] - v[j];
        if (cur < minv[j]) { minv[j] = cur; way[j] = j0; }
        if (minv[j] < delta) { delta = minv[j]; j1 = j; }
      }
      for (j = 0; j <= n; j++) {
        if (used[j]) { u[p[j]] += delta; v[j] -= delta; } else { minv[j] -= delta; }
      }
      j0 = j1;
    } while (p[j0] !== 0);
    do { var jw = way[j0]; p[j0] = p[jw]; j0 = jw; } while (j0);
  }
  var assign = [];
  for (j = 1; j <= n; j++) assign[p[j] - 1] = j - 1;
  return assign;
}
