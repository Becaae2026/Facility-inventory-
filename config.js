/*******************************************************************************
 * BEC FACILITY DEPARTMENT — INVENTORY BACKEND
 * Google Apps Script web app bound to the Facility Inventory spreadsheet.
 *
 * INSTALL
 *   1. Open the sheet ▸ Extensions ▸ Apps Script
 *   2. Replace the default Code.gs with this file ▸ Save
 *   3. Set CFG.TOKEN below to a private string (same string goes in the dashboard)
 *   4. Run  install()  once and approve the permission prompt
 *   5. Deploy ▸ New deployment ▸ Web app
 *        Execute as     : Me
 *        Who has access : Anyone
 *      Copy the /exec URL into the dashboard
 *
 * API
 *   GET  ?action=bootstrap&token=...&callback=fn   → all master data + logs (JSONP)
 *   POST {action:'postIn'   , token, header, lines}
 *   POST {action:'postOut'  , token, header, lines}
 *   POST {action:'voidLine' , token, log, docNo, itemCode, reason}
 *   POST {action:'addItem'  , token, item}
 *   POST {action:'backupNow', token}
 *
 * RULES
 *   • Rows are only ever appended. Cancellations flip Status to VOID.
 *   • Only input columns are written. Calculated columns are left untouched.
 *   • Every write is mirrored to an append-only Audit Log tab.
 ******************************************************************************/

var CFG = {
  TOKEN: 'bec-aae-fac-Aa92702689',   // must match API_TOKEN in config.js — do not change without updating both
  TZ: 'Asia/Kuwait',
  BACKUP_ROOT: 'Facility Inventory Backups',
  DAILY_KEEP_DAYS: 95,      // rolling daily snapshots
  ARCHIVE_KEEP_DAYS: 2010,  // month-end archives ≈ 5 years 6 months  (requirement: ≥ 5 years)
  DOC_PREFIX_IN: 'GRN',
  DOC_PREFIX_OUT: 'ISS'
};

/* Tab resolution — tolerant of renaming. Report tabs are excluded on purpose. */
var TAB_ALIASES = {
  items:    ['item master', 'items', 'item'],
  stockIn:  ['stock in', 'stockin', 'goods received', 'receipts'],
  stockOut: ['stock out', 'stockout', 'issues', 'issued'],
  branches: ['branches and departments', 'branches', 'branch'],
  vendors:  ['vendors', 'vendor', 'suppliers'],
  lists:    ['dropdown lists', 'dropdown', 'lists'],
  audit:    ['audit log', 'audit'],
  users:    ['users', 'user accounts', 'logins']
};
var TAB_EXCLUDE = ['consumption', 'summary', 'year to date', 'position', 'dashboard', 'chart'];

var ANCHORS = {
  items: 'item code',
  stockIn: 'document no',
  stockOut: 'document no',
  branches: 'branch code',
  vendors: 'vendor name',
  lists: 'category',
  audit: 'timestamp',
  users: 'username'
};

var USER_HEADERS = ['Username', 'Full name', 'Role', 'Status', 'Salt', 'Password hash',
                    'Created', 'Last login', 'Force reset'];
var SESSION_HOURS = 12;
var CURRENT_USER = '';

var AUDIT_HEADERS = ['Timestamp', 'Action', 'Document no', 'Item code', 'Quantity',
                     'Counterparty', 'User', 'Detail'];

/* ============================== ENTRY POINTS ============================== */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try {
    guard(p.token);
    switch (p.action) {
      case 'bootstrap': out = ok(bootstrap()); break;
      case 'ping':      out = ok({ pong: true, at: stamp() }); break;
      default:          out = ok(bootstrap());
    }
  } catch (err) {
    out = fail(err);
  }
  return wrap(out, p.callback);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (x) {}
  var out;
  try {
    guard(body.token);
    CURRENT_USER = '';
    switch (body.action) {
      /* open to anyone holding the API token */
      case 'login':          out = ok(login(body)); break;
      case 'logout':         out = ok(logout(body)); break;

      /* any signed-in user */
      case 'changePassword': out = ok(changePassword(body)); break;
      case 'postIn':         requireSession(body); out = ok(postIn(body)); break;
      case 'postOut':        requireSession(body); out = ok(postOut(body)); break;
      case 'voidLine':       requireSession(body); out = ok(voidLine(body)); break;
      case 'addItem':        requireSession(body); out = ok(addItem(body)); break;

      /* administrators only */
      case 'listUsers':      out = ok(listUsers(body)); break;
      case 'addUser':        out = ok(addUser(body)); break;
      case 'setUserStatus':  out = ok(setUserStatus(body)); break;
      case 'resetPassword':  out = ok(resetPassword(body)); break;
      case 'setOpening':     out = ok(setOpening(body)); break;
      case 'clearOpening':   out = ok(clearOpening(body)); break;
      case 'rollover':       out = ok(rollover(body)); break;
      case 'backupNow':      requireAdmin(body); out = ok(dailyBackup(true)); break;

      default: throw new Error('Unknown action: ' + body.action);
    }
  } catch (err) {
    out = fail(err);
  }
  return wrap(out, body.callback);
}

function guard(token) {
  if (CFG.TOKEN && token !== CFG.TOKEN) throw new Error('Access token rejected.');
}
function ok(data)  { return { ok: true,  data: data, at: stamp() }; }
function fail(err) { return { ok: false, error: String(err && err.message || err) }; }

function wrap(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================== SHEET PLUMBING ============================ */

function book() { return SpreadsheetApp.getActiveSpreadsheet(); }

function tab(key) {
  var sheets = book().getSheets();
  var pool = sheets.filter(function (s) {
    var n = s.getName().toLowerCase();
    if (key !== 'items' && TAB_EXCLUDE.some(function (x) { return n.indexOf(x) > -1; })) return false;
    return true;
  });
  var aliases = TAB_ALIASES[key] || [];
  var hit = null;
  aliases.forEach(function (a) {
    if (hit) return;
    hit = pool.filter(function (s) { return s.getName().toLowerCase() === a; })[0] || null;
  });
  aliases.forEach(function (a) {
    if (hit) return;
    hit = pool.filter(function (s) { return s.getName().toLowerCase().indexOf(a) === 0; })[0] || null;
  });
  aliases.forEach(function (a) {
    if (hit) return;
    hit = pool.filter(function (s) { return s.getName().toLowerCase().indexOf(a) > -1; })[0] || null;
  });
  if (!hit && key === 'audit') hit = createAuditTab();
  if (!hit && key === 'users') hit = createUsersTab();
  if (!hit) throw new Error('Tab not found for "' + key + '". Expected something like: ' + aliases[0]);
  return hit;
}

/** Finds the header row by scanning the first 15 rows for the anchor label. */
function headerRow(sheet, anchor) {
  var scan = sheet.getRange(1, 1, Math.min(15, sheet.getMaxRows()), sheet.getMaxColumns()).getValues();
  for (var r = 0; r < scan.length; r++) {
    for (var c = 0; c < scan[r].length; c++) {
      if (String(scan[r][c]).trim().toLowerCase() === anchor) return r + 1;
    }
  }
  throw new Error('Header "' + anchor + '" not found on tab ' + sheet.getName());
}

/** Returns {sheet, hr, cols:{lowercased header -> 1-based column}, headers:[]} */
function grid(key) {
  var sheet = tab(key);
  var hr = headerRow(sheet, ANCHORS[key]);
  var headers = sheet.getRange(hr, 1, 1, sheet.getMaxColumns()).getValues()[0];
  var cols = {};
  headers.forEach(function (h, i) {
    var k = String(h).trim().toLowerCase();
    if (k && !cols[k]) cols[k] = i + 1;
  });
  return { sheet: sheet, hr: hr, cols: cols, headers: headers };
}

/** Column lookup that accepts partial header names. */
function col(g, name) {
  var k = name.toLowerCase();
  if (g.cols[k]) return g.cols[k];
  var found = 0;
  Object.keys(g.cols).forEach(function (h) {
    if (!found && h.indexOf(k) === 0) found = g.cols[h];
  });
  if (!found) {
    Object.keys(g.cols).forEach(function (h) {
      if (!found && h.indexOf(k) > -1) found = g.cols[h];
    });
  }
  return found || 0;
}

/** Last row that actually holds data in the given key column. */
function lastDataRow(g, keyCol) {
  var max = g.sheet.getMaxRows();
  if (max <= g.hr) return g.hr;
  var vals = g.sheet.getRange(g.hr + 1, keyCol, max - g.hr, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return g.hr + 1 + i;
  }
  return g.hr;
}

/** Reads a tab into array-of-objects keyed by lowercased header. */
function readTab(key, keyHeader) {
  var g = grid(key);
  var kc = col(g, keyHeader);
  var last = lastDataRow(g, kc);
  if (last <= g.hr) return [];
  var width = g.sheet.getMaxColumns();
  var vals = g.sheet.getRange(g.hr + 1, 1, last - g.hr, width).getDisplayValues();
  var heads = g.headers.map(function (h) { return String(h).trim().toLowerCase(); });
  var rows = [];
  vals.forEach(function (r, i) {
    if (String(r[kc - 1]).trim() === '') return;
    var o = { _row: g.hr + 1 + i };
    heads.forEach(function (h, c) { if (h) o[h] = r[c]; });
    rows.push(o);
  });
  return rows;
}

/** Writes a value only when the target cell is currently blank (protects formulas). */
function fillIfBlank(sheet, row, column, value) {
  if (!column) return;
  var cell = sheet.getRange(row, column);
  if (String(cell.getDisplayValue()).trim() === '' && String(cell.getFormula()) === '') {
    cell.setValue(value);
  }
}

function num(v) {
  if (v === null || v === undefined) return 0;
  var s = String(v).replace(/[, ]/g, '').replace(/^\u2013|^-$/, '');
  if (s === '' || s === '-') return 0;
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function stamp() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm:ss'); }
function today() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }
function who() {
  if (CURRENT_USER) return CURRENT_USER;
  var u = '';
  try { u = Session.getActiveUser().getEmail(); } catch (e) {}
  return u || 'dashboard';
}

/* ============================== ACCOUNTS ================================= */

function createUsersTab() {
  var s = book().insertSheet('Users');
  s.getRange(1, 1).setValue('Users').setFontWeight('bold');
  s.getRange(2, 1).setValue('Managed from the dashboard. Passwords are stored only as salted hashes — ' +
                            'they cannot be read back, only reset.');
  s.getRange(4, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]).setFontWeight('bold');
  s.setFrozenRows(4);
  s.setColumnWidth(5, 220); s.setColumnWidth(6, 340);
  s.hideColumns(5, 2);
  return s;
}

function sha256(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
function newSalt() { return Utilities.getUuid().replace(/-/g, ''); }
function hashOf(salt, password) { return sha256(salt + '|' + password + '|bec-facility'); }

function userRows() { return readTab('users', 'username'); }

function findUser(username) {
  var want = String(username || '').trim().toLowerCase();
  var rows = userRows();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['username']).trim().toLowerCase() === want) return rows[i];
  }
  return null;
}

function newSession(username, role) {
  var token = Utilities.getUuid() + Utilities.getUuid().slice(0, 8);
  var payload = { u: username, r: role, exp: Date.now() + SESSION_HOURS * 3600000 };
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(payload), SESSION_HOURS * 3600);
  PropertiesService.getScriptProperties().setProperty('s_' + token, JSON.stringify(payload));
  return token;
}

function readSession(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get('s_' + token);
  if (!raw) raw = PropertiesService.getScriptProperties().getProperty('s_' + token);
  if (!raw) return null;
  var p;
  try { p = JSON.parse(raw); } catch (e) { return null; }
  if (!p.exp || p.exp < Date.now()) { killSession(token); return null; }
  return p;
}

function killSession(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('s_' + token);
  PropertiesService.getScriptProperties().deleteProperty('s_' + token);
}

/** Every write goes through here. Returns the session, sets CURRENT_USER. */
function requireSession(body) {
  var s = readSession(body && body.session);
  if (!s) throw new Error('Your session has ended. Sign in again.');
  CURRENT_USER = s.u;
  return s;
}
function requireAdmin(body) {
  var s = requireSession(body);
  if (String(s.r).toLowerCase() !== 'admin') throw new Error('That action is for administrators only.');
  return s;
}

function login(body) {
  var u = findUser(body.username);
  var bad = 'Username or password is not right.';
  if (!u) throw new Error(bad);
  if (String(u['status'] || 'ACTIVE').toUpperCase() !== 'ACTIVE') throw new Error('That account is disabled.');
  if (hashOf(u['salt'], String(body.password || '')) !== String(u['password hash'])) throw new Error(bad);

  var g = grid('users');
  g.sheet.getRange(u._row, col(g, 'last login')).setValue(stamp());
  CURRENT_USER = u['username'];
  audit('SIGN IN', '', '', 0, u['role'], 'Signed in');

  return {
    session: newSession(u['username'], u['role']),
    user: {
      username: u['username'],
      name: u['full name'] || u['username'],
      role: String(u['role'] || 'staff').toLowerCase(),
      mustReset: String(u['force reset'] || '').toUpperCase() === 'YES'
    },
    expiresIn: SESSION_HOURS
  };
}

function logout(body) {
  var s = readSession(body.session);
  if (s) { CURRENT_USER = s.u; audit('SIGN OUT', '', '', 0, '', 'Signed out'); }
  killSession(body.session);
  return { out: true };
}

function changePassword(body) {
  var s = requireSession(body);
  var u = findUser(s.u);
  if (!u) throw new Error('Account not found.');
  if (hashOf(u['salt'], String(body.oldPassword || '')) !== String(u['password hash'])) {
    throw new Error('The current password is not right.');
  }
  if (String(body.newPassword || '').length < 8) throw new Error('Use at least 8 characters.');
  writePassword(u._row, body.newPassword, false);
  audit('PASSWORD', '', '', 0, '', 'Changed own password');
  return { changed: true };
}

function writePassword(row, password, forceReset) {
  var g = grid('users');
  var salt = newSalt();
  g.sheet.getRange(row, col(g, 'salt')).setValue(salt);
  g.sheet.getRange(row, col(g, 'password hash')).setValue(hashOf(salt, String(password)));
  g.sheet.getRange(row, col(g, 'force reset')).setValue(forceReset ? 'YES' : '');
}

function listUsers(body) {
  requireAdmin(body);
  return userRows().map(function (u) {
    return {
      username: u['username'], name: u['full name'] || '',
      role: String(u['role'] || 'staff').toLowerCase(),
      status: String(u['status'] || 'ACTIVE').toUpperCase(),
      created: u['created'] || '', lastLogin: u['last login'] || '',
      mustReset: String(u['force reset'] || '').toUpperCase() === 'YES'
    };
  });
}

function addUser(body) {
  requireAdmin(body);
  var name = String(body.username || '').trim();
  if (!/^[A-Za-z0-9._-]{3,}$/.test(name)) {
    throw new Error('Username needs 3 or more characters — letters, numbers, dot, dash or underscore.');
  }
  if (findUser(name)) throw new Error('"' + name + '" already exists.');
  if (String(body.password || '').length < 8) throw new Error('Use at least 8 characters for the password.');

  var g = grid('users');
  var row = lastDataRow(g, col(g, 'username')) + 1;
  var s = g.sheet;
  s.getRange(row, col(g, 'username')).setValue(name);
  s.getRange(row, col(g, 'full name')).setValue(body.fullName || name);
  s.getRange(row, col(g, 'role')).setValue(String(body.role || 'staff').toLowerCase() === 'admin' ? 'admin' : 'staff');
  s.getRange(row, col(g, 'status')).setValue('ACTIVE');
  s.getRange(row, col(g, 'created')).setValue(stamp());
  writePassword(row, body.password, true);
  audit('USER ADDED', '', '', 0, body.role || 'staff', 'Account ' + name + ' created');
  return { username: name };
}

function setUserStatus(body) {
  var s = requireAdmin(body);
  var u = findUser(body.username);
  if (!u) throw new Error('Account not found.');
  var want = String(body.status || '').toUpperCase() === 'ACTIVE' ? 'ACTIVE' : 'DISABLED';
  if (want === 'DISABLED') {
    if (String(u['username']).toLowerCase() === String(s.u).toLowerCase()) {
      throw new Error('You cannot disable your own account.');
    }
    var admins = userRows().filter(function (r) {
      return String(r['role']).toLowerCase() === 'admin' &&
             String(r['status'] || 'ACTIVE').toUpperCase() === 'ACTIVE';
    });
    if (admins.length <= 1 && String(u['role']).toLowerCase() === 'admin') {
      throw new Error('That is the last active administrator. Add another before disabling this one.');
    }
  }
  var g = grid('users');
  g.sheet.getRange(u._row, col(g, 'status')).setValue(want);
  audit('USER ' + want, '', '', 0, u['username'], 'Account set to ' + want);
  return { username: u['username'], status: want };
}

function resetPassword(body) {
  requireAdmin(body);
  var u = findUser(body.username);
  if (!u) throw new Error('Account not found.');
  if (String(body.newPassword || '').length < 8) throw new Error('Use at least 8 characters.');
  writePassword(u._row, body.newPassword, true);
  audit('PASSWORD RESET', '', '', 0, u['username'], 'Password reset by administrator');
  return { username: u['username'] };
}

/* ========================== OPENING STOCK ADMIN ========================= */

/** Set the opening figure on one item. */
function setOpening(body) {
  requireAdmin(body);
  var g = grid('items');
  var cc = col(g, 'item code'), oc = col(g, 'opening stock');
  var last = lastDataRow(g, cc);
  var vals = g.sheet.getRange(g.hr + 1, cc, last - g.hr, 1).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(body.code).trim()) {
      var row = g.hr + 1 + i;
      var was = g.sheet.getRange(row, oc).getDisplayValue();
      g.sheet.getRange(row, oc).setValue(num(body.qty));
      audit('OPENING SET', '', body.code, num(body.qty), '', 'Opening changed from ' + (was || '0') +
            ' to ' + num(body.qty) + (body.reason ? ' | ' + body.reason : ''));
      SpreadsheetApp.flush();
      return { code: body.code, opening: num(body.qty) };
    }
  }
  throw new Error('Item code not found: ' + body.code);
}

/** Zero every opening figure. Movements are untouched. */
function clearOpening(body) {
  requireAdmin(body);
  if (String(body.confirm) !== 'CLEAR') throw new Error('Type CLEAR to confirm.');
  dailyBackup(true);
  var g = grid('items');
  var cc = col(g, 'item code'), oc = col(g, 'opening stock');
  var last = lastDataRow(g, cc);
  var n = last - g.hr;
  if (n < 1) throw new Error('No items on the master.');
  var zeros = [];
  for (var i = 0; i < n; i++) zeros.push([0]);
  g.sheet.getRange(g.hr + 1, oc, n, 1).setValues(zeros);
  audit('OPENING CLEARED', '', '', 0, '', 'All ' + n + ' opening figures set to zero' +
        (body.reason ? ' | ' + body.reason : ''));
  SpreadsheetApp.flush();
  return { cleared: n };
}

/**
 * Period rollover. Closing balance becomes the new opening figure, the two logs
 * are copied to dated archive tabs and then emptied. A backup runs first.
 */
function rollover(body) {
  requireAdmin(body);
  if (String(body.confirm) !== 'ROLLOVER') throw new Error('Type ROLLOVER to confirm.');
  var backup = dailyBackup('archive');

  var bal = balanceMap();
  var g = grid('items');
  var cc = col(g, 'item code'), oc = col(g, 'opening stock');
  var last = lastDataRow(g, cc);
  var codes = g.sheet.getRange(g.hr + 1, cc, last - g.hr, 1).getDisplayValues();
  var out = codes.map(function (r) {
    var c = String(r[0]).trim();
    return [c && bal[c] !== undefined ? bal[c] : 0];
  });
  g.sheet.getRange(g.hr + 1, oc, out.length, 1).setValues(out);

  var tag = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd');
  var moved = 0;
  ['stockIn', 'stockOut'].forEach(function (key) {
    var lg = grid(key);
    var dc = col(lg, 'document no');
    var lastRow = lastDataRow(lg, dc);
    if (lastRow <= lg.hr) return;
    var width = lg.sheet.getMaxColumns();
    var data = lg.sheet.getRange(lg.hr + 1, 1, lastRow - lg.hr, width).getValues();
    var name = (key === 'stockIn' ? 'Stock In' : 'Stock Out') + ' archive ' + tag;
    var arc = book().getSheetByName(name) || book().insertSheet(name);
    arc.clear();
    arc.getRange(1, 1, 1, width).setValues([lg.headers]).setFontWeight('bold');
    arc.getRange(2, 1, data.length, width).setValues(data);
    arc.setFrozenRows(1);
    lg.sheet.getRange(lg.hr + 1, 1, lastRow - lg.hr, width).clearContent();
    moved += data.length;
  });

  audit('ROLLOVER', tag, '', out.length, '',
        'Closing became opening on ' + out.length + ' items; ' + moved +
        ' log lines archived' + (body.reason ? ' | ' + body.reason : ''));
  SpreadsheetApp.flush();
  return { items: out.length, archived: moved, backup: backup };
}

/* ================================ READ API =============================== */

/** A real stock line: has a code and a name, and is not a totals row. */
function isRealItem(r){
  var code = String(r['item code'] || '').trim();
  var name = String(r['item name'] || '').trim();
  if(!code || !name) return false;
  if(/^(total|totals|sum|grand|subtotal)\b/i.test(code)) return false;
  if(/^(total|totals|sum|grand|subtotal)\b/i.test(name)) return false;
  return true;
}

function bootstrap() {
  var items = readTab('items', 'item code').filter(isRealItem).map(function (r) {
    return {
      code:     r['item code'],
      name:     r['item name'],
      category: r['category'],
      unit:     r['unit'],
      location: r['store location'] || '',
      opening:  num(firstOf(r, ['opening stock as at 01-aug-2026', 'opening stock'])),
      reorder:  num(r['reorder level']),
      cost:     num(r['standard cost'])
    };
  });

  var sIn = readTab('stockIn', 'document no').map(function (r) {
    return {
      doc: r['document no'], date: r['date'], vendor: r['vendor'], invoice: r['invoice no'],
      code: r['item code'], name: r['item name'], qty: num(r['quantity']),
      cost: num(r['unit cost']), remarks: r['remarks'] || '',
      status: (r['status'] || 'ACTIVE').toUpperCase(), posted: r['posted at'] || ''
    };
  });

  var sOut = readTab('stockOut', 'document no').map(function (r) {
    return {
      doc: r['document no'], date: r['date'], branch: r['branch code'],
      issuedTo: r['issued to'], requestedBy: r['requested by'],
      code: r['item code'], name: r['item name'], qty: num(r['quantity']),
      remarks: r['remarks'] || '', status: (r['status'] || 'ACTIVE').toUpperCase(),
      unit: r['unit'] || '', posted: r['posted at'] || ''
    };
  });

  var branches = readTab('branches', 'branch name').map(function (r) {
    return { code: r['branch code'] || '', name: r['branch name'], type: r['type'] || 'Branch' };
  });

  var vendors = readTab('vendors', 'vendor name').map(function (r) {
    return { name: r['vendor name'], contact: r['contact'] || '', phone: r['phone'] || '', notes: r['notes'] || '' };
  });

  var lists = { categories: [], units: [], statuses: [] };
  try {
    var lg = grid('lists');
    var lastL = lg.sheet.getMaxRows();
    var block = lg.sheet.getRange(lg.hr + 1, 1, Math.max(1, lastL - lg.hr), 3).getDisplayValues();
    block.forEach(function (r) {
      if (r[0]) lists.categories.push(r[0]);
      if (r[1]) lists.units.push(r[1]);
      if (r[2]) lists.statuses.push(r[2]);
    });
  } catch (e) {}

  return {
    items: items, stockIn: sIn, stockOut: sOut,
    branches: branches, vendors: vendors, lists: lists,
    sheetUrl: book().getUrl(), refreshed: stamp()
  };
}

function firstOf(row, keys) {
  for (var i = 0; i < keys.length; i++) {
    if (row[keys[i]] !== undefined) return row[keys[i]];
  }
  var hit = '';
  Object.keys(row).forEach(function (h) {
    if (!hit && h.indexOf(keys[keys.length - 1]) === 0) hit = row[h];
  });
  return hit;
}

/* Server-side balance map, used to validate issues. */
function balanceMap() {
  var b = {};
  readTab('items', 'item code').filter(isRealItem).forEach(function (r) {
    b[r['item code']] = num(firstOf(r, ['opening stock as at 01-aug-2026', 'opening stock']));
  });
  readTab('stockIn', 'document no').forEach(function (r) {
    if ((r['status'] || 'ACTIVE').toUpperCase() === 'VOID') return;
    if (b[r['item code']] === undefined) b[r['item code']] = 0;
    b[r['item code']] += num(r['quantity']);
  });
  readTab('stockOut', 'document no').forEach(function (r) {
    if ((r['status'] || 'ACTIVE').toUpperCase() === 'VOID') return;
    if (b[r['item code']] === undefined) b[r['item code']] = 0;
    b[r['item code']] -= num(r['quantity']);
  });
  return b;
}

/* =============================== WRITE API =============================== */

function nextDocNo(g, prefix) {
  var dc = col(g, 'document no');
  var last = lastDataRow(g, dc);
  var period = Utilities.formatDate(new Date(), CFG.TZ, 'yyyyMM');
  var head = prefix + '-' + period + '-';
  var max = 0;
  if (last > g.hr) {
    g.sheet.getRange(g.hr + 1, dc, last - g.hr, 1).getDisplayValues().forEach(function (r) {
      var v = String(r[0]);
      if (v.indexOf(head) === 0) {
        var n = parseInt(v.substring(head.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  }
  return head + ('0000' + (max + 1)).slice(-4);
}

function postIn(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = body.header || {}, lines = body.lines || [];
    if (!lines.length) throw new Error('Add at least one item line before posting.');
    if (!h.vendor) throw new Error('Pick a vendor.');

    var g = grid('stockIn');
    var itemIndex = {};
    readTab('items', 'item code').filter(isRealItem).forEach(function (r) { itemIndex[r['item code']] = r; });

    lines.forEach(function (l) {
      if (!itemIndex[l.code]) throw new Error('Item code not on the master: ' + l.code);
      if (num(l.qty) <= 0) throw new Error('Quantity must be greater than zero (' + l.code + ').');
    });

    var doc = nextDocNo(g, CFG.DOC_PREFIX_IN);
    var row = lastDataRow(g, col(g, 'document no'));
    var date = h.date || today();

    lines.forEach(function (l) {
      row++;
      var s = g.sheet;
      s.getRange(row, col(g, 'document no')).setValue(doc);
      s.getRange(row, col(g, 'date')).setValue(date);
      s.getRange(row, col(g, 'vendor')).setValue(h.vendor);
      s.getRange(row, col(g, 'invoice no')).setValue(h.invoice || '');
      s.getRange(row, col(g, 'item code')).setValue(l.code);
      s.getRange(row, col(g, 'quantity')).setValue(num(l.qty));
      s.getRange(row, col(g, 'unit cost')).setValue(num(l.cost));
      s.getRange(row, col(g, 'remarks')).setValue(l.remarks || h.remarks || '');
      s.getRange(row, col(g, 'status')).setValue('ACTIVE');
      fillIfBlank(s, row, col(g, 'item name'), itemIndex[l.code]['item name']);
      fillIfBlank(s, row, col(g, 'line value'), num(l.qty) * num(l.cost));
      fillIfBlank(s, row, col(g, 'posted at'), stamp());
      audit('RECEIPT', doc, l.code, num(l.qty), h.vendor,
            'Invoice ' + (h.invoice || '—') + ' | ' + (l.remarks || ''));
    });

    SpreadsheetApp.flush();
    return { docNo: doc, lines: lines.length };
  } finally {
    lock.releaseLock();
  }
}

function postOut(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = body.header || {}, lines = body.lines || [];
    if (!lines.length) throw new Error('Add at least one item line before posting.');
    if (!h.branch && !h.issuedTo) throw new Error('Pick the branch or department receiving the stock.');

    var g = grid('stockOut');
    var itemIndex = {};
    readTab('items', 'item code').filter(isRealItem).forEach(function (r) { itemIndex[r['item code']] = r; });
    var bal = balanceMap();

    /* Validate the whole document before writing anything. */
    var wanted = {};
    lines.forEach(function (l) {
      if (!itemIndex[l.code]) throw new Error('Item code not on the master: ' + l.code);
      if (num(l.qty) <= 0) throw new Error('Quantity must be greater than zero (' + l.code + ').');
      wanted[l.code] = (wanted[l.code] || 0) + num(l.qty);
    });
    var short = [];
    Object.keys(wanted).forEach(function (code) {
      var have = bal[code] || 0;
      if (wanted[code] > have) {
        short.push(itemIndex[code]['item name'] + ' (' + code + '): asked ' + wanted[code] + ', on hand ' + have);
      }
    });
    if (short.length) throw new Error('Not enough stock — nothing was posted.\n' + short.join('\n'));

    var doc = nextDocNo(g, CFG.DOC_PREFIX_OUT);
    var row = lastDataRow(g, col(g, 'document no'));
    var date = h.date || today();

    lines.forEach(function (l) {
      row++;
      var s = g.sheet;
      s.getRange(row, col(g, 'document no')).setValue(doc);
      s.getRange(row, col(g, 'date')).setValue(date);
      s.getRange(row, col(g, 'branch code')).setValue(h.branch || '');
      s.getRange(row, col(g, 'issued to')).setValue(h.issuedTo || '');
      s.getRange(row, col(g, 'requested by')).setValue(h.requestedBy || '');
      s.getRange(row, col(g, 'item code')).setValue(l.code);
      s.getRange(row, col(g, 'quantity')).setValue(num(l.qty));
      s.getRange(row, col(g, 'remarks')).setValue(l.remarks || h.remarks || '');
      s.getRange(row, col(g, 'status')).setValue('ACTIVE');
      fillIfBlank(s, row, col(g, 'item name'), itemIndex[l.code]['item name']);
      fillIfBlank(s, row, col(g, 'unit'), itemIndex[l.code]['unit']);
      fillIfBlank(s, row, col(g, 'posted at'), stamp());
      audit('ISSUE', doc, l.code, num(l.qty), h.issuedTo || h.branch,
            'Requested by ' + (h.requestedBy || '—') + ' | ' + (l.remarks || ''));
    });

    SpreadsheetApp.flush();
    return { docNo: doc, lines: lines.length };
  } finally {
    lock.releaseLock();
  }
}

function voidLine(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var key = body.log === 'out' ? 'stockOut' : 'stockIn';
    var g = grid(key);
    var dc = col(g, 'document no'), ic = col(g, 'item code'),
        sc = col(g, 'status'), qc = col(g, 'quantity'), rc = col(g, 'remarks');
    var last = lastDataRow(g, dc);
    if (last <= g.hr) throw new Error('That log is empty.');

    var vals = g.sheet.getRange(g.hr + 1, 1, last - g.hr, g.sheet.getMaxColumns()).getDisplayValues();
    var hitRow = 0, qty = 0;
    for (var i = 0; i < vals.length; i++) {
      if (vals[i][dc - 1] === body.docNo && vals[i][ic - 1] === body.itemCode &&
          String(vals[i][sc - 1]).toUpperCase() !== 'VOID') {
        hitRow = g.hr + 1 + i;
        qty = num(vals[i][qc - 1]);
        break;
      }
    }
    if (!hitRow) throw new Error('No active line found for ' + body.docNo + ' / ' + body.itemCode + '.');

    /* Voiding a receipt must not push the item negative. */
    if (key === 'stockIn') {
      var bal = balanceMap();
      if ((bal[body.itemCode] || 0) - qty < 0) {
        throw new Error('Voiding this receipt would take ' + body.itemCode +
                        ' below zero. Void the related issues first.');
      }
    }

    g.sheet.getRange(hitRow, sc).setValue('VOID');
    var old = g.sheet.getRange(hitRow, rc).getDisplayValue();
    g.sheet.getRange(hitRow, rc).setValue(
      (old ? old + ' | ' : '') + 'VOID ' + today() + ': ' + (body.reason || 'no reason given'));

    audit('VOID', body.docNo, body.itemCode, qty, key === 'stockIn' ? 'receipt' : 'issue',
          body.reason || 'no reason given');
    SpreadsheetApp.flush();
    return { voided: true, row: hitRow };
  } finally {
    lock.releaseLock();
  }
}

function addItem(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var it = body.item || {};
    if (!it.name) throw new Error('Item name is required.');
    var g = grid('items');
    var cc = col(g, 'item code');
    var last = lastDataRow(g, cc);

    var existing = g.sheet.getRange(g.hr + 1, 1, Math.max(1, last - g.hr), g.sheet.getMaxColumns()).getDisplayValues();
    var max = 0;
    existing.forEach(function (r) {
      var code = String(r[cc - 1]);
      if (/^FAC-\d+$/.test(code)) {
        var n = parseInt(code.split('-')[1], 10);
        if (n > max) max = n;
      }
      if (String(r[col(g, 'item name') - 1]).trim().toLowerCase() === it.name.trim().toLowerCase()) {
        throw new Error('"' + it.name + '" is already on the master as ' + code + '.');
      }
    });

    var code = 'FAC-' + ('0000' + (max + 1)).slice(-4);
    var row = last + 1;
    var s = g.sheet;
    s.getRange(row, cc).setValue(code);
    s.getRange(row, col(g, 'item name')).setValue(it.name);
    s.getRange(row, col(g, 'category')).setValue(it.category || '');
    s.getRange(row, col(g, 'unit')).setValue(it.unit || 'PCS');
    s.getRange(row, col(g, 'store location')).setValue(it.location || '');
    s.getRange(row, col(g, 'opening stock')).setValue(num(it.opening));
    s.getRange(row, col(g, 'reorder level')).setValue(num(it.reorder));
    s.getRange(row, col(g, 'standard cost')).setValue(num(it.cost));

    audit('NEW ITEM', code, code, num(it.opening), it.category || '',
          it.name + ' | opening ' + num(it.opening) + ' ' + (it.unit || 'PCS'));
    SpreadsheetApp.flush();
    return { code: code };
  } finally {
    lock.releaseLock();
  }
}

/* ============================== AUDIT LOG ================================ */

function createAuditTab() {
  var s = book().insertSheet('Audit Log');
  s.getRange(1, 1).setValue('Audit log').setFontWeight('bold');
  s.getRange(2, 1).setValue('Append only. Never edit, sort or delete rows on this tab — it is the permanent record.');
  s.getRange(4, 1, 1, AUDIT_HEADERS.length).setValues([AUDIT_HEADERS]).setFontWeight('bold');
  s.setFrozenRows(4);
  s.setColumnWidth(1, 160); s.setColumnWidth(8, 420);
  return s;
}

function audit(action, doc, code, qty, counterparty, detail) {
  try {
    var s = tab('audit');
    var hr = headerRow(s, 'timestamp');
    var row = Math.max(s.getLastRow(), hr) + 1;
    s.getRange(row, 1, 1, AUDIT_HEADERS.length).setValues([[
      stamp(), action, doc || '', code || '', qty || 0, counterparty || '', who(), detail || ''
    ]]);
  } catch (e) { /* audit must never block a posting */ }
}

/* =============================== BACKUP ================================== */

function folder(name, parent) {
  var root = parent || DriveApp.getRootFolder();
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function backupRoot() { return folder(CFG.BACKUP_ROOT); }

function tabToCsv(sheet) {
  var rows = sheet.getDataRange().getDisplayValues();
  return rows.map(function (r) {
    return r.map(function (c) {
      var v = String(c === null || c === undefined ? '' : c);
      return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
    }).join(',');
  }).join('\n');
}

/**
 * Runs nightly.
 *   • CSV snapshot of every tab → Backups/Daily/YYYY-MM/YYYY-MM-DD, kept 95 days
 *   • On the 1st of the month: full spreadsheet copy + CSV set → Backups/Archive/YYYY,
 *     kept CFG.ARCHIVE_KEEP_DAYS (≈5½ years), satisfying the 5-year retention rule
 *   • Audit Log tab inside the live sheet is never purged
 */
function dailyBackup(manual) {
  var root = backupRoot();
  var d = new Date();
  var ymd = Utilities.formatDate(d, CFG.TZ, 'yyyy-MM-dd');
  var ym  = Utilities.formatDate(d, CFG.TZ, 'yyyy-MM');
  var yyyy = Utilities.formatDate(d, CFG.TZ, 'yyyy');
  var dayOfMonth = parseInt(Utilities.formatDate(d, CFG.TZ, 'd'), 10);

  var daily = folder(ymd + (manual ? ' (manual ' + Utilities.formatDate(d, CFG.TZ, 'HHmm') + ')' : ''),
                     folder(ym, folder('Daily', root)));
  var made = 0;
  book().getSheets().forEach(function (s) {
    daily.createFile(s.getName().replace(/[\\/]/g, '-') + '.csv', tabToCsv(s), MimeType.CSV);
    made++;
  });

  var archived = false;
  if (dayOfMonth === 1 || manual === 'archive') {
    var arc = folder(yyyy, folder('Archive', root));
    var label = 'Facility Inventory — ' + Utilities.formatDate(d, CFG.TZ, 'yyyy-MM-dd');
    DriveApp.getFileById(book().getId()).makeCopy(label, arc);
    var arcCsv = folder(label + ' (CSV)', arc);
    book().getSheets().forEach(function (s) {
      arcCsv.createFile(s.getName().replace(/[\\/]/g, '-') + '.csv', tabToCsv(s), MimeType.CSV);
    });
    archived = true;
  }

  purge(folder('Daily', root), CFG.DAILY_KEEP_DAYS);
  purge(folder('Archive', root), CFG.ARCHIVE_KEEP_DAYS);

  audit('BACKUP', ymd, '', made, archived ? 'daily + archive' : 'daily',
        'Snapshot written to Drive ▸ ' + CFG.BACKUP_ROOT);
  return { date: ymd, filesWritten: made, monthEndArchive: archived };
}

function purge(root, keepDays) {
  var cutoff = new Date().getTime() - keepDays * 86400000;
  walk(root, function (f) {
    if (f.getDateCreated().getTime() < cutoff) { f.setTrashed(true); }
  });
  var kids = root.getFolders();
  while (kids.hasNext()) {
    var k = kids.next();
    if (!k.getFiles().hasNext() && !k.getFolders().hasNext() &&
        k.getDateCreated().getTime() < cutoff) k.setTrashed(true);
  }
}

function walk(f, fn) {
  var files = f.getFiles();
  while (files.hasNext()) fn(files.next());
  var subs = f.getFolders();
  while (subs.hasNext()) walk(subs.next(), fn);
}

/* ============================== INSTALLER ================================ */

function install() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(1).everyDays(1)
    .inTimezone(CFG.TZ).create();
  tab('audit');                  // creates the Audit Log tab if missing

  var seeded = '';
  try { tab('users'); } catch (e) { createUsersTab(); }
  if (!userRows().length) {
    var g = grid('users');
    var row = g.hr + 1;
    g.sheet.getRange(row, col(g, 'username')).setValue('admin');
    g.sheet.getRange(row, col(g, 'full name')).setValue('Administrator');
    g.sheet.getRange(row, col(g, 'role')).setValue('admin');
    g.sheet.getRange(row, col(g, 'status')).setValue('ACTIVE');
    g.sheet.getRange(row, col(g, 'created')).setValue(stamp());
    writePassword(row, 'ChangeMe123!', true);
    seeded = ' First sign-in: admin / ChangeMe123! — you will be asked to change it.';
  }

  var r = dailyBackup('archive'); // first snapshot + a baseline archive copy
  return 'Installed. Nightly backup at 01:00 Kuwait time.' + seeded +
         ' First snapshot: ' + JSON.stringify(r);
}

/** Handy for testing from the editor. */
function selfTest() {
  var b = bootstrap();
  Logger.log('items %s | stockIn %s | stockOut %s | branches %s | vendors %s',
    b.items.length, b.stockIn.length, b.stockOut.length, b.branches.length, b.vendors.length);
  return b.items.length;
}
