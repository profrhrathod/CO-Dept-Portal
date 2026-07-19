/***********************************************************************
 *  COMPUTER ENGG. DEPARTMENT EVENTS PORTAL  —  Backend  (v2.0)
 *  Dr. Punjabrao Deshmukh Polytechnic, Amravati
 *  ------------------------------------------------------------------
 *  ALL DATA IS STORED ONLY IN THE ATTACHED GOOGLE SHEET.
 *  Sheets : Users | Events | EventTypes | Registrations | Notices
 *           | Resources | Sessions
 *
 *  INSTALL / UPDATE:
 *    1. Google Sheet → Extensions → Apps Script
 *    2. Replace ALL contents of Code.gs with this file
 *    3. Replace ALL contents of Index.html with the supplied file
 *    4. Save, then Deploy → Manage deployments → Edit(✏) →
 *       Version: NEW VERSION → Deploy.   (Both files must be updated
 *       together; missing this step causes "hanging" buttons.)
 *    Sheets/columns are created automatically — no manual setup needed.
 *
 *  Default admin : admin@dpdpoly.ac.in / Admin@123  (change after login)
 ***********************************************************************/

var APP_VERSION   = '2.4';
var SESSION_HOURS = 12;
var DEFAULT_TYPES = ['FDP', 'QUIZ', 'Paper Presentation', 'Seminar'];


/* ------------------------------------------------------------------ */
/*  ENTRY & VERSION CHECK                                              */
/* ------------------------------------------------------------------ */

/*
 * This Code.gs now serves as a pure JSON API (for use with an externally
 * hosted, installable version of Index.html — see the PWA build).
 * Every function that used to be called via google.script.run is now
 * called by name through this dispatcher instead:
 *   POST body: { "action": "functionName", "args": [ ...same args as before... ] }
 *   GET  query: ?action=functionName&args=<JSON array, URL-encoded>
 * Both return the exact same JSON the function used to return to
 * withSuccessHandler(). The original HtmlService page-serving doGet()
 * has been removed since the page is now hosted separately (GitHub
 * Pages or similar) as index.html + manifest.json + sw.js.
 */
var API_MAP = {
  apiVersion: apiVersion,
  registerUser: registerUser,
  loginUser: loginUser,
  logoutUser: logoutUser,
  whoAmI: whoAmI,
  forgotPassword: forgotPassword,
  changePassword: changePassword,
  getPortalData: getPortalData,
  registerForEvent: registerForEvent,
  getPublicResources: getPublicResources,
  adminAddTopMenu: adminAddTopMenu,
  adminDeleteTopMenu: adminDeleteTopMenu,
  adminGetPendingUsers: adminGetPendingUsers,
  adminApproveUser: adminApproveUser,
  adminRejectUser: adminRejectUser,
  adminAddEventType: adminAddEventType,
  adminDeleteEventType: adminDeleteEventType,
  adminAddEvent: adminAddEvent,
  adminDeleteEvent: adminDeleteEvent,
  adminAddNotice: adminAddNotice,
  adminDeleteNotice: adminDeleteNotice,
  adminAddResource: adminAddResource,
  adminUpdateResource: adminUpdateResource,
  adminDeleteResource: adminDeleteResource,
  adminGetRegistrations: adminGetRegistrations,
  adminGetUsers: adminGetUsers,
  adminAddMenuTab: adminAddMenuTab,
  adminDeleteMenuTab: adminDeleteMenuTab,
  adminAddTabItem: adminAddTabItem,
  adminDeleteTabItem: adminDeleteTabItem,
  adminExportPdf: adminExportPdf
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dispatch_(action, args) {
  ensureSetup_();
  if (!API_MAP[action]) return { ok: false, msg: 'Function "' + action + '" is missing on the server. Please paste the latest Code.gs and Deploy → New version.' };
  try {
    var result = API_MAP[action].apply(null, args || []);
    return (result === undefined || result === null) ? { ok: true } : result;
  } catch (e) {
    return { ok: false, msg: 'Server error: ' + e.message };
  }
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  if (!action) {
    ensureSetup_();
    return json_({ ok: true, version: APP_VERSION, msg: 'Events Portal API is running.' });
  }
  var args = [];
  try { if (e.parameter.args) args = JSON.parse(e.parameter.args); } catch (err) {}
  return json_(dispatch_(action, args));
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  return json_(dispatch_(body.action, body.args));
}

function apiVersion() { ensureSetup_(); return { ok: true, version: APP_VERSION }; }


/* ------------------------------------------------------------------ */
/*  SELF-HEALING SETUP — runs automatically, safe to repeat            */
/* ------------------------------------------------------------------ */
function ensureSetup_() {
  var cache = CacheService.getScriptCache();
  if (cache.get('setup_ok_' + APP_VERSION)) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {
    'Users':         ['Email','Name','Role','Department','Mobile','Salt','PassHash','RegisteredOn','Status'],
    'Events':        ['EventID','Category','Title','EventDate','Time','Venue','Coordinator','Description','Link','PostedOn'],
    'EventTypes':    ['TypeName'],
    'Registrations': ['RegID','EventID','EventTitle','Category','UserEmail','UserName','Role','Department','Mobile','Timestamp'],
    'Notices':       ['NoticeID','Title','Details','Link','PostedOn'],
    'Resources':     ['ResourceID','MenuID','Title','Details','Link','PostedOn'],
    'TopMenus':      ['MenuID','MenuName'],
    'MenuTabs':      ['TabID','TabName','LinkURL','Visibility'],
    'TabItems':      ['ItemID','TabID','Title','Details','Link','PostedOn'],
    'Sessions':      ['Token','Email','Name','Role','Created']
  };
  for (var name in defs) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(defs[name]);
      sh.getRange(1, 1, 1, defs[name].length).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  // Users sheet migration: guarantee Status header in column 9
  var users = ss.getSheetByName('Users');
  if (users.getRange(1, 9).getValue() !== 'Status')
    users.getRange(1, 9).setValue('Status').setFontWeight('bold');

  // Seed default event types once
  var et = ss.getSheetByName('EventTypes');
  if (et.getLastRow() < 2)
    DEFAULT_TYPES.forEach(function (t) { et.appendRow([t]); });

  // Seed default home-page top menus once
  var tm = ss.getSheetByName('TopMenus');
  if (tm.getLastRow() < 2) {
    tm.appendRow(['TM-ACT', 'Various Activities']);
    tm.appendRow(['TM-CAL', 'Academic Calendar']);
    tm.appendRow(['TM-CRS', 'Courses (Subjects)']);
  }
  // MenuTabs migration: guarantee LinkURL / Visibility headers
  var mtsh = ss.getSheetByName('MenuTabs');
  if (mtsh.getRange(1, 3).getValue() !== 'LinkURL')
    mtsh.getRange(1, 3, 1, 2).setValues([['LinkURL', 'Visibility']]).setFontWeight('bold');

  // Migrate old Resources rows (Type text -> MenuID)
  var rs = ss.getSheetByName('Resources');
  rs.getRange(1, 2).setValue('MenuID');
  var rd = rs.getDataRange().getValues();
  var typeMap = { 'Activity': 'TM-ACT', 'Academic Calendar': 'TM-CAL', 'Course': 'TM-CRS' };
  for (var r = 1; r < rd.length; r++)
    if (typeMap[rd[r][1]]) rs.getRange(r + 1, 2).setValue(typeMap[rd[r][1]]);

  // Guarantee an admin account exists
  var data = users.getDataRange().getValues();
  var hasAdmin = false;
  for (var i = 1; i < data.length; i++) if (data[i][2] === 'admin') hasAdmin = true;
  if (!hasAdmin) {
    var salt = Utilities.getUuid();
    users.appendRow(['admin@dpdpoly.ac.in', 'Portal Administrator', 'admin', 'Administration',
                     '', salt, hash_(salt + 'Admin@123'), new Date(), 'Approved']);
  }
  cache.put('setup_ok_' + APP_VERSION, '1', 21600);
}

/*  Optional: run manually once if you want to force setup/migration.  */
function setupSheets() { CacheService.getScriptCache().remove('setup_ok_' + APP_VERSION); ensureSetup_(); }

/* ------------------------------------------------------------------ */
/*  HELPERS                                                            */
/* ------------------------------------------------------------------ */
function sheet_(n) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n); }

function hash_(t) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function findUser_(email) {
  var d = sheet_('Users').getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][0]).toLowerCase() === String(email).toLowerCase().trim()) {
      return { row: i + 1, email: d[i][0], name: d[i][1], role: d[i][2], dept: d[i][3],
               mobile: d[i][4], salt: d[i][5], hash: d[i][6],
               status: String(d[i][8] || '').trim() };
    }
  }
  return null;
}

function checkSession_(token) {
  if (!token) return null;
  var sh = sheet_('Sessions'), d = sh.getDataRange().getValues(), now = new Date();
  for (var i = 1; i < d.length; i++) {
    if (d[i][0] === token) {
      if ((now - new Date(d[i][4])) / 3600000 <= SESSION_HOURS)
        return { email: d[i][1], name: d[i][2], role: d[i][3] };
      sh.deleteRow(i + 1);
      return null;
    }
  }
  return null;
}

function requireAdmin_(token) {
  var s = checkSession_(token);
  if (!s || s.role !== 'admin') throw new Error('Admin access required. Please login as admin.');
  return s;
}

/* Google Apps Script CANNOT return Date objects to the browser — the whole
   response silently becomes null ("Empty response from server").  Sheets
   auto-converts typed times/dates into Date values, so every value returned
   to the client must pass through jsonSafe_().                              */
function fmtSmart_(d) {
  try {
    var tz = Session.getScriptTimeZone();
    // time-only cells are stored on the epoch date 30-12-1899
    if (d.getFullYear() < 1905) return Utilities.formatDate(d, tz, 'hh:mm a');
    if (d.getHours() === 0 && d.getMinutes() === 0)
      return Utilities.formatDate(d, tz, 'dd-MM-yyyy');
    return Utilities.formatDate(d, tz, 'dd-MM-yyyy hh:mm a');
  } catch (e) { return String(d); }
}

function jsonSafe_(o) {
  if (o === null || o === undefined) return '';
  if (Object.prototype.toString.call(o) === '[object Date]') return fmtSmart_(o);
  if (Array.isArray(o)) { var a = []; for (var i = 0; i < o.length; i++) a.push(jsonSafe_(o[i])); return a; }
  if (typeof o === 'object') { var r = {}; for (var k in o) r[k] = jsonSafe_(o[k]); return r; }
  return o;
}

function fmtDate_(v) {
  if (!v) return '';
  try { return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'dd-MM-yyyy'); }
  catch (e) { return String(v); }
}

function getTypes_() {
  var d = sheet_('EventTypes').getDataRange().getValues(), out = [];
  for (var i = 1; i < d.length; i++) if (String(d[i][0]).trim()) out.push(String(d[i][0]).trim());
  return out.length ? out : DEFAULT_TYPES.slice();
}

/* ------------------------------------------------------------------ */
/*  AUTH — registration (Pending), login (blocked until Approved)      */
/* ------------------------------------------------------------------ */
function registerUser(f) {
  try {
    ensureSetup_();
    if (!f || !f.name || !f.email || !f.password || !f.role || !f.dept)
      return { ok: false, msg: 'Please fill all required fields.' };
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(f.email))
      return { ok: false, msg: 'Please enter a valid e-mail address.' };
    if (String(f.password).length < 6)
      return { ok: false, msg: 'Password must be at least 6 characters.' };
    if (f.role !== 'teacher' && f.role !== 'student')
      return { ok: false, msg: 'Invalid role selected.' };

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (findUser_(f.email)) return { ok: false, msg: 'This e-mail is already registered.' };
      var salt = Utilities.getUuid();
      sheet_('Users').appendRow([f.email.trim(), f.name.trim(), f.role, f.dept.trim(),
        f.mobile || '', salt, hash_(salt + f.password), new Date(), 'Pending']);
    } finally { lock.releaseLock(); }
    return { ok: true, msg: 'Registration submitted. Your account will be activated after approval by the Admin.' };
  } catch (e) { return { ok: false, msg: 'Error: ' + e.message }; }
}

function loginUser(email, password) {
  try {
    ensureSetup_();
    var u = findUser_(email);
    if (!u || u.hash !== hash_(u.salt + password))
      return { ok: false, msg: 'Incorrect e-mail or password.' };
    if (u.role !== 'admin' && u.status !== 'Approved')
      return { ok: false, msg: 'Your registration is awaiting Admin approval. Please try again after approval.' };
    var token = Utilities.getUuid();
    sheet_('Sessions').appendRow([token, u.email, u.name, u.role, new Date()]);
    return jsonSafe_({ ok: true, token: token, user: { email: u.email, name: u.name, role: u.role } });
  } catch (e) { return { ok: false, msg: 'Error: ' + e.message }; }
}

function logoutUser(token) {
  try {
    var sh = sheet_('Sessions'), d = sh.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--) if (d[i][0] === token) sh.deleteRow(i + 1);
  } catch (e) {}
  return { ok: true };
}

function whoAmI(token) {
  ensureSetup_();
  var s = checkSession_(token);
  return s ? jsonSafe_({ ok: true, user: s }) : { ok: false };
}

function tempPass_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  var p = '';
  for (var i = 0; i < 10; i++) p += chars.charAt(Math.floor(Math.random() * chars.length));
  return p;
}

function forgotPassword(email) {
  try {
    ensureSetup_();
    email = String(email || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))
      return { ok: false, msg: 'Please enter a valid e-mail address.' };
    var u = findUser_(email);
    if (!u) return { ok: false, msg: 'This e-mail is not registered on the portal.' };
    if (u.role !== 'admin' && u.status !== 'Approved')
      return { ok: false, msg: 'Your registration is awaiting Admin approval. Password can be reset after approval.' };

    var cache = CacheService.getScriptCache();
    if (cache.get('fp_' + email.toLowerCase()))
      return { ok: false, msg: 'A new password was already sent recently. Please check your inbox and spam folder, or try again after 10 minutes.' };

    var temp = tempPass_();
    try {
      MailApp.sendEmail({
        to: u.email,
        subject: 'Computer Engg. Department Events Portal — New Password',
        body: 'Dear ' + u.name + ',\n\n'
          + 'As requested, your password for the Computer Engg. Department Events Portal, '
          + 'Dr. Punjabrao Deshmukh Polytechnic, Amravati has been reset.\n\n'
          + 'Your new password is:  ' + temp + '\n\n'
          + 'Please login with this password and change it immediately from the "My Account" section.\n'
          + 'If you did not request this reset, please inform the portal administrator.\n\n'
          + '— Portal Administrator\nComputer Engineering Department'
      });
    } catch (mailErr) {
      return { ok: false, msg: 'E-mail could not be sent (' + mailErr.message + '). Your existing password is unchanged. Please contact the Admin.' };
    }
    // change password only after the mail has gone out
    var salt = Utilities.getUuid();
    sheet_('Users').getRange(u.row, 6, 1, 2).setValues([[salt, hash_(salt + temp)]]);
    cache.put('fp_' + email.toLowerCase(), '1', 600);
    return { ok: true, msg: 'A new password has been sent to ' + u.email + '. Please check the inbox and spam folder, then login and change the password from My Account.' };
  } catch (e) { return { ok: false, msg: 'Error: ' + e.message }; }
}

function changePassword(token, oldPass, newPass) {
  try {
    var s = checkSession_(token);
    if (!s) return { ok: false, msg: 'Session expired. Please login again.' };
    if (String(newPass).length < 6) return { ok: false, msg: 'New password must be at least 6 characters.' };
    var u = findUser_(s.email);
    if (u.hash !== hash_(u.salt + oldPass)) return { ok: false, msg: 'Current password is incorrect.' };
    var salt = Utilities.getUuid();
    sheet_('Users').getRange(u.row, 6, 1, 2).setValues([[salt, hash_(salt + newPass)]]);
    return { ok: true, msg: 'Password changed successfully.' };
  } catch (e) { return { ok: false, msg: 'Error: ' + e.message }; }
}

/* ------------------------------------------------------------------ */
/*  PORTAL DATA (logged-in users)                                      */
/* ------------------------------------------------------------------ */
function getPortalData(token) {
  try {
    ensureSetup_();
    var s = checkSession_(token);
    if (!s) return { ok: false, msg: 'Session expired.' };

    var ev = sheet_('Events').getDataRange().getValues(), events = [];
    for (var i = 1; i < ev.length; i++) {
      if (!ev[i][0]) continue;
      events.push({ id: ev[i][0], category: ev[i][1], title: ev[i][2], date: fmtDate_(ev[i][3]),
                    time: ev[i][4], venue: ev[i][5], coordinator: ev[i][6],
                    description: ev[i][7], link: ev[i][8] });
    }
    events.reverse();

    var nt = sheet_('Notices').getDataRange().getValues(), notices = [];
    for (var j = 1; j < nt.length; j++) {
      if (!nt[j][0]) continue;
      notices.push({ id: nt[j][0], title: nt[j][1], details: nt[j][2],
                     link: nt[j][3], posted: fmtDate_(nt[j][4]) });
    }
    notices.reverse();

    var rg = sheet_('Registrations').getDataRange().getValues(), mine = [];
    for (var k = 1; k < rg.length; k++) {
      if (String(rg[k][4]).toLowerCase() === s.email.toLowerCase())
        mine.push({ eventId: rg[k][1], title: rg[k][2], category: rg[k][3], on: fmtDate_(rg[k][9]) });
    }
    var mt = sheet_('MenuTabs').getDataRange().getValues(), menuTabs = [];
    for (var m = 1; m < mt.length; m++) {
      if (!mt[m][0]) continue;
      var vis = String(mt[m][3] || 'All');
      if (s.role !== 'admin') {
        if (vis === 'Teachers only' && s.role !== 'teacher') continue;
        if (vis === 'Students only' && s.role !== 'student') continue;
      }
      menuTabs.push({ id: mt[m][0], name: mt[m][1], link: String(mt[m][2] || ''), visibility: vis });
    }

    var ti = sheet_('TabItems').getDataRange().getValues(), tabItems = [];
    for (var t = 1; t < ti.length; t++) {
      if (!ti[t][0]) continue;
      tabItems.push({ id: ti[t][0], tabId: ti[t][1], title: ti[t][2],
                      details: ti[t][3], link: ti[t][4], posted: fmtDate_(ti[t][5]) });
    }
    tabItems.reverse();

    return jsonSafe_({ ok: true, user: s, types: getTypes_(), events: events, notices: notices,
             myRegs: mine, menuTabs: menuTabs, tabItems: tabItems });
  } catch (e) { return { ok: false, msg: 'Error: ' + e.message }; }
}

function registerForEvent(token, eventId) {
  try {
    var s = checkSession_(token);
    if (!s) return { ok: false, msg: 'Session expired. Please login again.' };
    var u = findUser_(s.email);
    var ev = sheet_('Events').getDataRange().getValues(), event = null;
    for (var i = 1; i < ev.length; i++) if (ev[i][0] === eventId) { event = ev[i]; break; }
    if (!event) return { ok: false, msg: 'Event not found.' };

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var rg = sheet_('Registrations').getDataRange().getValues();
      for (var j = 1; j < rg.length; j++)
        if (rg[j][1] === eventId && String(rg[j][4]).toLowerCase() === s.email.toLowerCase())
          return { ok: false, msg: 'You are already registered for this event.' };
      sheet_('Registrations').appendRow(['REG-' + Date.now(), eventId, event[2], event[1],
        u.email, u.name, u.role, u.dept, u.mobile, new Date()]);
    } finally { lock.releaseLock(); }
    return { ok: true, msg: 'Registered successfully for "' + event[2] + '".' };
  } catch (e) { return { ok: false, msg: 'Error: ' + e.message }; }
}

/* ------------------------------------------------------------------ */
/*  PUBLIC — department resources for the home-page top menu           */
/* ------------------------------------------------------------------ */
function getPublicResources() {
  try {
    ensureSetup_();
    var md = sheet_('TopMenus').getDataRange().getValues(), menus = [];
    for (var m = 1; m < md.length; m++)
      if (md[m][0]) menus.push({ id: md[m][0], name: md[m][1] });

    var d = sheet_('Resources').getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++) {
      if (!d[i][0]) continue;
      out.push({ id: d[i][0], menuId: d[i][1], title: d[i][2],
                 details: d[i][3], link: d[i][4], posted: fmtDate_(d[i][5]) });
    }
    out.reverse();
    return jsonSafe_({ ok: true, menus: menus, rows: out });
  } catch (e) { return { ok: false, msg: e.message, menus: [], rows: [] }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — home-page top menu buttons (add / delete)                  */
/* ------------------------------------------------------------------ */
function adminAddTopMenu(token, name) {
  try {
    requireAdmin_(token);
    name = String(name || '').trim();
    if (!name) return { ok: false, msg: 'Menu name is required.' };
    var d = sheet_('TopMenus').getDataRange().getValues();
    for (var i = 1; i < d.length; i++)
      if (String(d[i][1]).toLowerCase() === name.toLowerCase())
        return { ok: false, msg: 'A top menu with this name already exists.' };
    sheet_('TopMenus').appendRow(['TM-' + Date.now(), name]);
    return { ok: true, msg: 'Top menu "' + name + '" added.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminDeleteTopMenu(token, id) {
  try {
    requireAdmin_(token);
    // cascade: remove contents of this menu first
    var rs = sheet_('Resources'), rd = rs.getDataRange().getValues();
    for (var i = rd.length - 1; i >= 1; i--)
      if (rd[i][1] === id) rs.deleteRow(i + 1);
    var tm = sheet_('TopMenus'), td = tm.getDataRange().getValues();
    for (var j = td.length - 1; j >= 1; j--)
      if (td[j][0] === id) { tm.deleteRow(j + 1); return { ok: true, msg: 'Top menu and its contents deleted.' }; }
    return { ok: false, msg: 'Top menu not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — approval of new registrations                              */
/* ------------------------------------------------------------------ */
function adminGetPendingUsers(token) {
  try {
    requireAdmin_(token);
    var d = sheet_('Users').getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][8] || '').trim() === 'Pending')
        out.push({ email: d[i][0], name: d[i][1], role: d[i][2],
                   dept: d[i][3], mobile: d[i][4], on: fmtDate_(d[i][7]) });
    }
    return jsonSafe_({ ok: true, rows: out });
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminApproveUser(token, email) {
  try {
    requireAdmin_(token);
    var u = findUser_(email);
    if (!u) return { ok: false, msg: 'User not found.' };
    sheet_('Users').getRange(u.row, 9).setValue('Approved');
    return { ok: true, msg: u.name + ' approved. The user can now login.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminRejectUser(token, email) {
  try {
    requireAdmin_(token);
    var u = findUser_(email);
    if (!u) return { ok: false, msg: 'User not found.' };
    if (u.role === 'admin') return { ok: false, msg: 'Admin account cannot be removed.' };
    sheet_('Users').deleteRow(u.row);
    return { ok: true, msg: 'Registration of ' + u.name + ' rejected and removed.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — event types (add / delete)                                 */
/* ------------------------------------------------------------------ */
function adminAddEventType(token, name) {
  try {
    requireAdmin_(token);
    name = String(name || '').trim();
    if (!name) return { ok: false, msg: 'Type name is required.' };
    var types = getTypes_();
    for (var i = 0; i < types.length; i++)
      if (types[i].toLowerCase() === name.toLowerCase())
        return { ok: false, msg: 'This event type already exists.' };
    sheet_('EventTypes').appendRow([name]);
    return { ok: true, msg: 'Event type "' + name + '" added.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminDeleteEventType(token, name) {
  try {
    requireAdmin_(token);
    var ev = sheet_('Events').getDataRange().getValues();
    for (var i = 1; i < ev.length; i++)
      if (String(ev[i][1]) === name)
        return { ok: false, msg: 'Cannot delete: events exist under "' + name + '". Delete those events first.' };
    var sh = sheet_('EventTypes'), d = sh.getDataRange().getValues();
    for (var j = d.length - 1; j >= 1; j--)
      if (String(d[j][0]).trim() === name) { sh.deleteRow(j + 1); return { ok: true, msg: 'Event type "' + name + '" deleted.' }; }
    return { ok: false, msg: 'Event type not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — events                                                     */
/* ------------------------------------------------------------------ */
function adminAddEvent(token, f) {
  try {
    requireAdmin_(token);
    if (!f || !f.title || !f.category || !f.date) return { ok: false, msg: 'Category, title and date are required.' };
    if (getTypes_().indexOf(f.category) === -1) return { ok: false, msg: 'Invalid category. Add it under Event Types first.' };
    sheet_('Events').appendRow(['EVT-' + Date.now(), f.category, f.title, f.date, f.time || '',
      f.venue || '', f.coordinator || '', f.description || '', f.link || '', new Date()]);
    return { ok: true, msg: 'Event added under ' + f.category + '.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminDeleteEvent(token, eventId) {
  try {
    requireAdmin_(token);
    var sh = sheet_('Events'), d = sh.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--)
      if (d[i][0] === eventId) { sh.deleteRow(i + 1); return { ok: true, msg: 'Event deleted.' }; }
    return { ok: false, msg: 'Event not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — notices                                                    */
/* ------------------------------------------------------------------ */
function adminAddNotice(token, f) {
  try {
    requireAdmin_(token);
    if (!f || !f.title) return { ok: false, msg: 'Notice title is required.' };
    sheet_('Notices').appendRow(['NTC-' + Date.now(), f.title, f.details || '', f.link || '', new Date()]);
    return { ok: true, msg: 'Notice published.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminDeleteNotice(token, id) {
  try {
    requireAdmin_(token);
    var sh = sheet_('Notices'), d = sh.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--)
      if (d[i][0] === id) { sh.deleteRow(i + 1); return { ok: true, msg: 'Notice removed.' }; }
    return { ok: false, msg: 'Notice not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — department resources (add / update / delete)               */
/* ------------------------------------------------------------------ */
function menuExists_(id) {
  var d = sheet_('TopMenus').getDataRange().getValues();
  for (var i = 1; i < d.length; i++) if (d[i][0] === id) return String(d[i][1]);
  return null;
}

function adminAddResource(token, f) {
  try {
    requireAdmin_(token);
    if (!f || !f.menuId || !f.title) return { ok: false, msg: 'Menu and title are required.' };
    var mname = menuExists_(f.menuId);
    if (!mname) return { ok: false, msg: 'Selected top menu no longer exists.' };
    sheet_('Resources').appendRow(['RSC-' + Date.now(), f.menuId, f.title,
      f.details || '', f.link || '', new Date()]);
    return { ok: true, msg: 'Entry published under "' + mname + '".' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminUpdateResource(token, f) {
  try {
    requireAdmin_(token);
    if (!f || !f.id || !f.menuId || !f.title) return { ok: false, msg: 'Menu and title are required.' };
    if (!menuExists_(f.menuId)) return { ok: false, msg: 'Selected top menu no longer exists.' };
    var sh = sheet_('Resources'), d = sh.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === f.id) {
        sh.getRange(i + 1, 2, 1, 4).setValues([[f.menuId, f.title, f.details || '', f.link || '']]);
        return { ok: true, msg: 'Entry updated.' };
      }
    }
    return { ok: false, msg: 'Entry not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminDeleteResource(token, id) {
  try {
    requireAdmin_(token);
    var sh = sheet_('Resources'), d = sh.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--)
      if (d[i][0] === id) { sh.deleteRow(i + 1); return { ok: true, msg: 'Entry removed.' }; }
    return { ok: false, msg: 'Entry not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — registrations & users                                      */
/* ------------------------------------------------------------------ */
function adminGetRegistrations(token, eventId) {
  try {
    requireAdmin_(token);
    var d = sheet_('Registrations').getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++) {
      if (eventId && d[i][1] !== eventId) continue;
      if (!d[i][0]) continue;
      out.push({ regId: d[i][0], eventTitle: d[i][2], category: d[i][3],
                 email: d[i][4], name: d[i][5], role: d[i][6],
                 dept: d[i][7], mobile: d[i][8], on: fmtDate_(d[i][9]) });
    }
    return jsonSafe_({ ok: true, rows: out });
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminGetUsers(token) {
  try {
    requireAdmin_(token);
    var d = sheet_('Users').getDataRange().getValues(), out = [];
    for (var i = 1; i < d.length; i++)
      out.push({ email: d[i][0], name: d[i][1], role: d[i][2], dept: d[i][3],
                 mobile: d[i][4], on: fmtDate_(d[i][7]), status: d[i][8] || '' });
    return jsonSafe_({ ok: true, rows: out });
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — custom menu tabs (decided by Admin) and their contents     */
/* ------------------------------------------------------------------ */
function adminAddMenuTab(token, f) {
  try {
    requireAdmin_(token);
    var name = String((f && f.name) || '').trim();
    var link = String((f && f.link) || '').trim();
    var visibility = String((f && f.visibility) || 'All');
    if (!name) return { ok: false, msg: 'Tab name is required.' };
    if (['All', 'Teachers only', 'Students only'].indexOf(visibility) === -1) visibility = 'All';
    if (link && !/^https?:\/\//i.test(link))
      return { ok: false, msg: 'The direct link must start with http:// or https://' };
    var reserved = ['events', 'notices', 'my registrations', 'admin panel'];
    if (reserved.indexOf(name.toLowerCase()) > -1)
      return { ok: false, msg: 'This name is reserved. Please choose another name.' };
    var d = sheet_('MenuTabs').getDataRange().getValues();
    for (var i = 1; i < d.length; i++)
      if (String(d[i][1]).toLowerCase() === name.toLowerCase())
        return { ok: false, msg: 'A menu tab with this name already exists.' };
    sheet_('MenuTabs').appendRow(['TAB-' + Date.now(), name, link, visibility]);
    return { ok: true, msg: 'Menu tab "' + name + '" added.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminDeleteMenuTab(token, tabId) {
  try {
    requireAdmin_(token);
    // remove the tab's contents first (cascade)
    var ti = sheet_('TabItems'), td = ti.getDataRange().getValues();
    for (var i = td.length - 1; i >= 1; i--)
      if (td[i][1] === tabId) ti.deleteRow(i + 1);
    var mt = sheet_('MenuTabs'), md = mt.getDataRange().getValues();
    for (var j = md.length - 1; j >= 1; j--)
      if (md[j][0] === tabId) { mt.deleteRow(j + 1); return { ok: true, msg: 'Menu tab and its contents deleted.' }; }
    return { ok: false, msg: 'Menu tab not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminAddTabItem(token, f) {
  try {
    requireAdmin_(token);
    if (!f || !f.tabId || !f.title) return { ok: false, msg: 'Menu tab and title are required.' };
    var md = sheet_('MenuTabs').getDataRange().getValues(), found = false;
    for (var i = 1; i < md.length; i++) if (md[i][0] === f.tabId) found = true;
    if (!found) return { ok: false, msg: 'Selected menu tab no longer exists.' };
    sheet_('TabItems').appendRow(['ITM-' + Date.now(), f.tabId, f.title,
      f.details || '', f.link || '', new Date()]);
    return { ok: true, msg: 'Content published on the menu tab.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

function adminDeleteTabItem(token, id) {
  try {
    requireAdmin_(token);
    var sh = sheet_('TabItems'), d = sh.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--)
      if (d[i][0] === id) { sh.deleteRow(i + 1); return { ok: true, msg: 'Content removed.' }; }
    return { ok: false, msg: 'Content not found.' };
  } catch (e) { return { ok: false, msg: e.message }; }
}

/* ------------------------------------------------------------------ */
/*  ADMIN — participant list PDF (event-wise)                          */
/* ------------------------------------------------------------------ */
var LOGO_LEFT_B64  = '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';
var LOGO_RIGHT_B64 = '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';

function escHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function adminExportPdf(token, eventId) {
  try {
    requireAdmin_(token);
    if (!eventId) return { ok: false, msg: 'Please select a specific event first.' };

    var ev = sheet_('Events').getDataRange().getValues(), event = null;
    for (var i = 1; i < ev.length; i++) if (ev[i][0] === eventId) { event = ev[i]; break; }
    if (!event) return { ok: false, msg: 'Event not found.' };

    var rg = sheet_('Registrations').getDataRange().getValues(), rows = [];
    for (var j = 1; j < rg.length; j++)
      if (rg[j][1] === eventId)
        rows.push([rg[j][5], rg[j][6], rg[j][7], rg[j][8]]);   // name, role, dept, mobile
    if (!rows.length) return { ok: false, msg: 'No registrations found for this event.' };

    var coord = String(event[6] || '').trim();
    var evDate = fmtDate_(event[3]);
    var body = '';
    for (var k = 0; k < rows.length; k++) {
      body += '<tr><td class="c">' + (k + 1) + '</td><td>' + escHtml_(rows[k][0]) + '</td>'
        + '<td class="c">' + escHtml_(rows[k][1]) + '</td><td>' + escHtml_(rows[k][2]) + '</td>'
        + '<td class="c">' + escHtml_(rows[k][3]) + '</td><td>&nbsp;</td></tr>';
    }

    var html = '<html><head><meta charset="UTF-8"><style>'
      + 'body{font-family:Arial,Helvetica,sans-serif;color:#111;font-size:12px;margin:24px}'
      + '.hd{width:100%;border-bottom:3px solid #1E3A5F;padding-bottom:8px}'
      + '.hd td{vertical-align:middle;border:none}'
      + '.hd img{width:70px;height:70px;border-radius:50%}'
      + '.soc{font-size:10px;letter-spacing:2px;text-transform:uppercase;color:#444;text-align:center}'
      + '.col{font-size:18px;font-weight:bold;color:#1E3A5F;text-align:center}'
      + '.dep{font-size:11px;letter-spacing:2px;text-transform:uppercase;color:#8a5b00;text-align:center}'
      + 'h2{font-size:14px;text-align:center;margin:14px 0 2px}'
      + '.sub{text-align:center;font-size:11px;color:#444;margin-bottom:12px}'
      + 'table.lst{width:100%;border-collapse:collapse}'
      + 'table.lst th,table.lst td{border:1px solid #555;padding:5px 6px;font-size:11px}'
      + 'table.lst th{background:#E9EFF6;color:#1E3A5F}'
      + 'td.c{text-align:center}'
      + '.sig{margin-top:60px;width:100%}'
      + '.sig td{border:none;font-size:11px}'
      + '.sigline{border-top:1px solid #111;width:220px;padding-top:4px;text-align:center}'
      + '</style></head><body>'
      + '<table class="hd"><tr>'
      + '<td style="width:80px"><img src="data:image/jpeg;base64,' + LOGO_LEFT_B64 + '"></td>'
      + '<td><div class="soc">Shree Shivaji Education Society, Amravati</div>'
      + '<div class="col">Dr. Punjabrao Deshmukh Polytechnic, Amravati</div>'
      + '<div class="dep">Computer Engg. Department Events Portal</div></td>'
      + '<td style="width:80px;text-align:right"><img src="data:image/jpeg;base64,' + LOGO_RIGHT_B64 + '"></td>'
      + '</tr></table>'
      + '<h2>Participant List &mdash; ' + escHtml_(event[2]) + ' (' + escHtml_(event[1]) + ')</h2>'
      + '<div class="sub">'
      + (evDate ? 'Date: ' + escHtml_(evDate) : '')
      + (event[4] ? ' &nbsp;|&nbsp; Time: ' + escHtml_(jsonSafe_(event[4])) : '')
      + (event[5] ? ' &nbsp;|&nbsp; Venue: ' + escHtml_(event[5]) : '')
      + '</div>'
      + '<table class="lst"><tr><th style="width:34px">Sr.</th><th>Name of Participant</th>'
      + '<th style="width:70px">Role</th><th>Department</th><th style="width:90px">Mobile</th>'
      + '<th style="width:120px">Signature</th></tr>' + body + '</table>'
      + '<table class="sig"><tr><td></td><td style="width:240px">'
      + '<div class="sigline">Signature of Coordinator' + (coord ? '<br><b>' + escHtml_(coord) + '</b>' : '') + '</div>'
      + '</td></tr></table>'
      + '</body></html>';

    var pdf = Utilities.newBlob(html, 'text/html', 'list.html').getAs('application/pdf');
    var fname = 'Participants_' + String(event[2]).replace(/[^A-Za-z0-9]+/g, '_').substring(0, 40) + '.pdf';
    return { ok: true, b64: Utilities.base64Encode(pdf.getBytes()), filename: fname };
  } catch (e) { return { ok: false, msg: 'PDF error: ' + e.message }; }
}
