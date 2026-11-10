/*************************************************************
 * BHT Complaint Management System – v36 (Layout, Filters & Co-Engineer Fix)
 *  - Column Order: 1st = Ticket, 2nd = Date (dd/mm/yy, no time) across Service Reports, MD Report & Complaints
 *  - Date Range Filter (From Date - To Date) added above Service Reports, MD Report & Complaints tables
 *  - Compact Search Bar + Clean Table Layout (no vertical letter wrapping / squished headers)
 *  - Coordinator Profile: Complaints shown in DESCENDING order (newest complaint at the top)
 *  - Assign Engineer: Removed "OTHER" login/phone/email form; added single "Other / Accompanying Person Name" field alongside primary engineer (e.g., VIVEK + 2nd person)
 *  - Live Sheet Sync: Automatic migration of old engineer names (Ramesh/Suresh/Mahesh -> Vivek/Sonu Pal/Himanshu) across existing records
 *************************************************************/

const SPREADSHEET_ID_MANUAL = '';

const PROP_KEY          = 'BHT_SHEET_ID';
const PROP_WEB_APP_URL  = 'WEB_APP_URL';
const PROP_SETUP_FLAG   = 'BHT_SETUP_V41';
const PROP_FOLDER_ID    = 'BHT_UPLOAD_FOLDER_ID';
const COORDINATOR_EMAIL = 'customercare@bhtindia.com';

const SHEETS = {
  USERS: 'Users',
  COMPLAINTS: 'Complaints',
  ENGINEERS: 'Engineers',
  SERVICE_REPORTS: 'ServiceReports',
  INVOICES: 'Invoices',
  NOTIFICATIONS: 'Notifications',
  COUNTERS: 'Counters',
  EMAIL_LOG: 'EmailLog',
  FILES: 'UploadedFiles'
};

const COMPLAINT_HEADERS = [
  'id', 'ticketNumber', 'companyName', 'address', 'location',
  'authorizedPerson', 'email', 'phone', 'equipment', 'serialNo', 'problemDescription',
  'status', 'assignedEngineerId', 'assignedEngineerName', 'assignedEngineerEmail',
  'warrantyStatus', 'materialRequired', 'materialDetails', 'remarks', 'billingDone', 'createdAt',
  'feedbackRating', 'feedbackComment', 'assignedAt', 'otherPersonName'
];

const SERVICE_REPORT_HEADERS = [
  'ticketNumber', 'customerName', 'engineerName', 'status',
  'serviceImages', 'equipmentPhotos', 'serviceCharges', 'kmDistance',
  'foodCharges', 'otherCharges', 'remarks', 'pendingReason', 'nextVisitDate',
  'submittedAt', 'serviceReportUrl', 'equipmentImageUrl', 'bookletNo'
];

const INVOICE_HEADERS = [
  'id', 'invoiceNumber', 'ticketId', 'customerName', 'amount', 'billCopyUrl', 'createdAt'
];

const FILES_HEADERS = [
  'id', 'timestamp', 'ticketNumber', 'category', 'fileName', 'mimeType', 'fileId', 'url', 'uploadedBy'
];

const NOTIFICATION_HEADERS = ['id', 'userRole', 'title', 'message', 'ticketId', 'read', 'createdAt', 'targetId'];

const PLACEHOLDER_EMAIL_PATTERN = /@(t|abc|example|test|x|y|foo|bar)\.com$/i;

const KEEP_CASE_KEYS = [
  'id', 'email', 'password', 'userRole',
  'assignedEngineerId', 'assignedEngineerEmail',
  'billCopyUrl', 'serviceImages', 'equipmentPhotos', 'serviceReportUrl', 'equipmentImageUrl',
  'createdAt', 'submittedAt', 'assignedAt', 'nextVisitDate', 'availability', 'fileId', 'url', 'mimeType'
];
function upperVal(v) { return typeof v === 'string' ? v.toUpperCase() : v; }
function upperObj(o) {
  var r = {};
  Object.keys(o).forEach(function(k) {
    r[k] = KEEP_CASE_KEYS.indexOf(k) > -1 ? o[k] : upperVal(o[k]);
  });
  return r;
}

const LEGACY_NAME_MAP = {
  'RAMESH KUMAR': 'VIVEK',
  'SURESH PATEL': 'SONU PAL',
  'MAHESH SINGH': 'HIMANSHU'
};
function cleanEngName(name) {
  var s = String(name || '').trim();
  var u = s.toUpperCase();
  if (LEGACY_NAME_MAP[u]) return LEGACY_NAME_MAP[u];
  Object.keys(LEGACY_NAME_MAP).forEach(function(oldK) {
    if (u.indexOf(oldK) > -1) {
      s = u.replace(oldK, LEGACY_NAME_MAP[oldK]);
      u = s.toUpperCase();
    }
  });
  return s;
}

function doGet(e) {
  try { ensureReady(); } catch (err) { Logger.log('doGet ensureReady: ' + err); }
  return HtmlService.createHtmlOutput(getFullHTML(e))
    .setTitle('BHT Service Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function isRealEmail(email) {
  if (!email) return false;
  var s = String(email).trim().toLowerCase();
  if (s.indexOf('@') === -1) return false;
  return !PLACEHOLDER_EMAIL_PATTERN.test(s);
}

function normalizePhoneDigits(p) {
  var d = String(p || '').replace(/\D/g, '');
  if (d.length > 10) d = d.slice(-10);
  return d;
}

var _ss = null, _shMemo = {}, _hdrMemo = {};

function getSpreadsheet() {
  if (_ss) return _ss;
  var props = PropertiesService.getScriptProperties();
  if (SPREADSHEET_ID_MANUAL && SPREADSHEET_ID_MANUAL.trim() !== '') {
    try {
      var m = SpreadsheetApp.openById(SPREADSHEET_ID_MANUAL.trim());
      if (m) {
        if (props.getProperty(PROP_KEY) !== m.getId()) props.setProperty(PROP_KEY, m.getId());
        return (_ss = m);
      }
    } catch (eManual) { Logger.log('Could not open SPREADSHEET_ID_MANUAL: ' + eManual.message); }
  }
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      if (props.getProperty(PROP_KEY) !== active.getId()) props.setProperty(PROP_KEY, active.getId());
      return (_ss = active);
    }
  } catch (e) {}
  var id = props.getProperty(PROP_KEY);
  if (id) {
    try { var s = SpreadsheetApp.openById(id); if (s) return (_ss = s); } catch (e) {}
  }
  var created = SpreadsheetApp.create('BHT Database');
  try { props.setProperty(PROP_KEY, created.getId()); } catch (e) {}
  return (_ss = created);
}

function getSheet(name) {
  if (_shMemo[name]) return _shMemo[name];
  var sh = getSpreadsheet().getSheetByName(name);
  if (sh) _shMemo[name] = sh;
  return sh;
}

function getHeaders(sh, name) {
  if (_hdrMemo[name]) return _hdrMemo[name];
  var lastCol = sh.getLastColumn();
  if (lastCol === 0) return [];
  var h = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(x) { return String(x).trim(); });
  _hdrMemo[name] = h;
  return h;
}

var _isReady = false;
function ensureReady() {
  if (_isReady) return;
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_SETUP_FLAG) === '1') { _isReady = true; return; }
  ensureSheets(getSpreadsheet());
  try { props.setProperty(PROP_SETUP_FLAG, '1'); } catch (e) {}
  _isReady = true;
  try { DriveApp.getRootFolder(); } catch(e) {}
}

function ensureSheets(ss) {
  createIfMissing(ss, SHEETS.USERS, ['id', 'name', 'email', 'password', 'phone', 'role']);
  createIfMissing(ss, SHEETS.COMPLAINTS, COMPLAINT_HEADERS);
  createIfMissing(ss, SHEETS.ENGINEERS, ['id', 'name', 'phone', 'email', 'availability']);
  createIfMissing(ss, SHEETS.SERVICE_REPORTS, SERVICE_REPORT_HEADERS);
  createIfMissing(ss, SHEETS.INVOICES, INVOICE_HEADERS);
  createIfMissing(ss, SHEETS.NOTIFICATIONS, NOTIFICATION_HEADERS);
  createIfMissing(ss, SHEETS.COUNTERS, ['key', 'value']);
  createIfMissing(ss, SHEETS.EMAIL_LOG, ['timestamp', 'to', 'subject', 'status', 'error', 'provider']);
  createIfMissing(ss, SHEETS.FILES, FILES_HEADERS);

  var defaultUsers = [
    ['u1', 'Coordinator',  'customercare@bhtindia.com', 'coord123',    '9000000002', 'COORDINATOR'],
    ['u2', 'VIVEK',        'services@bhtindia.com',  'vivek123',    '',           'ENGINEER'],
    ['u3', 'SONU PAL',     'services@bhtindia.com',  'sonupal123',  '',           'ENGINEER'],
    ['u4', 'HIMANSHU',     'services@bhtindia.com',  'himanshu123', '',           'ENGINEER'],
    ['u8', 'SONU JAISWAL', 'services@bhtindia.com',  'jaiswal123',  '',           'ENGINEER'],
    ['u5', 'Accounts',     'account@bhtindia.com',            'acc123',      '9000000004', 'ACCOUNTS'],
    ['u6', 'Admin',        'mail@bhtindia.com',          'admin123',    '9000000005', 'ADMIN'],
    ['u7', 'HR Manager',   'admin@bhtindia.com',          'hr123',       '9000000007', 'HR']
  ];
  var defaultEngs = [
    ['e1', 'VIVEK',        '', 'services@bhtindia.com', 'AVAILABLE'],
    ['e2', 'SONU PAL',     '', 'services@bhtindia.com', 'AVAILABLE'],
    ['e3', 'HIMANSHU',     '', 'services@bhtindia.com', 'AVAILABLE'],
    ['e4', 'SONU JAISWAL', '', 'services@bhtindia.com', 'AVAILABLE']
  ];

  upsertSeedRows(ss, SHEETS.USERS, defaultUsers);
  upsertSeedRows(ss, SHEETS.ENGINEERS, defaultEngs);
  migrateLegacyNamesInSheets(ss);
  clearAllCaches();
}

function migrateLegacyNamesInSheets(ss) {
  try {
    var shC = ss.getSheetByName(SHEETS.COMPLAINTS);
    if (shC && shC.getLastRow() >= 2) {
      var hC = getHeaders(shC, SHEETS.COMPLAINTS);
      var idxC = hC.indexOf('assignedEngineerName');
      if (idxC > -1) {
        var rC = shC.getRange(2, idxC + 1, shC.getLastRow() - 1, 1);
        var vC = rC.getValues(), chC = false;
        for (var i = 0; i < vC.length; i++) {
          var cleaned = cleanEngName(vC[i][0]);
          if (cleaned !== String(vC[i][0] || '').trim()) { vC[i][0] = cleaned; chC = true; }
        }
        if (chC) rC.setValues(vC);
      }
    }
    var shR = ss.getSheetByName(SHEETS.SERVICE_REPORTS);
    if (shR && shR.getLastRow() >= 2) {
      var hR = getHeaders(shR, SHEETS.SERVICE_REPORTS);
      var idxR = hR.indexOf('engineerName');
      if (idxR > -1) {
        var rR = shR.getRange(2, idxR + 1, shR.getLastRow() - 1, 1);
        var vR = rR.getValues(), chR = false;
        for (var j = 0; j < vR.length; j++) {
          var cleanedR = cleanEngName(vR[j][0]);
          if (cleanedR !== String(vR[j][0] || '').trim()) { vR[j][0] = cleanedR; chR = true; }
        }
        if (chR) rR.setValues(vR);
      }
    }
  } catch (e) {}
}

function upsertSeedRows(ss, name, rows) {
  var sh = ss.getSheetByName(name);
  if (!sh) return;
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    return;
  }
  var lastRow = sh.getLastRow();
  var existing = sh.getRange(2, 1, lastRow - 1, rows[0].length).getValues();
  var idMap = {};
  var oldNames = ['RAMESH KUMAR', 'SURESH PATEL', 'MAHESH SINGH'];
  var dummyPhones = ['9111111111', '9222222222', '9333333333', '9444444444'];

  for (var i = 0; i < existing.length; i++) {
    idMap[String(existing[i][0]).trim()] = { rowNum: i + 2, vals: existing[i] };
  }
  rows.forEach(function(r) {
    var hit = idMap[String(r[0]).trim()];
    if (hit) {
      var curName = String(hit.vals[1] || '').toUpperCase().trim();
      var curPhone = String(hit.vals[2] || '').trim();
      if (oldNames.indexOf(curName) > -1) {
        sh.getRange(hit.rowNum, 1, 1, r.length).setValues([r]);
      } else if (name === SHEETS.ENGINEERS && dummyPhones.indexOf(curPhone) > -1) {
        sh.getRange(hit.rowNum, 3).setValue('');
      }
    } else {
      sh.appendRow(r);
    }
  });
}

function createIfMissing(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#fff');
    sh.setFrozenRows(1);
  } else {
    const lastCol = sh.getLastColumn();
    if (lastCol === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#fff');
      sh.setFrozenRows(1);
      return;
    }
    const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
    headers.forEach(function(h) {
      var hh = String(h).trim();
      if (existing.indexOf(hh) === -1) {
        sh.getRange(1, sh.getLastColumn() + 1).setValue(hh).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#fff');
        existing.push(hh);
      }
    });
  }
}

var CHUNK_CHARS = 30000;
var CACHE_TTL   = 60;

function cacheGet(key) {
  try {
    var c = CacheService.getScriptCache();
    var n = parseInt(c.get(key + '_n'), 10);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    var m = c.getAll(keys), s = '';
    for (var j = 0; j < n; j++) {
      if (m[keys[j]] == null) return null;
      s += m[keys[j]];
    }
    return JSON.parse(s);
  } catch (e) { return null; }
}

function cachePut(key, obj) {
  try {
    var s = JSON.stringify(obj);
    var n = Math.ceil(s.length / CHUNK_CHARS) || 1;
    if (n > 80) return;
    var map = {};
    for (var i = 0; i < n; i++) map[key + '_' + i] = s.substr(i * CHUNK_CHARS, CHUNK_CHARS);
    map[key + '_n'] = String(n);
    CacheService.getScriptCache().putAll(map, CACHE_TTL);
  } catch (e) {}
}

function cacheDrop(key) {
  try {
    var c = CacheService.getScriptCache();
    var n = parseInt(c.get(key + '_n'), 10) || 0;
    var keys = [key + '_n'];
    for (var i = 0; i < n; i++) keys.push(key + '_' + i);
    c.removeAll(keys);
  } catch (e) {}
}

function clearAllCaches() {
  Object.keys(SHEETS).forEach(function(k) { cacheDrop('sh_' + SHEETS[k]); });
  try { CacheService.getScriptCache().remove('coordinator_emails'); } catch (e) {}
}

function readSheetRaw(name) {
  const sh = getSheet(name);
  if (!sh) return [];
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol === 0) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values.shift().map(function(h) { return String(h).trim(); });
  const tz = Session.getScriptTimeZone();
  var result = [];
  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var hasData = false;
    for (var c = 0; c < row.length; c++) {
      var cellVal = row[c];
      if (cellVal !== null && cellVal !== undefined && String(cellVal).trim() !== '') {
        hasData = true;
        break;
      }
    }
    if (!hasData) continue; // Skip blank / cleared ghost rows

    const o = {};
    for (var i = 0; i < headers.length; i++) {
      var v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm:ss');
      o[headers[i]] = (v !== null && v !== undefined) ? v : '';
    }

    // Skip ghost rows that don't have basic primary identifiers
    if (name === SHEETS.COMPLAINTS) {
      if (!String(o.ticketNumber || '').trim() && !String(o.companyName || '').trim() && !String(o.id || '').trim()) {
        continue;
      }
    } else if (name === SHEETS.SERVICE_REPORTS) {
      if (!String(o.ticketNumber || '').trim() && !String(o.customerName || '').trim() && !String(o.engineerName || '').trim()) {
        continue;
      }
    } else if (name === SHEETS.INVOICES) {
      if (!String(o.invoiceNumber || '').trim() && !String(o.ticketId || '').trim()) {
        continue;
      }
    } else if (name === SHEETS.USERS) {
      if (!String(o.name || '').trim() && !String(o.email || '').trim()) {
        continue;
      }
    } else if (name === SHEETS.ENGINEERS) {
      if (!String(o.name || '').trim()) {
        continue;
      }
    }

    if (o.assignedEngineerName) o.assignedEngineerName = cleanEngName(o.assignedEngineerName);
    if (o.engineerName) o.engineerName = cleanEngName(o.engineerName);
    if (name === SHEETS.ENGINEERS && o.name) o.name = cleanEngName(o.name);
    result.push(o);
  }
  return result;
}

function readSheet(name) {
  if (name === SHEETS.ENGINEERS || name === SHEETS.USERS) {
    return readSheetRaw(name);
  }
  var key = 'sh_' + name;
  var hit = cacheGet(key);
  if (hit) return hit;
  var data = readSheetRaw(name);
  cachePut(key, data);
  return data;
}

function appendRow(name, obj) {
  const sh = getSheet(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found');
  const headers = getHeaders(sh, name);
  if (!headers.length) throw new Error('Sheet "' + name + '" has no headers');
  const row = headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; });
  sh.appendRow(row);
  cacheDrop('sh_' + name);
  return obj;
}

function updateRowById(name, id, updates) {
  const sh = getSheet(name);
  if (!sh) return false;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const headers = getHeaders(sh, name);
  const idIdx = headers.indexOf('id');
  if (idIdx === -1) return false;
  const ids = sh.getRange(2, idIdx + 1, lastRow - 1, 1).getValues();
  const target = String(id);
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === target) {
      var range = sh.getRange(i + 2, 1, 1, headers.length);
      var row = range.getValues()[0], changed = false;
      for (var c = 0; c < headers.length; c++) {
        if (updates[headers[c]] !== undefined) { row[c] = updates[headers[c]]; changed = true; }
      }
      if (changed) range.setValues([row]);
      cacheDrop('sh_' + name);
      return true;
    }
  }
  return false;
}

function findById(name, id) {
  var t = String(id);
  var list = readSheet(name);
  for (var i = 0; i < list.length; i++) if (String(list[i].id) === t) return list[i];
  return undefined;
}

function generateTicketNumber() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {}
  try {
    const sh = getSheet(SHEETS.COUNTERS);
    if (!sh) return 'BHT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const lastRow = sh.getLastRow();
    if (lastRow < 2) { sh.appendRow(['TICKET', 1]); return 'BHT-0001'; }
    const data = sh.getRange(1, 1, lastRow, 2).getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === 'TICKET') {
        const next = (Number(data[i][1]) || 0) + 1;
        sh.getRange(i + 1, 2).setValue(next);
        return 'BHT-' + String(next).padStart(4, '0');
      }
    }
    sh.appendRow(['TICKET', 1]);
    return 'BHT-0001';
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function nowStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
function ok(d) { return { success: true, data: d }; }
function fail(m) { return { success: false, message: m }; }

function findEngineerIdForUser(user) {
  try { ensureReady(); } catch (e) {}
  try {
    var engs = readSheetRaw(SHEETS.ENGINEERS);
    var phone = normalizePhoneDigits(user && user.phone);
    var name = String(user && user.name || '').toLowerCase().trim();
    for (var i = 0; i < engs.length; i++) {
      var engPhone = normalizePhoneDigits(engs[i].phone);
      if ((name && String(engs[i].name || '').toLowerCase().trim() === name) ||
          (phone && engPhone && engPhone === phone)) {
        return String(engs[i].id).trim();
      }
    }
  } catch (e) {}
  return '';
}

var _folderMemo = null;

function getOrCreateUploadFolder() {
  if (_folderMemo) return _folderMemo;
  try {
    var props = PropertiesService.getScriptProperties();
    var fid = props.getProperty(PROP_FOLDER_ID);

    if (fid) {
      try {
        var existing = DriveApp.getFolderById(fid);
        if (existing) {
          _folderMemo = existing;
          return existing;
        }
      } catch (e) {}
    }

    var folders = DriveApp.getFoldersByName('BHT Uploads');
    var folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('BHT Uploads');
    }

    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    try {
      props.setProperty(PROP_FOLDER_ID, folder.getId());
    } catch (e) {}

    _folderMemo = folder;
    return folder;
  } catch (e) {
    Logger.log('getOrCreateUploadFolder error: ' + e.message);
    try {
      var root = DriveApp.getRootFolder();
      _folderMemo = root;
      return root;
    } catch (e2) {
      Logger.log('Root folder also failed: ' + e2.message);
      return null;
    }
  }
}

function saveBase64File(base64Data, fileName, mimeType, folder, meta) {
  if (!folder || !base64Data) {
    Logger.log('saveBase64File SKIP → folder=' + !!folder + ' | data length=' + (base64Data ? String(base64Data).length : 0));
    return null;
  }

  try {
    var name = (fileName || ('File_' + new Date().getTime())).replace(/[^\w\.\-]/g, '_');
    var mime = mimeType || 'image/jpeg';

    var cleanBase64 = String(base64Data)
      .replace(/^data:[^;]*;base64,/, '')
      .replace(/[\s\r\n]+/g, '');
    while (cleanBase64.length % 4 !== 0) {
      cleanBase64 += '=';
    }

    var bytes = Utilities.base64Decode(cleanBase64);
    if (!bytes || bytes.length === 0) {
      Logger.log('saveBase64File: empty bytes → ' + name);
      return null;
    }

    var blob = Utilities.newBlob(bytes, mime, name);
    var file = folder.createFile(blob);

    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}

    var fileId = file.getId();
    var fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;

    if (meta) {
      try {
        appendRow(SHEETS.FILES, {
          id:           Utilities.getUuid(),
          timestamp:    nowStr(),
          ticketNumber: meta.ticketNumber || '',
          category:     meta.category || 'OTHER',
          fileName:     name,
          mimeType:     mime,
          fileId:       fileId,
          url:          fileUrl,
          uploadedBy:   meta.uploadedBy || ''
        });
        Logger.log('UploadedFiles written: ' + name);
      } catch (eLog) {
        Logger.log('UploadedFiles error: ' + eLog.message);
      }
    }

    return file;
  } catch (e) {
    Logger.log('saveBase64File FATAL: ' + e.message + ' | ' + fileName);
    return null;
  }
}

function saveFilesArray(arr, folder, meta) {
  var urls = [];
  if (!arr || !arr.length) {
    Logger.log('saveFilesArray: empty array');
    return urls;
  }
  if (!folder) {
    Logger.log('saveFilesArray: folder is null');
    return urls;
  }

  Logger.log('saveFilesArray: processing ' + arr.length + ' files');

  for (var i = 0; i < arr.length; i++) {
    var item = arr[i];
    if (!item || !item.base64) continue;

    var f = saveBase64File(item.base64, item.name, item.type, folder, meta);
    if (f) {
      urls.push('https://drive.google.com/uc?export=view&id=' + f.getId());
    }
  }

  Logger.log('saveFilesArray finished → ' + urls.length + ' urls');
  return urls;
}

function getWebAppUrl() {
  try {
    var prop = PropertiesService.getScriptProperties().getProperty(PROP_WEB_APP_URL);
    if (prop && prop.indexOf('http') === 0) return prop;
  } catch (e) {}
  try {
    var url = ScriptApp.getService().getUrl();
    if (url && url.indexOf('/exec') > -1) {
      try { PropertiesService.getScriptProperties().setProperty(PROP_WEB_APP_URL, url); } catch (e) {}
    }
    return url || '';
  } catch (e) { return ''; }
}

/* ---------- AUTH ---------- */
function login(email, password) {
  var eN = String(email || '').toLowerCase().trim(), pN = String(password || '').trim();
  var HARDCODED = [
    { id: 'u1', name: 'Coordinator',  email: 'customercare@bhtindia.com', password: 'coord123',    phone: '9000000002', role: 'COORDINATOR' },
    { id: 'u2', name: 'VIVEK',        email: 'services@bhtindia.com',       password: 'vivek123',    phone: '',           role: 'ENGINEER', engineerId: 'e1' },
    { id: 'u3', name: 'SONU PAL',     email: 'services@bhtindia.com',     password: 'sonupal123',  phone: '',           role: 'ENGINEER', engineerId: 'e2' },
    { id: 'u4', name: 'HIMANSHU',     email: 'services@bhtindia.com',    password: 'himanshu123', phone: '',           role: 'ENGINEER', engineerId: 'e3' },
    { id: 'u8', name: 'SONU JAISWAL', email: 'services@bhtindia.com', password: 'jaiswal123',  phone: '',           role: 'ENGINEER', engineerId: 'e4' },
    { id: 'u5', name: 'Accounts',     email: 'account@bhtindia.com',            password: 'acc123',      phone: '9000000004', role: 'ACCOUNTS' },
    { id: 'u6', name: 'Admin',        email: 'mail@bhtindia.com',          password: 'admin123',    phone: '9000000005', role: 'ADMIN' },
    { id: 'u7', name: 'HR Manager',   email: 'admin@bhtindia.com',          password: 'hr123',       phone: '9000000007', role: 'HR' }
  ];
  for (var i = 0; i < HARDCODED.length; i++) {
    if ((HARDCODED[i].email === eN || HARDCODED[i].name.toLowerCase() === eN || (HARDCODED[i].role === 'ENGINEER' && eN === 'services@bhtindia.com' || eN === 'a91884424@gmail.com')) && HARDCODED[i].password === pN) {
      var base = Object.assign({}, HARDCODED[i]);
      return { success: true, data: base };
    }
  }
  try {
    ensureReady();
    var users = readSheetRaw(SHEETS.USERS);
    for (var k = 0; k < users.length; k++) {
      var uEmail = String(users[k].email).toLowerCase().trim();
      var uName = String(users[k].name).toLowerCase().trim();
      if ((uEmail === eN || uName === eN) && String(users[k].password).trim() === pN) {
        var base2 = { id: users[k].id, name: users[k].name, email: users[k].email, phone: users[k].phone, role: users[k].role };
        if (String(base2.role).toUpperCase() === 'ENGINEER') base2.engineerId = findEngineerIdForUser(base2);
        return { success: true, data: base2 };
      }
    }
  } catch (e) {}
  return { success: false, message: 'Invalid credentials' };
}

/* ---------- USER MANAGEMENT (ADMIN) ---------- */
function getAllUsers(role) {
  try {
    ensureReady();
    if (String(role).toUpperCase() !== 'ADMIN') return fail('Admin access only');
    return ok(readSheetRaw(SHEETS.USERS));
  } catch (e) { return fail(e.message); }
}

function saveUser(payload) {
  try {
    ensureReady();
    if (String(payload.roleCheck).toUpperCase() !== 'ADMIN') return fail('Admin access only');
    var data = upperObj({
      id: payload.id || Utilities.getUuid(),
      name: String(payload.name || '').trim(),
      email: String(payload.email || '').trim(),
      password: String(payload.password || '').trim(),
      phone: String(payload.phone || '').trim(),
      role: String(payload.role || 'ENGINEER').toUpperCase()
    });
    if (!data.name || !data.email || !data.password || !data.role) return fail('Please fill all required fields');

    var existing = payload.id ? findById(SHEETS.USERS, payload.id) : null;
    if (existing) {
      updateRowById(SHEETS.USERS, payload.id, data);
      if (data.role === 'ENGINEER') {
        var engs = readSheetRaw(SHEETS.ENGINEERS);
        var matchEng = engs.find(function(e) {
          return String(e.name).toUpperCase().trim() === String(existing.name).toUpperCase().trim() ||
                 (existing.phone && normalizePhoneDigits(e.phone) === normalizePhoneDigits(existing.phone));
        });
        if (matchEng) {
          updateRowById(SHEETS.ENGINEERS, matchEng.id, { name: data.name, phone: data.phone, email: data.email });
        }
      }
    } else {
      appendRow(SHEETS.USERS, data);
      if (data.role === 'ENGINEER') {
        appendRow(SHEETS.ENGINEERS, {
          id: 'e_' + Utilities.getUuid().substring(0,6),
          name: data.name,
          phone: data.phone,
          email: data.email,
          availability: 'AVAILABLE'
        });
      }
    }
    clearAllCaches();
    return ok(data);
  } catch (e) { return fail(e.message); }
}

function deleteUser(id, role) {
  try {
    ensureReady();
    if (String(role).toUpperCase() !== 'ADMIN') return fail('Admin access only');
    var sh = getSheet(SHEETS.USERS);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return fail('No users found');
    var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sh.deleteRow(i + 2);
        cacheDrop('sh_' + SHEETS.USERS);
        return ok({ id: id });
      }
    }
    return fail('User not found');
  } catch (e) { return fail(e.message); }
}

/* ---------- EMAILS ---------- */
var _emailLogBuf = [];
function logEmailAttempt(to, subject, status, errorMsg, provider) {
  _emailLogBuf.push([nowStr(), to || '', subject || '', status || '', errorMsg || '', provider || '']);
}
function flushEmailLog() {
  if (!_emailLogBuf.length) return;
  try {
    var sh = getSheet(SHEETS.EMAIL_LOG);
    if (!sh) {
      sh = getSpreadsheet().insertSheet(SHEETS.EMAIL_LOG);
      sh.getRange(1, 1, 1, 6).setValues([['timestamp', 'to', 'subject', 'status', 'error', 'provider']]);
    }
    sh.getRange(sh.getLastRow() + 1, 1, _emailLogBuf.length, 6).setValues(_emailLogBuf);
  } catch (e) {}
  _emailLogBuf = [];
}

function sendEmailSafe(to, subject, body, options) {
  options = options || {};
  var recip = String(to || '').trim();
  var recipLower = recip.toLowerCase();
  // ADMIN and HR never receive any email
  if (recipLower === 'mail@bhtindia.com' || recipLower === 'admin@bhtindia.com' || recipLower.indexOf('admin@') === 0 || recipLower.indexOf('hr@') === 0) {
    logEmailAttempt(recip, subject, 'SKIPPED_ADMIN_HR_ROLE', 'Admin and HR do not receive emails', '');
    return { sent: false, skipped: true, to: recip };
  }
  if (!recip || recip.indexOf('@') === -1) {
    logEmailAttempt(recip, subject, 'SKIPPED_INVALID', 'no @', '');
    return { sent: false, skipped: true, to: recip };
  }
  if (!isRealEmail(recip)) {
    logEmailAttempt(recip, subject, 'SKIPPED_PLACEHOLDER', 'placeholder', '');
    return { sent: false, skipped: true, to: recip };
  }
  try {
    GmailApp.sendEmail(recip, subject, body, options);
    logEmailAttempt(recip, subject, 'SENT', '', 'GmailApp');
    return { sent: true, to: recip, provider: 'GmailApp' };
  } catch (eG) {}
  try {
    MailApp.sendEmail(recip, subject, body, options);
    logEmailAttempt(recip, subject, 'SENT', '', 'MailApp');
    return { sent: true, to: recip, provider: 'MailApp' };
  } catch (eM) {
    logEmailAttempt(recip, subject, 'FAILED', eM.message, 'MailApp');
    return { sent: false, to: recip, error: eM.message };
  }
}

function getCoordinatorEmails() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('coordinator_emails');
  if (cached) { try { return JSON.parse(cached); } catch (e) {} }
  var emails = [];
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('COORDINATOR_EMAILS');
    if (prop) {
      prop.split(/[,;\s]+/).forEach(function(e) {
        e = e.trim();
        if (e && e.indexOf('@') > -1 && emails.indexOf(e) === -1) emails.push(e);
      });
    }
  } catch (e) {}
  try {
    readSheetRaw(SHEETS.USERS).forEach(function(u) {
      if (String(u.role).toUpperCase() === 'COORDINATOR' && u.email && String(u.email).indexOf('@') > -1 && emails.indexOf(u.email) === -1) {
        emails.push(u.email);
      }
    });
  } catch (e) {}
  if (COORDINATOR_EMAIL && COORDINATOR_EMAIL.indexOf('@') > -1 && emails.indexOf(COORDINATOR_EMAIL) === -1) {
    emails.push(COORDINATOR_EMAIL);
  }
  emails = emails.filter(function(em) {
    var eml = String(em || '').toLowerCase().trim();
    return eml && eml !== 'mail@bhtindia.com' && eml !== 'admin@bhtindia.com' && eml.indexOf('admin@') !== 0 && eml.indexOf('hr@') !== 0;
  });
  if (!emails.length) {
    var eff = Session.getEffectiveUser().getEmail();
    if (eff && eff.toLowerCase() !== 'mail@bhtindia.com' && eff.toLowerCase() !== 'admin@bhtindia.com' && eff.toLowerCase().indexOf('admin@') !== 0) {
      emails = [eff];
    }
  }
  try { cache.put('coordinator_emails', JSON.stringify(emails), 600); } catch (e) {}
  return emails;
}

function sendCustomerEmail(c) {
  if (!c.email || String(c.email).indexOf('@') === -1) return { sent: false };
  var url = getWebAppUrl();
  var body = 'Dear ' + (c.authorizedPerson || 'Customer') + ',\n\nThank you for registering your service complaint with BHT.\n\n' +
    'TICKET: ' + c.ticketNumber + '\nCompany: ' + c.companyName + '\nEquipment: ' + c.equipment + '\nProblem: ' + c.problemDescription + '\n' +
    (url ? '\nTrack your status here: ' + url + '?ticket=' + c.ticketNumber + '\n' : '') + '\nRegards,\nBHT Service Team';
  return sendEmailSafe(c.email, 'BHT Ticket Confirmed – ' + c.ticketNumber, body, { name: 'BHT Service' });
}

function sendCoordinatorEmail(c) {
  var recipients = getCoordinatorEmails();
  var url = getWebAppUrl();
  var body = 'NEW COMPLAINT REGISTERED – ' + c.ticketNumber + '\n\n' +
    'Company: ' + c.companyName + '\nContact Person: ' + c.authorizedPerson + '\nPhone: ' + c.phone + '\nEmail: ' + c.email + '\n' +
    'Equipment: ' + c.equipment + '\nProblem: ' + c.problemDescription + '\n' + (url ? '\nStaff Panel: ' + url + '\n' : '');
  return {
    results: recipients.map(function(to) {
      return sendEmailSafe(to, '🚨 [New Complaint] ' + c.ticketNumber, body, { name: 'BHT Alerts' });
    })
  };
}

function sendEngineerEmail(complaint, engineer) {
  if (!engineer.email || String(engineer.email).indexOf('@') === -1) return { sent: false };
  var url = getWebAppUrl();
  var teamDisplay = complaint.assignedEngineerName || engineer.name;
  var body = 'Hi ' + teamDisplay + ',\n\nA new job has been assigned to you. Full details below:\n\n' +
    '═══════════════════════════════════════════\n' +
    '  TICKET         : ' + complaint.ticketNumber + '\n' +
    '  COMPANY        : ' + complaint.companyName + '\n' +
    '  CONTACT PERSON : ' + complaint.authorizedPerson + '\n' +
    '  PHONE          : ' + complaint.phone + '\n' +
    '  EMAIL          : ' + complaint.email + '\n' +
    '  ADDRESS        : ' + complaint.address + ', ' + complaint.location + '\n' +
    '  EQUIPMENT      : ' + complaint.equipment + '\n' +
    '  SERIAL NO      : ' + (complaint.serialNo || '-') + '\n' +
    '  PROBLEM        : ' + complaint.problemDescription + '\n' +
    '═══════════════════════════════════════════\n\n' +
    'ASSIGNMENT DETAILS\n' +
    '  ASSIGNED TEAM     : ' + teamDisplay + '\n' +
    '  WARRANTY STATUS   : ' + (complaint.warrantyStatus || '-') + '\n' +
    '  MATERIAL REQUIRED : ' + (complaint.materialRequired || 'NO') + '\n' +
    '  MATERIAL DETAILS  : ' + (complaint.materialDetails || '-') + '\n' +
    '  REMARKS           : ' + (complaint.remarks || '-') + '\n' +
    (url ? '\nEngineer Portal: ' + url + '\n' : '');
  return sendEmailSafe(engineer.email, '🛠 [Job Assigned] ' + complaint.ticketNumber, body, { name: 'BHT Service' });
}

function sendCustomerAssignmentEmail(complaint, engineer) {
  if (!complaint.email || String(complaint.email).indexOf('@') === -1) return { sent: false };
  var url = getWebAppUrl();
  var teamDisplay = complaint.assignedEngineerName || engineer.name;
  var body = 'Dear ' + complaint.authorizedPerson + ',\n\nEngineer ' + teamDisplay + ' has been assigned to your ticket ' + complaint.ticketNumber + '.\n\n' +
    (url ? 'Track Status: ' + url + '?ticket=' + complaint.ticketNumber + '\n' : '');
  return sendEmailSafe(complaint.email, '✅ Engineer Assigned – ' + complaint.ticketNumber, body, { name: 'BHT Service' });
}

function sendServiceReportEmails(complaint, engineer, payload, serviceUrls, equipUrls) {
  var url = getWebAppUrl();
  var feedbackLink = url ? url + '?ticket=' + encodeURIComponent(complaint.ticketNumber) +
                           '&phone=' + encodeURIComponent(complaint.phone || '') + '&feedback=true' : '';
  var engineerName = complaint.assignedEngineerName || (engineer ? engineer.name : '-');
  var isComplete = payload.status === 'COMPLETED';
  var serviceCharge = Number(payload.serviceCharges) || 0;
  var totalFull = serviceCharge + (Number(payload.foodCharges) || 0) + (Number(payload.otherCharges) || 0);

  if (complaint.email && String(complaint.email).indexOf('@') > -1) {
    if (isComplete) {
      var cTxt = 'Dear ' + (complaint.authorizedPerson || 'Customer') + ',\n\n🎉 Your service request has been completed successfully.\n\n' +
        'Ticket     : ' + complaint.ticketNumber + '\n' +
        'Equipment  : ' + complaint.equipment + '\n' +
        'Engineer   : ' + engineerName + '\n' +
        'Service Charge: ₹' + serviceCharge + '\n' +
        (payload.remarks ? '\nWork Summary: ' + payload.remarks + '\n' : '') +
        (feedbackLink ? '\n\nPlease rate our service: ' + feedbackLink + '\n' : '') +
        '\n\nThank you for choosing BHT!\nRegards,\nBHT Service Team';
      var cHtml = '<div style="font-family:Manrope,Arial,sans-serif;max-width:640px;margin:0 auto;padding:20px;color:#1e1b4b;background:#faf8ff;">' +
        '<div style="background:linear-gradient(135deg,#8b5cf6,#6d28d9);color:#fff;padding:24px;border-radius:16px 16px 0 0;text-align:center;">' +
        '<h1 style="margin:0;font-size:22px;font-weight:700;">🎉 Service Completed</h1>' +
        '<p style="margin:6px 0 0;font-size:13px;">Ticket ' + complaint.ticketNumber + '</p></div>' +
        '<div style="background:#fff;border:1px solid #ede9fe;border-top:none;border-radius:0 0 16px 16px;padding:24px;">' +
        '<p>Dear <b>' + (complaint.authorizedPerson || 'Customer') + '</b>,</p>' +
        '<p>Your service request <b>' + complaint.ticketNumber + '</b> has been completed successfully. 🎉</p>' +
        '<table style="width:100%;border-collapse:collapse;border:1px solid #ede9fe;border-radius:10px;overflow:hidden;margin:16px 0;">' +
        '<tr><td style="padding:10px 12px;border-bottom:1px solid #f5f3ff;background:#faf8ff;color:#6b7280;font-size:13px;">Ticket</td><td style="padding:10px 12px;border-bottom:1px solid #f5f3ff;font-weight:600;text-align:right;font-size:13px;">' + complaint.ticketNumber + '</td></tr>' +
        '<tr><td style="padding:10px 12px;border-bottom:1px solid #f5f3ff;background:#faf8ff;color:#6b7280;font-size:13px;">Equipment</td><td style="padding:10px 12px;border-bottom:1px solid #f5f3ff;font-weight:600;text-align:right;font-size:13px;">' + complaint.equipment + '</td></tr>' +
        '<tr><td style="padding:10px 12px;border-bottom:1px solid #f5f3ff;background:#faf8ff;color:#6b7280;font-size:13px;">Engineer</td><td style="padding:10px 12px;border-bottom:1px solid #f5f3ff;font-weight:600;text-align:right;font-size:13px;">' + engineerName + '</td></tr>' +
        '<tr><td style="padding:12px;background:#ede9fe;color:#6d28d9;font-weight:700;font-size:14px;">Service Charge</td><td style="padding:12px;background:#ede9fe;color:#6d28d9;font-weight:700;text-align:right;font-size:14px;">₹' + serviceCharge + '</td></tr>' +
        '</table>' +
        (payload.remarks ? '<h3 style="font-size:15px;margin:16px 0 8px;color:#7c3aed;">Work Summary</h3><p style="background:#faf8ff;border-left:4px solid #8b5cf6;padding:10px 14px;margin:0;">' + payload.remarks + '</p>' : '') +
        (feedbackLink ? '<div style="background:#fef3c7;border:2px dashed #f59e0b;border-radius:16px;padding:22px;margin-top:24px;text-align:center;">' +
          '<h3 style="margin:0 0 8px;color:#b45309;font-size:17px;">🌟 We Value Your Feedback</h3>' +
          '<p style="margin:0 0 16px;color:#78350f;font-size:13px;">Please rate our service — it helps us improve!</p>' +
          '<a href="' + feedbackLink + '" style="display:inline-block;background:#f59e0b;color:#fff;padding:14px 32px;border-radius:10px;text-decoration:none;font-weight:700;font-size:15px;">⭐ RATE OUR SERVICE</a></div>' : '') +
        '<p style="margin-top:24px;">Thank you for choosing BHT!</p>' +
        '<p style="margin-top:24px;color:#6b7280;font-size:13px;">Warm regards,<br><b>BHT Service Team</b></p>' +
        '</div></div>';
      sendEmailSafe(complaint.email, '🎉 Service Completed – ' + complaint.ticketNumber, cTxt, { name: 'BHT Service', htmlBody: cHtml });
    } else {
      var pTxt = 'Dear ' + (complaint.authorizedPerson || 'Customer') + ',\n\n⚠️ Your service ticket ' + complaint.ticketNumber + ' is currently PENDING.\n\n' +
        'Reason: ' + (payload.pendingReason || '-') + '\n' +
        'Next Visit: ' + (payload.nextVisitDate || '-') + '\n\n' +
        'Thank you for your patience.\n\nRegards,\nBHT Service Team';
      sendEmailSafe(complaint.email, '⚠️ Service Pending – ' + complaint.ticketNumber, pTxt, { name: 'BHT Service' });
    }
  }

  getCoordinatorEmails().forEach(function(c) {
    var coTxt = 'Service update for ticket ' + complaint.ticketNumber + '\n\n' +
      'Customer: ' + complaint.companyName + '\n' +
      'Engineer: ' + engineerName + '\n' +
      'Status: ' + payload.status + '\n\n' +
      'Service Charge: ₹' + serviceCharge + '\n' +
      'Food Charges  : ₹' + (Number(payload.foodCharges) || 0) + '\n' +
      'Other Charges : ₹' + (Number(payload.otherCharges) || 0) + '\n' +
      'KM Distance   : ' + (Number(payload.kmDistance) || 0) + '\n' +
      'Total         : ₹' + totalFull + '\n' +
      (payload.remarks ? 'Remarks       : ' + payload.remarks + '\n' : '') +
      (payload.pendingReason ? 'Pending Reason: ' + payload.pendingReason + '\n' : '') +
      (payload.nextVisitDate ? 'Next Visit    : ' + payload.nextVisitDate + '\n' : '');
    sendEmailSafe(c, '[Coord] ' + (isComplete ? 'Completed' : 'Pending') + ' – ' + complaint.ticketNumber, coTxt, { name: 'BHT Alerts' });
  });

  if (isComplete && String(complaint.warrantyStatus).toUpperCase() === 'OUT_OF_WARRANTY') {
    try {
      readSheetRaw(SHEETS.USERS).forEach(function(u) {
        if (String(u.role).toUpperCase() === 'ACCOUNTS' && u.email && u.email.indexOf('@') > -1) {
          var accTxt = 'BILLING PENDING — Ticket ' + complaint.ticketNumber + '\n\n' +
            'Customer: ' + complaint.companyName + '\n' +
            'Service Charge: ₹' + serviceCharge + '\n' +
            'Food Charges  : ₹' + (Number(payload.foodCharges) || 0) + '\n' +
            'Other Charges : ₹' + (Number(payload.otherCharges) || 0) + '\n' +
            'Total Billable: ₹' + totalFull + '\n';
          sendEmailSafe(u.email, '💰 [Billing Pending] ' + complaint.ticketNumber, accTxt, { name: 'BHT Accounts' });
        }
      });
    } catch (e) {}
  }
}

function notifyCustomerBill(ticketId, invoice) {
  var c = findById(SHEETS.COMPLAINTS, ticketId);
  if (!c) return;
  var body = 'Dear ' + (c.authorizedPerson || 'Customer') + ',\n\n' +
    'Your invoice has been generated for ticket ' + c.ticketNumber + '.\n\n' +
    'Invoice No : ' + invoice.invoiceNumber + '\n' +
    'Amount     : ₹' + invoice.amount + '\n' +
    (invoice.billCopyUrl ? 'Bill Copy  : ' + invoice.billCopyUrl + '\n' : '') +
    '\nRegards,\nBHT Billing Team';
  if (c.email && String(c.email).indexOf('@') > -1) {
    sendEmailSafe(c.email, '🧾 Invoice Generated – ' + c.ticketNumber, body, { name: 'BHT Billing' });
  }
  try { addNotification('COORDINATOR', 'Invoice Generated', 'Ticket ' + c.ticketNumber + ' — Invoice ' + invoice.invoiceNumber + ' ₹' + invoice.amount, c.id); } catch (e) {}
  flushEmailLog();
}

function bgComplaintEmails(id) {
  try {
    ensureReady();
    var c = findById(SHEETS.COMPLAINTS, id);
    if (!c) return ok(false);
    try { sendCustomerEmail(c); } catch (e) {}
    try { sendCoordinatorEmail(c); } catch (e) {}
    flushEmailLog();
    return ok(true);
  } catch (e) { flushEmailLog(); return fail(e.message); }
}

function bgAssignEmails(id) {
  try {
    ensureReady();
    var c = findById(SHEETS.COMPLAINTS, id);
    if (!c) return ok(false);
    var eng = findById(SHEETS.ENGINEERS, c.assignedEngineerId) || {
      name: c.assignedEngineerName || 'Engineer',
      phone: '',
      email: c.assignedEngineerEmail || ''
    };
    if (eng) {
      try { sendEngineerEmail(c, eng); } catch (e) {}
      try { sendCustomerAssignmentEmail(c, eng); } catch (e) {}
    }
    flushEmailLog();
    return ok(true);
  } catch (e) { flushEmailLog(); return fail(e.message); }
}

function bgReportEmails(ticketId) {
  try {
    ensureReady();
    var c = findById(SHEETS.COMPLAINTS, ticketId);
    if (!c) return ok(false);
    var eng = findById(SHEETS.ENGINEERS, c.assignedEngineerId);
    var reps = readSheet(SHEETS.SERVICE_REPORTS).filter(function(r) { return String(r.ticketNumber) === String(c.ticketNumber); });
    var r = reps[reps.length - 1];
    if (!r) return ok(false);
    var payload = {
      status: r.status, serviceCharges: r.serviceCharges, kmDistance: r.kmDistance,
      foodCharges: r.foodCharges, otherCharges: r.otherCharges, remarks: r.remarks,
      pendingReason: r.pendingReason, nextVisitDate: r.nextVisitDate
    };
    var sUrls = String(r.serviceImages || '').split(/[\r\n]+/).filter(Boolean);
    var eUrls = String(r.equipmentPhotos || '').split(/[\r\n]+/).filter(Boolean);
    try { sendServiceReportEmails(c, eng, payload, sUrls, eUrls); } catch (e) {}
    flushEmailLog();
    return ok(true);
  } catch (e) { flushEmailLog(); return fail(e.message); }
}

/* ---------- PUBLIC API ---------- */
function _submissionFingerprint(p) {
  var s = [
    String(p.phone || '').replace(/\D/g, ''),
    String(p.companyName || '').trim().toLowerCase(),
    String(p.problemDescription || '').trim().toLowerCase(),
    String(p.equipment || '').trim().toLowerCase()
  ].join('|');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s);
  return 'sub_' + Utilities.base64EncodeWebSafe(digest).substring(0, 40);
}

function submitPublicComplaint(payload) {
  try {
    ensureReady();
    var cache = CacheService.getScriptCache();
    var fp = _submissionFingerprint(payload);
    var cached = cache.get(fp);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        parsed.duplicate = true;
        return ok(parsed);
      } catch (e) {}
    }
    const ticket = generateTicketNumber();
    const c = upperObj({
      id: Utilities.getUuid(),
      ticketNumber: ticket,
      companyName: String(payload.companyName || '').trim(),
      address: String(payload.address || '').trim(),
      location: String(payload.location || '').trim(),
      authorizedPerson: String(payload.authorizedPerson || '').trim(),
      email: String(payload.email || '').trim(),
      phone: String(payload.phone || '').trim(),
      equipment: String(payload.equipment || '').trim(),
      serialNo: String(payload.serialNo || '').trim(),
      problemDescription: String(payload.problemDescription || '').trim(),
      status: 'NEW',
      assignedEngineerId: '',
      assignedEngineerName: '',
      assignedEngineerEmail: '',
      warrantyStatus: '',
      materialRequired: '',
      materialDetails: '',
      remarks: '',
      billingDone: 'NO',
      createdAt: nowStr(),
      feedbackRating: 0,
      feedbackComment: '',
      assignedAt: '',
      otherPersonName: ''
    });
    appendRow(SHEETS.COMPLAINTS, c);

    addNotification('COORDINATOR', 'New Complaint', 'Ticket ' + ticket + ' created for ' + c.companyName, c.id);

    try { cache.put(fp, JSON.stringify({ ticketNumber: ticket, id: c.id, complaint: c }), 300); } catch (e) {}
    return ok({ ticketNumber: ticket, id: c.id, complaint: c });
  } catch (e) { return fail(e.message); }
}

function trackComplaint(ticket, phone) {
  try {
    ensureReady();
    const list = readSheetRaw(SHEETS.COMPLAINTS);
    var t = String(ticket || '').toUpperCase().trim();
    var p = normalizePhoneDigits(phone);
    if (!t && !p) return fail('Enter ticket number or mobile number.');
    var matches = list.filter(function(x) {
      var tm = !t || String(x.ticketNumber).toUpperCase().trim() === t;
      var pm = !p || normalizePhoneDigits(x.phone) === p;
      return tm && pm;
    });
    if (!matches.length) return fail('No complaint found for the given details.');
    const byId = {};
    const byName = {};
    readSheetRaw(SHEETS.ENGINEERS).forEach(function(e) {
      byId[String(e.id).trim()] = e;
      byName[String(e.name || '').toUpperCase().trim()] = e;
    });
    const results = matches.map(function(c) {
      var e = byId[String(c.assignedEngineerId || '').trim()] || byName[String(c.assignedEngineerName || '').toUpperCase().trim()];
      var baseEng = String(c.assignedEngineerName || (e ? e.name : '') || '').trim();
      if (e && e.name && c.otherPersonName && baseEng.indexOf(c.otherPersonName) === -1) {
        baseEng = e.name + ' + ' + c.otherPersonName;
      } else if (e && e.name && !c.otherPersonName && (!baseEng || LEGACY_NAME_MAP[baseEng.toUpperCase()])) {
        baseEng = e.name;
      }
      return {
        ticketNumber: c.ticketNumber,
        companyName: c.companyName,
        equipment: c.equipment,
        problem: c.problemDescription,
        status: c.status,
        engineer: cleanEngName(baseEng),
        assignedAt: c.assignedAt || '',
        createdAt: c.createdAt
      };
    });
    return ok({ complaints: results });
  } catch (e) { return fail(e.message); }
}

function submitFeedback(ticketNumber, rating, comment, phone) {
  try {
    ensureReady();
    var r = Math.min(5, Math.max(1, Math.round(Number(rating) || 5)));
    var tn = String(ticketNumber).toUpperCase().trim();
    var c = readSheet(SHEETS.COMPLAINTS).find(function(x) {
      return String(x.ticketNumber).toUpperCase().trim() === tn;
    });
    if (!c) return fail('Ticket not found');
    var provided = normalizePhoneDigits(phone);
    var registered = normalizePhoneDigits(c.phone);
    if (!provided || provided !== registered) return fail('Only the registered customer can submit feedback.');
    updateRowById(SHEETS.COMPLAINTS, c.id, { feedbackRating: r, feedbackComment: upperVal(comment || '') });
    try { addNotification('COORDINATOR', 'Feedback Received', 'Ticket ' + c.ticketNumber + ' rated ' + r + '/5', c.id); } catch (e) {}
    return ok({ success: true });
  } catch (e) { return fail(e.message); }
}

/* ---------- STAFF API ---------- */
function sortByTicketAsc(arr) {
  return (arr || []).slice().sort(function(a, b) {
    var na = parseInt(String(a.ticketNumber || a.ticketId || '').replace(/\D/g, ''), 10) || 0;
    var nb = parseInt(String(b.ticketNumber || b.ticketId || '').replace(/\D/g, ''), 10) || 0;
    if (na !== nb) return na - nb;
    return String(a.createdAt || a.submittedAt || '').localeCompare(String(b.createdAt || b.submittedAt || ''));
  });
}

function getAllComplaints(role, email, engineerId) {
  try {
    ensureReady();
    cacheDrop('sh_' + SHEETS.COMPLAINTS);
    var repCounts = {};
    readSheet(SHEETS.SERVICE_REPORTS).forEach(function(r) {
      var tn = String(r.ticketNumber || '').toUpperCase().trim();
      if (tn) repCounts[tn] = (repCounts[tn] || 0) + 1;
    });
    var engById = {};
    readSheetRaw(SHEETS.ENGINEERS).forEach(function(e) {
      engById[String(e.id || '').trim()] = e;
    });
    const all = sortByTicketAsc(readSheet(SHEETS.COMPLAINTS)).filter(function(c) {
      return c && (String(c.ticketNumber || '').trim() || String(c.companyName || '').trim() || String(c.id || '').trim());
    }).map(function(c) {
      var o = Object.assign({}, c);
      var tn = String(c.ticketNumber || '').toUpperCase().trim();
      o._visits = repCounts[tn] || 0;
      var engObj = engById[String(c.assignedEngineerId || '').trim()];
      if (engObj && engObj.name) {
        var extra = String(c.otherPersonName || '').trim();
        o.assignedEngineerName = extra ? (engObj.name + ' + ' + extra) : engObj.name;
      } else if (o.assignedEngineerName) {
        o.assignedEngineerName = cleanEngName(o.assignedEngineerName);
      }
      return o;
    });
    if (String(role).toUpperCase() === 'ADMIN') return ok(all);
    if (!role || role === 'COORDINATOR' || role === 'ACCOUNTS') return ok(all);
    if (role === 'ENGINEER') {
      var engId = String(engineerId || '').trim();
      return ok(all.filter(function(c) { return engId && String(c.assignedEngineerId || '').trim() === engId; }));
    }
    return ok(all);
  } catch (e) { return fail(e.message); }
}

function getComplaintById(id) {
  try {
    ensureReady();
    const c = findById(SHEETS.COMPLAINTS, id);
    return c ? ok(c) : fail('Complaint not found');
  } catch (e) { return fail(e.message); }
}

function getComplaintBundle(idOrTicket) {
  try {
    ensureReady();
    var c = findById(SHEETS.COMPLAINTS, idOrTicket);
    if (!c) {
      var list = readSheet(SHEETS.COMPLAINTS);
      for (var i = 0; i < list.length; i++) {
        if (String(list[i].ticketNumber).toUpperCase().trim() === String(idOrTicket).toUpperCase().trim()) {
          c = list[i];
          break;
        }
      }
    }
    if (!c) return fail('Complaint not found');
    var eng = c.assignedEngineerId ? (findById(SHEETS.ENGINEERS, c.assignedEngineerId) || null) : null;
    if (eng && eng.name) {
      var extra = String(c.otherPersonName || '').trim();
      c.assignedEngineerName = extra ? (eng.name + ' + ' + extra) : eng.name;
    } else if (c.assignedEngineerName) {
      c.assignedEngineerName = cleanEngName(c.assignedEngineerName);
    }
    if (!eng && c.assignedEngineerName) {
      eng = { id: c.assignedEngineerId || '', name: c.assignedEngineerName, phone: '', email: c.assignedEngineerEmail || '' };
    }
    var tn = String(c.ticketNumber || '').toUpperCase().trim();
    var reports = readSheet(SHEETS.SERVICE_REPORTS)
      .filter(function(r) { return String(r.ticketNumber).toUpperCase().trim() === tn; });
    var files = readSheet(SHEETS.FILES).filter(function(f) {
      return String(f.ticketNumber).toUpperCase().trim() === tn;
    });
    var invoices = readSheet(SHEETS.INVOICES).filter(function(i) {
      return String(i.ticketId).toUpperCase().trim() === tn;
    });
    return ok({ complaint: c, engineer: eng, reports: reports, invoices: invoices, files: files });
  } catch (e) { return fail(e.message); }
}

function updateComplaintDetails(id, data, role) {
  try {
    ensureReady();
    var r = String(role || '').toUpperCase();
    if (r !== 'ADMIN' && r !== 'COORDINATOR') return fail('Only Admin or Coordinator can edit complaint details');
    var c = findById(SHEETS.COMPLAINTS, id);
    if (!c) return fail('Complaint not found');
    data = data || {};
    var email = String(data.email || '').trim();
    if (email && email.indexOf('@') === -1) return fail('Please enter a valid email or leave blank');
    var updates = upperObj({
      companyName: String(data.companyName || '').trim(),
      authorizedPerson: String(data.authorizedPerson || '').trim(),
      email: email,
      phone: String(data.phone || '').trim(),
      address: String(data.address || '').trim(),
      location: String(data.location || '').trim(),
      equipment: String(data.equipment || '').trim(),
      serialNo: String(data.serialNo || '').trim(),
      problemDescription: String(data.problemDescription || '').trim()
    });
    if (!updates.companyName || !updates.authorizedPerson || !updates.phone || !updates.address || !updates.equipment || !updates.problemDescription) {
      return fail('Please fill all required fields');
    }
    updateRowById(SHEETS.COMPLAINTS, id, updates);
    return ok(Object.assign({}, c, updates));
  } catch (e) { return fail(e.message); }
}

function getAllEngineers() {
  try {
    ensureReady();
    cacheDrop('sh_' + SHEETS.ENGINEERS);
    return ok(readSheetRaw(SHEETS.ENGINEERS));
  } catch (e) { return fail(e.message); }
}

function assignEngineer(payload) {
  try {
    ensureReady();
    const complaint = findById(SHEETS.COMPLAINTS, payload.ticketId);
    if (!complaint) return fail('Ticket not found');

    const engineer = findById(SHEETS.ENGINEERS, payload.assignedEngineerId);
    if (!engineer) return fail('Engineer not found');

    var otherPerson = String(payload.otherPersonName || '').trim().toUpperCase();
    var combinedEngineerName = otherPerson ? (engineer.name + ' + ' + otherPerson) : engineer.name;

    var updates = upperObj({
      status: 'ASSIGNED',
      assignedEngineerId: engineer.id,
      assignedEngineerName: combinedEngineerName,
      assignedEngineerEmail: engineer.email || '',
      otherPersonName: otherPerson,
      warrantyStatus: payload.warrantyStatus,
      materialRequired: payload.materialRequired ? 'YES' : 'NO',
      materialDetails: payload.materialRequired ? payload.materialDetails : '',
      remarks: payload.remarks || '',
      assignedAt: nowStr()
    });
    updateRowById(SHEETS.COMPLAINTS, payload.ticketId, updates);

    addNotification('ENGINEER', 'Job Assigned', 'Ticket ' + complaint.ticketNumber + ' assigned to ' + combinedEngineerName + ' (' + complaint.companyName + ')', complaint.id, engineer.id);

    const updated = Object.assign({}, complaint, updates);
    var engOut = Object.assign({}, engineer, { name: combinedEngineerName });
    return ok({ engineer: engOut, complaint: updated });
  } catch (e) { return fail(e.message); }
}

function updateJobStatus(ticketId, status) {
  try {
    ensureReady();
    updateRowById(SHEETS.COMPLAINTS, ticketId, { status: status });
    return ok({ success: true, status: status });
  } catch (e) { return fail(e.message); }
}

function shortCloseComplaint(id, reason) {
  try {
    ensureReady();
    var r = String(reason || '').trim();
    if (!r) return fail('Please provide a reason for short closing');
    var c = findById(SHEETS.COMPLAINTS, id);
    if (!c) return fail('Complaint not found');
    var note = (c.remarks ? c.remarks + ' | ' : '') + 'SHORT CLOSE: ' + r;
    updateRowById(SHEETS.COMPLAINTS, id, { status: 'CLOSED', billingDone: 'YES', remarks: upperVal(note) });
    try { addNotification('COORDINATOR', 'Ticket Short-Closed', 'Ticket ' + c.ticketNumber + ' — Reason: ' + r, c.id); } catch (e) {}
    return ok({ id: id, reason: r });
  } catch (e) { return fail(e.message); }
}

function submitServiceReport(payload) {
  try {
    ensureReady();
    const complaint = findById(SHEETS.COMPLAINTS, payload.ticketId);
    if (!complaint) return fail('Complaint not found');

    if (complaint.status === 'COMPLETED' || complaint.status === 'CLOSED') {
      return fail('Job is already marked as ' + complaint.status + '. Cannot re-submit report.');
    }

    const engineer = findById(SHEETS.ENGINEERS, complaint.assignedEngineerId);
    var teamName = complaint.assignedEngineerName || (engineer ? engineer.name : 'ENGINEER');

    var hasFiles = (payload.serviceImagesArray && payload.serviceImagesArray.length) ||
                   (payload.equipmentPhotosArray && payload.equipmentPhotosArray.length);

    var folder = null;
    if (hasFiles) {
      folder = getOrCreateUploadFolder();
      if (!folder) Logger.log('CRITICAL: folder is null');
    }

    var serviceUrls = [];
    var equipUrls = [];

    if (payload.serviceImagesArray && payload.serviceImagesArray.length) {
      serviceUrls = saveFilesArray(payload.serviceImagesArray, folder, {
        ticketNumber: complaint.ticketNumber,
        category: 'SERVICE_IMAGE',
        uploadedBy: teamName
      });
    }

    if (payload.equipmentPhotosArray && payload.equipmentPhotosArray.length) {
      equipUrls = saveFilesArray(payload.equipmentPhotosArray, folder, {
        ticketNumber: complaint.ticketNumber,
        category: 'EQUIPMENT_IMAGE',
        uploadedBy: teamName
      });
    }

    const report = upperObj({
      ticketNumber: complaint.ticketNumber,
      customerName: complaint.companyName,
      engineerName: teamName,
      status: payload.status,
      serviceImages: serviceUrls.join('\n'),
      equipmentPhotos: equipUrls.join('\n'),
      serviceCharges: Number(payload.serviceCharges) || 0,
      kmDistance: Number(payload.kmDistance) || 0,
      foodCharges: Number(payload.foodCharges) || 0,
      otherCharges: Number(payload.otherCharges) || 0,
      remarks: String(payload.remarks || '').trim(),
      pendingReason: payload.pendingReason || '',
      nextVisitDate: payload.nextVisitDate || '',
      submittedAt: nowStr(),
      serviceReportUrl: '',
      equipmentImageUrl: equipUrls[0] || '',
      bookletNo: String(payload.bookletNo || '').trim()
    });

    appendRow(SHEETS.SERVICE_REPORTS, report);
    var newStatus = payload.status === 'COMPLETED' ? 'COMPLETED' : 'PENDING';
    updateRowById(SHEETS.COMPLAINTS, payload.ticketId, { status: newStatus });

    try {
      if (payload.status === 'COMPLETED') {
        var tot = (Number(payload.serviceCharges) || 0) + (Number(payload.foodCharges) || 0) + (Number(payload.otherCharges) || 0);
        addNotification('COORDINATOR', 'Job Completed', 'Ticket ' + complaint.ticketNumber + ' completed by ' + teamName + ' — Total ₹' + tot, complaint.id);
        if (String(complaint.warrantyStatus).toUpperCase() === 'OUT_OF_WARRANTY') {
          addNotification('ACCOUNTS', 'Billing Pending', 'Ticket ' + complaint.ticketNumber + ' (' + complaint.companyName + ') is ready for invoicing', complaint.id);
        }
      } else {
        addNotification('COORDINATOR', 'Job Pending', 'Ticket ' + complaint.ticketNumber + ' pending: ' + (payload.pendingReason || ''), complaint.id);
      }
      addNotification('HR', 'Service Report Submitted', 'Ticket ' + complaint.ticketNumber + ' by ' + teamName + ' — ' + payload.status, complaint.id);
    } catch (eN) {}

    return ok({ report: report, complaint: Object.assign({}, complaint, { status: newStatus }) });
  } catch (e) {
    Logger.log('submitServiceReport error: ' + e.message);
    return fail(e.message);
  }
}

function getAllServiceReports(role, email, engineerId) {
  try {
    ensureReady();
    var locMap = {};
    readSheet(SHEETS.COMPLAINTS).forEach(function(cc) { locMap[String(cc.ticketNumber).toUpperCase().trim()] = cc.location || ''; });
    const all = sortByTicketAsc(readSheet(SHEETS.SERVICE_REPORTS)).map(function(r) {
      var o = {};
      for (var k in r) o[k] = r[k];
      o.engineerName = cleanEngName(r.engineerName);
      o._location = locMap[String(r.ticketNumber).toUpperCase().trim()] || '';
      return o;
    });
    if (!role || role === 'ADMIN' || role === 'COORDINATOR' || role === 'HR') return ok(all.slice(0, 500));
    if (role === 'ENGINEER') {
      var myName = '';
      if (engineerId) {
        var e = findById(SHEETS.ENGINEERS, engineerId);
        if (e) myName = String(e.name).toLowerCase().trim();
      }
      return ok(all.filter(function(r) {
        var rn = String(r.engineerName || '').toLowerCase().trim();
        return !myName || rn === myName || rn.indexOf(myName + ' +') === 0;
      }).slice(0, 500));
    }
    if (role === 'ACCOUNTS') return ok(all.filter(function(r) { return String(r.status).toUpperCase() === 'COMPLETED'; }).slice(0, 500));
    return ok(all.slice(0, 500));
  } catch (e) { return fail(e.message); }
}

function getServiceReportsForTicket(ticketNumber) {
  try {
    ensureReady();
    var tn = String(ticketNumber || '').toUpperCase().trim();
    return ok(readSheet(SHEETS.SERVICE_REPORTS).filter(function(r) { return String(r.ticketNumber).toUpperCase().trim() === tn; }));
  } catch (e) { return fail(e.message); }
}

/* MD REPORT DATA (ADMIN ONLY) */
function getReportBundle(role) {
  try {
    ensureReady();
    if (String(role || '').toUpperCase() !== 'ADMIN') return fail('Admin access only');
    return ok({
      complaints: sortByTicketAsc(readSheetRaw(SHEETS.COMPLAINTS)),
      reports:    sortByTicketAsc(readSheetRaw(SHEETS.SERVICE_REPORTS)),
      invoices:   sortByTicketAsc(readSheetRaw(SHEETS.INVOICES)),
      engineers:  readSheetRaw(SHEETS.ENGINEERS)
    });
  } catch (e) { return fail(e.message); }
}

function createInvoice(payload) {
  try {
    ensureReady();
    var folder = payload.billBase64 ? getOrCreateUploadFolder() : null;
    var file = payload.billBase64 ? saveBase64File(payload.billBase64, payload.billName, payload.billType, folder, {
      ticketNumber: payload.ticketNumber || payload.ticketId,
      category: 'BILL',
      uploadedBy: payload.uploadedBy || 'ACCOUNTS'
    }) : null;
    var billUrl = file ? file.getUrl() : '';
    const inv = upperObj({
      id: Utilities.getUuid(),
      invoiceNumber: String(payload.invoiceNumber || '').trim(),
      ticketId: payload.ticketNumber || payload.ticketId,
      customerName: payload.customerName,
      amount: Number(payload.amount) || 0,
      billCopyUrl: billUrl,
      createdAt: nowStr()
    });
    appendRow(SHEETS.INVOICES, inv);
    updateRowById(SHEETS.COMPLAINTS, payload.ticketId, { status: 'CLOSED', billingDone: 'YES' });
    try { notifyCustomerBill(payload.ticketId, inv); } catch (e) {}
    return ok(inv);
  } catch (e) { return fail(e.message); }
}

function getAllInvoices() {
  try { ensureReady(); return ok(sortByTicketAsc(readSheet(SHEETS.INVOICES))); }
  catch (e) { return fail(e.message); }
}

/* ---------- NOTIFICATIONS ---------- */
function markNotificationReadOnServer(notifId) {
  try {
    ensureReady();
    var sh = getSheet(SHEETS.NOTIFICATIONS);
    if (!sh) return ok(false);
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return ok(false);
    var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      if (!notifId || String(data[i][0]).trim() === String(notifId).trim()) {
        sh.getRange(i + 2, 6).setValue('TRUE');
        if (notifId) break;
      }
    }
    cacheDrop('sh_' + SHEETS.NOTIFICATIONS);
    return ok(true);
  } catch (e) { return fail(e.message); }
}

function addNotification(userRole, title, message, ticketId, targetId) {
  try {
    ensureReady();
    appendRow(SHEETS.NOTIFICATIONS, {
      id: Utilities.getUuid(),
      userRole: String(userRole || '').toUpperCase(),
      title: title,
      message: message,
      ticketId: ticketId || '',
      read: 'FALSE',
      createdAt: nowStr(),
      targetId: targetId || ''
    });
  } catch (e) { Logger.log('addNotification error: ' + e.message); }
}

function getNotifications(role, engineerId) {
  try {
    ensureReady();
    var r = String(role || '').toUpperCase();
    var eid = String(engineerId || '').trim();
    return ok(readSheetRaw(SHEETS.NOTIFICATIONS)
      .filter(function(n) {
        var nr = String(n.userRole || '').toUpperCase().trim();
        if (r === 'ADMIN') return true;
        if (nr !== r && nr !== 'ALL') return false;
        var t = String(n.targetId || '').trim();
        if (r === 'ENGINEER' && t) return t === eid;
        return true;
      })
      .reverse().slice(0, 100));
  } catch (e) { return fail(e.message); }
}

/* ============================================================
   HTML UI — Aurora Violet Theme (v36 - Layout, Filters & Co-Engineer Fix)
   NOTE: String.raw keeps every backslash exactly as written.
============================================================ */
function getFullHTML(e) {
  var paramData = (e && e.parameter) ? JSON.stringify(e.parameter).replace(/</g, '\\u003c') : '{}';

  return String.raw`<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>BHT Service Management</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=Noto+Sans+Devanagari:wght@400;500;600;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
<style>
:root {
  --primary: #8b5cf6; --primary-hover: #7c3aed; --primary-dark: #6d28d9;
  --primary-light: #ede9fe; --primary-glow: rgba(139, 92, 246, 0.28);
  --accent: #f43f5e; --accent-light: #ffe4e6;
  --bg: #faf8ff; --bg-alt: #f5f3ff; --surface: #ffffff; --surface-alt: #fdfcff;
  --text-main: #1e1b4b; --text-soft: #4c1d95; --text-muted: #6b7280; --text-dim: #9ca3af;
  --border: #f0ecf9; --border-strong: #e5dbff;
  --success: #10b981; --success-hover: #059669; --success-light: #d1fae5;
  --danger: #ef4444; --danger-hover: #dc2626; --danger-light: #fee2e2;
  --warning: #f59e0b; --warning-hover: #d97706; --warning-light: #fef3c7;
  --radius: 18px; --radius-sm: 12px; --radius-lg: 24px; --radius-pill: 999px;
  --shadow-xs: 0 1px 2px rgba(30, 27, 75, 0.04);
  --shadow-sm: 0 2px 6px rgba(30, 27, 75, 0.05);
  --shadow-md: 0 4px 16px rgba(30, 27, 75, 0.07);
  --shadow-lg: 0 10px 32px rgba(30, 27, 75, 0.09);
  --shadow-xl: 0 20px 48px rgba(30, 27, 75, 0.12);
  --shadow-primary: 0 6px 20px rgba(139, 92, 246, 0.32);
  --font: 'Manrope', 'Noto Sans Devanagari', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
}
body.dark {
  --bg:#2a3040; --bg-alt:#333a4d; --surface:#353d52; --surface-alt:#3d465d;
  --text-main:#f1f5f9; --text-soft:#d5dbe8; --text-muted:#b4bdd0; --text-dim:#8e99b0;
  --border:#48516a; --border-strong:#56607c;
  --primary:#a78bfa; --primary-hover:#8b5cf6; --primary-dark:#7c3aed;
  --primary-light:#464566; --primary-glow:rgba(167,139,250,0.30);
  --accent-light:#5a3d4d; --success-light:#2f5a4c; --danger-light:#5e3d44; --warning-light:#5e5030;
  --shadow-xs:0 1px 2px rgba(0,0,0,0.15); --shadow-sm:0 2px 6px rgba(0,0,0,0.18);
  --shadow-md:0 4px 16px rgba(0,0,0,0.2); --shadow-lg:0 10px 32px rgba(0,0,0,0.22); --shadow-xl:0 20px 48px rgba(0,0,0,0.28);
  color-scheme: dark;
}
body.dark { background: var(--bg); }
body.dark .app-header { background: #2f3649; }
body.dark .sidebar { background: #2f3649; }
body.dark .nitem.unread, body.dark tr.unread-row td { background: #3f4863; }
body.dark .bell-badge { border-color: var(--surface); }
body.dark .pp-head { background: #3f4863; }
body.dark .loader-overlay { background: rgba(42,48,64,0.8); }
body.dark .alert-error { color:#fecaca; border-color:#8a4a52; }
body.dark .alert-warning { color:#fde68a; border-color:#8a7438; }
body.dark .alert-success { color:#a7f3d0; border-color:#3f7a66; }
body.dark .badge-new { background:#4a536b; color:#e2e8f0; }
body.dark .badge-assigned { background:#6b5a30; color:#fde68a; }
body.dark .badge-accepted { background:#52498a; color:#e0d7ff; }
body.dark .badge-onsite { background:#434f8a; color:#dbe2ff; }
body.dark .badge-pending { background:#7a424a; color:#ffd5d5; }
body.dark .badge-completed { background:#2f6b55; color:#c9f7e3; }
body.dark .badge-closed { background:#4a536b; color:#e2e8f0; }
body.dark .btn-pdf { background:#6b3f47; color:#fecaca; border-color:#8a4a52; }
body.dark .btn-print { background:#434f8a; color:#dbe2ff; border-color:#5a67a8; }
body.dark .btn-csv { background:#2f6b55; color:#c9f7e3; border-color:#3f7a66; }
body.dark .star-rating { color:#56607c; }
body.dark input::placeholder, body.dark textarea::placeholder { color: var(--text-dim); }
body.dark input[type=date]::-webkit-calendar-picker-indicator { filter: invert(0.8); }
body.dark .notif-popup { background:#353d52 !important; border-color:#48516a !important; }
* { margin:0; padding:0; box-sizing:border-box; font-family: var(--font); -webkit-font-smoothing: antialiased; }
body {
  background:
    radial-gradient(1200px 600px at 10% -10%, #f3e8ff 0%, transparent 45%),
    radial-gradient(900px 500px at 100% 100%, #ffe4e6 0%, transparent 45%),
    var(--bg);
  color: var(--text-main); line-height: 1.55; overflow-x: hidden;
  min-height: 100vh; font-weight: 500; transition: background-color .25s, color .25s;
}
.hidden { display: none !important; }
.admin-only { display: none !important; }
body.is-admin .admin-only { display: inline-flex !important; }
body.is-engineer .export-bar { display: none !important; }
@keyframes fadeIn { from { opacity: 0; transform: translateY(10px); } to { opacity: 1; transform: translateY(0); } }
@keyframes slideInRight { from { transform: translateX(120%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }
@keyframes spin { to { transform: rotate(360deg); } }
@keyframes slideUp { to { transform:translateY(0); } }
.animate-fade { animation: fadeIn 0.25s cubic-bezier(0.4, 0, 0.2, 1) forwards; }

input, select, textarea { width:100%; padding:13px 16px; border:1.5px solid var(--border-strong); border-radius: var(--radius-sm); font-size:14px; font-weight: 500; font-family: var(--font); background: var(--surface); transition: all 0.2s; color: var(--text-main); }
input::placeholder, textarea::placeholder { color: var(--text-dim); font-weight: 500; text-transform: none; }
input:hover, select:hover, textarea:hover { border-color: #d8caff; }
input:focus, select:focus, textarea:focus { outline:none; border-color: var(--primary); box-shadow: 0 0 0 4px var(--primary-glow); }
label { display:block; font-size:12px; font-weight: 700; margin-bottom:8px; color: var(--text-soft); text-transform: uppercase; letter-spacing: 0.6px; }

.main-content td, #trackResult td, #jobCards .card h4, #jobCards .card p, .nitem .ntitle, .nitem .nmsg, .pp-name,
input:not([type=email]):not([type=password]):not([type=file]):not([type=date]):not([type=number]), textarea { text-transform: uppercase; }
.keep-case, input[type=email], td.keep-case { text-transform: none !important; }

.btn { display:inline-flex; align-items:center; justify-content:center; gap:8px; padding:13px 24px; border-radius: var(--radius-sm); border:none; cursor:pointer; font-size:14px; font-weight: 700; font-family: var(--font); text-decoration:none; transition: all 0.22s cubic-bezier(0.4, 0, 0.2, 1); letter-spacing: -0.1px; white-space: nowrap; }
.btn:disabled { opacity:0.55; cursor:not-allowed; }
.btn-primary { background: linear-gradient(135deg, var(--primary) 0%, var(--primary-hover) 100%); color:#fff; box-shadow: var(--shadow-primary); }
.btn-primary:hover:not(:disabled) { transform: translateY(-2px); box-shadow: 0 10px 28px rgba(139, 92, 246, 0.42); }
.btn-secondary { background: var(--surface); color: var(--text-main); border: 1.5px solid var(--border-strong); }
.btn-secondary:hover:not(:disabled) { background: var(--primary-light); border-color: var(--primary); color: var(--primary-hover); }
.btn-danger { background: linear-gradient(135deg, var(--danger) 0%, var(--danger-hover) 100%); color:#fff; box-shadow: 0 4px 14px rgba(239,68,68,0.32); }
.btn-danger:hover:not(:disabled) { transform:translateY(-2px); }
.btn-success { background: linear-gradient(135deg, var(--success) 0%, var(--success-hover) 100%); color:#fff; box-shadow: 0 4px 14px rgba(16,185,129,0.32); }
.btn-success:hover:not(:disabled) { transform:translateY(-2px); }
.btn-wa { background: linear-gradient(135deg, #25D366 0%, #1eb755 100%); color:#fff; box-shadow: 0 4px 14px rgba(37,211,102,0.32); }
.btn-wa:hover:not(:disabled) { transform:translateY(-2px); }
.btn-warning { background: linear-gradient(135deg, var(--warning) 0%, var(--warning-hover) 100%); color:#fff; box-shadow: 0 4px 14px rgba(245,158,11,0.32); }
.btn-warning:hover:not(:disabled) { transform:translateY(-2px); }
.btn-block { width:100%; }
.btn-sm { padding: 8px 14px; font-size:12.5px; border-radius: 10px; }

.alert { padding:14px 18px; border-radius: var(--radius-sm); font-size:14px; margin-bottom:16px; font-weight: 600; display:flex; align-items:center; gap:10px; }
.alert-error { background: var(--danger-light); color: #991b1b; border: 1.5px solid #fecaca; }
.alert-warning { background: var(--warning-light); color: #92400e; border: 1.5px solid #fcd34d; }
.alert-success { background: var(--success-light); color: #065f46; border: 1.5px solid #6ee7b7; }

.card { background: var(--surface); border-radius: var(--radius); padding:26px; box-shadow: var(--shadow-md); border: 1px solid var(--border); margin-bottom:24px; animation: fadeIn 0.25s ease-out; }
.card-title { font-size:18px; font-weight: 800; margin-bottom:20px; display:flex; align-items:center; gap:10px; color: var(--text-main); letter-spacing: -0.3px; }

.table-wrap { width:100%; overflow-x: auto; border-radius: var(--radius-sm); border: 1px solid var(--border); background: var(--surface); }
table { width:100%; border-collapse: collapse; text-align:left; }
th, td { padding:10px 12px; border-bottom: 1px solid var(--border); font-size:12.5px; }
th { background: linear-gradient(180deg, var(--surface-alt) 0%, var(--surface) 100%); color: var(--text-muted); font-weight: 700; font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; }
tbody tr:last-child td { border-bottom: none; }
tbody tr:hover { background: var(--surface-alt); }
tr.unread-row td { background: #f8f5ff; }

/* Multi-visit & Sequence Gap Row Highlighting */
tbody tr.multi-visit-row td { background: #fef3c7 !important; border-color: #fde68a !important; }
body.dark tbody tr.multi-visit-row td { background: #5c4b23 !important; border-color: #78622d !important; }
tbody tr.seq-gap-row td { background: #ffe4e6 !important; border-color: #fecdd3 !important; }
body.dark tbody tr.seq-gap-row td { background: #5a2e38 !important; border-color: #7a3e4b !important; }
.multi-visit-pill { display:inline-flex; align-items:center; gap:4px; background:#d97706; color:#fff; padding:2px 8px; border-radius:999px; font-size:9.5px; font-weight:800; margin-top:3px; white-space:nowrap; }
.seq-gap-pill { display:inline-flex; align-items:center; gap:4px; background:#e11d48; color:#fff; padding:2px 8px; border-radius:999px; font-size:9.5px; font-weight:800; margin-top:3px; white-space:nowrap; }

.badge { display: inline-flex; align-items: center; gap: 5px; padding: 4px 10px; border-radius: var(--radius-pill); font-size: 10px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; }
.badge-new { background: #f1f5f9; color: #475569; }
.badge-assigned { background: #fef3c7; color: #92400e; }
.badge-accepted { background: #ede9fe; color: #6d28d9; }
.badge-onsite { background: #e0e7ff; color: #4338ca; }
.badge-pending { background: #fee2e2; color: #991b1b; }
.badge-completed { background: #d1fae5; color: #065f46; }
.badge-closed { background: #e2e8f0; color: #1e293b; }
.badge-notreq { background: #f3f4f6; color: #4b5563; }

/* Compact Toolbar & Date Range Filter Bar */
.toolbar { display:flex; flex-wrap:wrap; gap:8px; align-items:center; margin-bottom:14px; background: var(--surface); border:1px solid var(--border); padding:10px 12px; border-radius: var(--radius-sm); box-shadow: var(--shadow-xs); }
.toolbar .search-box { position:relative; flex: 0 1 220px; max-width: 240px; min-width: 160px; }
.toolbar .search-box i { position:absolute; left:12px; top:50%; transform:translateY(-50%); color: var(--text-dim); font-size:12.5px; pointer-events:none; }
.toolbar .search-box input { padding: 8px 12px 8px 34px; font-size:12.5px; height: 36px; }
.toolbar select, .toolbar input[type=date] { width:auto; flex: 0 1 auto; padding: 7px 10px; font-size:12.5px; height: 36px; }
.toolbar .date-range-inline { display:inline-flex; align-items:center; gap:6px; background:var(--bg-alt); padding:3px 8px; border-radius:10px; border:1px solid var(--border-strong); }
.toolbar .date-range-inline span { font-size:11px; font-weight:800; color:var(--text-soft); text-transform:uppercase; white-space:nowrap; }
.toolbar .date-range-inline input[type=date] { border:none; background:transparent; padding:4px 6px; height:28px; font-size:12px; width:125px; }
.toolbar .date-range-inline input[type=date]:focus { box-shadow:none; }
.toolbar .tb-count { font-size:11.5px; color: var(--text-muted); font-weight:700; margin-left:auto; white-space:nowrap; }
.toolbar .filter-group { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
.filter-toggle-btn.active { background: var(--primary-light); color: var(--primary-hover); border-color: var(--primary); }
.export-bar { display: inline-flex; gap:6px; align-items:center; }
.btn-pdf { background:#fee2e2; color:#b91c1c; border:1.5px solid #fecaca; }
.btn-pdf:hover { background:#fecaca; }
.btn-print { background:#e0e7ff; color:#3730a3; border:1.5px solid #c7d2fe; }
.btn-print:hover { background:#c7d2fe; }
.btn-csv { background:#d1fae5; color:#065f46; border:1.5px solid #a7f3d0; }
.btn-csv:hover { background:#a7f3d0; }

.public-wrap { max-width:840px; margin:0 auto; padding: 48px 20px; }
.public-header { text-align:center; margin-bottom:40px; animation: fadeIn 0.4s ease-out; }
.public-header h1 { font-size: 38px; font-weight: 800; letter-spacing: -1.2px; margin-bottom: 10px; background: linear-gradient(135deg, var(--primary-dark) 0%, var(--accent) 100%); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; }
.public-header p { color: var(--text-muted); font-size: 16px; font-weight: 500; }
.tabs-public { display:flex; justify-content:center; gap: 12px; margin-bottom:32px; }
.tabs-public button { padding: 13px 26px; border-radius: var(--radius-pill); border: 1.5px solid var(--border-strong); background: var(--surface); color: var(--text-muted); font-weight: 700; cursor:pointer; transition: all 0.22s; font-size: 14px; font-family: var(--font); }
.tabs-public button.active { background: linear-gradient(135deg, var(--primary) 0%, var(--primary-hover) 100%); color:#fff; border-color: var(--primary); box-shadow: var(--shadow-primary); }
.tabs-public button:hover:not(.active) { background: var(--primary-light); color: var(--primary-hover); border-color: var(--primary); }
.grid-row { display:grid; grid-template-columns: repeat(12, 1fr); gap: 20px; }
.col-6 { grid-column: span 6; }
.col-12 { grid-column: span 12; }

.pub-tools { position: fixed; top: 14px; right: 14px; z-index: 150; display:flex; gap:8px; }

.app-header { display:flex; justify-content:space-between; align-items:center; background: rgba(255,255,255,0.95); padding: 0 24px; height: 70px; border-bottom: 1px solid var(--border); position: fixed; top:0; left:0; right:0; z-index: 100; box-shadow: var(--shadow-sm); }
.app-header .header-left { display:flex; align-items:center; min-width:0; flex:1; }
.app-header h5 { font-size: 18px; font-weight: 800; display:flex; align-items:center; gap: 10px; white-space: nowrap; letter-spacing: -0.3px; }
.app-header h5 .header-logo { display: inline-flex; align-items: center; justify-content: center; width: 34px; height: 34px; border-radius: 10px; background: linear-gradient(135deg, var(--primary) 0%, var(--primary-hover) 100%); color: #fff; font-size: 16px; box-shadow: 0 4px 12px var(--primary-glow); }
.app-header h5 .header-title { background: linear-gradient(135deg, var(--primary-dark), var(--accent)); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; }
.app-header .user-badge { display:flex; align-items:center; gap:12px; flex-shrink: 0; }
.menu-toggle { display:none; background: var(--primary-light); border:none; font-size: 16px; color: var(--primary); cursor:pointer; margin-right:12px; padding: 9px 11px; border-radius: 10px; }
.hdr-item { position: relative; }
.icon-btn { position:relative; width:42px; height:42px; border-radius:12px; border:1.5px solid var(--border-strong); background: var(--surface); color: var(--text-soft); cursor:pointer; font-size:16px; font-weight:800; display:flex; align-items:center; justify-content:center; transition: all .18s; }
.icon-btn:hover { background: var(--primary-light); color: var(--primary); }
.bell-badge { position:absolute; top:-7px; right:-7px; min-width:20px; height:20px; padding:0 5px; border-radius:999px; background: var(--accent); color:#fff; font-size:11px; font-weight:800; display:flex; align-items:center; justify-content:center; border:2px solid #fff; line-height:1; }
.profile-btn { display:flex; align-items:center; gap:10px; background: var(--surface); border:1.5px solid var(--border-strong); border-radius:999px; padding:4px 14px 4px 4px; cursor:pointer; transition: all .18s; }
.profile-btn:hover { background: var(--primary-light); border-color: var(--primary); }
.avatar { width:34px; height:34px; border-radius:50%; background: linear-gradient(135deg, var(--primary), var(--accent)); color:#fff; font-weight:800; font-size:13px; display:flex; align-items:center; justify-content:center; flex-shrink:0; letter-spacing:0.3px; }
.avatar.lg { width:56px; height:56px; font-size:20px; }
.profile-meta { display:flex; flex-direction:column; line-height:1.2; text-align:left; max-width:150px; }
.profile-meta b { font-size:13px; font-weight:800; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; color: var(--text-main); text-transform: uppercase; }
.profile-meta small { font-size:10.5px; color: var(--text-muted); font-weight:800; text-transform:uppercase; letter-spacing:.5px; }
.profile-btn .fa-chevron-down { font-size:10px; color: var(--text-muted); }
.dd-panel { position:absolute; right:0; top:calc(100% + 12px); width:350px; max-width:calc(100vw - 24px); background: var(--surface); border:1px solid var(--border); border-radius: var(--radius); box-shadow: var(--shadow-xl); z-index:300; overflow:hidden; animation: fadeIn .15s; }
.dd-head { display:flex; justify-content:space-between; align-items:center; padding:14px 16px; border-bottom:1px solid var(--border); font-size:14px; }
.link-btn { background:none; border:none; color: var(--primary); font-weight:700; font-size:12px; cursor:pointer; }
.link-btn:hover { text-decoration: underline; }
.dd-body { max-height:360px; overflow-y:auto; }
.dd-foot { padding:12px; text-align:center; border-top:1px solid var(--border); background: var(--surface-alt); }
.dd-foot a { color: var(--primary); font-weight:700; font-size:13px; cursor:pointer; }
.nitem { padding:12px 16px; border-bottom:1px solid var(--border); cursor:pointer; position:relative; }
.nitem:hover { background: var(--surface-alt); }
.nitem.unread { background:#f8f5ff; padding-left:26px; }
.nitem.unread::before { content:''; position:absolute; left:11px; top:19px; width:8px; height:8px; border-radius:50%; background: var(--accent); }
.nitem .ntitle { font-weight:800; font-size:13px; }
.nitem .nmsg { font-size:12.5px; color: var(--text-muted); margin-top:2px; }
.nitem .ntime { font-size:11px; color: var(--text-dim); margin-top:4px; }
.nempty { padding:30px 16px; text-align:center; color: var(--text-muted); font-size:13px; }
.pp-head { display:flex; gap:14px; align-items:center; padding:18px 16px; border-bottom:1px solid var(--border); background: linear-gradient(135deg, #f5f3ff, #fff0f3); }
.pp-head .pp-name { font-weight:800; font-size:15px; }
.pp-head .pp-email { font-size:12px; color: var(--text-muted); word-break: break-all; text-transform: none; }
.pp-head .pp-role { display:inline-block; margin-top:4px; padding:2px 10px; border-radius:999px; background: var(--primary); color:#fff; font-size:10px; font-weight:800; letter-spacing:.6px; }
.pp-body { padding:14px 16px; }

.notif-popup { position:fixed; top:80px; right:20px; z-index:99999; background:#fff; border-left:5px solid var(--primary); border-radius:12px; padding:16px 20px; box-shadow:0 8px 32px rgba(30,27,75,0.2); max-width:340px; animation: slideInRight 0.3s ease; border-top:1px solid var(--border); border-right:1px solid var(--border); border-bottom:1px solid var(--border); }
.notif-popup .np-title { font-weight:800; font-size:14px; color:var(--text-main); margin-bottom:4px; padding-right:20px; }
.notif-popup .np-msg { font-size:13px; color:var(--text-muted); }
.notif-popup .np-close { position:absolute; top:8px; right:8px; background:none; border:none; font-size:18px; cursor:pointer; color:var(--text-muted); line-height:1; }

.app-layout { display:flex; margin-top: 70px; height: calc(100vh - 70px); }
.sidebar { width: 250px; background: linear-gradient(180deg, #ffffff 0%, var(--surface-alt) 100%); border-right: 1px solid var(--border); padding: 24px 14px; display: flex; flex-direction: column; gap: 4px; transition: transform 0.3s ease; flex-shrink: 0; overflow-y: auto; }
.sidebar a { display:flex; align-items:center; gap:12px; padding: 12px 16px; border-radius: var(--radius-sm); color: var(--text-muted); font-weight: 700; text-decoration:none; font-size: 13.5px; cursor:pointer; transition: all 0.18s; letter-spacing: -0.1px; }
.sidebar a i { font-size: 16px; width: 20px; text-align:center; }
.sidebar a:hover { background: var(--primary-light); color: var(--primary-hover); }
.sidebar a.active { background: linear-gradient(135deg, var(--primary) 0%, var(--primary-hover) 100%); color: #fff; box-shadow: 0 6px 16px var(--primary-glow); }
.side-label { display:none; padding:14px 16px 4px; font-size:10.5px; font-weight:800; letter-spacing:1px; text-transform:uppercase; color:var(--text-dim); }
body.is-admin .side-label.admin-only { display:block !important; }
.main-content { flex: 1; padding: 22px 20px; overflow-y: auto; background: transparent; min-width: 0; scroll-behavior: smooth; }
.page { display: none; }
.page.active { display: block; animation: fadeIn 0.2s ease-out; }
.page-title { font-size: 24px; font-weight: 800; margin-bottom: 18px; color: var(--text-main); display:flex; align-items: center; gap: 10px; letter-spacing: -0.6px; }

.stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px; }
.stat-card { background: var(--surface); padding: 16px 18px; border-radius: var(--radius); border: 1px solid var(--border); box-shadow: var(--shadow-sm); display: flex; align-items: center; gap: 14px; transition: transform 0.2s, box-shadow 0.2s; position: relative; overflow: hidden; }
.stat-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-lg); }
.stat-icon { width: 48px; height: 48px; border-radius: 13px; display: flex; align-items: center; justify-content: center; font-size: 20px; color: #fff; flex-shrink: 0; box-shadow: 0 6px 14px rgba(30, 27, 75, 0.15); }
.bg-blue { background: linear-gradient(135deg, #8b5cf6 0%, #6366f1 100%); }
.bg-orange { background: linear-gradient(135deg, #f59e0b 0%, #f43f5e 100%); }
.bg-red { background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%); }
.bg-green { background: linear-gradient(135deg, #10b981 0%, #059669 100%); }
.stat-info { min-width: 0; flex: 1; }
.stat-info .lbl { display: block; font-size: 10.5px; font-weight: 800; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.7px; white-space: nowrap; }
.stat-info .val { display: block; font-size: 25px; font-weight: 800; color: var(--text-main); margin-top: 2px; line-height: 1.1; letter-spacing: -0.6px; }

.modal-overlay { position: fixed; top:0; left:0; width:100%; height:100%; background: rgba(30, 27, 75, 0.55); z-index: 9999; display: flex; justify-content: center; align-items: center; padding: 16px; animation: fadeIn 0.2s forwards; }
.modal-content { background: var(--surface); width:100%; max-width: 500px; border-radius: var(--radius-lg); padding: 32px; box-shadow: var(--shadow-xl); transform: translateY(20px); animation: slideUp 0.3s forwards; max-height: 90vh; overflow-y: auto; border: 1px solid var(--border); }
.modal-header { display:flex; justify-content:space-between; align-items:center; margin-bottom: 24px; }
.modal-header h3 { font-size: 20px; font-weight: 800; letter-spacing: -0.4px; }
.modal-close { background: var(--bg-alt); border:none; font-size: 18px; color: var(--text-muted); cursor:pointer; width: 36px; height: 36px; border-radius: 10px; display: flex; align-items: center; justify-content: center; }
.modal-close:hover { background: var(--primary-light); color: var(--primary); }

.star-rating { display:flex; justify-content:center; gap: 12px; font-size: 42px; color: #e5dbff; cursor:pointer; margin: 22px 0; }
.star-rating i { transition: color 0.15s, transform 0.15s; }
.star-rating i:hover { transform: scale(1.15); }
.star-rating i.active { color: var(--warning); filter: drop-shadow(0 2px 6px rgba(245,158,11,0.4)); }

.photo-links { display:flex; flex-wrap: wrap; gap: 6px; justify-content: flex-start; }
.photo-links a { display: inline-flex; align-items: center; justify-content: center; width: 32px; height: 32px; background: var(--primary-light); color: var(--primary); border-radius: 8px; text-decoration: none; font-size: 13px; transition: all 0.18s; }
.photo-links a:hover { background: var(--primary); color: #fff; transform: scale(1.08); }
.wa-float { position: fixed; bottom: 30px; right: 30px; background: var(--surface); padding: 20px; border-radius: var(--radius); box-shadow: var(--shadow-xl); border-left: 5px solid #25D366; z-index: 9000; animation: slideInRight 0.4s cubic-bezier(0.4, 0, 0.2, 1); }
.wa-float .wa-title { font-weight: 800; color: var(--text-main); margin-bottom: 8px; font-size: 15px; }
.wa-float .wa-actions { display:flex; flex-direction: column; gap: 8px; margin-top: 12px; }
.loader-overlay { position: fixed; top:0; left:0; width:100%; height:100%; background: rgba(250, 248, 255, 0.8); z-index: 10000; display:flex; justify-content:center; align-items:center; }
.spinner { width: 52px; height: 52px; border: 4px solid var(--primary-light); border-top-color: var(--primary); border-right-color: var(--accent); border-radius: 50%; animation: spin 0.9s linear infinite; }
.skeleton { color: var(--text-muted); padding: 12px; font-size: 13px; }

.app-layout > * { min-width: 0; }

/* MD REPORT COMPACT FILTER BAR */
.md-bar { display:flex; gap:10px; flex-wrap:wrap; align-items:center; background:var(--surface); border:1px solid var(--border); padding:10px 12px; border-radius:var(--radius-sm); box-shadow:var(--shadow-xs); margin-bottom:10px; }
.md-bar .search-box { position:relative; flex:0 1 230px; max-width:250px; min-width:170px; }
.md-bar .search-box i { position:absolute; left:12px; top:50%; transform:translateY(-50%); color:var(--text-dim); font-size:12.5px; pointer-events:none; }
.md-bar .search-box input { padding:8px 12px 8px 34px; font-size:12.5px; height:36px; }
.md-bar select { width:auto; flex:0 0 140px; padding:7px 10px; font-size:12.5px; height:36px; }
.md-bar .date-range-inline { display:inline-flex; align-items:center; gap:6px; background:var(--bg-alt); padding:3px 8px; border-radius:10px; border:1px solid var(--border-strong); }
.md-bar .date-range-inline span { font-size:11px; font-weight:800; color:var(--text-soft); text-transform:uppercase; white-space:nowrap; }
.md-bar .date-range-inline input[type=date] { border:none; background:transparent; padding:4px 6px; height:28px; font-size:12px; width:125px; }
.md-bar .date-range-inline input[type=date]:focus { box-shadow:none; }
.md-chips { display:flex; gap:6px; flex-wrap:wrap; }
.md-chip { padding:6px 14px; border-radius:999px; border:1.5px solid var(--border-strong); background:var(--surface); color:var(--text-muted); font-weight:700; font-size:12px; cursor:pointer; font-family:var(--font); }
.md-chip.active { background:linear-gradient(135deg,var(--primary),var(--primary-hover)); color:#fff; border-color:var(--primary); box-shadow:var(--shadow-primary); }
.md-sec { border:1px solid var(--border); border-radius:14px; padding:16px 18px; margin-bottom:16px; background:var(--surface); }
.md-sec h4 { font-size:13px; font-weight:800; text-transform:uppercase; letter-spacing:.7px; color:var(--primary-hover); margin-bottom:12px; display:flex; align-items:center; gap:8px; }
.md-kv { display:grid; grid-template-columns:repeat(auto-fit,minmax(200px,1fr)); gap:12px 20px; }
.md-kv .k { font-size:10.5px; font-weight:800; color:var(--text-muted); text-transform:uppercase; letter-spacing:.6px; }
.md-kv .v { font-size:14px; font-weight:600; margin-top:2px; word-break:break-word; }
.md-visit { display:flex; gap:14px; padding:14px 0; border-bottom:1px dashed var(--border-strong); }
.md-visit:last-child { border-bottom:none; }
.md-vno { width:34px; height:34px; border-radius:50%; background:linear-gradient(135deg,var(--primary),var(--accent)); color:#fff; font-weight:800; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
.md-visit-body { flex:1; min-width:0; }
.md-visits-badge { display:inline-flex; align-items:center; justify-content:center; min-width:26px; height:24px; padding:0 7px; border-radius:999px; background:var(--primary-light); color:var(--primary-hover); font-weight:800; font-size:11.5px; }
.md-visits-badge.multi { background:#f59e0b; color:#fff; box-shadow:0 2px 8px rgba(245,158,11,0.4); }

@media (max-width: 900px) {
  .stats-grid { grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 20px; }
  .stat-card { padding: 14px 16px; gap: 12px; border-radius: 14px; min-height: 74px; }
  .stat-icon { width: 42px; height: 42px; border-radius: 12px; font-size: 17px; }
  .stat-info .lbl { font-size: 10px; letter-spacing: 0.4px; }
  .stat-info .val { font-size: 22px; margin-top: 1px; }
  .app-header { padding: 0 14px; }
  .app-header h5 .header-logo { display: none; }
  .app-header h5 .header-title { font-size: 15px; }
  .app-header .user-badge { gap: 8px; }
  .icon-btn { width:38px; height:38px; }
  .profile-meta { max-width: 110px; }
  .menu-toggle { display: block; }
  .sidebar { position: fixed; top: 70px; left: 0; height: calc(100vh - 70px); transform: translateX(-100%); z-index: 99; box-shadow: var(--shadow-xl); }
  .sidebar.open { transform: translateX(0); }
  .main-content { padding: 20px 16px; }
  .col-6 { grid-column: span 12; }
  .page-title { font-size: 22px; }
  .wa-float { bottom: 20px; right: 20px; left: 20px; }
  .notif-popup { left: 12px; right: 12px; max-width: none; }
  .toolbar .search-box, .md-bar .search-box { flex: 1 1 100%; max-width: none; }
}

@media (max-width: 600px) {
  .stats-grid { grid-template-columns: 1fr; gap: 10px; }
  .stat-card { padding: 13px 16px; gap: 14px; min-height: 68px; border-radius: 14px; }
  .stat-icon { width: 46px; height: 46px; border-radius: 12px; font-size: 19px; flex-shrink: 0; }
  .stat-info { display: flex; align-items: center; justify-content: space-between; gap: 12px; width: 100%; min-width: 0; }
  .stat-info .lbl { font-size: 13px; letter-spacing: 0.3px; font-weight: 700; text-transform: none; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
  .stat-info .val { font-size: 24px; margin: 0; flex-shrink: 0; }
  .card { padding: 16px; border-radius: 14px; }
  .grid-row { gap: 12px; }
  .public-header h1 { font-size: 30px; }
  .profile-meta, .profile-btn .fa-chevron-down { display: none; }
  .profile-btn { padding: 3px; }
  .dd-panel { position: fixed; top: 76px; left: 12px; right: 12px; width: auto; max-width: none; }
  .toolbar select, .toolbar input[type=date] { flex: 1 1 45%; }
  .toolbar .tb-count { margin-left: 0; width: 100%; }
  table.detail-table, table.detail-table tbody, table.detail-table tr, table.detail-table th, table.detail-table td { display: block; width: 100% !important; min-width: 0; }
  table.detail-table tr { padding: 12px 4px; border-bottom: 1px solid var(--border); }
  table.detail-table tr:last-child { border-bottom: none; }
  table.detail-table th { background: transparent !important; color: var(--text-muted); font-size: 10.5px; font-weight: 800; letter-spacing: 0.6px; text-transform: uppercase; padding: 2px 0 0; border: none; }
  table.detail-table td { padding: 2px 0 10px; border: none !important; font-size: 14px; word-break: break-word; }
  .table-wrap-cards { border: none !important; background: transparent !important; overflow: visible !important; }
  .table-wrap-cards .responsive-table { display: block; }
  .table-wrap-cards .responsive-table thead { display: none; }
  .table-wrap-cards .responsive-table tbody { display: block; }
  .table-wrap-cards .responsive-table tr { display: block; background: var(--surface); border: 1px solid var(--border); border-radius: 14px; margin-bottom: 12px; padding: 12px 14px; box-shadow: var(--shadow-sm); }
  .table-wrap-cards .responsive-table tr {
    display: flex;
    flex-direction: column;
    gap: 6px;
    background: var(--surface);
    border: 1px solid var(--border-strong);
    border-radius: 14px;
    margin-bottom: 14px;
    padding: 14px;
    box-shadow: var(--shadow-sm);
  }
  .table-wrap-cards .responsive-table td {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 10px;
    padding: 6px 0;
    border-bottom: 1px dashed var(--border);
    text-align: right;
    font-size: 13px;
    width: 100% !important;
  }
  .table-wrap-cards .responsive-table td:last-child {
    border-bottom: none;
    padding-top: 10px;
    justify-content: flex-end;
  }
  .table-wrap-cards .responsive-table td::before {
    content: attr(data-label);
    font-weight: 800;
    color: var(--text-muted);
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    text-align: left;
    flex-shrink: 0;
    max-width: 40%;
  }
  .table-wrap-cards .responsive-table td > * {
    max-width: 60%;
    word-break: break-word;
  }
  .table-wrap-cards .responsive-table td[data-label="Action"]::before {
    display: none;
  }
  .table-wrap-cards .responsive-table td[data-label="Action"] {
    justify-content: flex-end;
  }
  .table-wrap-cards .responsive-table td[data-label="Action"] .btn {
    width: 100%;
    padding: 10px 16px;
  }
  /* Admin MD Report Mobile Card: Show only main details (Ticket, Date, Company & Location, Warranty, Engineer, Status), hide dense fields from card, full info in View Full Details */
  #page-mdreport .table-wrap-cards .responsive-table .md-extra-cell {
    display: none !important;
  }
  #page-mdreport .table-wrap-cards .responsive-table tr {
    padding: 16px;
    background: var(--surface);
    border: 1.5px solid var(--border-strong);
    border-radius: 16px;
    margin-bottom: 14px;
    box-shadow: var(--shadow-md);
  }
  #page-mdreport .table-wrap-cards .responsive-table td {
    padding: 8px 0;
    font-size: 13.5px;
  }
  .table-wrap-cards .responsive-table tr.empty-row { text-align: center; padding: 28px; color: var(--text-muted); }
  .table-wrap-cards .responsive-table tr.empty-row td { display: block; text-align: center; border: none; }
  .table-wrap-cards .responsive-table tr.empty-row td::before { display: none; }
}

/* ===== CLEAN TABLE GRID (No vertical letter-by-letter wrapping!) ===== */
:root { --grid:#cfc6f3; }
body.dark { --grid:#66718f; }
@media (min-width: 601px) {
  .table-wrap {
    border: 1px solid var(--grid);
    overflow-x: auto;
    scrollbar-width: thin;
    scrollbar-color: var(--primary-light) transparent;
  }
  .table-wrap::-webkit-scrollbar {
    height: 7px;
  }
  .table-wrap::-webkit-scrollbar-thumb {
    background: var(--primary-light);
    border-radius: 4px;
  }
  table { border-collapse: collapse; width: 100%; min-width: 1000px; table-layout: auto; }
  th, td {
    border: 1px solid var(--grid) !important;
    padding: 8px 10px;
    font-size: 12px;
    vertical-align: middle;
    text-align: left;
    word-break: normal;
    overflow-wrap: break-word;
  }
  th {
    background: var(--primary-light);
    color: var(--text-soft);
    font-size: 10.5px;
    letter-spacing: 0.3px;
    white-space: nowrap !important;
  }
  td .btn-sm { padding: 6px 10px; font-size: 11.5px; }
  td .badge { font-size: 9.5px; padding: 4px 8px; }
  tbody tr:nth-child(even) td { background: var(--surface-alt); }
  tbody tr.unread-row:nth-child(n) td { background: #f8f5ff; }
  body.dark tbody tr.unread-row:nth-child(n) td { background: #3f4863; }
  tr.empty-row td { text-align: center; padding: 28px; }
  td .photo-links { justify-content: flex-start; }
  td[data-label="Ticket"],
  td[data-label="Date"],
  td[data-label="Reg. Date"],
  td[data-label="Assigned On"],
  td[data-label="Status"],
  td[data-label="Billing"],
  td[data-label="Bill"],
  td[data-label="Action"],
  td[data-label="Phone"],
  td[data-label="Engineer"] { white-space: nowrap; }
}
</style>

<script>
var SERVER_PARAMS = ${paramData};
var WEB_URL = ${JSON.stringify(getWebAppUrl() || '')};
</script>
</head>
<body>

<div id="globalLoader" class="loader-overlay hidden"><div class="spinner"></div></div>

<div id="pubTools" class="pub-tools">
  <button class="icon-btn js-lang" onclick="toggleLang()" title="Language"></button>
  <button class="icon-btn js-theme" onclick="toggleTheme()" title="Dark mode"></button>
</div>

<div id="publicView">
  <div class="public-wrap">
    <div class="public-header">
      <h1>BHT Service Portal</h1>
      <p>Register and track your service requests effortlessly.</p>
    </div>
    <div class="tabs-public">
      <button id="pubTabForm" class="active" onclick="pubShow('form')"><i class="fa-solid fa-plus-circle"></i> Register</button>
      <button id="pubTabTrack" onclick="pubShow('track')"><i class="fa-solid fa-magnifying-glass"></i> Track Status</button>
    </div>

    <div id="pubForm" class="card animate-fade">
      <h2 class="card-title"><i class="fa-solid fa-clipboard-list" style="color:var(--primary);"></i> Complaint Details</h2>
      <div class="grid-row">
        <div class="col-6"><label>Company Name *</label><input id="cf_company" placeholder="ABC Corp"></div>
        <div class="col-6"><label>Authorized Person *</label><input id="cf_person" placeholder="John Doe"></div>
        <div class="col-6"><label>Email (Optional)</label><input id="cf_email" type="email" placeholder="you@company.com"></div>
        <div class="col-6"><label>Phone *</label><input id="cf_phone" placeholder="9876543210"></div>
        <div class="col-12"><label>Address *</label><input id="cf_address" placeholder="Street, Area"></div>
        <div class="col-6"><label>Location / City *</label><input id="cf_location" placeholder="Mumbai"></div>
        <div class="col-6"><label>Equipment *</label>
          <select id="cf_equipment"><option value="">Select Equipment</option><option value="ELECTRIC WIRE HOIST">ELECTRIC WIRE HOIST</option><option value="DOUBLE- SGIB TYPE">DOUBLE- SGIB TYPE</option><option value="SINGLE- SGIB TYPE">SINGLE- SGIB TYPE</option><option value="EOT CRANE - DOUBLE G BOX TYPE">EOT CRANE - DOUBLE G BOX TYPE</option><option value="EOT CRANE - SINGLE G BOX TYPE">EOT CRANE - SINGLE G BOX TYPE</option><option value="CHAIN HOIST">CHAIN HOIST</option><option value="CHAIN PULLY BLOCK">CHAIN PULLY BLOCK</option><option value="OTHER">OTHER</option></select>
        </div>
        <div class="col-6"><label>Serial No (Optional)</label><input id="cf_serial" placeholder="SN-12345"></div>
        <div class="col-12"><label>Problem Description *</label><textarea id="cf_problem" rows="3" placeholder="Describe the issue..."></textarea></div>
      </div>
      <button id="btnSubmit" class="btn btn-primary btn-block" style="margin-top:24px;" onclick="submitPublicComplaint()"><i class="fa-solid fa-paper-plane"></i> Submit Complaint</button>
    </div>

    <div id="pubTrack" class="card hidden animate-fade">
      <h2 class="card-title"><i class="fa-solid fa-location-dot" style="color:var(--primary);"></i> Track Your Request</h2>
      <p style="color:var(--text-muted); margin-bottom:20px; font-size:14px;">Enter your ticket number and the registered mobile number to view your complaint status.</p>
      <div class="grid-row">
        <div class="col-6"><label>Ticket Number</label><input id="tr_ticket" placeholder="BHT-0001"></div>
        <div class="col-6"><label>Mobile Number</label><input id="tr_phone" placeholder="9876543210"></div>
      </div>
      <button class="btn btn-primary btn-block" style="margin-top:24px;" onclick="trackComplaint()"><i class="fa-solid fa-magnifying-glass"></i> Track Now</button>
      <div style="margin-top:24px;">
        <div id="trackResult"></div>
      </div>
    </div>

    <div id="pubThanks" class="card hidden animate-fade" style="text-align:center; padding:48px 24px;">
      <div style="width:80px; height:80px; background:linear-gradient(135deg, var(--success), var(--success-hover)); color:#fff; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:40px; margin:0 auto 24px; box-shadow:0 10px 24px rgba(16,185,129,0.35);"><i class="fa-solid fa-check"></i></div>
      <h2 style="font-size:28px; margin-bottom:8px; font-weight:800; letter-spacing:-0.6px;">Success!</h2>
      <p style="color:var(--text-muted); margin-bottom:24px;">Your complaint has been registered.</p>
      <div style="background:var(--bg-alt); border:1.5px dashed var(--border-strong); padding:20px; border-radius:14px; margin-bottom:24px;">
        <p style="font-size:12px; color:var(--text-muted); text-transform:uppercase; font-weight:800; letter-spacing:1.2px;">Ticket Number</p>
        <div id="thanksTicket" style="font-size:36px; font-weight:800; color:var(--primary); margin-top:8px; letter-spacing:-1px;">BHT-0000</div>
      </div>
      <div style="display:flex; gap:12px; justify-content:center; flex-wrap:wrap;">
        <button class="btn btn-secondary" onclick="pubShow('form')">Register Another</button>
      </div>
    </div>

    <div style="text-align:center; margin-top:40px;">
      <a style="color:var(--text-muted); font-size:14px; cursor:pointer; font-weight:700; text-decoration:none;" onclick="showStaffLogin()">Staff Login <i class="fa-solid fa-arrow-right"></i></a>
    </div>
  </div>
</div>

<div id="staffLogin" class="hidden" style="min-height:100vh; display:flex; align-items:center; justify-content:center; padding:20px;">
  <div class="card animate-fade" style="width:100%; max-width:440px; padding:40px;">
    <div style="text-align:center; margin-bottom:32px;">
      <div style="width:64px; height:64px; background:linear-gradient(135deg, var(--primary), var(--primary-hover)); color:#fff; border-radius:18px; display:flex; align-items:center; justify-content:center; font-size:26px; margin:0 auto 16px; box-shadow:var(--shadow-primary);"><i class="fa-solid fa-shield-halved"></i></div>
      <h2 style="font-size:24px; font-weight:800; letter-spacing:-0.5px;">Staff Portal</h2>
      <p style="color:var(--text-muted); font-size:14px; margin-top:4px;">Sign in to manage operations</p>
    </div>
    <div style="margin-bottom:16px;"><label>Email Address or Name</label><input id="loginEmail" type="text" class="keep-case" placeholder="you@company.com or VIVEK"></div>
    <div style="margin-bottom:24px;"><label>Password</label><input id="loginPassword" type="password" placeholder="••••••••" onkeydown="if(event.key==='Enter')doLogin()"></div>
    <button class="btn btn-primary btn-block" onclick="doLogin()">Sign In</button>
    <div style="text-align:center; margin-top:24px; font-size:13px; color:var(--text-muted);">
      <a style="cursor:pointer; text-decoration:underline;" onclick="backToPublic()">Return to Public Site</a>
    </div>
  </div>
</div>

<div id="appShell" class="hidden">
  <header class="app-header">
    <div class="header-left">
      <button class="menu-toggle" onclick="toggleSidebar()"><i class="fa-solid fa-bars"></i></button>
      <h5>
        <i class="fa-solid fa-screwdriver-wrench header-logo"></i>
        <span class="header-title">BHT System</span>
      </h5>
    </div>
    <div class="user-badge">
      <button class="icon-btn js-lang" onclick="toggleLang()" title="Language"></button>
      <button class="icon-btn js-theme" onclick="toggleTheme()" title="Dark mode"></button>
      <div class="hdr-item">
        <button class="icon-btn" onclick="toggleNotifPanel(event)" title="Notifications"><i class="fa-solid fa-bell"></i><span id="notifBadge" class="bell-badge hidden">0</span></button>
        <div class="dd-panel hidden" id="notifPanel">
          <div class="dd-head"><b><i class="fa-solid fa-bell" style="color:var(--primary);"></i> Notifications</b><button class="link-btn" onclick="markAllRead()">Mark all read</button></div>
          <div class="dd-body" id="notifPanelBody"></div>
          <div class="dd-foot"><a onclick="openNotifPage()">View all notifications</a></div>
        </div>
      </div>
      <div class="hdr-item">
        <button class="profile-btn" onclick="toggleProfilePanel(event)" title="Profile">
          <span class="avatar" id="hdrAvatar">U</span>
          <span class="profile-meta"><b id="hdrName">User</b><small id="hdrRole">ROLE</small></span>
          <i class="fa-solid fa-chevron-down"></i>
        </button>
        <div class="dd-panel hidden" id="profilePanel" style="width:300px;">
          <div class="pp-head">
            <span class="avatar lg" id="ppAvatar">U</span>
            <div style="min-width:0;">
              <div class="pp-name" id="ppName"></div>
              <div class="pp-email" id="ppEmail"></div>
              <span class="pp-role" id="ppRole"></span>
            </div>
          </div>
          <div class="pp-body">
            <button class="btn btn-danger btn-block btn-sm" onclick="doLogout()"><i class="fa-solid fa-arrow-right-from-bracket"></i> Logout</button>
          </div>
        </div>
      </div>
    </div>
  </header>

  <div class="app-layout">
    <aside class="sidebar" id="sidebar">
      <a data-page="mdreport" class="admin-only" onclick="goTo('mdreport')"><i class="fa-solid fa-crown"></i> MD Report</a>

      <div class="side-label admin-only">Coordinator</div>
      <a data-page="complaints" onclick="goTo('complaints')"><i class="fa-solid fa-clipboard-list"></i> Complaints</a>

      <div class="side-label admin-only">Engineer / HR</div>
      <a data-page="jobs" onclick="goTo('jobs')"><i class="fa-solid fa-screwdriver-wrench"></i> My Jobs</a>
      <a data-page="serviceReports" onclick="goTo('serviceReports')"><i class="fa-solid fa-file-signature"></i> Service Reports</a>

      <div class="side-label admin-only">Accounts</div>
      <a data-page="billing" onclick="goTo('billing')"><i class="fa-solid fa-money-bill-wave"></i> Billing</a>
      <a data-page="invoices" onclick="goTo('invoices')"><i class="fa-solid fa-receipt"></i> Invoices</a>

      <div class="side-label admin-only">Admin</div>
      <a data-page="users" class="admin-only" onclick="goTo('users')"><i class="fa-solid fa-users"></i> Users</a>
      <a data-page="notifications" onclick="goTo('notifications')"><i class="fa-solid fa-bell"></i> Notifications</a>
    </aside>

    <main class="main-content" id="mainContent">

      <!-- ============ MD REPORT ============ -->
      <div class="page" id="page-mdreport">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:14px;">
          <div>
            <h2 class="page-title" style="margin:0;"><i class="fa-solid fa-crown" style="color:var(--warning);"></i> MD Report</h2>
            <p style="color:var(--text-muted); font-size:12.5px; margin-top:2px;">Complete ticket-wise report in ascending sequence: customer, site, equipment, warranty, engineer, visits, charges/KM, and billing</p>
          </div>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <button class="btn btn-secondary btn-sm filter-toggle-btn" id="md_ftog" onclick="toggleMDFilters()" title="Show / Hide Search & Filters"><i class="fa-solid fa-filter"></i> <span>Filters</span></button>
            <button class="btn btn-secondary btn-sm" onclick="loadMD()"><i class="fa-solid fa-rotate-right"></i> Refresh</button>
            <span class="export-bar">
              <button class="btn btn-pdf btn-sm" onclick="exportView('md','pdf')"><i class="fa-solid fa-file-pdf"></i> PDF</button>
              <button class="btn btn-print btn-sm" onclick="exportView('md','print')"><i class="fa-solid fa-print"></i> Print</button>
              <button class="btn btn-csv btn-sm" onclick="exportView('md','csv')"><i class="fa-solid fa-file-excel"></i> Excel</button>
            </span>
          </div>
        </div>

        <div id="md_filter_container">
        <div class="md-bar">
          <div class="search-box">
            <i class="fa-solid fa-magnifying-glass"></i>
            <input id="md_q" data-lbl="Search" placeholder="Search ticket, company..." oninput="renderMD()">
          </div>
          <div class="date-range-inline">
            <span><i class="fa-regular fa-calendar"></i> From</span>
            <input type="date" id="md_from" data-lbl="From Date" onchange="renderMD()">
            <span>To</span>
            <input type="date" id="md_to" data-lbl="To Date" onchange="renderMD()">
          </div>
          <select id="md_eng" data-lbl="Engineer" onchange="renderMD()"><option value="">All Engineers</option></select>
          <select id="md_per" data-lbl="Period" onchange="renderMD()">
            <option value="">All Time</option>
            <option value="today">Today</option>
            <option value="7">Last 7 Days</option>
            <option value="30">Last 30 Days</option>
            <option value="month">This Month</option>
            <option value="year">This Year</option>
          </select>
          <button class="btn btn-secondary btn-sm" onclick="resetMDFilters()" title="Reset Filters"><i class="fa-solid fa-rotate-left"></i></button>
        </div>
        <div class="md-chips" id="md_chips">
          <button class="md-chip active" onclick="mdSetStatus(this,'ALL')">All</button>
          <button class="md-chip" onclick="mdSetStatus(this,'OPEN')">Open</button>
          <button class="md-chip" onclick="mdSetStatus(this,'PENDING')">Pending</button>
          <button class="md-chip" onclick="mdSetStatus(this,'COMPLETED')">Completed</button>
          <button class="md-chip" onclick="mdSetStatus(this,'CLOSED')">Closed</button>
          <button class="md-chip" onclick="mdSetStatus(this,'MULTI_VISIT')">🔥 2+ Visits</button>
          <button class="md-chip" onclick="mdSetStatus(this,'SEQ_GAP')">⚠️ Pending Sequence Gap</button>
        </div>
        </div>
        <div id="md_cnt" style="font-size:12px; font-weight:700; color:var(--text-muted); margin:6px 0 10px;"></div>

        <div class="table-wrap table-wrap-cards">
          <table class="responsive-table">
            <thead>
              <tr>
                <th>Ticket</th>
                <th>Date</th>
                <th>Company & Location</th>
                <th>Site Person</th>
                <th>Equipment & Problem</th>
                <th>Warranty</th>
                <th>Engineer</th>
                <th>Visits</th>
                <th>Status</th>
                <th>Charges / KM</th>
                <th>Bill</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody id="md_body"></tbody>
          </table>
        </div>
      </div>

      <!-- ============ COMPLAINTS ============ -->
      <div class="page" id="page-complaints">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px;">
          <h2 class="page-title" style="margin:0;">Complaints</h2>
          <button class="btn btn-secondary btn-sm" onclick="refreshComplaints()"><i class="fa-solid fa-rotate-right"></i> Refresh</button>
        </div>
        <div class="stats-grid" id="statsComplaints"></div>
        <div class="toolbar" id="tb_cmp"></div>
        <div class="table-wrap table-wrap-cards">
          <table class="responsive-table">
            <thead>
              <tr>
                <th>Ticket</th><th>Date</th><th>Company</th><th>Person</th><th>Phone</th><th>Location</th>
                <th>Warranty</th><th>Status</th><th>Engineer</th><th>Visits</th><th>Billing</th><th>Action</th>
              </tr>
            </thead>
            <tbody id="complaintTableBody"></tbody>
          </table>
        </div>
      </div>

      <div class="page" id="page-complaintDetail">
        <h2 class="page-title">Complaint Details</h2>
        <div id="complaintDetailCard"></div>
        <div id="complaintWaPanel"></div>
        <div class="card" style="margin-top:24px;">
          <h3 class="card-title"><i class="fa-solid fa-clock-rotate-left" style="color:var(--primary);"></i> Service History</h3>
          <div class="toolbar" id="tb_his"></div>
          <div id="complaintServiceReports"></div>
        </div>
      </div>

      <div class="page" id="page-assignEngineer">
        <h2 class="page-title">Assign Engineer</h2>
        <div class="card" style="max-width:600px; margin:0 auto;">
          <div class="grid-row">
            <div class="col-12">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
                <label style="margin:0;">Engineer *</label>
                <button type="button" class="link-btn" onclick="reloadEngineersList()" title="Sync latest list from Google Sheet"><i class="fa-solid fa-rotate-right"></i> Sync from Sheet</button>
              </div>
              <select id="ae_engineer"><option value="">Select Engineer</option></select>
            </div>
            <div class="col-12">
              <label>Other / Accompanying Person Name (Optional)</label>
              <input id="ae_otherPerson" placeholder="e.g. Second person name alongside selected engineer">
            </div>
            <div class="col-12"><label>Warranty Status *</label>
              <select id="ae_warranty">
                <option value="UNDER_WARRANTY">Under Warranty</option>
                <option value="OUT_OF_WARRANTY">Out of Warranty</option>
                <option value="UNDER_AMC">Under AMC</option>
                <option value="NEW_EQUIPMENT_INSTALLED">New Equipment Installed</option>
                <option value="NEW_ORDER_DIMENSION_SITE_VISIT">New Order Dimension / Site Visit</option>
              </select>
            </div>
            <div class="col-12"><label>Material Required?</label>
              <select id="ae_matReq" onchange="toggleMaterialDetails()"><option value="false">No</option><option value="true">Yes</option></select>
            </div>
            <div class="col-12 hidden" id="ae_matWrap"><label>Material Details</label><input id="ae_matDet" placeholder="e.g. Capacitor"></div>
            <div class="col-12"><label>Remarks</label><input id="ae_remarks" placeholder="Any note..."></div>
          </div>
          <div style="display:flex; gap:12px; margin-top:24px;">
            <button class="btn btn-primary" style="flex:1;" onclick="submitAssign(this)">Assign Job</button>
            <button class="btn btn-secondary" style="flex:1;" onclick="goTo(&quot;complaints&quot;)">Cancel</button>
          </div>
        </div>
      </div>

      <div class="page" id="page-jobs">
        <h2 class="page-title">My Jobs</h2>
        <div class="stats-grid" id="statsJobs"></div>
        <div class="tabs" style="display:flex; gap:8px; margin-bottom:16px; border-bottom:1px solid var(--border); padding-bottom:12px; overflow-x:auto;">
          <button class="btn btn-primary btn-sm" onclick="switchJobTab(this,'New')">New</button>
          <button class="btn btn-secondary btn-sm" onclick="switchJobTab(this,'Active')">Active</button>
          <button class="btn btn-secondary btn-sm" onclick="switchJobTab(this,'Pending')">Pending</button>
          <button class="btn btn-secondary btn-sm" onclick="switchJobTab(this,'Completed')">Completed</button>
        </div>
        <div class="toolbar" id="tb_job"></div>
        <div id="jobCards" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(280px, 1fr)); gap:20px;"></div>
      </div>

      <div class="page" id="page-jobDetail">
        <h2 class="page-title">Job Details</h2>
        <div id="jobDetailCard"></div>
      </div>

      <div class="page" id="page-serviceReport">
        <h2 class="page-title">Submit Service Report</h2>
        <div class="card" style="max-width:720px; margin:0 auto;">
          <div class="grid-row">
            <div class="col-12">
              <label>Status</label>
              <select id="sr_status" onchange="toggleReportFields()"><option value="COMPLETED">Complete</option><option value="PENDING">Pending</option></select>
            </div>
          </div>

          <div id="sr_completeFields" style="margin-top:20px;">
            <div class="grid-row">
              <div class="col-6"><label>Booklet No</label><input type="text" id="sr_bookletNo" placeholder="e.g. SB-123"></div>
              <div class="col-6"><label>Service Photos</label><input type="file" id="sr_serviceImages" accept="image/*" multiple></div>
              <div class="col-6"><label>Equipment Photos</label><input type="file" id="sr_equipPhotos" accept="image/*" multiple></div>
              <div class="col-6"><label>Service Charge (₹)</label><input type="number" id="sr_serviceCharge" value="0"></div>
              <div class="col-6"><label>KM Distance</label><input type="number" id="sr_km" value="0"></div>
              <div class="col-6"><label>Food Charges (₹)</label><input type="number" id="sr_food" value="0"></div>
              <div class="col-6"><label>Other Charges (₹)</label><input type="number" id="sr_other" value="0"></div>
              <div class="col-12"><label>Remarks</label><textarea id="sr_remarks" rows="2"></textarea></div>
            </div>
          </div>

          <div id="sr_pendingFields" style="display:none; margin-top:20px;">
            <div class="grid-row">
              <div class="col-12"><label>Equipment Photos</label><input type="file" id="sr_equipPhotosPend" accept="image/*" multiple></div>
              <div class="col-6"><label>KM Distance</label><input type="number" id="sr_pkm" value="0"></div>
              <div class="col-6"><label>Food Charges (₹)</label><input type="number" id="sr_pfood" value="0"></div>
              <div class="col-12">
                <label>Pending Reason</label>
                <select id="sr_pendingReason"><option value="">Select Reason</option><option value="Site not clear / not ready">Site not clear / not ready</option><option value="Site not accessible / closed">Site not accessible / closed</option><option value="Spare part not available">Spare part not available</option><option value="Customer / Authorized person not available">Customer / Authorized person not available</option><option value="Power / Electricity issue at site">Power / Electricity issue at site</option><option value="Work under progress / Next visit required">Work under progress / Next visit required</option><option value="Other">Other</option></select>
              </div>
              <div class="col-12"><label>Next Visit Date</label><input type="date" id="sr_nextVisit"></div>
            </div>
          </div>

          <div style="display:flex; gap:12px; margin-top:32px;">
            <button class="btn btn-primary" style="flex:1;" onclick="submitServiceReport(this)">Submit Report</button>
            <button class="btn btn-secondary" style="flex:1;" onclick="goTo(&quot;jobs&quot;)">Cancel</button>
          </div>
        </div>
      </div>

      <div class="page" id="page-serviceReports">
        <h2 class="page-title">Service Reports Log</h2>
        <div class="toolbar" id="tb_rep"></div>
        <div class="table-wrap table-wrap-cards">
          <table class="responsive-table">
            <thead>
              <tr>
                <th>Ticket</th><th>Date</th><th>Customer</th><th>Location</th><th>Engineer</th>
                <th>Status</th><th>Amount / Charges / KM</th><th>Photos</th><th>Action</th>
              </tr>
            </thead>
            <tbody id="serviceReportsBody"></tbody>
          </table>
        </div>
      </div>

      <div class="page" id="page-billing">
        <h2 class="page-title">Billing</h2>
        <div class="stats-grid" id="statsBilling"></div>
        <div class="toolbar" id="tb_bil"></div>
        <div class="table-wrap table-wrap-cards">
          <table class="responsive-table">
            <thead><tr><th>Ticket</th><th>Date</th><th>Company</th><th>Equipment</th><th>Action</th></tr></thead>
            <tbody id="pendingBillsBody"></tbody>
          </table>
        </div>
        <div id="bl_result" style="margin-top:32px;"></div>
      </div>

      <div class="page" id="page-invoices">
        <h2 class="page-title">Invoice History</h2>
        <div class="toolbar" id="tb_inv"></div>
        <div class="table-wrap table-wrap-cards">
          <table class="responsive-table">
            <thead><tr><th>Ticket No</th><th>Date</th><th>Invoice No</th><th>Customer</th><th>Amount</th><th>Bill Copy</th></tr></thead>
            <tbody id="invoiceTableBody"></tbody>
          </table>
        </div>
      </div>

      <div class="page" id="page-users">
        <h2 class="page-title">User Management</h2>
        <div style="display:flex; justify-content:flex-end; margin-bottom:12px;">
          <button class="btn btn-primary btn-sm" onclick="openUserModal()"><i class="fa-solid fa-plus"></i> Add User</button>
        </div>
        <div class="table-wrap table-wrap-cards">
          <table class="responsive-table">
            <thead><tr><th>Name</th><th>Email</th><th>Password</th><th>Phone</th><th>Role</th><th>Action</th></tr></thead>
            <tbody id="usersTableBody"></tbody>
          </table>
        </div>
      </div>

      <div class="page" id="page-notifications">
        <h2 class="page-title">Notifications</h2>
        <div class="toolbar" id="tb_ntf"></div>
        <div style="display:flex; justify-content:flex-end; margin-bottom:12px;">
          <button class="btn btn-secondary btn-sm" onclick="markAllRead()"><i class="fa-solid fa-check-double"></i> Mark all as read</button>
        </div>
        <div class="table-wrap table-wrap-cards">
          <table class="responsive-table">
            <thead><tr><th>Title</th><th>Message</th><th>Date</th></tr></thead>
            <tbody id="notifBody"></tbody>
          </table>
        </div>
      </div>
    </main>
  </div>
</div>

<!-- MODALS -->
<div id="mdModal" class="modal-overlay hidden" onclick="if(event.target===this)closeMDModal()">
  <div class="modal-content" style="max-width:940px; padding:0; overflow:hidden;">
    <div class="modal-header" style="padding:18px 24px; margin:0; background:var(--surface-alt); border-bottom:1px solid var(--border);">
      <h3 id="mdm_title" style="margin:0;">Ticket</h3>
      <button class="modal-close" onclick="closeMDModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <div id="mdm_body" style="padding:22px 24px; max-height:76vh; overflow-y:auto;"></div>
  </div>
</div>

<div id="billDetailsModal" class="modal-overlay hidden">
  <div class="modal-content" style="max-width: 640px; padding: 0; overflow: hidden;">
    <div class="modal-header" style="padding: 20px 24px; margin: 0; background: var(--surface-alt); border-bottom: 1px solid var(--border);">
      <h3 style="margin: 0; font-size: 18px;"><i class="fa-solid fa-file-invoice-dollar" style="color:var(--primary);"></i> Create Bill</h3>
      <button class="modal-close" onclick="closeBillDetailsModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <div style="padding: 24px;" id="billDetailsContent"></div>
  </div>
</div>

<div id="billModal" class="modal-overlay hidden">
  <div class="modal-content">
    <div class="modal-header">
      <h3><i class="fa-solid fa-file-invoice-dollar" style="color:var(--primary);"></i> Upload Invoice</h3>
      <button class="modal-close" onclick="closeBillModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <div class="grid-row">
      <div class="col-12"><label>Invoice Number *</label><input id="inv_num" placeholder="INV-1234"></div>
      <div class="col-12"><label>Amount (₹) *</label><input type="number" id="inv_amount" value="0"></div>
      <div class="col-12"><label>Upload Bill Copy (PDF/Image) *</label><input type="file" id="inv_file" accept="image/*,application/pdf"></div>
    </div>
    <div style="display:flex; gap:12px; margin-top:24px;">
      <button class="btn btn-success" style="flex:1;" onclick="submitNewInvoice(this)">Submit</button>
      <button class="btn btn-secondary" style="flex:1;" onclick="closeBillModal()">Cancel</button>
    </div>
  </div>
</div>

<div id="editModal" class="modal-overlay hidden">
  <div class="modal-content" style="max-width:640px;">
    <div class="modal-header">
      <h3><i class="fa-solid fa-pen-to-square" style="color:var(--primary);"></i> Edit Customer Complaint</h3>
      <button class="modal-close" onclick="closeEditModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <div class="grid-row">
      <div class="col-6"><label>Company Name *</label><input id="ed_company"></div>
      <div class="col-6"><label>Authorized Person *</label><input id="ed_person"></div>
      <div class="col-6"><label>Email (Optional)</label><input id="ed_email" type="email"></div>
      <div class="col-6"><label>Phone *</label><input id="ed_phone"></div>
      <div class="col-12"><label>Address *</label><input id="ed_address"></div>
      <div class="col-6"><label>Location / City *</label><input id="ed_location"></div>
      <div class="col-6"><label>Equipment *</label><input id="ed_equipment" list="ed_equipList"><datalist id="ed_equipList"><option value="ELECTRIC WIRE HOIST"></option><option value="DOUBLE- SGIB TYPE"></option><option value="SINGLE- SGIB TYPE"></option><option value="EOT CRANE - DOUBLE G BOX TYPE"></option><option value="EOT CRANE - SINGLE G BOX TYPE"></option><option value="CHAIN HOIST"></option><option value="CHAIN PULLY BLOCK"></option><option value="OTHER"></option></datalist></div>
      <div class="col-12"><label>Serial No (Optional)</label><input id="ed_serial"></div>
      <div class="col-12"><label>Problem Description *</label><textarea id="ed_problem" rows="3"></textarea></div>
    </div>
    <div style="display:flex; gap:12px; margin-top:24px;">
      <button class="btn btn-primary" style="flex:1;" onclick="submitEditComplaint(this)"><i class="fa-solid fa-floppy-disk"></i> Save Changes</button>
      <button class="btn btn-secondary" style="flex:1;" onclick="closeEditModal()">Cancel</button>
    </div>
  </div>
</div>

<div id="shortCloseModal" class="modal-overlay hidden">
  <div class="modal-content" style="max-width:480px;">
    <div class="modal-header">
      <h3><i class="fa-solid fa-lock" style="color:var(--danger);"></i> Short Close Ticket</h3>
      <button class="modal-close" onclick="closeShortCloseModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <p style="color:var(--text-muted); margin-bottom:16px;">Please provide a valid reason for short closing this ticket. This is required.</p>
    <textarea id="sc_reason" rows="4" placeholder="Reason for short close (required)" style="text-transform:none;"></textarea>
    <div style="display:flex; gap:12px; margin-top:20px;">
      <button class="btn btn-danger" style="flex:1;" onclick="submitShortClose(this)"><i class="fa-solid fa-lock"></i> Short Close</button>
      <button class="btn btn-secondary" style="flex:1;" onclick="closeShortCloseModal()">Cancel</button>
    </div>
  </div>
</div>

<div id="feedbackModal" class="modal-overlay hidden">
  <div class="modal-content" style="text-align:center;">
    <div class="modal-header" style="justify-content:center; position:relative;">
      <h3><i class="fa-solid fa-star" style="color:var(--warning);"></i> Rate Our Service</h3>
      <button class="modal-close" style="position:absolute; right:0; top:0;" onclick="closeFeedbackModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <p style="color:var(--text-muted); margin-bottom:16px;">Ticket: <b id="fb_ticket_display"></b></p>
    <div id="starContainer" class="star-rating">
      <i class="fa-solid fa-star" data-val="1"></i>
      <i class="fa-solid fa-star" data-val="2"></i>
      <i class="fa-solid fa-star" data-val="3"></i>
      <i class="fa-solid fa-star" data-val="4"></i>
      <i class="fa-solid fa-star" data-val="5"></i>
    </div>
    <div style="text-align:left; margin-bottom:12px;">
      <label>Registered Mobile Number *</label>
      <input id="fb_phone" placeholder="Mobile number used during registration" inputmode="numeric">
    </div>
    <p style="font-size:12px; color:var(--text-muted); margin-bottom:16px; text-align:left;">Only the registered customer can submit feedback.</p>
    <textarea id="fb_comment" rows="3" placeholder="Tell us about your experience..." style="margin-bottom:20px; text-transform:none;"></textarea>
    <button class="btn btn-primary btn-block" onclick="submitFeedbackForm(this)">Submit Feedback</button>
    <button class="btn btn-secondary btn-block" style="margin-top:8px;" onclick="closeFeedbackModal()">Cancel</button>
  </div>
</div>

<div id="userModal" class="modal-overlay hidden">
  <div class="modal-content">
    <div class="modal-header">
      <h3><i class="fa-solid fa-user-gear" style="color:var(--primary);"></i> <span id="um_title">Add User</span></h3>
      <button class="modal-close" onclick="closeUserModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <input type="hidden" id="um_id">
    <div class="grid-row">
      <div class="col-12"><label>Full Name *</label><input id="um_name" placeholder="John Doe"></div>
      <div class="col-12"><label>Email *</label><input id="um_email" type="email" placeholder="user@company.com"></div>
      <div class="col-12"><label>Password *</label><input id="um_password" type="text" placeholder="password123"></div>
      <div class="col-12"><label>Phone (Optional for Engineers)</label><input id="um_phone" placeholder="9876543210"></div>
      <div class="col-12"><label>Role *</label>
        <select id="um_role">
          <option value="ENGINEER">ENGINEER</option>
          <option value="COORDINATOR">COORDINATOR</option>
          <option value="ACCOUNTS">ACCOUNTS</option>
          <option value="HR">HR</option>
          <option value="ADMIN">ADMIN</option>
        </select>
      </div>
    </div>
    <div style="display:flex; gap:12px; margin-top:24px;">
      <button class="btn btn-primary" style="flex:1;" onclick="saveUserData(this)"><i class="fa-solid fa-floppy-disk"></i> Save</button>
      <button class="btn btn-secondary" style="flex:1;" onclick="closeUserModal()">Cancel</button>
    </div>
  </div>
</div>

<div id="serviceReportModal" class="modal-overlay hidden">
  <div class="modal-content" style="max-width:640px;">
    <div class="modal-header">
      <h3><i class="fa-solid fa-file-signature" style="color:var(--primary);"></i> Service Report Details</h3>
      <button class="modal-close" onclick="closeSRModal()"><i class="fa-solid fa-xmark"></i></button>
    </div>
    <div id="sr_modal_content"></div>
    <div style="text-align:right; margin-top:20px;">
      <button class="btn btn-secondary" onclick="closeSRModal()">Close</button>
    </div>
  </div>
</div>

<script>
window.onerror = function(msg, url, line) {
  var d = document.createElement('div');
  d.style.cssText = 'background:#ef4444;color:#fff;padding:12px 16px;position:fixed;top:0;left:0;right:0;z-index:999999;font-family:sans-serif;font-size:13px;box-shadow:0 4px 12px rgba(0,0,0,0.2);display:flex;justify-content:space-between;align-items:center;';
  d.innerHTML = '<span><b>Error:</b> ' + String(msg) + ' (line ' + line + ')</span><button style="background:none;border:none;color:#fff;font-size:16px;cursor:pointer;" onclick="this.parentElement.remove()">×</button>';
  document.body.appendChild(d);
  setTimeout(function(){ try { d.remove(); } catch(e){} }, 8000);
};

var App = {
  user: null, currentPage: 'complaints', pageParams: {},
  currentComplaint: null, currentEngineer: null, currentJob: null,
  currentFeedbackTicket: null, currentFeedbackRating: 0, currentFeedbackPhone: '', lastPublic: null,
  cache: { complaints: null, engineers: null, reports: null, invoices: null, users: null },
  trackData: [], historyData: [], readMem: null,
  freshLogin: false
};

var allJobs = [], jobTab = 'New', billingComplaint = null, billingReport = null;
var NOTIFS = [], NOTIF_TIMER = null, NOTIF_PREV = -1;
var VIEW = {};
var shortCloseId = null;
var MD_DATA = null, MD_ROWS = [], MD_STATUS = 'ALL';

function $(id) { return document.getElementById(id); }
function showLoader(v) { $('globalLoader').classList.toggle('hidden', !v); }
function isAdmin() { return !!(App.user && String(App.user.role).toUpperCase() === 'ADMIN'); }
function canEditComplaint() { return !!(App.user && ['ADMIN','COORDINATOR'].indexOf(String(App.user.role).toUpperCase()) > -1); }

/* ---------- Date Formatting (dd/mm/yy without time) & Sequence Sorting ---------- */
function fmtDate(val) {
  if (!val) return '-';
  var s = String(val).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3] + '/' + m[2] + '/' + m[1].slice(-2);
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var dd = ('0' + d.getDate()).slice(-2);
    var mm = ('0' + (d.getMonth() + 1)).slice(-2);
    var yy = String(d.getFullYear()).slice(-2);
    return dd + '/' + mm + '/' + yy;
  }
  return s.substring(0, 10);
}

function ticketNumVal(str) {
  var m = String(str || '').match(/(\d+)/);
  return m ? parseInt(m[1], 10) : 0;
}

function sortTicketsAsc(list) {
  return (list || []).slice().sort(function(a, b) {
    var na = ticketNumVal(a.ticketNumber || a.ticketId);
    var nb = ticketNumVal(b.ticketNumber || b.ticketId);
    if (na !== nb) return na - nb;
    return String(a.createdAt || a.submittedAt || '').localeCompare(String(b.createdAt || b.submittedAt || ''));
  });
}

function sortTicketsDesc(list) {
  return (list || []).slice().sort(function(a, b) {
    var na = ticketNumVal(a.ticketNumber || a.ticketId);
    var nb = ticketNumVal(b.ticketNumber || b.ticketId);
    if (na !== nb) return nb - na;
    return String(b.createdAt || b.submittedAt || '').localeCompare(String(a.createdAt || a.submittedAt || ''));
  });
}

function computeMaxDoneTicket(list) {
  var maxDone = 0;
  (list || []).forEach(function(c) {
    var st = String(c.status || '').toUpperCase();
    if (st === 'COMPLETED' || st === 'CLOSED') {
      var n = ticketNumVal(c.ticketNumber);
      if (n > maxDone) maxDone = n;
    }
  });
  return maxDone;
}

function formatWarrantyLabel(w) {
  var s = String(w || '').toUpperCase().trim();
  var map = {
    'UNDER_WARRANTY': 'UNDER WARRANTY',
    'IN_WARRANTY': 'UNDER WARRANTY',
    'OUT_OF_WARRANTY': 'OUT OF WARRANTY',
    'UNDER_AMC': 'UNDER AMC',
    'AMC': 'UNDER AMC',
    'NEW_EQUIPMENT_INSTALLED': 'NEW EQUIPMENT INSTALLED',
    'NEW_ORDER_DIMENSION_SITE_VISIT': 'NEW ORDER DIMENSION / SITE VISIT'
  };
  return map[s] || s || '-';
}

function warrantyMatchesFilter(itemW, filterW) {
  if (!filterW) return true;
  var iw = String(itemW || '').toUpperCase().trim();
  var fw = String(filterW || '').toUpperCase().trim();
  if (fw === 'UNDER_WARRANTY') return iw === 'UNDER_WARRANTY' || iw === 'IN_WARRANTY';
  if (fw === 'UNDER_AMC') return iw === 'UNDER_AMC' || iw === 'AMC';
  return iw === fw;
}

/* ---------- Audio & Notification ---------- */
var _audioCtx = null;
function initAudio() {
  try {
    if (!_audioCtx) _audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    if (_audioCtx && _audioCtx.state === 'suspended') _audioCtx.resume();
  } catch (e) {}
}
document.addEventListener('click', initAudio, { once: true });
document.addEventListener('touchstart', initAudio, { once: true });

function playNotifSound() {
  try {
    initAudio();
    if (!_audioCtx) return;
    var t = _audioCtx.currentTime;
    var notes = [880, 1175];
    notes.forEach(function(freq, i) {
      var osc = _audioCtx.createOscillator();
      var g = _audioCtx.createGain();
      osc.connect(g); g.connect(_audioCtx.destination);
      osc.frequency.value = freq;
      osc.type = 'sine';
      var start = t + i * 0.14;
      g.gain.setValueAtTime(0.0001, start);
      g.gain.exponentialRampToValueAtTime(0.35, start + 0.02);
      g.gain.exponentialRampToValueAtTime(0.0001, start + 0.25);
      osc.start(start);
      osc.stop(start + 0.28);
    });
  } catch (e) {}
}

function showNotifPopup(n) {
  try {
    var el = document.createElement('div');
    el.className = 'notif-popup';
    el.innerHTML = '<button class="np-close" onclick="this.parentElement.remove()">×</button>' +
      '<div class="np-title">🔔 ' + esc(n.title) + '</div>' +
      '<div class="np-msg">' + esc(n.message) + '</div>';
    document.body.appendChild(el);
    setTimeout(function() {
      try {
        el.style.opacity = '0';
        el.style.transition = 'opacity 0.4s';
        setTimeout(function() { try { el.remove(); } catch(e){} }, 400);
      } catch (e) {}
    }, 7000);
  } catch (e) {}
}

function requestBrowserNotifPermission() {
  try {
    if ('Notification' in window && Notification.permission === 'default') {
      Notification.requestPermission();
    }
  } catch (e) {}
}
function showBrowserNotif(n) {
  try {
    if ('Notification' in window && Notification.permission === 'granted' && document.hidden) {
      var b = new Notification('BHT: ' + n.title, { body: n.message });
      b.onclick = function() { try { window.focus(); b.close(); } catch(e){} };
    }
  } catch (e) {}
}

/* ---------- Theme & Language ---------- */
var LANG = 'en', THEME = 'light';
function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }

var HI = {
  'BHT Service Portal':'BHT सेवा पोर्टल', 'Register and track your service requests effortlessly.':'अपनी सेवा शिकायतें आसानी से दर्ज करें और ट्रैक करें।',
  'Register':'पंजीकरण', 'Track Status':'स्थिति ट्रैक करें', 'Complaint Details':'शिकायत विवरण',
  'Company Name *':'कंपनी का नाम *', 'Authorized Person *':'अधिकृत व्यक्ति *', 'Email *':'ईमेल *', 'Phone *':'फ़ोन *',
  'Address *':'पता *', 'Location / City *':'स्थान / शहर *', 'Equipment *':'उपकरण *', 'Select Equipment':'उपकरण चुनें',
  'AC Unit':'एसी यूनिट', 'Generator':'जनरेटर', 'Compressor':'कंप्रेसर', 'Other':'अन्य',
  'Serial No (Optional)':'सीरियल नं. (वैकल्पिक)', 'Problem Description *':'समस्या का विवरण *', 'Submit Complaint':'शिकायत जमा करें',
  'Track Your Request':'अपना अनुरोध ट्रैक करें',
  'Enter your ticket number and the registered mobile number to view your complaint status.':'अपनी शिकायत की स्थिति देखने के लिए टिकट नंबर और पंजीकृत मोबाइल नंबर दर्ज करें।',
  'Ticket Number':'टिकट नंबर', 'Mobile Number':'मोबाइल नंबर', 'Track Now':'अभी ट्रैक करें', 'Success!':'सफल!',
  'Your complaint has been registered.':'आपकी शिकायत दर्ज हो गई है।', 'Register Another':'एक और दर्ज करें',
  'Staff Login':'स्टाफ लॉगिन', 'Staff Portal':'स्टाफ पोर्टल', 'Sign in to manage operations':'संचालन प्रबंधित करने के लिए साइन इन करें',
  'Email Address':'ईमेल पता', 'Password':'पासवर्ड', 'Sign In':'साइन इन', 'Return to Public Site':'सार्वजनिक साइट पर लौटें',
  'BHT System':'BHT सिस्टम', 'Notifications':'सूचनाएं', 'Mark all read':'सभी पढ़ा हुआ करें', 'View all notifications':'सभी सूचनाएं देखें', 'Logout':'लॉग आउट',
  'Complaints':'शिकायतें', 'My Jobs':'मेरे कार्य', 'Service Reports':'सेवा रिपोर्ट', 'Billing':'बिलिंग', 'Invoices':'चालान', 'Users':'उपयोगकर्ता',
  'MD Report':'एमडी रिपोर्ट', 'Coordinator':'समन्वयक', 'Engineer / HR':'इंजीनियर / एचआर', 'Accounts':'लेखा', 'Admin':'एडमिन',
  'New':'नया', 'Assigned':'सौंपा गया', 'Pending':'लंबित', 'Completed':'पूर्ण', 'Active':'सक्रिय',
  'Refresh':'रिफ्रेश', 'Ticket':'टिकट', 'Company':'कंपनी', 'Equipment':'उपकरण', 'Status':'स्थिति', 'Engineer':'इंजीनियर', 'Action':'कार्रवाई', 'Open':'खोलें',
  'Service History':'सेवा इतिहास', 'Assign Engineer':'इंजीनियर नियुक्त करें', 'Warranty Status':'वारंटी स्थिति', 'In Warranty':'वारंटी में',
  'Out of Warranty':'वारंटी से बाहर', 'Material Required?':'सामग्री चाहिए?', 'No':'नहीं', 'Yes':'हाँ', 'Material Details':'सामग्री विवरण',
  'Remarks':'टिप्पणी', 'Assign Job':'कार्य सौंपें', 'Cancel':'रद्द करें', 'Job Details':'कार्य विवरण', 'Submit Service Report':'सेवा रिपोर्ट जमा करें',
  'Complete':'पूर्ण', 'Submit Report':'रिपोर्ट जमा करें', 'Booklet No':'बुकलेट नं.', 'Service Photos':'सेवा फोटो', 'Equipment Photos':'उपकरण फोटो',
  'Service Charge (₹)':'सेवा शुल्क (₹)', 'KM Distance':'किमी दूरी', 'Food Charges (₹)':'भोजन शुल्क (₹)', 'Other Charges (₹)':'अन्य शुल्क (₹)',
  'Pending Reason':'लंबित का कारण', 'Select Reason':'कारण चुनें', 'Spare part not available':'पुर्ज़ा उपलब्ध नहीं', 'Customer not available':'ग्राहक उपलब्ध नहीं',
  'Site not accessible':'साइट तक पहुँच नहीं', 'Next Visit Date':'अगली विज़िट की तारीख', 'Service Reports Log':'सेवा रिपोर्ट लॉग',
  'Customer':'ग्राहक', 'Booklet':'बुकलेट', 'Charges':'शुल्क', 'Photos':'फोटो', 'Submitted':'जमा किया', 'Invoice History':'चालान इतिहास',
  'Invoice No':'चालान नं.', 'Ticket No':'टिकट नं.', 'Amount':'राशि', 'Bill Copy':'बिल की प्रति', 'Date':'तारीख',
  'Upload Invoice':'चालान अपलोड करें', 'Invoice Number *':'चालान नंबर *', 'Amount (₹) *':'राशि (₹) *', 'Upload Bill Copy (PDF/Image) *':'बिल की प्रति अपलोड करें (PDF/इमेज) *',
  'Submit':'जमा करें', 'Rate Our Service':'हमारी सेवा को रेट करें', 'Registered Mobile Number *':'पंजीकृत मोबाइल नंबर *',
  'Only the registered customer can submit feedback.':'केवल पंजीकृत ग्राहक ही फीडबैक दे सकता है।', 'Submit Feedback':'फीडबैक जमा करें',
  'Warranty':'वारंटी', 'PDF':'PDF', 'Print':'प्रिंट', 'CSV':'Excel',
  'Mark all as read':'सभी को पढ़ा हुआ चिह्नित करें', 'Title':'शीर्षक', 'Message':'संदेश', 'Reset':'रीसेट',
  'Edit Customer Complaint':'ग्राहक शिकायत संपादित करें', 'Save Changes':'बदलाव सहेजें', 'Edit Details':'विवरण संपादित करें',
  'Complaints Report':'शिकायत रिपोर्ट', 'Person':'व्यक्ति', 'Email':'ईमेल', 'Phone':'फ़ोन', 'Address':'पता', 'Problem':'समस्या', 'Not specified':'निर्दिष्ट नहीं',
  'Unassigned':'नियुक्त नहीं', 'Rating':'रेटिंग', 'Feedback':'फीडबैक', 'Material Required':'सामग्री आवश्यक', 'None':'कोई नहीं', 'Standard':'मानक',
  'Short Close':'शॉर्ट क्लोज़', 'WhatsApp Customer':'ग्राहक को व्हाट्सऐप', 'Back to List':'सूची पर वापस', 'Back to Jobs':'कार्यों पर वापस',
  'Accept Job':'कार्य स्वीकारें', 'Mark On-Site':'साइट पर पहुँचे', 'Open Job':'कार्य खोलें', 'Create Bill':'बिल बनाएं',
  'Notify Customer':'ग्राहक को सूचित करें', 'Notify Engineer':'इंजीनियर को सूचित करें', 'Close':'बंद करें', 'Report Saved Successfully':'रिपोर्ट सफलतापूर्वक सहेजी गई',
  'All Status':'सभी स्थिति', 'All Warranty':'सभी वारंटी', 'All Equipment':'सभी उपकरण', 'All Engineer':'सभी इंजीनियर', 'All Customer':'सभी ग्राहक',
  'All':'सभी', 'Select Engineer':'इंजीनियर चुनें',
  'NEW':'नया', 'ASSIGNED':'सौंपा गया', 'ACCEPTED':'स्वीकृत', 'ON_SITE':'साइट पर', 'PENDING':'लंबित', 'COMPLETED':'पूर्ण', 'CLOSED':'बंद',
  'UNDER_WARRANTY':'वारंटी में', 'IN_WARRANTY':'वारंटी में', 'OUT_OF_WARRANTY':'वारंटी से बाहर', 'UNDER_AMC':'एएमसी में', 'NEW_EQUIPMENT_INSTALLED':'नया उपकरण स्थापित', 'NEW_ORDER_DIMENSION_SITE_VISIT':'नया ऑर्डर / साइट विज़िट',
  'Site Person':'साइट व्यक्ति', 'Location':'स्थान', 'Assigned On':'सौंपने की तारीख', 'Visits':'विज़िट', 'Last Visit':'अंतिम विज़िट', 'Bill':'बिल', 'View':'देखें',
  'All Time':'सभी समय', 'Today':'आज', 'Last 7 Days':'पिछले 7 दिन', 'Last 30 Days':'पिछले 30 दिन', 'This Month':'इस माह', 'This Year':'इस वर्ष',
  'No complaints registered yet':'अभी तक कोई शिकायत दर्ज नहीं', 'No complaints match your search / filters':'आपकी खोज से कोई शिकायत मेल नहीं खाती',
  'No service reports found':'कोई सेवा रिपोर्ट नहीं मिली', 'No reports match your search / filters':'आपकी खोज से कोई रिपोर्ट मेल नहीं खाती',
  'No pending bills found':'कोई लंबित बिल नहीं', 'No invoices recorded yet':'अभी कोई चालान दर्ज नहीं', 'No notifications found':'कोई सूचना नहीं मिली', 'No notifications yet':'अभी कोई सूचना नहीं',
  'No jobs found in this section.':'इस अनुभाग में कोई कार्य नहीं।', 'No service reports submitted yet.':'अभी तक कोई सेवा रिपोर्ट जमा नहीं।', 'Loading history…':'इतिहास लोड हो रहा है…',
  'Only Out-of-Warranty jobs require invoicing.':'केवल वारंटी से बाहर के कार्यों का चालान आवश्यक है।', 'Search':'खोजें',
  'Please wait...':'कृपया प्रतीक्षा करें...', 'Language':'भाषा', 'Dark mode':'डार्क मोड', 'Profile':'प्रोफ़ाइल',
  'Please fill in all required fields marked with *':'कृपया * वाले सभी आवश्यक फ़ील्ड भरें', 'Signed out successfully':'सफलतापूर्वक साइन आउट हुए',
  'Engineer assigned successfully':'इंजीनियर सफलतापूर्वक नियुक्त किया गया', 'Please select an engineer':'कृपया इंजीनियर चुनें',
  'Please enter your email and password':'कृपया ईमेल और पासवर्ड दर्ज करें', 'Invalid credentials':'अमान्य क्रेडेंशियल',
  'Describe the issue...':'समस्या बताएं...', 'Tell us about your experience...':'अपना अनुभव बताएं...', 'Any note...':'कोई नोट...',
  'NOT REQ':'आवश्यक नहीं'
};
var HI_REV = null;
var HI_PAT = [
  [/^(\d+) of (\d+) records(.*)$/, function(m) { return m[1] + ' / ' + m[2] + ' रिकॉर्ड' + m[3].replace(' • Total ', ' • कुल '); }],
  [/^New notification: (.+)$/, function(m) { return 'नई सूचना: ' + m[1]; }],
  [/^Assigned to (.+)$/, function(m) { return 'सौंपा गया: ' + m[1]; }],
  [/^Status updated to (.+)$/, function(m) { return 'स्थिति अपडेट: ' + m[1]; }],
  [/^Jobs - (.+)$/, function(m) { return 'कार्य - ' + m[1]; }]
];
var ORIG = new WeakMap(), OBS = null;

function trText(s) {
  var t = s.trim();
  if (!t) return null;
  var h = null;
  if (Object.prototype.hasOwnProperty.call(HI, t)) h = HI[t];
  else {
    for (var i = 0; i < HI_PAT.length; i++) { var m = t.match(HI_PAT[i][0]); if (m) { h = HI_PAT[i][1](m); break; } }
  }
  if (h === null) return null;
  return s.replace(t, function() { return h; });
}
function revText(s) {
  var t = s.trim();
  if (!t) return null;
  if (!HI_REV) { HI_REV = {}; Object.keys(HI).forEach(function(k) { if (!HI_REV[HI[k]]) HI_REV[HI[k]] = k; }); }
  if (Object.prototype.hasOwnProperty.call(HI_REV, t)) { var o = HI_REV[t]; return s.replace(t, function() { return o; }); }
  return null;
}
function transNode(n) {
  var p = n.parentNode;
  if (!p || /^(SCRIPT|STYLE|TEXTAREA)$/.test(p.nodeName)) return;
  var cur = n.nodeValue;
  if (LANG === 'hi') {
    var t = trText(cur);
    if (t !== null && t !== cur) { if (!ORIG.has(n)) ORIG.set(n, cur); n.nodeValue = t; }
  } else {
    if (ORIG.has(n)) { n.nodeValue = ORIG.get(n); ORIG.delete(n); }
    else { var r = revText(cur); if (r !== null && r !== cur) n.nodeValue = r; }
  }
}
function transAttrs(el) {
  ['placeholder', 'title'].forEach(function(a) {
    if (!el.hasAttribute || !el.hasAttribute(a)) return;
    var v = el.getAttribute(a), k = 'data-o-' + a;
    if (LANG === 'hi') {
      if (Object.prototype.hasOwnProperty.call(HI, v)) { el.setAttribute(k, v); el.setAttribute(a, HI[v]); }
    } else if (el.hasAttribute(k)) { el.setAttribute(a, el.getAttribute(k)); el.removeAttribute(k); }
  });
}
function transTree(root) {
  if (!root) return;
  if (root.nodeType === 3) return transNode(root);
  if (root.nodeType !== 1) return;
  var w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null), list = [], n;
  while ((n = w.nextNode())) list.push(n);
  list.forEach(transNode);
  transAttrs(root);
  root.querySelectorAll('[placeholder],[title],[data-o-placeholder],[data-o-title]').forEach(transAttrs);
}
function obsOn() { if (OBS) OBS.observe(document.body, { childList: true, subtree: true, characterData: true }); }
function startLangObserver() {
  if (OBS || !window.MutationObserver) return;
  OBS = new MutationObserver(function(recs) {
    if (LANG !== 'hi') return;
    OBS.disconnect();
    recs.forEach(function(r) {
      if (r.type === 'characterData') transNode(r.target);
      else r.addedNodes.forEach(function(x) { transTree(x); });
    });
    obsOn();
  });
  obsOn();
}
function applyLang() {
  if (OBS) OBS.disconnect();
  transTree(document.body);
  obsOn();
}
function updateToolIcons() {
  document.querySelectorAll('.js-lang').forEach(function(b) { b.textContent = LANG === 'hi' ? 'EN' : 'हिं'; });
  document.querySelectorAll('.js-theme').forEach(function(b) { b.innerHTML = THEME === 'dark' ? '<i class="fa-solid fa-sun"></i>' : '<i class="fa-solid fa-moon"></i>'; });
}
function setTheme(t) { THEME = t; document.body.classList.toggle('dark', t === 'dark'); lsSet('bht_theme', t); updateToolIcons(); }
function toggleTheme() { setTheme(THEME === 'dark' ? 'light' : 'dark'); }
function setLang(l) {
  LANG = l; lsSet('bht_lang', l);
  document.documentElement.lang = l === 'hi' ? 'hi' : 'en';
  applyLang(); updateToolIcons();
}
function toggleLang() { setLang(LANG === 'hi' ? 'en' : 'hi'); }
function initPrefs() {
  var t = lsGet('bht_theme');
  THEME = t === 'dark' ? 'dark' : 'light';
  document.body.classList.toggle('dark', THEME === 'dark');
  LANG = lsGet('bht_lang') === 'hi' ? 'hi' : 'en';
  updateToolIcons();
  startLangObserver();
  if (LANG === 'hi') { document.documentElement.lang = 'hi'; applyLang(); }
}

function lockBtn(btn) {
  if (!btn || btn.dataset.locked === '1') return false;
  btn.dataset.locked = '1'; btn.dataset.orig = btn.innerHTML;
  btn.disabled = true; btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Please wait...';
  return true;
}
function unlockBtn(btn) {
  if (!btn) return;
  btn.dataset.locked = ''; btn.disabled = false;
  if (btn.dataset.orig) btn.innerHTML = btn.dataset.orig;
}

function toast(msg, type) {
  var el = document.createElement('div');
  el.style.cssText = 'position:fixed; top:20px; left:50%; transform:translateX(-50%); z-index:99999; padding:13px 24px; border-radius:30px; color:#fff; font-weight:700; font-size:14px; box-shadow:0 8px 24px rgba(30,27,75,0.18); display:flex; align-items:center; gap:8px; max-width:92vw; font-family:Manrope,sans-serif;';
  el.style.background = type === 'error' ? 'linear-gradient(135deg,#ef4444,#dc2626)' : (type === 'success' ? 'linear-gradient(135deg,#10b981,#059669)' : 'linear-gradient(135deg,#8b5cf6,#7c3aed)');
  var icon = type === 'error' ? 'fa-circle-exclamation' : (type === 'success' ? 'fa-circle-check' : 'fa-circle-info');
  el.innerHTML = '<i class="fa-solid ' + icon + '"></i> ' + esc(msg);
  document.body.appendChild(el);
  setTimeout(function() { el.style.opacity = '0'; el.style.transition = 'opacity 0.3s'; setTimeout(function(){ el.remove(); }, 300); }, 4500);
}

function callServer(fnName, args, onSuccess, onError, silent) {
  args = Array.isArray(args) ? args : [];
  var done = false;
  var timer = silent ? null : setTimeout(function() {
    if (done) return; done = true; showLoader(false);
    toast('Taking longer than usual. If you already submitted, use Track Status before retrying.', 'info');
    if (onError) onError({ message: 'timeout', soft: true });
  }, 60000);

  if (typeof google === 'undefined' || !google.script || !google.script.run) {
    if (timer) clearTimeout(timer);
    if (!silent) { showLoader(false); toast('Google Script environment not ready.', 'error'); }
    if (onError) onError({ message: 'no-env' });
    return;
  }

  var runner = google.script.run
    .withSuccessHandler(function(res) { if (done) return; done = true; if (timer) clearTimeout(timer); if (onSuccess) onSuccess(res); })
    .withFailureHandler(function(err) {
      if (done) return; done = true; if (timer) clearTimeout(timer);
      if (!silent) { showLoader(false); toast(err.message || 'Operation failed', 'error'); }
      if (onError) onError(err);
    });

  try {
    var n = args.length;
    if      (n === 0) runner[fnName]();
    else if (n === 1) runner[fnName](args[0]);
    else if (n === 2) runner[fnName](args[0], args[1]);
    else if (n === 3) runner[fnName](args[0], args[1], args[2]);
    else if (n === 4) runner[fnName](args[0], args[1], args[2], args[3]);
    else              runner[fnName](args[0], args[1], args[2], args[3], args[4]);
  } catch (callErr) {
    done = true; if (timer) clearTimeout(timer);
    if (!silent) { showLoader(false); toast('Call error: ' + callErr.message, 'error'); }
    if (onError) onError(callErr);
  }
}

function bg(fnName, args) { callServer(fnName, args, function(){}, function(){}, true); }

function esc(v) {
  return String(v == null ? '' : v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

function toggleSidebar() { $('sidebar').classList.toggle('open'); }

/* ---------- Filters ---------- */
function F(k) { return function(r) { return r[k]; }; }
function txt(v) { return String(v == null ? '' : v).toLowerCase(); }
function matchSearch(obj, q, fields) {
  q = txt(q).trim();
  if (!q) return true;
  var hay = fields.map(function(f) { return txt(obj[f]); }).join(' ');
  return q.split(/\s+/).every(function(p) { return hay.indexOf(p) > -1; });
}
function inRange(dateStr, from, to) {
  var d = String(dateStr || '').substring(0, 10);
  if (!from && !to) return true;
  if (!d) return false;
  if (from && d < from) return false;
  if (to && d > to) return false;
  return true;
}
function uniq(list, f) {
  var m = {}, o = [];
  (list || []).forEach(function(r) {
    var v = typeof f === 'function' ? f(r) : r[f];
    v = String(v == null ? '' : v).trim();
    if (v && !m[v]) { m[v] = 1; o.push(v); }
  });
  return o.sort();
}
function fv(key, name) { var e = $(key + '_' + name); return e ? e.value : ''; }
function fillDyn(key, name, vals) {
  var s = $(key + '_' + name);
  if (!s) return;
  var cur = s.value, lbl = s.getAttribute('data-lbl') || '';
  s.innerHTML = '<option value="">All ' + esc(lbl) + '</option>' + vals.map(function(v) { return '<option value="' + esc(v) + '">' + esc(v) + '</option>'; }).join('');
  if (cur && vals.indexOf(cur) > -1) s.value = cur;
}
function setCount(key, n, total, extra) {
  var e = $(key + '_cnt');
  if (e) e.textContent = n + ' of ' + total + ' records' + (extra ? ' • ' + extra : '');
}
function money(n) { return '₹' + Number(n || 0).toLocaleString('en-IN'); }

var TOOLBARS = {
  cmp: { render: 'renderComplaintsList', ph: 'Search ticket, company...', exp: 'complaints', dates: true,
         selects: [
           ['eng','Engineer',null],
           ['status','Status',['NEW','ASSIGNED','ACCEPTED','ON_SITE','PENDING','COMPLETED','CLOSED']],
           ['warranty','Warranty',['UNDER_WARRANTY','OUT_OF_WARRANTY','UNDER_AMC','NEW_EQUIPMENT_INSTALLED','NEW_ORDER_DIMENSION_SITE_VISIT']]
         ] },
  job: { render: 'renderJobCards', ph: 'Search jobs...', exp: 'jobs', dates: false, selects: [] },
  rep: { render: 'renderReportsList', ph: 'Search reports...', exp: 'reports', dates: true,
         selects: [['eng','Engineer',null], ['status','Status',['COMPLETED','PENDING']]] },
  bil: { render: 'renderBillingList', ph: 'Search bills...', exp: 'billing', dates: true, selects: [] },
  inv: { render: 'renderInvoicesList', ph: 'Search invoice, ticket...', exp: 'invoices', dates: true,
         selects: [['cust','Customer',null]] },
  ntf: { render: 'renderNotifPage', ph: 'Search notifications...', exp: 'notifications', dates: true,
         selects: [['title','Type',null]] },
  his: { render: 'renderHistoryList', ph: 'Search service history...', exp: '', dates: true,
         selects: [['status','Status',['COMPLETED','PENDING']]] }
};

function buildToolbar(key) {
  var t = TOOLBARS[key], el = $('tb_' + key);
  if (!el) return;
  var h = '<div class="search-box"><i class="fa-solid fa-magnifying-glass"></i><input id="' + key + '_q" data-lbl="Search" placeholder="' + t.ph + '" oninput="' + t.render + '()"></div>';
  if (t.dates) {
    h += '<div class="date-range-inline">' +
         '<span><i class="fa-regular fa-calendar"></i> From</span>' +
         '<input type="date" id="' + key + '_from" data-lbl="From Date" title="From date" onchange="' + t.render + '()">' +
         '<span>To</span>' +
         '<input type="date" id="' + key + '_to" data-lbl="To Date" title="To date" onchange="' + t.render + '()">' +
         '</div>';
  }
  var hasSelects = (t.selects && t.selects.length);
  if (hasSelects) {
    h += '<div class="filter-group" id="' + key + '_filters">';
    (t.selects || []).forEach(function(s) {
      h += '<select id="' + key + '_' + s[0] + '" data-lbl="' + s[1] + '"' + (s[2] ? '' : ' data-dyn="1"') + ' onchange="' + t.render + '()"><option value="">All ' + s[1] + '</option>' +
           (s[2] ? s[2].map(function(o) { return '<option value="' + o + '">' + formatWarrantyLabel(o) + '</option>'; }).join('') : '') + '</select>';
    });
    h += '</div>';
  }
  h += '<button class="btn btn-secondary btn-sm" data-k="' + key + '" onclick="resetToolbar(this.dataset.k)" title="Reset filters"><i class="fa-solid fa-rotate-left"></i></button>';
  h += '<span class="tb-count" id="' + key + '_cnt"></span>';
  if (t.exp) {
    h += '<span class="export-bar">' +
      '<button class="btn btn-pdf btn-sm" data-k="' + t.exp + '" data-m="pdf" onclick="exportView(this.dataset.k,this.dataset.m)"><i class="fa-solid fa-file-pdf"></i> PDF</button>' +
      '<button class="btn btn-print btn-sm" data-k="' + t.exp + '" data-m="print" onclick="exportView(this.dataset.k,this.dataset.m)"><i class="fa-solid fa-print"></i> Print</button>' +
      '<button class="btn btn-csv btn-sm" data-k="' + t.exp + '" data-m="csv" onclick="exportView(this.dataset.k,this.dataset.m)"><i class="fa-solid fa-file-excel"></i> Excel</button>' +
      '</span>';
  }
  el.innerHTML = h;
}
function clearToolbarValues(key) {
  var el = $('tb_' + key);
  if (!el) return;
  el.querySelectorAll('input').forEach(function(x) { if (x.type !== 'button') x.value = ''; });
  el.querySelectorAll('select').forEach(function(s) {
    if (s.getAttribute('data-dyn') === '1') s.innerHTML = '<option value="">All ' + esc(s.getAttribute('data-lbl') || '') + '</option>';
    s.value = '';
  });
}
function resetToolbar(key) {
  clearToolbarValues(key);
  var fn = window[TOOLBARS[key].render];
  if (typeof fn === 'function') fn();
}
function buildAllToolbars() { Object.keys(TOOLBARS).forEach(buildToolbar); }

/* ---------- Export ---------- */
function setView(key, title, cols, rows, note) { VIEW[key] = { title: title, cols: cols, rows: rows, note: note || '' }; }
function collectNote(containerId) {
  var el = $(containerId);
  if (!el) return '';
  var out = [];
  el.querySelectorAll('input,select').forEach(function(x) {
    if (!x.value) return;
    if (x.closest('.hidden')) return;
    var l = x.getAttribute('data-lbl') || x.id, v = x.value;
    if (x.tagName === 'SELECT') v = x.options[x.selectedIndex].text;
    out.push(l + ': ' + v);
  });
  return out.join(' | ');
}

var COLS = {
  complaints: [
    ['Ticket',F('ticketNumber')],
    ['Date',function(c){return fmtDate(c.createdAt);}],
    ['Company',F('companyName')],
    ['Person',F('authorizedPerson')],
    ['Phone',F('phone')],
    ['Location',F('location')],
    ['Equipment',F('equipment')],
    ['Warranty',function(c){return formatWarrantyLabel(c.warrantyStatus);}],
    ['Status',F('status')],
    ['Engineer',function(c){return c.assignedEngineerName||'';}],
    ['Visits',function(c){return c._visits||0;}]
  ],
  jobs: [
    ['Ticket',F('ticketNumber')],
    ['Date',function(c){return fmtDate(c.createdAt);}],
    ['Company',F('companyName')],
    ['Person',F('authorizedPerson')],
    ['Phone',F('phone')],
    ['Equipment',F('equipment')],
    ['Status',F('status')],
    ['Address',function(c){return (c.address||'')+', '+(c.location||'');}]
  ],
  reports: [
    ['Ticket',F('ticketNumber')],
    ['Date',function(r){return fmtDate(r.submittedAt);}],
    ['Customer',F('customerName')],
    ['Location',F('_location')],
    ['Engineer',F('engineerName')],
    ['Status',F('status')],
    ['Service (₹)',function(r){return Number(r.serviceCharges)||0;}],
    ['Food (₹)',function(r){return Number(r.foodCharges)||0;}],
    ['Other (₹)',function(r){return Number(r.otherCharges)||0;}],
    ['Total (₹)',function(r){return (Number(r.serviceCharges)||0)+(Number(r.foodCharges)||0)+(Number(r.otherCharges)||0);}],
    ['KM',function(r){return Number(r.kmDistance)||0;}],
    ['Remarks',function(r){return r.remarks||r.pendingReason||'';}]
  ],
  billing: [
    ['Ticket',F('ticketNumber')],
    ['Date',function(c){return fmtDate(c.createdAt);}],
    ['Company',F('companyName')],
    ['Equipment',F('equipment')],
    ['Warranty',function(c){return formatWarrantyLabel(c.warrantyStatus);}],
    ['Address',function(c){return (c.address||'')+', '+(c.location||'');}]
  ],
  invoices: [
    ['Ticket',F('ticketId')],
    ['Date',function(i){return fmtDate(i.createdAt);}],
    ['Invoice No',F('invoiceNumber')],
    ['Customer',F('customerName')],
    ['Amount (₹)',F('amount')]
  ],
  notifications: [
    ['Title',F('title')],
    ['Message',F('message')],
    ['Date',function(n){return fmtDate(n.createdAt);}]
  ],
  md: [
    ['Ticket',F('ticketNumber')],
    ['Date',function(r){return fmtDate(r.createdAt);}],
    ['Company',F('companyName')],
    ['Site Person',F('authorizedPerson')],
    ['Phone',F('phone')],
    ['Email',F('email')],
    ['Location',F('location')],
    ['Address',F('address')],
    ['Equipment',F('equipment')],
    ['Serial No',F('serialNo')],
    ['Problem',F('problemDescription')],
    ['Warranty',function(r){return formatWarrantyLabel(r.warrantyStatus);}],
    ['Material Req',F('materialRequired')],
    ['Material Details',F('materialDetails')],
    ['Engineer',function(r){return r.assignedEngineerName||'';}],
    ['Assigned On',function(r){return fmtDate(r.assignedAt);}],
    ['Visits',F('_visits')],
    ['Last Visit',function(r){return fmtDate(r._lastVisit);}],
    ['Status',F('status')],
    ['Service Charge (₹)',F('_svcCharge')],
    ['Food Charge (₹)',F('_foodCharge')],
    ['Other Charge (₹)',F('_otherCharge')],
    ['Total Amount (₹)',F('_service')],
    ['Total KM',F('_kmTotal')],
    ['Bill Status',F('_billText')],
    ['Invoice No',F('_invNo')],
    ['Invoice (₹)',F('_invAmt')],
    ['Remarks',F('remarks')]
  ]
};

function exportView(key, mode) {
  if (App.user && String(App.user.role).toUpperCase() === 'ENGINEER') {
    return toast('Engineers cannot export data', 'error');
  }
  var v = VIEW[key];
  if (!v || !v.rows || !v.rows.length) return toast('No data to export', 'error');

  var head = v.cols.map(function(c) { return c[0]; });
  var rowsToExport = v.rows.slice();

  var body = rowsToExport.map(function(r) { return v.cols.map(function(c) { var x = c[1](r); return x == null ? '' : String(x); }); });
  var meta = 'Generated: ' + fmtDate(new Date()) + ' | By: ' + App.user.name + ' | Records: ' + body.length + (v.note ? ' | ' + v.note : '');
  if (mode === 'print') return printTable(v.title, head, body, meta);
  if (mode === 'csv') return downloadCsv(v.title, head, body);
  makePdf(v.title, head, body, meta);
}

function printTable(title, head, body, meta) {
  var w = window.open('', '_blank');
  if (!w) return toast('Popup blocked – allow popups for this site and try again', 'error');
  var h = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + esc(title) + '</title><style>' +
    'body{font-family:Arial,sans-serif;padding:20px;color:#111}h2{margin:0 0 4px;color:#4c1d95}p{margin:0 0 14px;font-size:11px;color:#555}' +
    'table{width:100%;border-collapse:collapse;font-size:9.5px}th,td{border:1px solid #ccc;padding:4px 6px;text-align:left;vertical-align:top}' +
    'th{background:#ede9fe;color:#4c1d95}tr:nth-child(even) td{background:#faf8ff}@media print{@page{size:landscape;margin:8mm}}' +
    '</style></head><body><h2>BHT – ' + esc(title) + '</h2><p>' + esc(meta) + '</p><table><thead><tr>' +
    head.map(function(x) { return '<th>' + esc(x) + '</th>'; }).join('') + '</tr></thead><tbody>' +
    body.map(function(r) { return '<tr>' + r.map(function(x) { return '<td>' + esc(x) + '</td>'; }).join('') + '</tr>'; }).join('') +
    '</tbody></table></body></html>';
  w.document.open(); w.document.write(h); w.document.close();
  setTimeout(function() { try { w.focus(); w.print(); } catch (e) {} }, 500);
}

function downloadCsv(title, head, body) {
  var cell = function(s) { return '"' + String(s).replace(/"/g, '""') + '"'; };
  var csv = '\ufeff' + [head].concat(body).map(function(r) { return r.map(cell).join(','); }).join('\r\n');
  try {
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = fileStamp(title) + '.csv';
    document.body.appendChild(a); a.click();
    setTimeout(function() { try { URL.revokeObjectURL(a.href); a.remove(); } catch (e) {} }, 500);
    toast('Excel/CSV downloaded', 'success');
  } catch (e) { toast('Download failed: ' + e.message, 'error'); }
}
function fileStamp(title) {
  var d = new Date();
  var p = function(n) { return (n < 10 ? '0' : '') + n; };
  return 'BHT_' + String(title).replace(/[^a-z0-9]+/gi, '_') + '_' + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate());
}

function loadScript(src, cb) {
  var s = document.createElement('script');
  s.src = src; s.onload = function() { cb(null); };
  s.onerror = function() { cb(new Error('Could not load ' + src)); };
  document.head.appendChild(s);
}
function ensurePdfLib(cb) {
  if (window.jspdf && window.jspdf.jsPDF && window.__autoTableReady) return cb(null);
  var base = 'https://cdnjs.cloudflare.com/ajax/libs/';
  var step2 = function() {
    loadScript(base + 'jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js', function(err) {
      if (err) return cb(err);
      window.__autoTableReady = true; cb(null);
    });
  };
  if (window.jspdf && window.jspdf.jsPDF) return step2();
  loadScript(base + 'jspdf/2.5.1/jspdf.umd.min.js', function(err) { if (err) return cb(err); step2(); });
}
function pdfSafe(s) { return String(s == null ? '' : s).replace(/₹/g, 'Rs.').replace(/[^\x00-\xFF]/g, ''); }
function makePdf(title, head, body, meta) {
  showLoader(true);
  ensurePdfLib(function(err) {
    showLoader(false);
    if (err) { toast('PDF library failed to load. Use Print → Save as PDF instead.', 'error'); return; }
    try {
      var jsPDF = window.jspdf.jsPDF;
      var doc = new jsPDF({ orientation: head.length > 6 ? 'landscape' : 'portrait', unit: 'mm', format: 'a4' });
      doc.setFontSize(14); doc.setTextColor(76, 29, 149);
      doc.text(pdfSafe('BHT - ' + title), 10, 12);
      doc.setFontSize(8); doc.setTextColor(90);
      var lines = doc.splitTextToSize(pdfSafe(meta), doc.internal.pageSize.getWidth() - 20);
      doc.text(lines, 10, 18);
      doc.autoTable({
        head: [head.map(pdfSafe)],
        body: body.map(function(r) { return r.map(pdfSafe); }),
        startY: 20 + lines.length * 3.5,
        styles: { fontSize: 5.8, cellPadding: 1.2, overflow: 'linebreak' },
        headStyles: { fillColor: [139, 92, 246], textColor: 255 },
        alternateRowStyles: { fillColor: [250, 248, 255] },
        margin: { left: 6, right: 6 }
      });
      var name = fileStamp(title) + '.pdf';
      try { doc.save(name); toast('PDF downloaded', 'success'); }
      catch (e2) { window.open(doc.output('bloburl'), '_blank'); }
    } catch (e) { toast('PDF error: ' + e.message, 'error'); }
  });
}

/* ---------- Top bar profile/notifications ---------- */
function initials(n) {
  var p = String(n || '').trim().split(/\s+/).filter(Boolean);
  if (!p.length) return 'U';
  return (p[0].charAt(0) + (p.length > 1 ? p[p.length - 1].charAt(0) : '')).toUpperCase();
}
function setHeaderUser() {
  var u = App.user; if (!u) return;
  $('hdrAvatar').textContent = initials(u.name);
  $('hdrName').textContent = u.name || 'User';
  $('hdrRole').textContent = u.role || '';
  $('ppAvatar').textContent = initials(u.name);
  $('ppName').textContent = u.name || '';
  $('ppEmail').textContent = u.email || '';
  $('ppRole').textContent = u.role || '';
}
function closePanels() {
  var a = $('notifPanel'), b = $('profilePanel');
  if (a) a.classList.add('hidden');
  if (b) b.classList.add('hidden');
}
function toggleProfilePanel(ev) {
  ev.stopPropagation();
  var p = $('profilePanel'), wasHidden = p.classList.contains('hidden');
  closePanels();
  if (wasHidden) p.classList.remove('hidden');
}
function toggleNotifPanel(ev) {
  ev.stopPropagation();
  var p = $('notifPanel'), wasHidden = p.classList.contains('hidden');
  closePanels();
  if (wasHidden) { p.classList.remove('hidden'); renderBell(); fetchNotifs(); }
}
document.addEventListener('click', function(e) { if (!e.target.closest('.hdr-item')) closePanels(); });

function readKey() { return 'bht_read_' + (App.user ? App.user.id : 'x'); }
function readIds() {
  var s = null;
  try { s = localStorage.getItem(readKey()); } catch (e) {}
  if (s == null) s = App.readMem || '[]';
  try { return JSON.parse(s) || []; } catch (e) { return []; }
}
function saveIds(a) {
  a = a.slice(-400);
  var s = JSON.stringify(a);
  App.readMem = s;
  try { localStorage.setItem(readKey(), s); } catch (e) {}
}
function isUnread(n, ids) {
  if (String(n.read).toUpperCase() === 'TRUE') return false;
  return ids.indexOf(String(n.id)) === -1;
}
function markRead(id) {
  var ids = readIds();
  if (ids.indexOf(String(id)) === -1) { ids.push(String(id)); saveIds(ids); }
  for (var i = 0; i < NOTIFS.length; i++) {
    if (String(NOTIFS[i].id) === String(id)) { NOTIFS[i].read = 'TRUE'; break; }
  }
  NOTIF_PREV = unreadCount();
  renderBell();
  if (App.currentPage === 'notifications') renderNotifPage();
  bg('markNotificationReadOnServer', [id]);
}
function markAllRead() {
  var ids = readIds();
  NOTIFS.forEach(function(n) {
    if (ids.indexOf(String(n.id)) === -1) ids.push(String(n.id));
    n.read = 'TRUE';
  });
  saveIds(ids);
  NOTIF_PREV = 0;
  renderBell();
  if (App.currentPage === 'notifications') renderNotifPage();
  bg('markNotificationReadOnServer', ['']);
}
function unreadCount() {
  var ids = readIds();
  return NOTIFS.filter(function(n) { return isUnread(n, ids); }).length;
}
function renderBell() {
  var cnt = unreadCount(), b = $('notifBadge');
  if (b) { b.textContent = cnt > 99 ? '99+' : cnt; b.classList.toggle('hidden', cnt === 0); }
  var body = $('notifPanelBody');
  if (!body) return;
  var ids = readIds();
  body.innerHTML = NOTIFS.length ? NOTIFS.slice(0, 8).map(function(n) {
    return '<div class="nitem' + (isUnread(n, ids) ? ' unread' : '') + '" data-id="' + esc(n.id) + '" data-t="' + esc(n.ticketId) + '" onclick="openNotif(this.dataset.id,this.dataset.t)">' +
      '<div class="ntitle">' + esc(n.title) + '</div><div class="nmsg">' + esc(n.message) + '</div><div class="ntime">' + esc(fmtDate(n.createdAt)) + '</div></div>';
  }).join('') : '<div class="nempty"><i class="fa-regular fa-bell-slash" style="font-size:24px; display:block; margin-bottom:8px;"></i>No notifications yet</div>';
}
function openNotif(id, ticketId) {
  markRead(id);
  closePanels();
  var role = App.user ? App.user.role : '';
  if (ticketId && (role === 'COORDINATOR' || role === 'ADMIN')) goTo('complaintDetail', { id: ticketId });
  else if (ticketId && role === 'ENGINEER') goTo('jobDetail', { id: ticketId });
  else if (role === 'ACCOUNTS') goTo('billing');
  else goTo('notifications');
}
function openNotifPage() { closePanels(); goTo('notifications'); }

function notifArgs() { return [App.user.role, App.user.engineerId || '']; }

function fetchNotifs() {
  if (!App.user) return;
  callServer('getNotifications', notifArgs(), function(res) {
    if (!res || !res.success) return;
    var firstLoad = (NOTIF_PREV === -1);
    var oldIds = {};
    NOTIFS.forEach(function(n) { oldIds[String(n.id)] = 1; });
    NOTIFS = res.data || [];
    App.freshLogin = false;

    if (!firstLoad) {
      var readList = readIds();
      var fresh = NOTIFS.filter(function(n) {
        return !oldIds[String(n.id)] && readList.indexOf(String(n.id)) === -1;
      });
      if (fresh.length) {
        playNotifSound();
        showNotifPopup(fresh[0]);
        showBrowserNotif(fresh[0]);
      }
    }
    NOTIF_PREV = unreadCount();
    renderBell();
    if (App.currentPage === 'notifications') renderNotifPage();
  }, null, true);
}
function startNotifPolling() {
  stopNotifPolling();
  NOTIF_PREV = -1;
  fetchNotifs();
  NOTIF_TIMER = setInterval(fetchNotifs, 60000);
}
function stopNotifPolling() { if (NOTIF_TIMER) { clearInterval(NOTIF_TIMER); NOTIF_TIMER = null; } }

document.addEventListener('visibilitychange', function() {
  if (!document.hidden && App.user) fetchNotifs();
});

function onShow_notifications() {
  if (NOTIFS.length) { fillDyn('ntf', 'title', uniq(NOTIFS, 'title')); renderNotifPage(); }
  else showLoader(true);
  callServer('getNotifications', notifArgs(), function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    NOTIFS = res.data || [];
    fillDyn('ntf', 'title', uniq(NOTIFS, 'title'));
    renderBell();
    if (App.currentPage === 'notifications') renderNotifPage();
  });
}
function renderNotifPage() {
  var q = fv('ntf', 'q'), ti = fv('ntf', 'title'), df = fv('ntf', 'from'), dt = fv('ntf', 'to');
  var ids = readIds();
  var list = NOTIFS.filter(function(n) {
    if (ti && n.title !== ti) return false;
    if (!inRange(n.createdAt, df, dt)) return false;
    return matchSearch(n, q, ['title', 'message', 'createdAt']);
  });
  setCount('ntf', list.length, NOTIFS.length);
  setView('notifications', 'Notifications', COLS.notifications, list, collectNote('tb_ntf'));
  $('notifBody').innerHTML = list.length ? list.map(function(n) {
    var un = isUnread(n, ids);
    return '<tr class="' + (un ? 'unread-row' : '') + '">' +
      '<td data-label="Title"><b>' + (un ? '<i class="fa-solid fa-circle" style="color:var(--accent); font-size:8px; margin-right:6px;"></i>' : '') + esc(n.title) + '</b></td>' +
      '<td data-label="Message">' + esc(n.message) + '</td>' +
      '<td data-label="Date" style="color:var(--text-muted); font-size:12px;">' + esc(fmtDate(n.createdAt)) + '</td></tr>';
  }).join('') : '<tr class="empty-row"><td colspan="3">No notifications found</td></tr>';
}

/* ---------- Local cache helpers ---------- */
function findCachedComplaint(id) {
  var l = App.cache.complaints; if (!l) return null;
  for (var i = 0; i < l.length; i++) if (String(l[i].id) === String(id)) return l[i];
  return null;
}
function findCachedEngineer(id) {
  var l = App.cache.engineers; if (!l || !id) return null;
  for (var i = 0; i < l.length; i++) if (String(l[i].id).trim() === String(id).trim()) return l[i];
  return null;
}
function patchCached(id, fields) {
  var c = findCachedComplaint(id);
  if (c) for (var k in fields) c[k] = fields[k];
}

function normalizePhone(p) {
  var d = String(p || '').replace(/\D/g, '');
  if (d.length === 10) d = '91' + d;
  return d.length > 9 ? d : '';
}
function openWhatsApp(phone, msg) {
  var p = normalizePhone(phone);
  if (!p) return toast('Invalid phone number for WhatsApp', 'error');
  window.open('https://wa.me/' + p + '?text=' + encodeURIComponent(msg), '_blank');
}

function waNotifyAssignCustomer(phone, person, ticket, engName, engPhone) {
  openWhatsApp(phone, 'Dear ' + (person || 'Customer') + ',\n\nYour service ticket ' + ticket + ' has been assigned to ' + (engName || 'our Engineer') + '.\n\nRegards,\nBHT Service');
}
function waNotifyAssignEngineer(engPhone, engName, ticket, company, person, clientPhone, problem, extras) {
  extras = extras || {};
  var msg = 'Hi ' + (engName || 'Engineer') + ',\n\n' +
    '📋 NEW JOB ASSIGNED\n' +
    '════════════════════════\n' +
    'Ticket: ' + ticket + '\n' +
    'Company: ' + company + '\n' +
    'Contact: ' + person + ' (' + clientPhone + ')\n' +
    'Email: ' + (extras.email || '-') + '\n' +
    'Address: ' + (extras.address || '') + ', ' + (extras.location || '') + '\n' +
    'Equipment: ' + (extras.equipment || '') + '\n' +
    'Serial No: ' + (extras.serialNo || '-') + '\n' +
    'Problem: ' + problem + '\n' +
    '════════════════════════\n' +
    'Warranty: ' + formatWarrantyLabel(extras.warrantyStatus) + '\n' +
    'Material Required: ' + (extras.materialRequired || 'NO') + '\n' +
    'Material Details: ' + (extras.materialDetails || '-') + '\n' +
    'Remarks: ' + (extras.remarks || '-') + '\n\n' +
    'Regards,\nBHT Operations';
  openWhatsApp(engPhone, msg);
}

var WA_CTX = null;
function waDigits(p) { var d = String(p || '').replace(/\D/g, ''); if (d.length > 10) d = d.slice(-10); return d; }
function waFeedbackLink(c) {
  var base = (typeof WEB_URL !== 'undefined' && WEB_URL) ? WEB_URL : '';
  if (!base) return '';
  return base + '?ticket=' + encodeURIComponent(c.ticketNumber) + '&phone=' + encodeURIComponent(waDigits(c.phone)) + '&feedback=true';
}
function waAssigned(c, e) {
  var en = c.assignedEngineerName || (e && e.name) || 'our Engineer';
  openWhatsApp(c.phone, 'Dear ' + (c.authorizedPerson || 'Customer') + ',\n\n' +
    '👷 *ENGINEER ASSIGNED*\n━━━━━━━━━━━━━━━━━━\n' +
    'Ticket: ' + c.ticketNumber + '\n' +
    'Equipment: ' + (c.equipment || '-') + '\n' +
    'Engineer: ' + en + '\n\n' +
    'Our engineer will visit your site soon.\n\nRegards,\nBHT Service');
}
function waWork(c, p) {
  var done = String(p.status).toUpperCase() === 'COMPLETED';
  var eng = p.engineerName || c.assignedEngineerName || 'Our Engineer';
  var dt = fmtDate(p.submittedAt || new Date());
  var out = String(c.warrantyStatus || '').toUpperCase() === 'OUT_OF_WARRANTY';
  var svc = Number(p.serviceCharges) || 0, food = Number(p.foodCharges) || 0, oth = Number(p.otherCharges) || 0;
  var m = 'Dear ' + (c.authorizedPerson || 'Customer') + ',\n\n';
  if (done) {
    m += '✅ *SERVICE WORK COMPLETED*\n━━━━━━━━━━━━━━━━━━\n' +
      'Ticket: ' + c.ticketNumber + '\nEquipment: ' + (c.equipment || '-') + '\nEngineer: ' + eng + '\nVisit Date: ' + dt + '\n';
    if (p.bookletNo) m += 'Service Booklet No: ' + p.bookletNo + '\n';
    if (p.remarks) m += 'Work Done: ' + p.remarks + '\n';
    if (out) {
      m += '\n💰 *Service Charges*\nService: ₹' + svc + '\nFood: ₹' + food + '\nOther: ₹' + oth + '\n*Total: ₹' + (svc + food + oth) + '*\n\nYour invoice will be shared separately.';
    } else {
      m += '\nNo invoice required (' + formatWarrantyLabel(c.warrantyStatus) + ').';
    }
  } else {
    m += '⏳ *SERVICE UPDATE - WORK PENDING*\n━━━━━━━━━━━━━━━━━━\n' +
      'Ticket: ' + c.ticketNumber + '\nEquipment: ' + (c.equipment || '-') + '\nEngineer: ' + eng + '\nVisit Date: ' + dt + '\n' +
      'Reason: ' + (p.pendingReason || '-') + '\nNext Visit: ' + fmtDate(p.nextVisitDate) + '\n\nThank you for your patience.';
  }
  m += '\n\nRegards,\nBHT Service';
  openWhatsApp(c.phone, m);
}
function waBill(c, inv) {
  openWhatsApp(c.phone, 'Dear ' + (c.authorizedPerson || 'Customer') + ',\n\n' +
    '🧾 *INVOICE DETAILS*\n━━━━━━━━━━━━━━━━━━\n' +
    'Ticket: ' + c.ticketNumber + '\n' +
    'Invoice No: ' + inv.invoiceNumber + '\n' +
    'Amount: ₹' + inv.amount + '\n' +
    'Date: ' + fmtDate(inv.createdAt) + '\n' +
    (inv.billCopyUrl ? 'Bill Copy: ' + inv.billCopyUrl + '\n' : '') +
    '\nKindly arrange the payment at the earliest.\n\nRegards,\nBHT Accounts');
}
function waFinal(c) {
  var link = waFeedbackLink(c);
  openWhatsApp(c.phone, 'Dear ' + (c.authorizedPerson || 'Customer') + ',\n\n' +
    '🎉 *TICKET CLOSED - THANK YOU*\n━━━━━━━━━━━━━━━━━━\n' +
    'Ticket: ' + c.ticketNumber + '\n' +
    'Equipment: ' + (c.equipment || '-') + '\n' +
    'Engineer: ' + (c.assignedEngineerName || '-') + '\n\n' +
    'Service work and billing are completed. We hope you are satisfied with our service.\n\n' +
    '⭐ *Please rate our service*\n' +
    (link ? link + '\n' : 'Reply with your rating (1 to 5 stars) and feedback.\n') +
    '\nThank you for choosing BHT!\n\nRegards,\nBHT Service Team');
}
function waCtxSet(c, e, reps, invs) {
  var r = (reps || []).slice().sort(function(a, b) { return String(a.submittedAt || '').localeCompare(String(b.submittedAt || '')); });
  WA_CTX = { c: c, e: e, reps: r, invs: invs || [] };
}
function waPanelHtml() {
  var x = WA_CTX; if (!x) return '';
  var c = x.c, st = String(c.status || '').toUpperCase();
  var items = [
    ['assigned', '1. Engineer Assigned', !!c.assignedEngineerName, 'Engineer not assigned yet'],
    ['work', '2. Work Update', x.reps.length > 0, 'No service report yet'],
    ['bill', '3. Bill Details', x.invs.length > 0, 'Bill not created yet'],
    ['final', '4. Final + Feedback', (st === 'COMPLETED' || st === 'CLOSED'), 'Available after work is completed']
  ];
  return '<div class="card" style="margin-top:24px;"><h3 class="card-title"><i class="fa-brands fa-whatsapp" style="color:#25D366;"></i> WhatsApp to Customer</h3>' +
    '<div style="display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:10px;">' +
    items.map(function(it) {
      return '<button class="btn btn-wa btn-sm" data-s="' + it[0] + '" ' + (it[2] ? '' : 'disabled title="' + esc(it[3]) + '"') + ' onclick="waStage(this.dataset.s)"><i class="fa-brands fa-whatsapp"></i> ' + it[1] + '</button>';
    }).join('') + '</div>' +
    '<p style="font-size:12px; color:var(--text-muted); margin-top:12px;">Each stage sends a different message so the customer is never confused. Step 4 includes the rating / feedback link.</p></div>';
}
function renderWaPanel(c, e, reps, invs) {
  var el = $('complaintWaPanel'); if (!el) return;
  var role = App.user ? App.user.role : '';
  if (role !== 'COORDINATOR' && role !== 'ADMIN') { el.innerHTML = ''; return; }
  waCtxSet(c, e, reps, invs);
  el.innerHTML = waPanelHtml();
}
function waStage(s) {
  var x = WA_CTX; if (!x) return;
  if (s === 'assigned') return waAssigned(x.c, x.e);
  if (s === 'work') { var r = x.reps[x.reps.length - 1]; if (!r) return toast('No service report yet', 'error'); return waWork(x.c, r); }
  if (s === 'bill') { if (!x.invs.length) return toast('Bill not created yet', 'error'); return waBill(x.c, x.invs[0]); }
  if (s === 'final') return waFinal(x.c);
}

function showAssignWaBox(c, e) {
  var box = document.createElement('div'); box.className = 'wa-float';
  var divTitle = document.createElement('div'); divTitle.className = 'wa-title';
  divTitle.innerHTML = '<i class="fa-solid fa-circle-check" style="color:var(--success);"></i> Assigned to ' + esc(c.assignedEngineerName || e.name);
  var divActions = document.createElement('div'); divActions.className = 'wa-actions';
  var btnCust = document.createElement('button'); btnCust.className = 'btn btn-wa btn-sm';
  btnCust.innerHTML = '<i class="fa-brands fa-whatsapp"></i> Notify Customer';
  btnCust.onclick = function() { waAssigned(c, e); };
  divActions.appendChild(btnCust);
  if (e && e.phone) {
    var btnEng = document.createElement('button'); btnEng.className = 'btn btn-wa btn-sm';
    btnEng.innerHTML = '<i class="fa-brands fa-whatsapp"></i> Notify Engineer';
    btnEng.onclick = function() {
      waNotifyAssignEngineer(e.phone, c.assignedEngineerName || e.name, c.ticketNumber, c.companyName, c.authorizedPerson, c.phone, c.problemDescription, {
        email: c.email, address: c.address, location: c.location, equipment: c.equipment,
        serialNo: c.serialNo, warrantyStatus: c.warrantyStatus,
        materialRequired: c.materialRequired, materialDetails: c.materialDetails, remarks: c.remarks
      });
    };
    divActions.appendChild(btnEng);
  }
  var btnClose = document.createElement('button'); btnClose.className = 'btn btn-secondary btn-sm';
  btnClose.innerHTML = '<i class="fa-solid fa-xmark"></i> Close';
  btnClose.onclick = function() { box.remove(); };
  divActions.appendChild(btnClose);
  box.appendChild(divTitle); box.appendChild(divActions);
  document.body.appendChild(box);
  setTimeout(function() { try { box.remove(); } catch(e){} }, 25000);
}

function showServiceReportWaBox(c, payload) {
  var box = document.createElement('div'); box.className = 'wa-float';
  var divTitle = document.createElement('div'); divTitle.className = 'wa-title';
  divTitle.innerHTML = '<i class="fa-solid fa-circle-check" style="color:var(--success);"></i> Report Saved Successfully';
  var divActions = document.createElement('div'); divActions.className = 'wa-actions';
  var btnCust = document.createElement('button'); btnCust.className = 'btn btn-wa btn-sm';
  btnCust.innerHTML = '<i class="fa-brands fa-whatsapp"></i> WhatsApp Customer';
  btnCust.onclick = function() { waWork(c, payload); };
  var btnClose = document.createElement('button'); btnClose.className = 'btn btn-secondary btn-sm';
  btnClose.innerHTML = '<i class="fa-solid fa-xmark"></i> Close';
  btnClose.onclick = function() { box.remove(); };
  divActions.appendChild(btnCust); divActions.appendChild(btnClose);
  box.appendChild(divTitle); box.appendChild(divActions);
  document.body.appendChild(box);
  setTimeout(function() { try { box.remove(); } catch(e){} }, 25000);
}

function showBillWaBox(c, inv) {
  var box = document.createElement('div'); box.className = 'wa-float';
  var divTitle = document.createElement('div'); divTitle.className = 'wa-title';
  divTitle.innerHTML = '<i class="fa-solid fa-circle-check" style="color:var(--success);"></i> Invoice Saved';
  var divActions = document.createElement('div'); divActions.className = 'wa-actions';
  var btnCust = document.createElement('button'); btnCust.className = 'btn btn-wa btn-sm';
  btnCust.innerHTML = '<i class="fa-brands fa-whatsapp"></i> Notify Customer';
  btnCust.onclick = function() { waBill(c, inv); };
  var btnClose = document.createElement('button'); btnClose.className = 'btn btn-secondary btn-sm';
  btnClose.innerHTML = '<i class="fa-solid fa-xmark"></i> Close';
  btnClose.onclick = function() { box.remove(); };
  divActions.appendChild(btnCust); divActions.appendChild(btnClose);
  box.appendChild(divTitle); box.appendChild(divActions);
  document.body.appendChild(box);
  setTimeout(function() { try { box.remove(); } catch(e){} }, 25000);
}

function renderPhotoLinks(str) {
  if (!str || str === '[]') return '<span style="color:#9ca3af">-</span>';
  var urls = String(str).split(/[\r\n]+/).filter(function(u){ return u.trim(); });
  if (!urls.length) return '<span style="color:#9ca3af">-</span>';
  return '<div class="photo-links" style="display:flex; gap:4px; flex-wrap:wrap; justify-content:flex-start;">' + urls.map(function(u) {
    return '<a href="' + esc(u) + '" target="_blank" title="View Photo" style="padding:0; background:none; border:none; width:auto; height:auto;">' +
           '<img src="' + esc(u) + '" style="width:32px; height:32px; object-fit:cover; border-radius:6px; border:1px solid var(--border); display:block;" onerror="this.style.display=&quot;none&quot;; this.nextElementSibling.style.display=&quot;inline-flex&quot;; this.nextElementSibling.style.width=&quot;32px&quot;; this.nextElementSibling.style.height=&quot;32px&quot;;">' +
           '<i class="fa-solid fa-image" style="display:none; align-items:center; justify-content:center; background:var(--primary-light); color:var(--primary); border-radius:6px; font-size:13px;"></i></a>';
  }).join('') + '</div>';
}
function renderAllReportPhotos(r) {
  var combined = [r.serviceImages || '', r.equipmentPhotos || ''].filter(Boolean).join('\n');
  return renderPhotoLinks(combined);
}
function renderDocLink(url) {
  return url ? '<div class="photo-links"><a href="' + esc(url) + '" target="_blank" title="Open Document"><i class="fa-solid fa-file-pdf"></i></a></div>' : '-';
}
function renderCharges(r) {
  var svc = Number(r.serviceCharges) || 0;
  var food = Number(r.foodCharges) || 0;
  var oth = Number(r.otherCharges) || 0;
  var km = Number(r.kmDistance) || 0;
  var tot = svc + food + oth;
  return '<b>Total: ₹' + tot + '</b><div style="font-size:10.5px; color:var(--text-muted); margin-top:2px; white-space:nowrap;">Svc: ₹' + svc + ' | Food: ₹' + food + ' | Oth: ₹' + oth + ' | KM: ' + km + ' km</div>';
}
function statusBadge(s) {
  s = String(s || 'NEW').toUpperCase();
  var map = { NEW:'badge-new', ASSIGNED:'badge-assigned', ACCEPTED:'badge-accepted', ON_SITE:'badge-onsite', PENDING:'badge-pending', COMPLETED:'badge-completed', CLOSED:'badge-closed', 'NOT REQ':'badge-notreq' };
  return '<span class="badge ' + (map[s] || 'badge-new') + '">' + esc(s) + '</span>';
}

function renderStatsInto(containerId, list) {
  var el = $(containerId);
  if (!el) return;
  var newC = 0, assignedC = 0, pendingC = 0, completedC = 0;
  for (var i = 0; i < list.length; i++) {
    if (!list[i] || (!list[i].ticketNumber && !list[i].companyName && !list[i].id)) continue;
    var s = list[i].status;
    if (s === 'NEW') newC++;
    else if (s === 'ASSIGNED' || s === 'ACCEPTED' || s === 'ON_SITE') assignedC++;
    else if (s === 'PENDING') pendingC++;
    else if (s === 'COMPLETED' || s === 'CLOSED') completedC++;
  }
  el.innerHTML =
    '<div class="stat-card"><div class="stat-icon bg-blue"><i class="fa-solid fa-plus"></i></div><div class="stat-info"><span class="lbl">New</span><span class="val">' + newC + '</span></div></div>' +
    '<div class="stat-card"><div class="stat-icon bg-orange"><i class="fa-solid fa-user-check"></i></div><div class="stat-info"><span class="lbl">Assigned</span><span class="val">' + assignedC + '</span></div></div>' +
    '<div class="stat-card"><div class="stat-icon bg-red"><i class="fa-solid fa-clock"></i></div><div class="stat-info"><span class="lbl">Pending</span><span class="val">' + pendingC + '</span></div></div>' +
    '<div class="stat-card"><div class="stat-icon bg-green"><i class="fa-solid fa-circle-check"></i></div><div class="stat-info"><span class="lbl">Completed</span><span class="val">' + completedC + '</span></div></div>';
}

function openFeedbackModal(ticketNumber, phone) {
  $('fb_ticket_display').textContent = ticketNumber;
  App.currentFeedbackTicket = ticketNumber; App.currentFeedbackPhone = phone || ''; App.currentFeedbackRating = 5;
  $('fb_comment').value = ''; $('fb_phone').value = phone || '';
  updateStars(5); $('feedbackModal').classList.remove('hidden');
}
function closeFeedbackModal() { $('feedbackModal').classList.add('hidden'); }
function updateStars(val) {
  document.querySelectorAll('#starContainer i').forEach(function(star) {
    var v = parseInt(star.getAttribute('data-val'), 10);
    if (v <= val) star.classList.add('active'); else star.classList.remove('active');
  });
}
function submitFeedbackForm(btn) {
  if (!App.currentFeedbackRating || App.currentFeedbackRating < 1) return toast('Please select a star rating', 'error');
  var phone = $('fb_phone').value.trim();
  if (!phone) return toast('Please enter your registered mobile number', 'error');
  if (!lockBtn(btn)) return;
  var comment = $('fb_comment').value.trim();
  showLoader(true);
  callServer('submitFeedback', [App.currentFeedbackTicket, App.currentFeedbackRating, comment, phone], function(res) {
    showLoader(false); unlockBtn(btn);
    if (!res.success) return toast(res.message, 'error');
    toast('Thank you for rating our service!', 'success');
    closeFeedbackModal();
    if (!App.user) {
      $('tr_ticket').value = App.currentFeedbackTicket;
      $('tr_phone').value = phone;
      pubShow('track');
      trackComplaint();
    }
  }, function() { showLoader(false); unlockBtn(btn); });
}

function pubShow(w) {
  $('pubTabForm').classList.toggle('active', w === 'form');
  $('pubTabTrack').classList.toggle('active', w === 'track');
  $('pubForm').classList.toggle('hidden', w !== 'form');
  $('pubTrack').classList.toggle('hidden', w !== 'track');
  $('pubThanks').classList.add('hidden');
}

function submitPublicComplaint() {
  var btn = $('btnSubmit');
  if (btn.dataset.locked === '1') return;
  var p = {
    companyName: $('cf_company').value.trim(), authorizedPerson: $('cf_person').value.trim(),
    email: $('cf_email').value.trim(), phone: $('cf_phone').value.trim(),
    address: $('cf_address').value.trim(), location: $('cf_location').value.trim(),
    equipment: $('cf_equipment').value, serialNo: $('cf_serial').value.trim(),
    problemDescription: $('cf_problem').value.trim()
  };
  if (!p.companyName || !p.authorizedPerson || !p.phone || !p.address || !p.equipment || !p.problemDescription)
    return toast('Please fill in all required fields marked with *', 'error');

  lockBtn(btn); showLoader(true);
  callServer('submitPublicComplaint', [p], function(res) {
    showLoader(false); unlockBtn(btn);
    if (!res.success) return toast(res.message || 'Submission failed', 'error');
    App.lastPublic = res.data.complaint;
    $('thanksTicket').textContent = res.data.ticketNumber;
    $('pubForm').classList.add('hidden'); $('pubTrack').classList.add('hidden'); $('pubThanks').classList.remove('hidden');
    ['cf_company','cf_person','cf_email','cf_phone','cf_address','cf_location','cf_serial','cf_problem'].forEach(function(id){ $(id).value = ''; });
    $('cf_equipment').value = '';
    if (res.data.duplicate) toast('This complaint was already registered — using the same ticket.', 'info');
    else bg('bgComplaintEmails', [res.data.id]);
  }, function() { showLoader(false); unlockBtn(btn); });
}

function trackComplaint() {
  var t = $('tr_ticket').value.trim(), p = $('tr_phone').value.trim();
  if (!t && !p) return toast('Enter ticket number or mobile number', 'error');
  showLoader(true);
  callServer('trackComplaint', [t, p], function(res) {
    showLoader(false);
    if (!res.success) {
      App.trackData = [];
      $('trackResult').innerHTML = '<div class="alert alert-error"><i class="fa-solid fa-circle-exclamation"></i> ' + esc(res.message) + '</div>';
      return;
    }
    var list = res.data.complaints;
    if (!list || !list.length) {
      App.trackData = [];
      $('trackResult').innerHTML = '<div class="alert alert-warning"><i class="fa-solid fa-triangle-exclamation"></i> No complaints found.</div>';
      return;
    }
    App.trackData = sortTicketsAsc(list);
    renderTrackList();
  }, function(err) { showLoader(false); });
}

function renderTrackList() {
  var list = App.trackData || [];
  if (!list.length) { $('trackResult').innerHTML = ''; return; }
  $('trackResult').innerHTML = '<div class="table-wrap table-wrap-cards"><table class="responsive-table"><thead><tr><th>Ticket</th><th>Date</th><th>Company</th><th>Equipment</th><th>Status</th><th>Assigned Engineer</th></tr></thead><tbody>' +
    list.map(function(d) {
      var engDisplay = d.engineer ? ('<b>' + esc(d.engineer) + '</b>') : '<span style="color:var(--text-muted);">Unassigned</span>';
      return '<tr>' +
        '<td data-label="Ticket"><b>' + esc(d.ticketNumber) + '</b></td>' +
        '<td data-label="Date">' + esc(fmtDate(d.createdAt)) + '</td>' +
        '<td data-label="Company">' + esc(d.companyName) + '</td>' +
        '<td data-label="Equipment">' + esc(d.equipment) + '</td>' +
        '<td data-label="Status">' + statusBadge(d.status) + '</td>' +
        '<td data-label="Assigned Engineer">' + engDisplay + '</td></tr>';
    }).join('') + '</tbody></table></div>';
}

function showStaffLogin() { $('publicView').classList.add('hidden'); $('appShell').classList.add('hidden'); $('staffLogin').classList.remove('hidden'); }
function backToPublic() { $('staffLogin').classList.add('hidden'); $('appShell').classList.add('hidden'); $('publicView').classList.remove('hidden'); }

function doLogin() {
  var e = $('loginEmail').value.trim(), p = $('loginPassword').value.trim();
  if (!e || !p) return toast('Please enter your email and password', 'error');
  showLoader(true);
  callServer('login', [e, p], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message || 'Login failed', 'error');
    App.user = res.data;
    App.freshLogin = true;
    try { localStorage.setItem('bht_staff', JSON.stringify({ user: res.data, ts: Date.now() })); } catch (err) {}
    requestBrowserNotifPermission();
    enterApp();
  }, function() { showLoader(false); });
}

function doLogout() {
  try { localStorage.removeItem('bht_cmp_' + (App.user ? App.user.id : 'x')); } catch (e) {}
  try { localStorage.removeItem('bht_staff'); } catch (e) {}
  stopNotifPolling();
  closePanels();

  App.user = null; App.freshLogin = false;
  App.currentPage = 'complaints';
  App.pageParams = {};
  App.currentComplaint = null; App.currentEngineer = null; App.currentJob = null;
  App.currentFeedbackTicket = null; App.currentFeedbackRating = 0; App.currentFeedbackPhone = '';
  App.lastPublic = null;
  App.cache = { complaints: null, engineers: null, reports: null, invoices: null, users: null };
  App.trackData = []; App.historyData = []; App.readMem = null;
  allJobs = []; jobTab = 'New'; billingComplaint = null; billingReport = null;
  NOTIFS = []; NOTIF_PREV = -1; VIEW = {}; MD_DATA = null; MD_ROWS = []; MD_STATUS = 'ALL';
  document.body.classList.remove('is-admin');
  document.body.classList.remove('is-engineer');
  renderBell();

  var jobBtns = document.querySelectorAll('#page-jobs .tabs button');
  jobBtns.forEach(function(b, idx) { b.className = idx === 0 ? 'btn btn-primary btn-sm' : 'btn btn-secondary btn-sm'; });

  ['complaintTableBody','complaintDetailCard','complaintServiceReports','jobCards','jobDetailCard','serviceReportsBody','pendingBillsBody','bl_result','invoiceTableBody','notifBody','usersTableBody','md_body'].forEach(function(id) {
    var el = $(id); if (el) el.innerHTML = '';
  });
  Object.keys(TOOLBARS).forEach(clearToolbarValues);
  if ($('md_q')) $('md_q').value = '';
  if ($('md_per')) $('md_per').value = '';
  if ($('md_from')) $('md_from').value = '';
  if ($('md_to')) $('md_to').value = '';
  $('trackResult').innerHTML = '';

  document.querySelectorAll('.wa-float, .notif-popup').forEach(function(el) { try { el.remove(); } catch (e) {} });

  showLoader(false);
  ['billModal','billDetailsModal','feedbackModal','editModal','shortCloseModal','userModal','serviceReportModal','mdModal'].forEach(function(id) {
    var m = $(id); if (m) m.classList.add('hidden');
  });

  var app = $('appShell'); if (app) app.classList.add('hidden');
  var login = $('staffLogin'); if (login) login.classList.add('hidden');
  var pub = $('publicView'); if (pub) pub.classList.remove('hidden');
  var pt = $('pubTools'); if (pt) pt.classList.remove('hidden');

  document.querySelectorAll('.sidebar a').forEach(function(a){ a.classList.remove('active'); });
  document.querySelectorAll('.page').forEach(function(p){ p.classList.remove('active'); });

  var sb = $('sidebar'); if (sb) sb.classList.remove('open');
  if ($('loginEmail')) $('loginEmail').value = '';
  if ($('loginPassword')) $('loginPassword').value = '';

  pubShow('form');
  try { window.scrollTo({ top: 0, behavior: 'instant' }); } catch (e) { window.scrollTo(0, 0); }
  toast('Signed out successfully', 'success');
}

function enterApp() {
  $('publicView').classList.add('hidden');
  $('staffLogin').classList.add('hidden');
  $('pubTools').classList.add('hidden');
  $('appShell').classList.remove('hidden');
  setHeaderUser();
  var role = App.user.role;
  document.body.classList.toggle('is-admin', role === 'ADMIN');
  document.body.classList.toggle('is-engineer', role === 'ENGINEER');

  var show = {
    mdreport: role === 'ADMIN',
    complaints: ['COORDINATOR','ADMIN','ENGINEER'].includes(role),
    jobs: ['ENGINEER','ADMIN'].includes(role),
    serviceReports: ['COORDINATOR','ADMIN','ENGINEER','HR'].includes(role),
    billing: ['ACCOUNTS','ADMIN'].includes(role),
    invoices: ['ACCOUNTS','ADMIN'].includes(role),
    users: role === 'ADMIN',
    notifications: true
  };
  Object.keys(show).forEach(function(k) {
    var a = document.querySelector('.sidebar a[data-page="' + k + '"]');
    if (a) a.style.display = show[k] ? 'flex' : 'none';
  });
  var landing = 'complaints';
  if (role === 'ADMIN') landing = 'mdreport';
  else if (role === 'ENGINEER') landing = 'jobs';
  else if (role === 'ACCOUNTS') landing = 'billing';
  else if (role === 'HR') landing = 'serviceReports';

  goTo(landing);
  startNotifPolling();
  if (role === 'COORDINATOR' || role === 'ADMIN' || role === 'ENGINEER') {
    callServer('getAllEngineers', [], function(r) { if (r && r.success) App.cache.engineers = r.data || []; }, null, true);
  }
  callServer('getAllServiceReports', [App.user.role, App.user.email, App.user.engineerId || ''], function(r) { if (r && r.success && !App.cache.reports) App.cache.reports = r.data || []; }, null, true);
  if (role === 'ADMIN' || role === 'ACCOUNTS') {
    callServer('getAllInvoices', [], function(r) { if (r && r.success && !App.cache.invoices) App.cache.invoices = r.data || []; }, null, true);
  }
  if (role === 'ADMIN') {
    callServer('getAllUsers', [role], function(r) { if (r && r.success) App.cache.users = r.data || []; }, null, true);
  }
}

window.addEventListener('DOMContentLoaded', function() {
  initPrefs();
  buildAllToolbars();
  if (LANG === 'hi') applyLang();

  document.querySelectorAll('#starContainer i').forEach(function(star) {
    star.addEventListener('mouseover', function() { updateStars(parseInt(this.getAttribute('data-val'), 10)); });
    star.addEventListener('click', function() {
      App.currentFeedbackRating = parseInt(this.getAttribute('data-val'), 10);
      updateStars(App.currentFeedbackRating);
    });
  });
  var container = $('starContainer');
  if (container) container.addEventListener('mouseleave', function() { updateStars(App.currentFeedbackRating || 5); });

  var urlParams = new URLSearchParams(window.location.search);
  var ticketParam = (typeof SERVER_PARAMS !== 'undefined' && SERVER_PARAMS.ticket) ? SERVER_PARAMS.ticket : urlParams.get('ticket');
  var phoneParam  = (typeof SERVER_PARAMS !== 'undefined' && SERVER_PARAMS.phone) ? SERVER_PARAMS.phone : urlParams.get('phone');
  var isFeedback  = (typeof SERVER_PARAMS !== 'undefined' && String(SERVER_PARAMS.feedback).toLowerCase() === 'true') || urlParams.get('feedback') === 'true';

  if (ticketParam) {
    $('publicView').classList.remove('hidden');
    pubShow('track');
    $('tr_ticket').value = ticketParam;
    if (phoneParam) $('tr_phone').value = phoneParam;
    trackComplaint();
    if (isFeedback) setTimeout(function() { openFeedbackModal(ticketParam, phoneParam || ''); }, 600);
    return;
  }

  var saved = null;
  try { saved = localStorage.getItem('bht_staff'); } catch (e) {}
  if (saved) {
    try {
      var parsed = JSON.parse(saved);
      if (parsed && parsed.user && (Date.now() - (parsed.ts || 0)) < 30 * 24 * 60 * 60 * 1000) {
        App.user = parsed.user;
        App.freshLogin = false;
        enterApp();
        return;
      } else {
        try { localStorage.removeItem('bht_staff'); } catch (e2) {}
      }
    } catch (e) {}
  }

  $('publicView').classList.remove('hidden');
  if (typeof google !== 'undefined' && google.script && google.script.url) {
    google.script.url.getLocation(function(loc) {
      if (loc && loc.parameter && loc.parameter.ticket && !ticketParam) {
        pubShow('track');
        $('tr_ticket').value = loc.parameter.ticket;
        if (loc.parameter.phone) $('tr_phone').value = loc.parameter.phone;
        trackComplaint();
        if (String(loc.parameter.feedback).toLowerCase() === 'true') {
          setTimeout(function() { openFeedbackModal(loc.parameter.ticket, loc.parameter.phone || ''); }, 600);
        }
      }
    });
  }
});

function goTo(page, params) {
  App.currentPage = page; App.pageParams = params || {};
  closePanels();
  document.querySelectorAll('.page').forEach(function(p){ p.classList.remove('active'); });
  var el = $('page-' + page); if (el) el.classList.add('active');
  document.querySelectorAll('.sidebar a').forEach(function(a){ a.classList.remove('active'); });
  var link = document.querySelector('.sidebar a[data-page="' + page + '"]'); if (link) link.classList.add('active');
  if (window.innerWidth <= 900) $('sidebar').classList.remove('open');
  var mc = $('mainContent');
  if (mc) mc.scrollTop = 0;
  try { window.scrollTo({ top: 0, behavior: 'smooth' }); } catch (e) { window.scrollTo(0, 0); }
  var h = window['onShow_' + page];
  if (typeof h === 'function') h(App.pageParams);
}

function loadComplaints(cb) { callServer('getAllComplaints', [App.user.role, App.user.email, App.user.engineerId || ''], cb); }

function fetchComplaintsSWR(page, render) {
  var ck = 'bht_cmp_' + (App.user ? App.user.id : 'x');
  if (!App.cache.complaints) {
    var saved = lsGet(ck);
    if (saved) { try { App.cache.complaints = JSON.parse(saved); } catch (e) {} }
  }
  var had = !!App.cache.complaints;
  if (had) render(App.cache.complaints); else showLoader(true);
  loadComplaints(function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    App.cache.complaints = sortTicketsAsc(res.data || []);
    try { lsSet(ck, JSON.stringify(App.cache.complaints.slice(0, 400))); } catch (e) {}
    if (App.currentPage === page) render(App.cache.complaints);
  });
}

/* ============================================================
   MD REPORT (Admin Profile Table Details + Charges/KM + Date Range + Ascending)
============================================================ */
function onShow_mdreport() {
  if (!isAdmin()) return goTo('complaints');
  initMDFilterToggle();
  if (MD_DATA) { buildMD(); renderMD(); }
  loadMD();
}

function loadMD() {
  if (!MD_DATA) showLoader(true);
  callServer('getReportBundle', [App.user.role], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    MD_DATA = res.data;
    if (res.data.engineers) App.cache.engineers = res.data.engineers;
    buildMD();
    renderMD();
  }, function() { showLoader(false); });
}

function reportTotal(r) {
  return (Number(r.serviceCharges) || 0) + (Number(r.foodCharges) || 0) + (Number(r.otherCharges) || 0);
}

function buildMD() {
  var d = MD_DATA || { complaints: [], reports: [], invoices: [], engineers: [] };
  var repMap = {}, invMap = {}, engMap = {};
  (d.engineers || App.cache.engineers || []).forEach(function(e) {
    engMap[String(e.id || '').trim()] = e;
    engMap[String(e.name || '').toUpperCase().trim()] = e;
  });
  (d.reports || []).forEach(function(r) {
    var k = String(r.ticketNumber || '').toUpperCase().trim();
    if (!repMap[k]) repMap[k] = [];
    repMap[k].push(r);
  });
  (d.invoices || []).forEach(function(i) {
    var k = String(i.ticketId || '').toUpperCase().trim();
    if (!invMap[k]) invMap[k] = [];
    invMap[k].push(i);
  });

  var sortedComplaints = sortTicketsAsc(d.complaints || []);
  var maxDone = computeMaxDoneTicket(sortedComplaints);

  MD_ROWS = sortedComplaints.filter(function(c) {
    return c && (c.ticketNumber || c.companyName || c.id);
  }).map(function(c) {
    var k = String(c.ticketNumber || '').toUpperCase().trim();
    var reps = (repMap[k] || []).slice().sort(function(a, b) { return String(a.submittedAt || '').localeCompare(String(b.submittedAt || '')); });
    var invs = invMap[k] || [];
    var svcCharge = 0, foodCharge = 0, otherCharge = 0, kmTotal = 0;
    reps.forEach(function(r) {
      svcCharge += (Number(r.serviceCharges) || 0);
      foodCharge += (Number(r.foodCharges) || 0);
      otherCharge += (Number(r.otherCharges) || 0);
      kmTotal += (Number(r.kmDistance) || 0);
    });
    var service = svcCharge + foodCharge + otherCharge;
    var inv = invs.length ? invs[0] : null;
    var billText;
    if (String(c.warrantyStatus || '').toUpperCase() !== 'OUT_OF_WARRANTY') billText = 'NOT REQ';
    else if (c.billingDone === 'YES') billText = 'DONE';
    else billText = 'PENDING';

    var engObj = engMap[String(c.assignedEngineerId || '').trim()] || engMap[String(c.assignedEngineerName || '').toUpperCase().trim()];
    var st = String(c.status || '').toUpperCase();
    var isDone = (st === 'COMPLETED' || st === 'CLOSED');
    var tNum = ticketNumVal(c.ticketNumber);

    var o = {};
    for (var key in c) o[key] = c[key];
    if (engObj && engObj.name) {
      var extraP = String(c.otherPersonName || '').trim();
      o.assignedEngineerName = extraP ? (engObj.name + ' + ' + extraP) : engObj.name;
    }
    o._reps = reps;
    o._inv = inv;
    o._invs = invs;
    o._visits = reps.length;
    o._lastVisit = reps.length ? String(reps[reps.length - 1].submittedAt || '') : '';
    o._svcCharge = svcCharge;
    o._foodCharge = foodCharge;
    o._otherCharge = otherCharge;
    o._kmTotal = kmTotal;
    o._service = service;
    o._billText = billText;
    o._invNo = inv ? inv.invoiceNumber : '';
    o._invAmt = inv ? (Number(inv.amount) || 0) : '';
    o._isMultiVisit = reps.length >= 2;
    o._isSeqGap = !isDone && tNum > 0 && tNum < maxDone;
    return o;
  });
  fillDyn('md', 'eng', uniq(MD_ROWS, function(r){ return r.assignedEngineerName ? r.assignedEngineerName.split(' +')[0] : ''; }));
}

function mdSetStatus(btn, s) {
  MD_STATUS = s;
  document.querySelectorAll('#md_chips .md-chip').forEach(function(b) { b.classList.remove('active'); });
  btn.classList.add('active');
  renderMD();
}

function toggleMDFilters() {
  var c = document.getElementById("md_filter_container"), btn = document.getElementById("md_ftog");
  if (!c) return;
  var show = c.style.display === 'none';
  c.style.display = show ? 'block' : 'none';
  if (btn) btn.classList.toggle('active', show);
  lsSet('bht_md_ftog', show ? '1' : '0');
}

function initMDFilterToggle() {
  var saved = lsGet('bht_md_ftog');
  var c = document.getElementById("md_filter_container"), btn = document.getElementById("md_ftog");
  if (!c) return;
  if (saved === '0') {
    c.style.display = 'none';
    if (btn) btn.classList.remove('active');
  } else {
    c.style.display = 'block';
    if (btn) btn.classList.add('active');
  }
}

function resetMDFilters() {
  if ($('md_q')) $('md_q').value = '';
  if ($('md_from')) $('md_from').value = '';
  if ($('md_to')) $('md_to').value = '';
  if ($('md_per')) $('md_per').value = '';
  MD_STATUS = 'ALL';
  var chips = document.querySelectorAll('#md_chips .md-chip');
  chips.forEach(function(b, i) { b.classList.toggle('active', i === 0); });
  renderMD();
}

function mdPeriodOk(dateStr, per) {
  if (!per) return true;
  var d = String(dateStr || '').substring(0, 10);
  if (!d) return false;
  var now = new Date();
  var p = function(n) { return (n < 10 ? '0' : '') + n; };
  var today = now.getFullYear() + '-' + p(now.getMonth() + 1) + '-' + p(now.getDate());
  if (per === 'today') return d === today;
  if (per === 'month') return d.substring(0, 7) === today.substring(0, 7);
  if (per === 'year') return d.substring(0, 4) === today.substring(0, 4);
  var days = Number(per) || 0;
  var from = new Date(now.getTime() - days * 86400000);
  var fs = from.getFullYear() + '-' + p(from.getMonth() + 1) + '-' + p(from.getDate());
  return d >= fs;
}

function mdStatusOk(c) {
  var s = String(c.status || '').toUpperCase();
  if (MD_STATUS === 'ALL') return true;
  if (MD_STATUS === 'OPEN') return ['NEW','ASSIGNED','ACCEPTED','ON_SITE'].indexOf(s) > -1;
  if (MD_STATUS === 'MULTI_VISIT') return c._isMultiVisit;
  if (MD_STATUS === 'SEQ_GAP') return c._isSeqGap;
  return s === MD_STATUS;
}

function renderMD() {
  var q = fv('md', 'q'), per = fv('md', 'per'), df = fv('md', 'from'), dt = fv('md', 'to'), en = fv('md', 'eng');
  var list = sortTicketsAsc(MD_ROWS.filter(function(c) {
    if (!mdStatusOk(c)) return false;
    if (en && String(c.assignedEngineerName || '').indexOf(en) === -1) return false;
    if (!mdPeriodOk(c.createdAt, per)) return false;
    if (!inRange(c.createdAt, df, dt)) return false;
    return matchSearch(c, q, ['ticketNumber','companyName','authorizedPerson','phone','email','location','address','equipment','serialNo','assignedEngineerName','status','_invNo','remarks','problemDescription','warrantyStatus','otherPersonName']);
  }));
  var total = 0, totalKm = 0;
  list.forEach(function(c) { total += c._service; totalKm += c._kmTotal; });
  var cnt = $('md_cnt');
  if (cnt) cnt.textContent = list.length + ' of ' + MD_ROWS.length + ' records (Ascending Sequence) • Total Amount ' + money(total) + ' • Total Distance ' + totalKm + ' km';

  var note = [];
  if (MD_STATUS !== 'ALL') note.push('Status: ' + MD_STATUS);
  if (df || dt) note.push('Date: ' + (df || 'start') + ' to ' + (dt || 'end'));
  if (per) note.push('Period: ' + $('md_per').options[$('md_per').selectedIndex].text);
  if (q) note.push('Search: ' + q);
  setView('md', 'MD Report', COLS.md, list, note.join(' | '));

  $('md_body').innerHTML = list.length ? list.map(function(c) {
    var billCell;
    if (c._billText === 'NOT REQ') billCell = '<span class="badge badge-notreq">NOT REQ</span>';
    else if (c._billText === 'DONE') billCell = '<span class="badge badge-completed">DONE</span>' + (c._invNo ? '<div style="font-size:10.5px; margin-top:2px; white-space:nowrap;"><b>' + esc(c._invNo) + '</b> (' + money(c._invAmt) + ')</div>' : '');
    else billCell = '<span class="badge badge-pending">PENDING</span>';

    var rowClass = c._isMultiVisit ? 'multi-visit-row' : (c._isSeqGap ? 'seq-gap-row' : '');
    var ticketFlags = '';
    if (c._isMultiVisit) ticketFlags += '<div><span class="multi-visit-pill"><i class="fa-solid fa-fire"></i> ' + c._visits + ' VISITS</span></div>';
    if (c._isSeqGap) ticketFlags += '<div><span class="seq-gap-pill"><i class="fa-solid fa-triangle-exclamation"></i> REMAINING IN SEQ</span></div>';

    var chargesBreakdown = '<b>Total: ' + money(c._service) + '</b>' +
      '<div style="font-size:10.5px; color:var(--text-muted); margin-top:2px; white-space:nowrap;">' +
      'Svc: ' + money(c._svcCharge) + ' | Food: ' + money(c._foodCharge) + '</div>' +
      '<div style="font-size:10.5px; color:var(--text-muted); white-space:nowrap;">' +
      'Oth: ' + money(c._otherCharge) + ' | <b>KM: ' + c._kmTotal + ' km</b></div>';

    return '<tr class="' + rowClass + '">' +
      '<td data-label="Ticket"><b>' + esc(c.ticketNumber) + '</b>' + ticketFlags + '</td>' +
      '<td data-label="Date">' + esc(fmtDate(c.createdAt)) + '</td>' +
      '<td data-label="Company & Location"><b>' + esc(c.companyName) + '</b><div style="font-size:11px; color:var(--text-muted);">' + esc(c.location || c.address) + '</div></td>' +
      '<td class="md-extra-cell" data-label="Site Person"><b>' + esc(c.authorizedPerson) + '</b><div style="font-size:11px; color:var(--text-muted);">' + esc(c.phone) + '</div></td>' +
      '<td class="md-extra-cell" data-label="Equipment & Problem"><b>' + esc(c.equipment) + '</b><div style="font-size:11px; color:var(--text-muted);">' + esc(c.problemDescription) + '</div></td>' +
      '<td data-label="Warranty"><b>' + esc(formatWarrantyLabel(c.warrantyStatus)) + '</b>' + (c.materialRequired === 'YES' ? '<div style="font-size:10.5px; color:var(--text-muted);">Mat: ' + esc(c.materialDetails || 'YES') + '</div>' : '') + '</td>' +
      '<td data-label="Engineer"><b>' + esc(c.assignedEngineerName || 'Unassigned') + '</b>' + (c.assignedAt ? '<div style="font-size:10.5px; color:var(--text-muted);">' + esc(fmtDate(c.assignedAt)) + '</div>' : '') + '</td>' +
      '<td class="md-extra-cell" data-label="Visits"><span class="md-visits-badge' + (c._isMultiVisit ? ' multi' : '') + '">' + c._visits + '</span>' + (c._lastVisit ? '<div style="font-size:10px; color:var(--text-muted); margin-top:2px;">' + esc(fmtDate(c._lastVisit)) + '</div>' : '') + '</td>' +
      '<td data-label="Status">' + statusBadge(c.status) + '</td>' +
      '<td class="md-extra-cell" data-label="Charges / KM">' + chargesBreakdown + '</td>' +
      '<td class="md-extra-cell" data-label="Bill">' + billCell + '</td>' +
      '<td data-label="Action"><button class="btn btn-primary btn-sm btn-block" data-id="' + esc(c.id) + '" onclick="openMDModal(this.dataset.id)"><i class="fa-solid fa-eye"></i> View Full Details</button></td></tr>';
  }).join('') : '<tr class="empty-row"><td colspan="12">' + (MD_ROWS.length ? 'No records match your search' : 'No data yet') + '</td></tr>';
}

function mdKV(k, v) {
  return '<div><div class="k">' + esc(k) + '</div><div class="v">' + (v === '' || v == null ? '-' : v) + '</div></div>';
}

function openMDModal(id) {
  var c = null;
  for (var i = 0; i < MD_ROWS.length; i++) if (String(MD_ROWS[i].id) === String(id)) { c = MD_ROWS[i]; break; }
  if (!c) return;
  $('mdm_title').innerHTML = '<i class="fa-solid fa-crown" style="color:var(--warning);"></i> ' + esc(c.ticketNumber) + ' &nbsp; ' + statusBadge(c.status);

  var h = '';
  h += '<div class="md-sec"><h4><i class="fa-solid fa-building"></i> Customer & Site</h4><div class="md-kv">' +
    mdKV('Company', esc(c.companyName)) +
    mdKV('Site Person', esc(c.authorizedPerson)) +
    mdKV('Phone', esc(c.phone)) +
    mdKV('Email', '<span class="keep-case">' + esc(c.email) + '</span>') +
    mdKV('Address', esc((c.address || '') + (c.location ? ', ' + c.location : ''))) +
    mdKV('Equipment', esc(c.equipment)) +
    mdKV('Serial No', esc(c.serialNo)) +
    mdKV('Registered On', esc(fmtDate(c.createdAt))) +
    '</div><div style="margin-top:12px;"><div class="md-kv"><div style="grid-column:1/-1;"><div class="k">Problem</div><div class="v">' + esc(c.problemDescription) + '</div></div></div></div></div>';

  h += '<div class="md-sec"><h4><i class="fa-solid fa-user-gear"></i> Assignment & Warranty</h4><div class="md-kv">' +
    mdKV('Engineer / Team', esc(c.assignedEngineerName || 'Unassigned')) +
    mdKV('Other Person', esc(c.otherPersonName || '-')) +
    mdKV('Assigned On', esc(fmtDate(c.assignedAt))) +
    mdKV('Warranty Status', esc(formatWarrantyLabel(c.warrantyStatus))) +
    mdKV('Material Required', esc(c.materialRequired || 'NO')) +
    mdKV('Material Details', esc(c.materialDetails || '-')) +
    mdKV('Remarks', esc(c.remarks || '-')) +
    '</div></div>';

  h += '<div class="md-sec"><h4><i class="fa-solid fa-coins"></i> Cumulative Amount / Charges / KM Summary</h4><div class="md-kv">' +
    mdKV('Service Charges', money(c._svcCharge)) +
    mdKV('Food Charges', money(c._foodCharge)) +
    mdKV('Other Charges', money(c._otherCharge)) +
    mdKV('Total KM Distance', esc(c._kmTotal + ' km')) +
    mdKV('Total Amount', '<b>' + money(c._service) + '</b>') +
    mdKV('Total Visits', '<b>' + c._visits + '</b>') +
    '</div></div>';

  h += '<div class="md-sec"><h4><i class="fa-solid fa-route"></i> Site Visits (' + c._visits + ')</h4>';
  if (!c._reps.length) {
    h += '<p style="color:var(--text-muted);">No visit reported yet.</p>';
  } else {
    c._reps.forEach(function(r, idx) {
      h += '<div class="md-visit"><div class="md-vno">' + (idx + 1) + '</div><div class="md-visit-body">' +
        '<div style="display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-bottom:8px;"><b>' + esc(fmtDate(r.submittedAt)) + '</b>' + statusBadge(r.status) + '<span style="color:var(--text-muted); font-size:12px;">by ' + esc(r.engineerName || '-') + '</span></div>' +
        '<div class="md-kv">' +
        mdKV('Booklet No', esc(r.bookletNo || '-')) +
        mdKV('Service Charge', money(r.serviceCharges)) +
        mdKV('Food', money(r.foodCharges)) +
        mdKV('Other', money(r.otherCharges)) +
        mdKV('KM Distance', esc((r.kmDistance || 0) + ' km')) +
        mdKV('Visit Total', '<b>' + money(reportTotal(r)) + '</b>') +
        (r.status === 'PENDING' ? mdKV('Pending Reason', esc(r.pendingReason || '-')) + mdKV('Next Visit', esc(fmtDate(r.nextVisitDate))) : '') +
        '</div>' +
        (r.remarks ? '<div style="margin-top:8px; font-size:13px;"><span class="k" style="font-size:10.5px; font-weight:800; color:var(--text-muted);">REMARKS:</span> ' + esc(r.remarks) + '</div>' : '') +
        '<div style="margin-top:10px;">' + renderAllReportPhotos(r) + '</div>' +
        '</div></div>';
    });
  }
  h += '</div>';

  h += '<div class="md-sec"><h4><i class="fa-solid fa-file-invoice-dollar"></i> Billing</h4>';
  if (c._billText === 'NOT REQ') {
    h += '<p style="color:var(--text-muted);">Billing not required (' + esc(formatWarrantyLabel(c.warrantyStatus)) + ').</p>';
  } else if (c._invs.length) {
    c._invs.forEach(function(inv) {
      h += '<div class="md-kv" style="margin-bottom:10px;">' +
        mdKV('Invoice No', esc(inv.invoiceNumber)) +
        mdKV('Amount', '<b>' + money(inv.amount) + '</b>') +
        mdKV('Date', esc(fmtDate(inv.createdAt))) +
        mdKV('Bill Copy', inv.billCopyUrl ? '<a href="' + esc(inv.billCopyUrl) + '" target="_blank" style="color:var(--primary); font-weight:700;"><i class="fa-solid fa-file-pdf"></i> Open Bill</a>' : '-') +
        '</div>';
    });
  } else if (c.billingDone === 'YES') {
    h += '<p><span class="badge badge-completed">DONE</span> <span style="color:var(--text-muted);">(closed without invoice / short closed)</span></p>';
  } else {
    h += '<p><span class="badge badge-pending">PENDING</span> <span style="color:var(--text-muted);">Service total so far: ' + money(c._service) + '</span></p>';
  }
  h += '</div>';

  if (Number(c.feedbackRating) > 0) {
    h += '<div class="md-sec"><h4><i class="fa-solid fa-star"></i> Customer Feedback</h4>' +
      '<div style="color:#f59e0b; font-weight:800; font-size:16px;">' + '⭐'.repeat(Math.min(5, Number(c.feedbackRating))) + ' (' + esc(c.feedbackRating) + '/5)</div>' +
      (c.feedbackComment ? '<p style="margin-top:6px;">' + esc(c.feedbackComment) + '</p>' : '') + '</div>';
  }

  waCtxSet(c, findCachedEngineer(c.assignedEngineerId), c._reps, c._invs);
  h += waPanelHtml();
  $('mdm_body').innerHTML = h;
  $('mdModal').classList.remove('hidden');
}
function closeMDModal() { $('mdModal').classList.add('hidden'); }

/* ---------- Complaints (Coordinator Profile: DESCENDING Order so Newest Complaint is at Top) ---------- */
function refreshComplaints() { App.cache.complaints = null; onShow_complaints(); }
function onShow_complaints() {
  fetchComplaintsSWR('complaints', function(list) {
    renderComplaintsList();
  });
}
function renderComplaintsList() {
  var isCoordRole = App.user && String(App.user.role).toUpperCase() === 'COORDINATOR';
  var rawAll = (App.cache.complaints || []).filter(function(c) {
    return c && (String(c.ticketNumber || '').trim() || String(c.companyName || '').trim() || String(c.id || '').trim());
  });
  var all = isCoordRole ? sortTicketsDesc(rawAll) : sortTicketsAsc(rawAll);
  var maxDone = computeMaxDoneTicket(rawAll);

  var q = fv('cmp','q'), st = fv('cmp','status'), wr = fv('cmp','warranty'), df = fv('cmp','from'), dt = fv('cmp','to'), en = fv('cmp','eng');
  fillDyn('cmp', 'eng', uniq(all, function(r){ return r.assignedEngineerName ? r.assignedEngineerName.split(' +')[0] : ''; }));
  var list = all.filter(function(c) {
    if (!c || (!String(c.ticketNumber || '').trim() && !String(c.companyName || '').trim() && !String(c.id || '').trim())) return false;
    if (en && String(c.assignedEngineerName || '').indexOf(en) === -1) return false;
    if (st && String(c.status).toUpperCase() !== st) return false;
    if (!warrantyMatchesFilter(c.warrantyStatus, wr)) return false;
    if (!inRange(c.createdAt, df, dt)) return false;
    return matchSearch(c, q, ['ticketNumber','companyName','authorizedPerson','phone','email','equipment','serialNo','problemDescription','assignedEngineerName','location','address','status','warrantyStatus','otherPersonName']);
  });
  setCount('cmp', list.length, all.length, isCoordRole ? 'Newest First' : 'Ascending Sequence');
  setView('complaints', 'Complaints Report', COLS.complaints, list, collectNote('tb_cmp'));
  renderStatsInto('statsComplaints', list);

  $('complaintTableBody').innerHTML = list.length ? list.map(function(c) {
    var engName = c.assignedEngineerName || '-';
    var billStat = '';
    if (String(c.warrantyStatus).toUpperCase() !== 'OUT_OF_WARRANTY') {
      billStat = '<span class="badge badge-notreq">NOT REQ</span>';
    } else if (c.billingDone === 'YES') {
      billStat = '<span class="badge badge-completed">DONE</span>';
    } else {
      billStat = '<span class="badge badge-pending">PENDING</span>';
    }

    var visits = Number(c._visits) || 0;
    var isMulti = visits >= 2;
    var stUp = String(c.status || '').toUpperCase();
    var isDone = (stUp === 'COMPLETED' || stUp === 'CLOSED');
    var tNum = ticketNumVal(c.ticketNumber);
    var isSeqGap = !isDone && tNum > 0 && tNum < maxDone;

    var rowClass = isMulti ? 'multi-visit-row' : (isSeqGap ? 'seq-gap-row' : '');
    var flagsHtml = '';
    if (isMulti) flagsHtml += '<div><span class="multi-visit-pill"><i class="fa-solid fa-fire"></i> ' + visits + ' VISITS</span></div>';
    if (isSeqGap) flagsHtml += '<div><span class="seq-gap-pill"><i class="fa-solid fa-triangle-exclamation"></i> REMAINING IN SEQ</span></div>';

    return '<tr class="' + rowClass + '">' +
      '<td data-label="Ticket"><b>' + esc(c.ticketNumber) + '</b>' + flagsHtml + '</td>' +
      '<td data-label="Date">' + esc(fmtDate(c.createdAt)) + '</td>' +
      '<td data-label="Company"><b>' + esc(c.companyName) + '</b></td>' +
      '<td data-label="Person">' + esc(c.authorizedPerson) + '</td>' +
      '<td data-label="Phone">' + esc(c.phone) + '</td>' +
      '<td data-label="Location">' + esc(c.location) + '</td>' +
      '<td data-label="Warranty">' + esc(formatWarrantyLabel(c.warrantyStatus)) + '</td>' +
      '<td data-label="Status">' + statusBadge(c.status) + '</td>' +
      '<td data-label="Engineer"><b>' + esc(engName) + '</b></td>' +
      '<td data-label="Visits"><span class="md-visits-badge' + (isMulti ? ' multi' : '') + '">' + visits + '</span></td>' +
      '<td data-label="Billing">' + billStat + '</td>' +
      '<td data-label="Action"><button class="btn btn-primary btn-sm" data-id="' + esc(c.id) + '" onclick="openComplaint(this.dataset.id)"><i class="fa-solid fa-user-plus"></i> Open</button></td></tr>';
  }).join('') : '<tr class="empty-row"><td colspan="12">' + (all.length ? 'No complaints match your search / filters' : 'No complaints registered yet') + '</td></tr>';
}
function openComplaint(id) { goTo('complaintDetail', { id: id }); }

function renderComplaintHistory(list) {
  App.historyData = list || [];
  clearToolbarValues('his');
  renderHistoryList();
}
function renderHistoryList() {
  var all = App.historyData || [];
  var q = fv('his','q'), st = fv('his','status'), df = fv('his','from'), dt = fv('his','to');
  var list = all.filter(function(r) {
    if (st && r.status !== st) return false;
    if (!inRange(r.submittedAt, df, dt)) return false;
    return matchSearch(r, q, ['status','bookletNo','remarks','pendingReason','engineerName','submittedAt']);
  });
  setCount('his', list.length, all.length);
  if (!all.length) { $('complaintServiceReports').innerHTML = '<p style="color:var(--text-muted); padding:12px;">No service reports submitted yet.</p>'; return; }
  if (!list.length) { $('complaintServiceReports').innerHTML = '<p style="color:var(--text-muted); padding:12px;">No matching history.</p>'; return; }
  $('complaintServiceReports').innerHTML = '<div class="table-wrap table-wrap-cards"><table class="responsive-table"><thead><tr><th>Date</th><th>Status</th><th>Booklet</th><th>Charges / KM</th><th>Photos</th><th>Remarks</th></tr></thead><tbody>' +
    list.map(function(r) {
      return '<tr>' +
        '<td data-label="Date">' + esc(fmtDate(r.submittedAt)) + '</td>' +
        '<td data-label="Status">' + statusBadge(r.status) + '</td>' +
        '<td data-label="Booklet">' + esc(r.bookletNo || '-') + '</td>' +
        '<td data-label="Charges / KM">' + renderCharges(r) + '</td>' +
        '<td data-label="Photos">' + renderAllReportPhotos(r) + '</td>' +
        '<td data-label="Remarks">' + esc(r.remarks || r.pendingReason || '-') + '</td></tr>';
    }).join('') + '</tbody></table></div>';
}

function onShow_complaintDetail(p) {
  clearToolbarValues('his');
  if ($('complaintWaPanel')) $('complaintWaPanel').innerHTML = '';
  var cached = findCachedComplaint(p.id);
  if (cached) {
    var ce = findCachedEngineer(cached.assignedEngineerId);
    App.currentComplaint = cached; App.currentEngineer = ce;
    renderComplaintDetail(cached, ce);
    $('complaintServiceReports').innerHTML = '<p class="skeleton">Loading history…</p>';
  } else { showLoader(true); }
  callServer('getComplaintBundle', [p.id], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    if (App.currentPage !== 'complaintDetail') return;
    var c = res.data.complaint, e = res.data.engineer;
    App.currentComplaint = c; App.currentEngineer = e;
    renderComplaintDetail(c, e);
    renderComplaintHistory(res.data.reports);
    renderWaPanel(c, e, res.data.reports, res.data.invoices);
  });
}

function renderComplaintDetail(c, e) {
  var st = String(c.status).toUpperCase();
  var isFinished = st === 'COMPLETED' || st === 'CLOSED';
  var engDisplay = (c.assignedEngineerName || (e && e.name) || 'Unassigned');
  var rowsHtml = '';
  rowsHtml += '<tr><th>Ticket</th><td><b>' + esc(c.ticketNumber) + '</b></td><th>Status</th><td>' + statusBadge(c.status) + '</td></tr>';
  rowsHtml += '<tr><th>Registered Date</th><td>' + esc(fmtDate(c.createdAt)) + '</td><th>Assigned Date</th><td>' + esc(fmtDate(c.assignedAt)) + '</td></tr>';
  rowsHtml += '<tr><th>Company</th><td>' + esc(c.companyName) + '</td><th>Person</th><td>' + esc(c.authorizedPerson) + '</td></tr>';
  rowsHtml += '<tr><th>Email</th><td class="keep-case">' + esc(c.email) + '</td><th>Phone</th><td>' + esc(c.phone) + '</td></tr>';
  rowsHtml += '<tr><th>Address</th><td>' + esc(c.address + ', ' + c.location) + '</td><th>Equipment</th><td>' + esc(c.equipment) + '</td></tr>';
  rowsHtml += '<tr><th>Problem</th><td>' + esc(c.problemDescription) + '</td><th>Warranty</th><td>' + esc(formatWarrantyLabel(c.warrantyStatus)) + '</td></tr>';
  rowsHtml += '<tr><th>Assigned Engineer(s)</th><td><b>' + esc(engDisplay) + '</b></td><th>Remarks</th><td>' + esc(c.remarks || '-') + '</td></tr>';
  rowsHtml += '<tr><th>Billing Status</th><td colspan="3">' + (c.billingDone === 'YES' ? '<span class="badge badge-completed">DONE</span> <a href="#" onclick="goTo(&quot;invoices&quot;); return false;" style="margin-left:10px; font-size:12px; color:var(--primary); text-decoration:underline;">View Bill Copy</a>' : '<span class="badge badge-pending">PENDING</span>') + '</td></tr>';

  if (c.feedbackRating > 0) {
    rowsHtml += '<tr><th>Rating</th><td colspan="3"><div style="color:#f59e0b; font-weight:bold;">' + '⭐'.repeat(Math.min(5, c.feedbackRating)) + ' (' + c.feedbackRating + '/5)</div></td></tr>';
    if (c.feedbackComment) rowsHtml += '<tr><th>Feedback</th><td colspan="3">' + esc(c.feedbackComment) + '</td></tr>';
  }
  var html = '<div class="card"><div class="table-wrap"><table class="detail-table"><tbody>' + rowsHtml + '</tbody></table></div>' +
    '<div style="display:flex; gap:12px; justify-content:center; margin-top:24px; flex-wrap:wrap;">' +
    (canEditComplaint() && !isFinished ? '<button class="btn btn-warning" onclick="openEditComplaint()"><i class="fa-solid fa-pen-to-square"></i> Edit Details</button>' : '') +
    (['COORDINATOR','ADMIN'].includes(App.user.role) && !isFinished ? '<button class="btn btn-primary" data-id="' + esc(c.id) + '" onclick="goTo(&quot;assignEngineer&quot;, {id: this.dataset.id})"><i class="fa-solid fa-user-plus"></i> Assign Engineer</button>' : '') +
    (['COORDINATOR','ADMIN'].includes(App.user.role) && !isFinished ? '<button class="btn btn-danger" data-id="' + esc(c.id) + '" onclick="openShortCloseModal(this.dataset.id)"><i class="fa-solid fa-lock"></i> Short Close</button>' : '') +
    '<button class="btn btn-secondary" onclick="goTo(&quot;complaints&quot;)">Back to List</button></div></div>';
  $('complaintDetailCard').innerHTML = html;
}

function openEditComplaint() {
  var c = App.currentComplaint;
  if (!c || !canEditComplaint()) return;
  $('ed_company').value = c.companyName || '';
  $('ed_person').value = c.authorizedPerson || '';
  $('ed_email').value = c.email || '';
  $('ed_phone').value = c.phone || '';
  $('ed_address').value = c.address || '';
  $('ed_location').value = c.location || '';
  $('ed_equipment').value = c.equipment || '';
  $('ed_serial').value = c.serialNo || '';
  $('ed_problem').value = c.problemDescription || '';
  $('editModal').classList.remove('hidden');
}
function closeEditModal() { $('editModal').classList.add('hidden'); }
function submitEditComplaint(btn) {
  var c = App.currentComplaint;
  if (!c || !canEditComplaint()) return;
  var p = {
    companyName: $('ed_company').value.trim(), authorizedPerson: $('ed_person').value.trim(),
    email: $('ed_email').value.trim(), phone: $('ed_phone').value.trim(),
    address: $('ed_address').value.trim(), location: $('ed_location').value.trim(),
    equipment: $('ed_equipment').value.trim(), serialNo: $('ed_serial').value.trim(),
    problemDescription: $('ed_problem').value.trim()
  };
  if (!p.companyName || !p.authorizedPerson || !p.phone || !p.address || !p.equipment || !p.problemDescription)
    return toast('Please fill all required fields', 'error');
  if (!lockBtn(btn)) return;
  showLoader(true);
  callServer('updateComplaintDetails', [c.id, p, App.user.role], function(res) {
    showLoader(false); unlockBtn(btn);
    if (!res.success) return toast(res.message, 'error');
    patchCached(c.id, res.data);
    App.currentComplaint = res.data;
    MD_DATA = null;
    closeEditModal();
    renderComplaintDetail(res.data, App.currentEngineer);
    toast('Complaint details updated', 'success');
  }, function() { showLoader(false); unlockBtn(btn); });
}

function openShortCloseModal(id) {
  shortCloseId = id;
  $('sc_reason').value = '';
  $('shortCloseModal').classList.remove('hidden');
  setTimeout(function() { try { $('sc_reason').focus(); } catch(e){} }, 150);
}
function closeShortCloseModal() { $('shortCloseModal').classList.add('hidden'); shortCloseId = null; }
function submitShortClose(btn) {
  var reason = $('sc_reason').value.trim();
  if (!reason) return toast('Please provide a valid reason', 'error');
  if (!shortCloseId) return;
  if (!lockBtn(btn)) return;
  showLoader(true);
  callServer('shortCloseComplaint', [shortCloseId, reason], function(res) {
    showLoader(false); unlockBtn(btn);
    if (res.success) {
      patchCached(shortCloseId, { status: 'CLOSED', billingDone: 'YES' });
      MD_DATA = null;
      toast('Ticket short-closed successfully', 'success');
      closeShortCloseModal();
      goTo('complaints');
    } else {
      toast(res.message, 'error');
    }
  }, function() { showLoader(false); unlockBtn(btn); });
}

/* Updated Engineer Select: Shows clean Engineer Name without "OTHER" option; separate "Other Person Name" field below */
function fillEngineerSelect(list) {
  var sel = $('ae_engineer'), cur = sel.value;
  var opts = '<option value="">Select Engineer</option>' + (list || []).map(function(e) {
    return '<option value="' + esc(e.id) + '">' + esc(e.name) + '</option>';
  }).join('');
  sel.innerHTML = opts;
  if (cur) sel.value = cur;
}
function reloadEngineersList() {
  showLoader(true);
  callServer('getAllEngineers', [], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    App.cache.engineers = res.data || [];
    fillEngineerSelect(App.cache.engineers);
    toast('Synced latest engineers from Sheet', 'success');
  });
}
function onShow_assignEngineer() {
  toggleMaterialDetails();
  if ($('ae_otherPerson')) $('ae_otherPerson').value = '';
  if (App.cache.engineers) fillEngineerSelect(App.cache.engineers);
  else showLoader(true);
  callServer('getAllEngineers', [], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    App.cache.engineers = res.data || [];
    fillEngineerSelect(App.cache.engineers);
  });
}
function toggleMaterialDetails() { $('ae_matWrap').classList.toggle('hidden', $('ae_matReq').value !== 'true'); }
function submitAssign(btn) {
  var selVal = $('ae_engineer').value;
  if (!selVal) return toast('Please select an engineer', 'error');
  var p = {
    ticketId: App.pageParams.id,
    assignedEngineerId: selVal,
    otherPersonName: $('ae_otherPerson') ? $('ae_otherPerson').value.trim() : '',
    warrantyStatus: $('ae_warranty').value,
    materialRequired: $('ae_matReq').value === 'true',
    materialDetails: $('ae_matDet').value,
    remarks: $('ae_remarks').value
  };
  if (!lockBtn(btn)) return;
  showLoader(true);
  callServer('assignEngineer', [p], function(res) {
    showLoader(false); unlockBtn(btn);
    if (!res.success) return toast(res.message, 'error');
    toast('Engineer assigned successfully', 'success');
    MD_DATA = null;
    if (res.data.complaint) patchCached(p.ticketId, res.data.complaint);
    if (res.data.complaint && res.data.engineer) showAssignWaBox(res.data.complaint, res.data.engineer);
    bg('bgAssignEmails', [p.ticketId]);
    goTo('complaints');
  }, function() { showLoader(false); unlockBtn(btn); });
}

/* ---------- Jobs ---------- */
function onShow_jobs() {
  fetchComplaintsSWR('jobs', function(list) {
    allJobs = sortTicketsAsc(list || []);
    renderStatsInto('statsJobs', allJobs);
    renderJobCards();
  });
}
function switchJobTab(btn, t) {
  jobTab = t;
  document.querySelectorAll('#page-jobs .tabs button').forEach(function(b){ b.className = 'btn btn-secondary btn-sm'; });
  btn.className = 'btn btn-primary btn-sm';
  renderJobCards();
}
function renderJobCards() {
  var q = fv('job','q'), eq = fv('job','equip');
  var base = allJobs.filter(function(j) {
    if (!j || (!j.ticketNumber && !j.companyName && !j.id)) return false;
    if (jobTab === 'New') return j.status === 'ASSIGNED';
    if (jobTab === 'Active') return ['ACCEPTED','ON_SITE'].includes(j.status);
    if (jobTab === 'Pending') return j.status === 'PENDING';
    return ['COMPLETED','CLOSED'].includes(j.status);
  });
  var list = sortTicketsAsc(base.filter(function(j) {
    if (eq && j.equipment !== eq) return false;
    return matchSearch(j, q, ['ticketNumber','companyName','authorizedPerson','phone','equipment','address','location','problemDescription','serialNo']);
  }));
  setCount('job', list.length, base.length);
  setView('jobs', 'Jobs - ' + jobTab, COLS.jobs, list, collectNote('tb_job'));
  $('jobCards').innerHTML = list.length ? list.map(function(j) {
    var visits = Number(j._visits) || 0;
    return '<div class="card" style="margin:0;' + (visits >= 2 ? 'border:2px solid #f59e0b;' : '') + '">' +
      '<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">' +
      '<h4 style="font-size:18px; margin:0; font-weight:800; letter-spacing:-0.3px;">' + esc(j.ticketNumber) + '</h4>' +
      '<span style="font-size:11.5px; color:var(--text-muted);">' + esc(fmtDate(j.createdAt)) + '</span></div>' +
      (visits >= 2 ? '<div style="margin-bottom:8px;"><span class="multi-visit-pill"><i class="fa-solid fa-fire"></i> ' + visits + ' VISITS</span></div>' : '') +
      '<p style="color:var(--text-muted); font-size:14px; margin-bottom:12px;"><i class="fa-solid fa-building"></i> ' + esc(j.companyName) + '</p>' +
      statusBadge(j.status) +
      '<div style="margin-top:16px; display:flex; gap:8px;">' +
      '<button class="btn btn-primary btn-sm" style="flex:1;" data-id="' + esc(j.id) + '" onclick="goTo(&quot;jobDetail&quot;, {id: this.dataset.id})">Open Job</button>' +
      '</div></div>';
  }).join('') : '<p style="grid-column:1/-1; text-align:center; padding:32px; color:var(--text-muted);">No jobs found in this section.</p>';
}
function onShow_jobDetail(p) {
  var cached = findCachedComplaint(p.id);
  if (cached) renderJobDetail(cached); else showLoader(true);
  callServer('getComplaintById', [p.id], function(res) {
    showLoader(false);
    if (res.success && App.currentPage === 'jobDetail') renderJobDetail(res.data);
  });
}
function renderJobDetail(j) {
  App.currentJob = j;
  var isFinished = j.status === 'COMPLETED' || j.status === 'CLOSED';
  var statusActions = '';
  if (j.status === 'ASSIGNED') statusActions = '<button class="btn btn-warning" data-id="' + esc(j.id) + '" onclick="changeJobStatus(this.dataset.id, &quot;ACCEPTED&quot;)"><i class="fa-solid fa-handshake"></i> Accept Job</button>';
  else if (j.status === 'ACCEPTED') statusActions = '<button class="btn btn-warning" data-id="' + esc(j.id) + '" onclick="changeJobStatus(this.dataset.id, &quot;ON_SITE&quot;)"><i class="fa-solid fa-location-dot"></i> Mark On-Site</button>';

  var html = '<div class="card"><div class="table-wrap"><table class="detail-table"><tbody>' +
    '<tr><th>Ticket</th><td>' + esc(j.ticketNumber) + '</td><th>Status</th><td>' + statusBadge(j.status) + '</td></tr>' +
    '<tr><th>Date</th><td>' + esc(fmtDate(j.createdAt)) + '</td><th>Assigned On</th><td>' + esc(fmtDate(j.assignedAt)) + '</td></tr>' +
    '<tr><th>Company</th><td>' + esc(j.companyName) + '</td><th>Person</th><td>' + esc(j.authorizedPerson) + ' (' + esc(j.phone) + ')</td></tr>' +
    '<tr><th>Email</th><td colspan="3" class="keep-case">' + esc(j.email || '-') + '</td></tr>' +
    '<tr><th>Address</th><td colspan="3">' + esc(j.address + ', ' + j.location) + '</td></tr>' +
    '<tr><th>Equipment</th><td>' + esc(j.equipment) + '</td><th>Serial No</th><td>' + esc(j.serialNo || '-') + '</td></tr>' +
    '<tr><th>Problem</th><td colspan="3">' + esc(j.problemDescription) + '</td></tr>' +
    '<tr><th>Warranty</th><td>' + esc(formatWarrantyLabel(j.warrantyStatus)) + '</td><th>Material Required</th><td>' + esc(j.materialDetails || 'None') + '</td></tr>' +
    '<tr><th>Assigned Team</th><td><b>' + esc(j.assignedEngineerName || '-') + '</b></td><th>Remarks</th><td>' + esc(j.remarks || '-') + '</td></tr></tbody></table></div>' +
    '<div style="display:flex; gap:12px; justify-content:center; margin-top:24px; flex-wrap:wrap;">' +
    statusActions +
    (!isFinished ? '<button class="btn btn-primary" data-id="' + esc(j.id) + '" onclick="goTo(&quot;serviceReport&quot;, {id: this.dataset.id})"><i class="fa-solid fa-wrench"></i> Submit Service Report</button>' : '') +
    '<button class="btn btn-secondary" onclick="goTo(&quot;jobs&quot;)">Back to Jobs</button></div></div>';
  $('jobDetailCard').innerHTML = html;
}
function changeJobStatus(ticketId, newStatus) {
  showLoader(true);
  callServer('updateJobStatus', [ticketId, newStatus], function(res) {
    showLoader(false);
    if (res.success) {
      patchCached(ticketId, { status: newStatus });
      MD_DATA = null;
      toast('Status updated to ' + newStatus, 'success');
      goTo('jobDetail', { id: ticketId });
    }
  });
}

/* ---------- Service report ---------- */
function onShow_serviceReport() {
  $('sr_status').value = 'COMPLETED'; toggleReportFields();
  ['sr_bookletNo','sr_remarks','sr_pendingReason','sr_nextVisit'].forEach(function(id){ var el = $(id); if (el) el.value = ''; });
  ['sr_serviceCharge','sr_km','sr_food','sr_other','sr_pkm','sr_pfood'].forEach(function(id){ var el = $(id); if (el) el.value = '0'; });
  ['sr_serviceImages','sr_equipPhotos','sr_equipPhotosPend'].forEach(function(id){ var el = $(id); if (el) el.value = ''; });
}
function toggleReportFields() {
  var s = $('sr_status').value;
  $('sr_completeFields').style.display = s === 'COMPLETED' ? 'block' : 'none';
  $('sr_pendingFields').style.display = s === 'PENDING' ? 'block' : 'none';
}
function readFileAsBase64(f, cb) {
  if (!f) return cb('');
  var r = new FileReader();
  r.onload = function(e) {
    var res = String(e.target && e.target.result || '');
    cb(res.indexOf(',') > -1 ? res.split(',')[1] : res);
  };
  r.onerror = function() { cb(''); };
  r.readAsDataURL(f);
}

function resizeImage(f, mw, mh) {
  return new Promise(function(resolve) {
    if (!f || !f.type || !f.type.startsWith('image/')) return resolve(f);
    var url = null;
    try { url = URL.createObjectURL(f); } catch (e) { return resolve(f); }
    var img = new Image();
    img.onerror = function() { try { URL.revokeObjectURL(url); } catch (e) {} resolve(f); };
    img.onload = function() {
      try {
        var w = img.width, h = img.height;
        if (w > mw || h > mh) {
          var ra = Math.min(mw / w, mh / h);
          w = Math.round(w * ra);
          h = Math.round(h * ra);
        }
        var c = document.createElement('canvas');
        c.width = w; c.height = h;
        c.getContext('2d').drawImage(img, 0, 0, w, h);
        try { URL.revokeObjectURL(url); } catch (e) {}
        c.toBlob(function(b) {
          if (!b) return resolve(f);
          resolve(new File([b], String(f.name || 'photo').replace(/\.[^.]+$/, '') + '.jpg', { type: 'image/jpeg' }));
        }, 'image/jpeg', 0.55);
      } catch (err) { resolve(f); }
    };
    img.src = url;
  });
}

function processFiles(fList, cb) {
  if (!fList || !fList.length) return cb([]);
  var ps = [];
  for (var i = 0; i < fList.length; i++) {
    (function(f) {
      ps.push(new Promise(function(res) {
        var p = (f.type && f.type.startsWith('image/')) ? resizeImage(f, 800, 800) : Promise.resolve(f);
        p.then(function(pf) {
          readFileAsBase64(pf, function(b) {
            if (!b) return res(null);
            res({
              base64: b,
              name: pf.name || 'file.jpg',
              type: pf.type || 'image/jpeg'
            });
          });
        });
      }));
    })(fList[i]);
  }
  Promise.all(ps).then(function(arr) {
    cb(arr.filter(Boolean));
  });
}
function submitServiceReport(btn) {
  var s = $('sr_status').value;
  var p = { ticketId: App.pageParams.id, status: s };
  var svc = null, eqp = null;
  if (s === 'COMPLETED') {
    p.bookletNo = $('sr_bookletNo').value;
    p.serviceCharges = Number($('sr_serviceCharge').value) || 0;
    p.kmDistance = Number($('sr_km').value) || 0;
    p.foodCharges = Number($('sr_food').value) || 0;
    p.otherCharges = Number($('sr_other').value) || 0;
    p.remarks = $('sr_remarks').value;
    svc = $('sr_serviceImages').files; eqp = $('sr_equipPhotos').files;
  } else {
    if (!$('sr_pendingReason').value || !$('sr_nextVisit').value) return toast('Please provide pending reason and next visit date', 'error');
    p.kmDistance = Number($('sr_pkm').value) || 0;
    p.foodCharges = Number($('sr_pfood').value) || 0;
    p.pendingReason = $('sr_pendingReason').value;
    p.nextVisitDate = $('sr_nextVisit').value;
    eqp = $('sr_equipPhotosPend').files;
  }
  if (!lockBtn(btn)) return;
  showLoader(true);
  var ts = [];
  ts.push(new Promise(function(res) { processFiles(svc, function(arr) { p.serviceImagesArray = arr; res(); }); }));
  ts.push(new Promise(function(res) { processFiles(eqp, function(arr) { p.equipmentPhotosArray = arr; res(); }); }));
  Promise.all(ts).then(function() {
    callServer('submitServiceReport', [p], function(res) {
      showLoader(false); unlockBtn(btn);
      if (!res.success) return toast(res.message, 'error');
      toast('Service report submitted — customer and coordinator will be notified.', 'success');
      patchCached(p.ticketId, { status: (s === 'COMPLETED' ? 'COMPLETED' : 'PENDING') });
      App.cache.reports = null; App.cache.complaints = null; MD_DATA = null;
      if (res.data.complaint) showServiceReportWaBox(res.data.complaint, p);
      bg('bgReportEmails', [p.ticketId]);
      goTo('jobs');
    }, function() { showLoader(false); unlockBtn(btn); });
  });
}

/* ---------- Service reports log (Column 1 = Ticket, Column 2 = Date) ---------- */
function onShow_serviceReports() {
  if (App.cache.reports) { renderReportsList(); }
  else showLoader(true);
  callServer('getAllServiceReports', [App.user.role, App.user.email, App.user.engineerId || ''], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    App.cache.reports = sortTicketsAsc(res.data || []);
    if (App.currentPage === 'serviceReports') renderReportsList();
  });
}
function renderReportsList() {
  var all = sortTicketsAsc(App.cache.reports || []);
  fillDyn('rep', 'eng', uniq(all, function(r){ return r.engineerName ? r.engineerName.split(' +')[0] : ''; }));
  var q = fv('rep','q'), st = fv('rep','status'), en = fv('rep','eng'), df = fv('rep','from'), dt = fv('rep','to');
  var list = all.filter(function(r) {
    if (!r || (!r.ticketNumber && !r.customerName)) return false;
    if (st && String(r.status).toUpperCase() !== st) return false;
    if (en && r.engineerName !== en) return false;
    if (!inRange(r.submittedAt, df, dt)) return false;
    return matchSearch(r, q, ['ticketNumber','customerName','engineerName','status','bookletNo','remarks','pendingReason','submittedAt','_location']);
  });
  var total = 0;
  list.forEach(function(r) { total += reportTotal(r); });
  setCount('rep', list.length, all.length, 'Total ' + money(total));
  setView('reports', 'Service Reports', COLS.reports, list, collectNote('tb_rep'));

  $('serviceReportsBody').innerHTML = list.length ? list.map(function(r) {
    var loc = '';
    if (App.cache.complaints) {
      var cmp = App.cache.complaints.find(function(c){ return c.ticketNumber === r.ticketNumber; });
      if (cmp) loc = cmp.location;
    }

    return '<tr>' +
      '<td data-label="Ticket"><b>' + esc(r.ticketNumber) + '</b></td>' +
      '<td data-label="Date">' + esc(fmtDate(r.submittedAt)) + '</td>' +
      '<td data-label="Customer"><b>' + esc(r.customerName) + '</b></td>' +
      '<td data-label="Location">' + esc(r._location || loc) + '</td>' +
      '<td data-label="Engineer"><b>' + esc(r.engineerName) + '</b></td>' +
      '<td data-label="Status">' + statusBadge(r.status) + '</td>' +
      '<td data-label="Amount / Charges / KM">' + renderCharges(r) + '</td>' +
      '<td data-label="Photos">' + renderAllReportPhotos(r) + '</td>' +
      '<td data-label="Action"><button class="btn btn-secondary btn-sm" data-t="' + esc(r.ticketNumber) + '" onclick="viewServiceReport(this.dataset.t)"><i class="fa-solid fa-eye"></i> View</button></td></tr>';
  }).join('') : '<tr class="empty-row"><td colspan="9">' + (all.length ? 'No reports match your search / filters' : 'No service reports found') + '</td></tr>';
}

function viewServiceReport(ticketNumber) {
  showLoader(true);
  callServer('getComplaintBundle', [ticketNumber], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    var c = res.data.complaint;
    var e = res.data.engineer;
    var r = res.data.reports[res.data.reports.length - 1] || {};

    var html = '<div class="table-wrap"><table class="detail-table"><tbody>' +
      '<tr><th>Ticket</th><td>' + esc(ticketNumber) + '</td><th>Status</th><td>' + statusBadge(r.status || 'PENDING') + '</td></tr>' +
      '<tr><th>Date</th><td>' + esc(fmtDate(r.submittedAt)) + '</td><th>Booklet No</th><td>' + esc(r.bookletNo || '-') + '</td></tr>' +
      '<tr><th>Customer</th><td>' + esc(c ? c.companyName : r.customerName) + '</td><th>Location</th><td>' + esc(c ? c.location : '-') + '</td></tr>' +
      '<tr><th>Engineer / Team</th><td colspan="3"><b>' + esc(r.engineerName || (c && c.assignedEngineerName) || (e && e.name) || '-') + '</b></td></tr>' +
      '<tr><th>Service Charge</th><td>₹' + (r.serviceCharges || 0) + '</td><th>KM Distance</th><td>' + (r.kmDistance || 0) + ' km</td></tr>' +
      '<tr><th>Food Charges</th><td>₹' + (r.foodCharges || 0) + '</td><th>Other Charges</th><td>₹' + (r.otherCharges || 0) + '</td></tr>' +
      '<tr><th>Total Amount</th><td colspan="3"><b>₹' + reportTotal(r) + '</b></td></tr>' +
      '<tr><th>Remarks</th><td colspan="3">' + esc(r.remarks || r.pendingReason || '-') + '</td></tr>' +
      '<tr><th>Service Photos</th><td colspan="3">' + renderPhotoLinks(r.serviceImages) + '</td></tr>' +
      '<tr><th>Equipment Photos</th><td colspan="3">' + renderPhotoLinks(r.equipmentPhotos) + '</td></tr>' +
      '</tbody></table></div>';

    $('sr_modal_content').innerHTML = html;
    $('serviceReportModal').classList.remove('hidden');
  });
}
function closeSRModal() { $('serviceReportModal').classList.add('hidden'); }

/* ---------- Billing ---------- */
function onShow_billing() {
  $('bl_result').innerHTML = '';
  fetchComplaintsSWR('billing', function(all) {
    renderBillingList();
  });
}
function renderBillingList() {
  var all = sortTicketsAsc(App.cache.complaints || []);
  var q = fv('bil','q'), eq = fv('bil','equip'), df = fv('bil','from'), dt = fv('bil','to');
  var pending = all.filter(function(c) {
    if (!c || (!c.ticketNumber && !c.companyName)) return false;
    return String(c.status).toUpperCase() === 'COMPLETED' &&
           String(c.warrantyStatus).toUpperCase() === 'OUT_OF_WARRANTY' &&
           String(c.billingDone || 'NO').toUpperCase() !== 'YES';
  });
  var list = pending.filter(function(c) {
    if (eq && c.equipment !== eq) return false;
    if (!inRange(c.createdAt, df, dt)) return false;
    return matchSearch(c, q, ['ticketNumber','companyName','equipment','address','location','authorizedPerson','phone']);
  });
  renderStatsInto('statsBilling', all);
  setCount('bil', list.length, pending.length);
  setView('billing', 'Pending Bills', COLS.billing, list, collectNote('tb_bil'));
  $('pendingBillsBody').innerHTML = list.length ? list.map(function(c) {
    return '<tr>' +
      '<td data-label="Ticket"><b>' + esc(c.ticketNumber) + '</b></td>' +
      '<td data-label="Date">' + esc(fmtDate(c.createdAt)) + '</td>' +
      '<td data-label="Company">' + esc(c.companyName) + '</td>' +
      '<td data-label="Equipment">' + esc(c.equipment) + '</td>' +
      '<td data-label="Action"><button class="btn btn-primary btn-sm" data-id="' + esc(c.id) + '" data-t="' + esc(c.ticketNumber) + '" onclick="openBillFor(this.dataset.id, this.dataset.t)"><i class="fa-solid fa-file-invoice"></i> Create Bill</button></td></tr>';
  }).join('') : '<tr class="empty-row"><td colspan="5">' + (pending.length ? 'No bills match your search / filters' : 'No pending bills found') + '</td></tr>';
}
function openBillFor(id, ticketNumber) {
  showLoader(true);
  callServer('getServiceReportsForTicket', [ticketNumber], function(res) {
    showLoader(false);
    var reports = res.data || [];
    var report = reports[reports.length - 1] || null;
    var cached = findCachedComplaint(id);
    if (cached) { billingComplaint = cached; billingReport = report; renderBilling(cached, report); return; }
    callServer('getComplaintById', [id], function(r2) {
      if (!r2.success) return toast(r2.message, 'error');
      billingComplaint = r2.data; billingReport = report;
      renderBilling(r2.data, report);
    });
  });
}

function renderBilling(c, r) {
  var isOut = String(c.warrantyStatus).toUpperCase() === 'OUT_OF_WARRANTY';
  var chargesHtml = '';
  if (r) {
    var total = reportTotal(r);
    chargesHtml = '<tr><th>SERVICE CHARGE</th><td>₹' + (r.serviceCharges || 0) + '</td><th>FOOD CHARGES</th><td>₹' + (r.foodCharges || 0) + '</td></tr>' +
                  '<tr><th>OTHER CHARGES</th><td>₹' + (r.otherCharges || 0) + '</td><th>KM DISTANCE</th><td>' + (r.kmDistance || 0) + ' km</td></tr>' +
                  '<tr><th>TOTAL BILLABLE</th><td colspan="3"><b>₹' + total + '</b></td></tr>';
  }

  var imagesHtml = '';
  if (r) {
    var sImgs = String(r.serviceImages || '').split(/[\r\n]+/).filter(Boolean);
    var eImgs = String(r.equipmentPhotos || '').split(/[\r\n]+/).filter(Boolean);
    if (sImgs.length || eImgs.length) {
      imagesHtml += '<tr><th>SERVICE IMAGES</th><td colspan="3"><div style="display:flex; gap:10px; flex-wrap:wrap; padding:10px 0;">';
      sImgs.forEach(function(url) {
        imagesHtml += '<a href="' + esc(url) + '" target="_blank"><img src="' + esc(url) + '" style="width:60px; height:60px; object-fit:cover; border-radius:6px; border:1px solid var(--border);" alt="Service"></a>';
      });
      eImgs.forEach(function(url) {
        imagesHtml += '<a href="' + esc(url) + '" target="_blank"><img src="' + esc(url) + '" style="width:60px; height:60px; object-fit:cover; border-radius:6px; border:1px solid var(--border);" alt="Equipment"></a>';
      });
      imagesHtml += '</div></td></tr>';
    }
  }

  var html = '<div class="table-wrap" style="border:none;"><table class="detail-table" style="border:1px solid var(--border-strong);"><tbody>' +
    '<tr><th>TICKET</th><td>' + esc(c.ticketNumber) + '</td><th>COMPANY</th><td>' + esc(c.companyName) + '</td></tr>' +
    '<tr><th>ADDRESS</th><td colspan="3">' + esc(c.address + ', ' + c.location) + '</td></tr>' +
    '<tr><th>WARRANTY</th><td colspan="3">' + esc(formatWarrantyLabel(c.warrantyStatus)) + '</td></tr>' +
    chargesHtml +
    imagesHtml +
    '</tbody></table></div>';

  if (!isOut) html += '<div style="color:var(--danger); text-align:center; margin-top:16px; font-weight:700;"><i class="fa-solid fa-triangle-exclamation"></i> Only Out-of-Warranty jobs require invoicing.</div>';
  else html += '<div style="display:flex; gap:12px; justify-content:center; margin-top:24px; flex-wrap:wrap;"><button class="btn btn-success" onclick="openBillModal()"><i class="fa-solid fa-upload"></i> Upload Invoice</button><button class="btn btn-secondary" onclick="closeBillDetailsModal()">Cancel</button></div>';

  $('billDetailsContent').innerHTML = html;
  $('billDetailsModal').classList.remove('hidden');
}

function closeBillDetailsModal() {
  $('billDetailsModal').classList.add('hidden');
}

function openBillModal() {
  closeBillDetailsModal();
  $('inv_num').value = '';
  $('inv_amount').value = '';
  $('inv_file').value = '';
  $('billModal').classList.remove('hidden');
}
function closeBillModal() { $('billModal').classList.add('hidden'); }
function submitNewInvoice(btn) {
  var n = $('inv_num').value.trim(), a = $('inv_amount').value, f = $('inv_file').files[0];
  if (!n || !a || !f) return toast('Please enter invoice number, amount, and attach copy', 'error');
  if (!lockBtn(btn)) return;
  showLoader(true);
  var prep = (f.type && f.type.startsWith('image/')) ? resizeImage(f, 1600, 1600) : Promise.resolve(f);
  prep.then(function(pf) {
    readFileAsBase64(pf, function(b) {
      var p = { ticketId: billingComplaint.id, ticketNumber: billingComplaint.ticketNumber, customerName: billingComplaint.companyName,
                invoiceNumber: n, amount: a, billBase64: b, billName: pf.name, billType: pf.type, uploadedBy: (App.user && App.user.name) || 'ACCOUNTS' };
      callServer('createInvoice', [p], function(res) {
        showLoader(false); unlockBtn(btn);
        if (!res.success) return toast(res.message, 'error');
        patchCached(billingComplaint.id, { status: 'CLOSED', billingDone: 'YES' });
        App.cache.invoices = null; MD_DATA = null;
        toast('Invoice uploaded and recorded', 'success');
        closeBillModal();
        if (billingComplaint) showBillWaBox(billingComplaint, res.data);
        goTo('invoices');
      }, function() { showLoader(false); unlockBtn(btn); });
    });
  });
}

/* ---------- Invoices ---------- */
function onShow_invoices() {
  if (App.cache.invoices) { fillDyn('inv', 'cust', uniq(App.cache.invoices, 'customerName')); renderInvoicesList(); }
  else showLoader(true);
  callServer('getAllInvoices', [], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    App.cache.invoices = sortTicketsAsc(res.data || []);
    fillDyn('inv', 'cust', uniq(App.cache.invoices, 'customerName'));
    if (App.currentPage === 'invoices') renderInvoicesList();
  });
}
function renderInvoicesList() {
  var all = sortTicketsAsc(App.cache.invoices || []);
  var q = fv('inv','q'), cu = fv('inv','cust'), df = fv('inv','from'), dt = fv('inv','to');
  var list = all.filter(function(i) {
    if (!i || (!i.invoiceNumber && !i.ticketId)) return false;
    if (cu && i.customerName !== cu) return false;
    if (!inRange(i.createdAt, df, dt)) return false;
    return matchSearch(i, q, ['invoiceNumber','ticketId','customerName','amount','createdAt']);
  });
  var total = 0; list.forEach(function(i) { total += Number(i.amount) || 0; });
  setCount('inv', list.length, all.length, 'Total ' + money(total));
  setView('invoices', 'Invoices', COLS.invoices, list, collectNote('tb_inv'));
  $('invoiceTableBody').innerHTML = list.length ? list.map(function(i) {
    return '<tr>' +
      '<td data-label="Ticket No"><b>' + esc(i.ticketId) + '</b></td>' +
      '<td data-label="Date">' + esc(fmtDate(i.createdAt)) + '</td>' +
      '<td data-label="Invoice No"><b>' + esc(i.invoiceNumber) + '</b></td>' +
      '<td data-label="Customer">' + esc(i.customerName) + '</td>' +
      '<td data-label="Amount">₹' + Number(i.amount || 0).toFixed(2) + '</td>' +
      '<td data-label="Bill Copy">' + renderDocLink(i.billCopyUrl) + '</td></tr>';
  }).join('') : '<tr class="empty-row"><td colspan="6">' + (all.length ? 'No invoices match your search / filters' : 'No invoices recorded yet') + '</td></tr>';
}

/* ---------- Users (Admin) ---------- */
function onShow_users() {
  if (!isAdmin()) return goTo('complaints');
  if (App.cache.users) renderUsersList(); else showLoader(true);
  callServer('getAllUsers', [App.user.role], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    App.cache.users = res.data || [];
    renderUsersList();
  });
}
function renderUsersList() {
  var list = App.cache.users || [];
  $('usersTableBody').innerHTML = list.length ? list.map(function(u) {
    return '<tr>' +
      '<td data-label="Name"><b>' + esc(u.name) + '</b></td>' +
      '<td data-label="Email" class="keep-case">' + esc(u.email) + '</td>' +
      '<td data-label="Password" class="keep-case"><code>' + esc(u.password) + '</code></td>' +
      '<td data-label="Phone">' + esc(u.phone || '-') + '</td>' +
      '<td data-label="Role">' + esc(u.role) + '</td>' +
      '<td data-label="Action"><button class="btn btn-secondary btn-sm" data-id="' + esc(u.id) + '" onclick="openUserModal(this.dataset.id)"><i class="fa-solid fa-pen"></i> Edit</button> ' +
      '<button class="btn btn-danger btn-sm" data-id="' + esc(u.id) + '" onclick="deleteUser(this.dataset.id)"><i class="fa-solid fa-trash"></i></button></td></tr>';
  }).join('') : '<tr class="empty-row"><td colspan="6">No users found</td></tr>';
}
function openUserModal(id) {
  var u = (App.cache.users || []).find(function(x) { return String(x.id) === String(id); });
  $('um_id').value = id || '';
  $('um_name').value = u ? u.name : '';
  $('um_email').value = u ? u.email : '';
  $('um_password').value = u ? u.password : '';
  $('um_phone').value = u ? u.phone : '';
  $('um_role').value = u ? u.role : 'ENGINEER';
  $('um_title').textContent = u ? 'Edit User' : 'Add User';
  $('userModal').classList.remove('hidden');
}
function closeUserModal() { $('userModal').classList.add('hidden'); }
function saveUserData(btn) {
  var p = {
    id: $('um_id').value,
    name: $('um_name').value.trim(),
    email: $('um_email').value.trim(),
    password: $('um_password').value.trim(),
    phone: $('um_phone').value.trim(),
    role: $('um_role').value,
    roleCheck: App.user.role
  };
  if (!p.name || !p.email || !p.password) return toast('Please fill all required fields', 'error');
  if (!lockBtn(btn)) return;
  showLoader(true);
  callServer('saveUser', [p], function(res) {
    showLoader(false); unlockBtn(btn);
    if (!res.success) return toast(res.message, 'error');
    toast('User saved successfully', 'success');
    App.cache.users = null;
    App.cache.engineers = null;
    closeUserModal();
    onShow_users();
  }, function() { showLoader(false); unlockBtn(btn); });
}
function deleteUser(id) {
  if (!confirm('Are you sure you want to delete this user?')) return;
  showLoader(true);
  callServer('deleteUser', [id, App.user.role], function(res) {
    showLoader(false);
    if (!res.success) return toast(res.message, 'error');
    toast('User deleted', 'success');
    App.cache.users = null;
    onShow_users();
  });
}
</script>
</body>
</html>`;
}