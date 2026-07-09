// ══════════════════════════════════════════════════════════
// КОНСТАНТЫ
// ══════════════════════════════════════════════════════════
const SPREADSHEET_ID = '1D0qPA-d5pRfLrWkDjZDawdkGYeDNzw_7_9cN4UZ6PdQ';
const SHEET_STORAGE  = 'storage';
const SHEET_USERS    = 'users';

// ══════════════════════════════════════════════════════════
// ТОЧКА ВХОДА - отдаёт HTML пользователю
// ══════════════════════════════════════════════════════════
function doGet() {
  try {
    setupSheets();
  } catch (e) {
    Logger.log('Ошибка при инициализации листов: ' + e.message);
  }
  
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('MK Tours — Реестр и калькулятор')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

// ══════════════════════════════════════════════════════════
// АВТОРИЗАЦИЯ
// ══════════════════════════════════════════════════════════
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Пользователь не авторизован');
  const user = registerOrUpdateUser_(email);
  writeAuditLog_('user', user.user_id, 'login', null, null, null, JSON.stringify({ email, role: user.role }), email, new Date());
  return { email: user.email, role: user.role, is_active: user.is_active };
}

function _registerUser(email) {
  return registerOrUpdateUser_(email);
}

function registerOrUpdateUser_(email) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ensureSheet_(ss, SHEET_USERS, SHEETS_STRUCTURE.users);
  normalizeUsersSheet_(sheet);
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('email');
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()) {
      sheet.getRange(i + 1, headers.indexOf('last_seen_at') + 1).setValue(now);
      const user = rowToObject_(headers, sheet.getRange(i + 1, 1, 1, headers.length).getValues()[0]);
      if (user.is_active === '' || user.is_active === null) user.is_active = true;
      return user;
    }
  }

  const row = [Utilities.getUuid(), email, 'viewer', now, now, true];
  sheet.appendRow(row);
  return rowToObject_(SHEETS_STRUCTURE.users, row);
}

function normalizeUsersSheet_(sheet) {
  const headers = SHEETS_STRUCTURE.users;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];
  const hasNewSchema = currentHeaders.indexOf('email') === 1 && currentHeaders.indexOf('role') === 2;
  if (!hasNewSchema) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), headers.length)).getValues();
  const normalized = rows.map(row => {
    const first = String(row[0] || '');
    const second = String(row[1] || '');
    if (first.indexOf('@') > -1 && second.indexOf('@') === -1) {
      return [Utilities.getUuid(), first, 'viewer', row[1] || new Date(), row[2] || new Date(), true];
    }
    return [
      row[0] || Utilities.getUuid(),
      row[1] || '',
      row[2] || 'viewer',
      row[3] || new Date(),
      row[4] || new Date(),
      row[5] === '' ? true : row[5]
    ];
  });
  sheet.getRange(2, 1, normalized.length, headers.length).setValues(normalized);
}

function getUserRole_(email) {
  const user = registerOrUpdateUser_(email);
  if (!user.is_active) throw new Error('Пользователь отключён');
  return user.role || 'viewer';
}

function requireRole_(allowedRoles) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Не авторизован');
  const role = getUserRole_(email);
  if (allowedRoles.indexOf(role) === -1) {
    throw new Error('Недостаточно прав. Требуется роль: ' + allowedRoles.join(', '));
  }
  return { email, role };
}

// ══════════════════════════════════════════════════════════
// ХРАНИЛИЩЕ - замена window.storage
// ══════════════════════════════════════════════════════════

// Аналог window.storage.set(key, value)
function storageSet(key, valueJson) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Не авторизован');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ensureSheet_(ss, SHEET_STORAGE, ['email', 'key', 'value_json', 'updated_at']);
  const data  = sheet.getDataRange().getValues();
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][1] === key) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 3, 1, 2).setValues([[valueJson, new Date()]]);
    } else {
      sheet.appendRow([email, key, valueJson, new Date()]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Аналог window.storage.get(key)
function storageGet(key) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Не авторизован');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ensureSheet_(ss, SHEET_STORAGE, ['email', 'key', 'value_json', 'updated_at']);
  const data  = sheet.getDataRange().getValues();

  let lastValue = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1] === key) {
      lastValue = data[i][2];
    }
  }
  if (lastValue === null) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === key) lastValue = data[i][2];
    }
  }
  return lastValue === null ? null : { value: lastValue };
}

// Загрузить все общие данные за один запрос
function storageGetAll() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Не авторизован');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ensureSheet_(ss, SHEET_STORAGE, ['email', 'key', 'value_json', 'updated_at']);
  const data  = sheet.getDataRange().getValues();

  const result = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1]) {
      result[data[i][1]] = data[i][2];
    }
  }
  return result;
}

// ══════════════════════════════════════════════════════════
// НОВАЯ МОДЕЛЬ ДАННЫХ - Листы для MVP
// ══════════════════════════════════════════════════════════

const SHEETS_STRUCTURE = {
  tours: ['tour_id', 'tour_code', 'title', 'date_start', 'date_end', 'direction', 'guide', 'currency', 'status', 'version', 'created_by', 'created_at', 'updated_by', 'updated_at', 'archived_by', 'archived_at'],
  tour_calculations: ['calculation_id', 'tour_id', 'version', 'input_json', 'result_json', 'created_by', 'created_at'],
  audit_log: ['event_id', 'entity_type', 'entity_id', 'action', 'field_name', 'old_value', 'new_value', 'snapshot_json', 'user_email', 'created_at'],
  exports: ['export_id', 'tour_id', 'tour_code', 'file_name', 'file_url', 'exported_by', 'exported_at'],
  settings: ['key', 'value', 'updated_at'],
  users: ['user_id', 'email', 'role', 'created_at', 'last_seen_at', 'is_active']
};

function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  for (const [sheetName, headers] of Object.entries(SHEETS_STRUCTURE)) {
    ensureSheet_(ss, sheetName, headers);
  }
  
  Logger.log('Sheets initialized');
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  return sheet;
}

function getSheet_(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(name);
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, idx) => {
    obj[h] = row[idx];
  });
  return obj;
}

// ── Туры ──

function listTours(filters = {}) {
  requireRole_(['admin', 'editor', 'viewer']);
  const sheet = getSheet_('tours');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tours = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tour = rowToObject_(headers, row);
    
    if (tour.status !== 'archived' || filters.includeArchived) {
      tours.push(tour);
    }
  }
  
  return tours;
}

function getTour(tourId) {
  requireRole_(['admin', 'editor', 'viewer']);
  const result = getTour_(tourId);
  return result ? result.tour : null;
}

function getTour_(tourId) {
  const sheet = getSheet_('tours');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === tourId) {
      const tour = rowToObject_(headers, row);
      return { tour, rowIndex: i + 1 };
    }
  }
  return null;
}

function createTour(payload) {
  const { email } = requireRole_(['admin', 'editor']);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  
  try {
    const sheet = getSheet_('tours');
    const tourId = Utilities.getUuid();
    const now = new Date();
    
    const row = [
      tourId,
      payload.tour_code || '',
      payload.title || '',
      payload.date_start || '',
      payload.date_end || '',
      payload.direction || '',
      payload.guide || '',
      payload.currency || 'EUR',
      payload.status || 'draft',
      1,
      email,
      now,
      email,
      now,
      '',
      ''
    ];
    
    sheet.appendRow(row);
    writeAuditLog_('tour', tourId, 'create', null, null, null, JSON.stringify(payload), email, now);
    
    return { ok: true, tourId, version: 1 };
  } finally {
    lock.releaseLock();
  }
}

function updateTour(tourId, payload, expectedVersion) {
  const { email } = requireRole_(['admin', 'editor']);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  
  try {
    const result = getTour_(tourId);
    if (!result) throw new Error('Тур не найден');
    
    const { tour, rowIndex } = result;
    
    if (Number(tour.version) !== Number(expectedVersion)) {
      throw new Error('Тур был изменён другим пользователем. Обновите данные перед сохранением.');
    }
    
    const sheet = getSheet_('tours');
    const now = new Date();
    const nextVersion = Number(tour.version) + 1;
    
    const updates = {
      ...tour,
      ...payload,
      version: nextVersion,
      updated_by: email,
      updated_at: now
    };
    
    const headers = SHEETS_STRUCTURE.tours;
    const newRow = headers.map(h => updates[h] ?? '');
    
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([newRow]);
    writeAuditLog_('tour', tourId, 'update', null, null, JSON.stringify(tour), JSON.stringify(updates), email, now);
    
    return { ok: true, version: nextVersion };
  } finally {
    lock.releaseLock();
  }
}

function archiveTour(tourId, expectedVersion) {
  const { email } = requireRole_(['admin']);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  
  try {
    const result = getTour_(tourId);
    if (!result) throw new Error('Тур не найден');
    
    const { tour, rowIndex } = result;
    
    if (Number(tour.version) !== Number(expectedVersion)) {
      throw new Error('Версия не совпадает');
    }
    
    const sheet = getSheet_('tours');
    const now = new Date();
    
    const nextVersion = Number(tour.version) + 1;
    sheet.getRange(rowIndex, 9).setValue('archived');
    sheet.getRange(rowIndex, 10).setValue(nextVersion);
    sheet.getRange(rowIndex, 13).setValue(email);
    sheet.getRange(rowIndex, 14).setValue(now);
    sheet.getRange(rowIndex, 15).setValue(email);
    sheet.getRange(rowIndex, 16).setValue(now);
    
    writeAuditLog_('tour', tourId, 'archive', null, null, null, null, email, now);
    
    return { ok: true, version: nextVersion };
  } finally {
    lock.releaseLock();
  }
}

function restoreTour(tourId, expectedVersion) {
  const { email } = requireRole_(['admin']);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const result = getTour_(tourId);
    if (!result) throw new Error('Тур не найден');
    const { tour, rowIndex } = result;
    if (Number(tour.version) !== Number(expectedVersion)) {
      throw new Error('Версия не совпадает');
    }

    const sheet = getSheet_('tours');
    const now = new Date();
    const nextVersion = Number(tour.version) + 1;
    sheet.getRange(rowIndex, 9).setValue('active');
    sheet.getRange(rowIndex, 10).setValue(nextVersion);
    sheet.getRange(rowIndex, 13).setValue(email);
    sheet.getRange(rowIndex, 14).setValue(now);
    sheet.getRange(rowIndex, 15, 1, 2).setValues([['', '']]);
    writeAuditLog_('tour', tourId, 'restore', null, null, null, JSON.stringify(tour), email, now);
    return { ok: true, version: nextVersion };
  } finally {
    lock.releaseLock();
  }
}

function duplicateTour(tourId) {
  const { email } = requireRole_(['admin', 'editor']);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const result = getTour_(tourId);
    if (!result) throw new Error('Тур не найден');
    const source = result.tour;
    const sheet = getSheet_('tours');
    const now = new Date();
    const newTourId = Utilities.getUuid();
    const row = [
      newTourId,
      (source.tour_code || '') + '-copy',
      (source.title || 'Тур') + ' — копия',
      source.date_start || '',
      source.date_end || '',
      source.direction || '',
      source.guide || '',
      source.currency || 'EUR',
      'draft',
      1,
      email,
      now,
      email,
      now,
      '',
      ''
    ];
    sheet.appendRow(row);

    const latestCalc = getLatestCalculation(tourId);
    if (latestCalc) {
      const calcSheet = getSheet_('tour_calculations');
      calcSheet.appendRow([
        Utilities.getUuid(),
        newTourId,
        1,
        JSON.stringify(latestCalc.input || {}),
        JSON.stringify(latestCalc.result || {}),
        email,
        now
      ]);
    }

    writeAuditLog_('tour', newTourId, 'create', null, null, null, JSON.stringify({ duplicated_from: tourId }), email, now);
    return { ok: true, tourId: newTourId, version: 1 };
  } finally {
    lock.releaseLock();
  }
}
// ── Расчёты ──

function saveTourCalculation(tourId, input, result, expectedVersion) {
  const { email } = requireRole_(['admin', 'editor']);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  
  try {
    const tourResult = getTour_(tourId);
    if (!tourResult) throw new Error('Тур не найден');
    
    const { tour } = tourResult;
    
    if (Number(tour.version) !== Number(expectedVersion)) {
      throw new Error('Тур был изменён');
    }
    
    const sheet = getSheet_('tour_calculations');
    const calcId = Utilities.getUuid();
    const now = new Date();
    const calcVersion = getNextCalculationVersion_(tourId);
    
    sheet.appendRow([
      calcId,
      tourId,
      calcVersion,
      JSON.stringify(input),
      JSON.stringify(result),
      email,
      now
    ]);
    
    writeAuditLog_('calculation', tourId, 'calculation_save', null, null, null, JSON.stringify({ calculation_id: calcId, version: calcVersion, result }), email, now);
    
    return { ok: true, calcId, calculationVersion: calcVersion, tourVersion: Number(tour.version) };
  } finally {
    lock.releaseLock();
  }
}

function getNextCalculationVersion_(tourId) {
  const sheet = getSheet_('tour_calculations');
  const data = sheet.getDataRange().getValues();
  let maxVersion = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === tourId) maxVersion = Math.max(maxVersion, Number(data[i][2]) || 0);
  }
  return maxVersion + 1;
}

function getLatestCalculation(tourId) {
  requireRole_(['admin', 'editor', 'viewer']);
  const sheet = getSheet_('tour_calculations');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  let latest = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === tourId) {
      if (!latest || data[i][2] > latest[2]) {
        latest = data[i];
      }
    }
  }
  
  if (latest) {
    return {
      calculation_id: latest[0],
      tour_id: latest[1],
      version: latest[2],
      input: JSON.parse(latest[3]),
      result: JSON.parse(latest[4]),
      created_by: latest[5],
      created_at: latest[6]
    };
  }
  
  return null;
}

// ── История ──

function writeAuditLog_(entityType, entityId, action, fieldName, oldValue, newValue, snapshotJson, userEmail, createdAt) {
  const sheet = getSheet_('audit_log');
  if (!sheet) return;
  
  const eventId = Utilities.getUuid();
  sheet.appendRow([
    eventId,
    entityType,
    entityId,
    action,
    fieldName || '',
    oldValue || '',
    newValue || '',
    snapshotJson || '',
    userEmail,
    createdAt || new Date()
  ]);
}

function getTourHistory(tourId) {
  requireRole_(['admin', 'editor', 'viewer']);
  const sheet = getSheet_('audit_log');
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const history = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === tourId) {
      history.push({
        event_id: data[i][0],
        entity_type: data[i][1],
        entity_id: data[i][2],
        action: data[i][3],
        field_name: data[i][4],
        old_value: data[i][5],
        new_value: data[i][6],
        snapshot_json: data[i][7],
        user_email: data[i][8],
        created_at: data[i][9]
      });
    }
  }
  
  return history;
}

function exportTourToExcel(tourId) {
  const { email } = requireRole_(['admin', 'editor', 'viewer']);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const tourResult = getTour_(tourId);
    if (!tourResult) throw new Error('Тур не найден');

    const tour = tourResult.tour;
    const calc = getLatestCalculation(tourId);
    const history = getTourHistory(tourId);
    const now = new Date();
    const fileName = sanitizeFileName_((tour.tour_code || tour.title || 'tour') + '_' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm')) + '.xlsx';

    const temp = SpreadsheetApp.create(fileName.replace(/\.xlsx$/i, ''));
    const infoSheet = temp.getSheets()[0];
    infoSheet.setName('Тур');
    infoSheet.getRange(1, 1, 10, 2).setValues([
      ['Код тура', tour.tour_code || ''],
      ['Название', tour.title || ''],
      ['Даты', (tour.date_start || '') + ' - ' + (tour.date_end || '')],
      ['Направление', tour.direction || ''],
      ['Гид', tour.guide || ''],
      ['Валюта', tour.currency || ''],
      ['Статус', tour.status || ''],
      ['Автор', tour.created_by || ''],
      ['Дата создания', tour.created_at || ''],
      ['Дата изменения', tour.updated_at || '']
    ]);

    writeJsonSheet_(temp.insertSheet('Расчёт'), calc ? calc.input : {});
    writeJsonSheet_(temp.insertSheet('Итоги'), calc ? calc.result : {});
    const historySheet = temp.insertSheet('История');
    const historyHeaders = SHEETS_STRUCTURE.audit_log;
    historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]);
    if (history.length) {
      historySheet.getRange(2, 1, history.length, historyHeaders.length).setValues(
        history.map(event => historyHeaders.map(h => event[h] || ''))
      );
    }

    SpreadsheetApp.flush();
    const xlsxBlob = DriveApp.getFileById(temp.getId()).getBlob().getAs(MimeType.MICROSOFT_EXCEL).setName(fileName);
    const xlsxFile = DriveApp.createFile(xlsxBlob);
    DriveApp.getFileById(temp.getId()).setTrashed(true);

    const exportId = Utilities.getUuid();
    getSheet_('exports').appendRow([exportId, tourId, tour.tour_code || '', fileName, xlsxFile.getUrl(), email, now]);
    writeAuditLog_('export', tourId, 'export', null, null, null, JSON.stringify({ export_id: exportId, file_name: fileName, file_url: xlsxFile.getUrl() }), email, now);

    return { ok: true, exportId, fileName, fileUrl: xlsxFile.getUrl() };
  } finally {
    lock.releaseLock();
  }
}

function writeJsonSheet_(sheet, obj) {
  const rows = flattenObject_(obj || {}).map(pair => [pair[0], pair[1]]);
  sheet.getRange(1, 1, 1, 2).setValues([['Параметр', 'Значение']]);
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function flattenObject_(obj, prefix) {
  const rows = [];
  Object.keys(obj || {}).forEach(key => {
    const value = obj[key];
    const path = prefix ? prefix + '.' + key : key;
    if (value && typeof value === 'object' && !Array.isArray(value) && !(value instanceof Date)) {
      rows.push.apply(rows, flattenObject_(value, path));
    } else {
      rows.push([path, typeof value === 'object' ? JSON.stringify(value) : value]);
    }
  });
  return rows;
}

function sanitizeFileName_(name) {
  return String(name || 'tour').replace(/[\\/:*?"<>|]/g, '_').substring(0, 120);
}

