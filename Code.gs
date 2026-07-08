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
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('MK Tours Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ══════════════════════════════════════════════════════════
// АВТОРИЗАЦИЯ
// ══════════════════════════════════════════════════════════
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Пользователь не авторизован');
  _registerUser(email);
  return { email };
}

function _registerUser(email) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_USERS);
  const data  = sheet.getDataRange().getValues();
  const found = data.slice(1).some(row => row[0] === email);
  if (!found) {
    sheet.appendRow([email, new Date(), new Date()]);
  } else {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.getRange(i + 1, 3).setValue(new Date());
        break;
      }
    }
  }
}

// ══════════════════════════════════════════════════════════
// ХРАНИЛИЩЕ - замена window.storage
// ══════════════════════════════════════════════════════════

// Аналог window.storage.set(key, value)
function storageSet(key, valueJson) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Не авторизован');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STORAGE);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1] === key) {
      sheet.getRange(i + 1, 3).setValue(valueJson);
      return { ok: true };
    }
  }
  sheet.appendRow([email, key, valueJson]);
  return { ok: true };
}

// Аналог window.storage.get(key)
function storageGet(key) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Не авторизован');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STORAGE);
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email && data[i][1] === key) {
      return { value: data[i][2] };
    }
  }
  return null;
}

// Загрузить все данные пользователя за один запрос
function storageGetAll() {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Не авторизован');

  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_STORAGE);
  const data  = sheet.getDataRange().getValues();

  const result = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      result[data[i][1]] = data[i][2];
    }
  }
  return result;
}
