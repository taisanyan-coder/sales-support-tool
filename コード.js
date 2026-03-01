const APP_SPREADSHEET_ID = '1WaV0wT57jeOOqb8TaXlzE-KoE2FXWGq6d4UKlyxmlgg';
const TZ = 'Asia/Tokyo';
const CATEGORIES = ['契約・請求', '営業・トラブル', 'その他'];
const STATUSES = ['未対応', '対応中', '完了'];
const LINK_SHEET_NAME = 'LINK';

function doGet(e) {
  const page = e && e.parameter ? e.parameter.page : '';
  if (page === 'utilities') {
    return HtmlService.createHtmlOutputFromFile('utilities');
  }
  if (page === 'nameFormatter') {
    return HtmlService.createHtmlOutputFromFile('nameFormatter');
  }
  return HtmlService.createHtmlOutputFromFile('index');
}

function getLinks() {
  try {
    var ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(LINK_SHEET_NAME);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol <= 1) return [];

    var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var map = {};
    for (var i = 0; i < header.length; i++) {
      var h = String(header[i] || '').trim();
      if (h) map[h] = i;
    }

    if (map.label == null || map.url == null || map.order == null || map.enabled == null) {
      return [];
    }

    var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var rows = [];

    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      if (row[map.enabled] !== true) continue;

      var label = String(row[map.label] || '').trim();
      var url = String(row[map.url] || '').trim();
      if (!label || !url) continue;

      var ord = Number(row[map.order]);
      if (!isFinite(ord)) ord = 999999;

      rows.push({ label: label, url: url, order: ord });
    }

    rows.sort(function(a, b) {
      if (a.order !== b.order) return a.order - b.order;
      return String(a.label).localeCompare(String(b.label));
    });

    return rows.map(function(x) { return { label: x.label, url: x.url }; });
  } catch (e) {
    return [];
  }
}

function initPage() {
  assertSheetsAndColumns_();
  return {
    links: getLinks(),
    companies: loadCompanies_(),
    categories: CATEGORIES.slice(),
    statuses: STATUSES.slice(),
    actions: listActions(),
    today_ymd: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')
  };
}

function listActions() {
  assertSheetsAndColumns_();
  const ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Actions');
  const colMap = buildColumnIndexMap_(sheet);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const actions = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[colMap.is_deleted - 1] === true) continue;

    actions.push({
      action_id: String(row[colMap.action_id - 1] || '').trim(),
      created_at: String(row[colMap.created_at - 1] || '').trim(),
      updated_at: String(row[colMap.updated_at - 1] || '').trim(),
      due_date: normalizeDate_(row[colMap.due_date - 1]),
      company_name: String(row[colMap.company_name - 1] || '').trim(),
      staff_name: String(row[colMap.staff_name - 1] || '').trim(),
      category: String(row[colMap.category - 1] || '').trim(),
      status: String(row[colMap.status - 1] || '').trim(),
      note: String(row[colMap.note - 1] || '').trim(),
      completed_at: String(row[colMap.completed_at - 1] || '').trim()
    });
  }

  actions.sort(function(a, b) {
    const ad = a.due_date || '9999-12-31';
    const bd = b.due_date || '9999-12-31';
    if (ad < bd) return -1;
    if (ad > bd) return 1;
    const ac = a.created_at || '';
    const bc = b.created_at || '';
    if (ac < bc) return -1;
    if (ac > bc) return 1;
    return 0;
  });

  return actions;
}

function createAction(payload) {
  assertSheetsAndColumns_();
  const p = payload || {};
  const companyName = String(p.company_name || '').trim();
  const staffName = String(p.staff_name || '').trim();
  const category = String(p.category || '').trim();
  const status = String(p.status || '').trim() || '未対応';
  const note = String(p.note || '').trim();
  const dueRaw = p.due_date;

  if (!companyName) throw new Error('company_name は必須です。');
  if (dueRaw === null || dueRaw === undefined || String(dueRaw).trim() === '') throw new Error('due_date は必須です。');
  if (!note) throw new Error('note は必須です。');
  if (CATEGORIES.indexOf(category) === -1) throw new Error('INVALID_CATEGORY: ' + category);
  if (STATUSES.indexOf(status) === -1) throw new Error('INVALID_STATUS: ' + status);

  const dueDate = normalizeDateToDate_(dueRaw);
  const now = nowJst_();
  const completedAt = status === '完了' ? now : '';

  const ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Actions');
  const colMap = buildColumnIndexMap_(sheet);
  const row = new Array(sheet.getLastColumn()).fill('');

  row[colMap.action_id - 1] = generateActionId_();
  row[colMap.created_at - 1] = now;
  row[colMap.updated_at - 1] = now;
  row[colMap.due_date - 1] = dueDate;
  row[colMap.company_name - 1] = companyName;
  row[colMap.staff_name - 1] = staffName;
  row[colMap.category - 1] = category;
  row[colMap.status - 1] = status;
  row[colMap.note - 1] = note;
  row[colMap.completed_at - 1] = completedAt;
  row[colMap.is_deleted - 1] = false;
  row[colMap.deleted_at - 1] = '';

  sheet.appendRow(row);
  return listActions();
}

function updateAction(actionId, patch) {
  assertSheetsAndColumns_();
  const id = String(actionId || '').trim();
  if (!id) throw new Error('actionId は必須です。');
  if (!patch || Object.keys(patch).length === 0) throw new Error('patch は必須です。');

  const ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Actions');
  const colMap = buildColumnIndexMap_(sheet);
  const rowIndex = findActionRow_(sheet, colMap, id);
  if (rowIndex === -1) throw new Error('ACTION_NOT_FOUND: ' + id);

  const row = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (row[colMap.is_deleted - 1] === true) throw new Error('ACTION_DELETED: ' + id);

  const now = nowJst_();
  const currentStatus = String(row[colMap.status - 1] || '').trim();

  if (Object.prototype.hasOwnProperty.call(patch, 'company_name')) {
    sheet.getRange(rowIndex, colMap.company_name).setValue(String(patch.company_name || '').trim());
  }
  if (Object.prototype.hasOwnProperty.call(patch, 'staff_name')) {
    sheet.getRange(rowIndex, colMap.staff_name).setValue(String(patch.staff_name || '').trim());
  }
  if (Object.prototype.hasOwnProperty.call(patch, 'category')) {
    const category = String(patch.category || '').trim();
    if (CATEGORIES.indexOf(category) === -1) throw new Error('INVALID_CATEGORY: ' + category);
    sheet.getRange(rowIndex, colMap.category).setValue(category);
  }
  if (Object.prototype.hasOwnProperty.call(patch, 'due_date')) {
    if (patch.due_date === null || patch.due_date === undefined || String(patch.due_date).trim() === '') throw new Error('due_date は必須です。');
    sheet.getRange(rowIndex, colMap.due_date).setValue(normalizeDateToDate_(patch.due_date));
  }
  if (Object.prototype.hasOwnProperty.call(patch, 'note')) {
    const note = String(patch.note || '').trim();
    if (!note) throw new Error('note は必須です。');
    sheet.getRange(rowIndex, colMap.note).setValue(note);
  }
  if (Object.prototype.hasOwnProperty.call(patch, 'status')) {
    const nextStatus = String(patch.status || '').trim();
    if (STATUSES.indexOf(nextStatus) === -1) throw new Error('INVALID_STATUS: ' + nextStatus);
    sheet.getRange(rowIndex, colMap.status).setValue(nextStatus);

    if (currentStatus !== '完了' && nextStatus === '完了') {
      sheet.getRange(rowIndex, colMap.completed_at).setValue(now);
    } else if (currentStatus === '完了' && nextStatus !== '完了') {
      sheet.getRange(rowIndex, colMap.completed_at).setValue('');
    }
  }

  sheet.getRange(rowIndex, colMap.updated_at).setValue(now);
  return listActions();
}

function deleteAction(actionId) {
  assertSheetsAndColumns_();
  const id = String(actionId || '').trim();
  if (!id) throw new Error('actionId は必須です。');

  const ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Actions');
  const colMap = buildColumnIndexMap_(sheet);
  const rowIndex = findActionRow_(sheet, colMap, id);
  if (rowIndex === -1) throw new Error('ACTION_NOT_FOUND: ' + id);

  const now = nowJst_();
  sheet.getRange(rowIndex, colMap.is_deleted).setValue(true);
  sheet.getRange(rowIndex, colMap.deleted_at).setValue(now);
  sheet.getRange(rowIndex, colMap.updated_at).setValue(now);
  return listActions();
}

function getCompanyContacts(companyName) {
  assertSheetsAndColumns_();
  const result = {
    contact_contract_billing: '',
    contact_sales_trouble: ''
  };

  const name = String(companyName || '').trim();
  if (!name) return result;

  const ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Companies');
  const colMap = buildColumnIndexMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const company = String(row[colMap.company_name - 1] || '').trim();
    if (company === name) {
      result.contact_contract_billing = String(row[colMap.contact_contract_billing - 1] || '').trim();
      result.contact_sales_trouble = String(row[colMap.contact_sales_trouble - 1] || '').trim();
      return result;
    }
  }
  return result;
}

function assertSheetsAndColumns_() {
  const ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
  const companies = ss.getSheetByName('Companies');
  const actions = ss.getSheetByName('Actions');

  if (!companies) throw new Error('SHEET_MISSING: Companies');
  if (!actions) throw new Error('SHEET_MISSING: Actions');

  const requiredCompanies = ['company_id', 'company_name', 'contact_contract_billing', 'contact_sales_trouble', 'memo_company'];
  const requiredActions = ['action_id', 'created_at', 'updated_at', 'due_date', 'company_name', 'staff_name', 'category', 'status', 'note', 'completed_at', 'is_deleted', 'deleted_at'];

  validateHeaders_(companies, 'Companies', requiredCompanies);
  validateHeaders_(actions, 'Actions', requiredActions);
}

function buildColumnIndexMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = lastColumn > 0 ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
  const map = {};

  for (let i = 0; i < headers.length; i++) {
    const key = String(headers[i] || '').trim();
    if (!key) continue;
    if (Object.prototype.hasOwnProperty.call(map, key)) {
      throw new Error('DUPLICATE_HEADER: ' + key);
    }
    map[key] = i + 1;
  }
  return map;
}

function findActionRow_(sheet, colMap, actionId) {
  const id = String(actionId || '').trim();
  if (!id) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  const values = sheet.getRange(2, colMap.action_id, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      return i + 2;
    }
  }
  return -1;
}

function normalizeDate_(value) {
  try {
    if (value instanceof Date && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
    }
    const s = String(value || '').trim();
    if (!s) return '';
    const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    return '';
  } catch (e) {
    return '';
  }
}

function normalizeDateToDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const s = String(value || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('due_date の形式が不正です。');

  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const dt = new Date(y, mo - 1, d);

  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) {
    throw new Error('due_date の形式が不正です。');
  }
  return dt;
}

function nowJst_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function generateActionId_() {
  const ts = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
  const rand = ('0000' + Math.floor(Math.random() * 10000)).slice(-4);
  return 'A_' + ts + '_' + rand;
}

function loadCompanies_() {
  const ss = SpreadsheetApp.openById(APP_SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Companies');
  const colMap = buildColumnIndexMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, colMap.company_name, lastRow - 1, 1).getValues();
  const set = {};
  for (let i = 0; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    if (name) set[name] = true;
  }
  return Object.keys(set).sort();
}

function validateHeaders_(sheet, sheetName, requiredColumns) {
  const map = buildColumnIndexMap_(sheet);
  for (let i = 0; i < requiredColumns.length; i++) {
    const col = requiredColumns[i];
    if (!Object.prototype.hasOwnProperty.call(map, col)) {
      throw new Error('COLUMNS_MISSING: ' + sheetName + ': ' + col);
    }
  }
}