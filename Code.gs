/**
 * ===== POS Web App Backend (Google Apps Script) =====
 * วิธีติดตั้ง: ดูไฟล์ "คู่มือติดตั้ง.md" ที่แนบมาด้วย
 *
 * สคริปต์นี้เปลี่ยน Google Sheet ให้เป็นฐานข้อมูล + API สำหรับเว็บแอป POS
 * โดยจะสร้างชีตที่จำเป็นให้อัตโนมัติในครั้งแรกที่เรียกใช้งาน
 */

// ---------- ตั้งค่า ----------
// ตั้งรหัสลับเอง (ต้องตรงกับที่กรอกในหน้า "ตั้งค่า" ของเว็บแอป)
const API_TOKEN = 'change-this-secret-token';

const SHEET_NAMES = {
  CONFIG: 'Config',
  PRODUCTS: 'Products',
  USERS: 'Users',
  SALES: 'Sales'
};

// ---------- Helper ----------
function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name, headers) {
  const ss = getSS_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureSheets_() {
  getSheet_(SHEET_NAMES.CONFIG, ['key', 'value']);
  getSheet_(SHEET_NAMES.PRODUCTS, ['id', 'name', 'category', 'price', 'cost', 'stock', 'sku', 'active']);
  getSheet_(SHEET_NAMES.USERS, ['id', 'username', 'pin', 'role', 'active']);
  getSheet_(SHEET_NAMES.SALES, ['id', 'datetime', 'cashier', 'items_json', 'subtotal', 'discount', 'tax', 'total', 'payment_method']);

  // สร้างผู้ใช้ผู้ดูแลระบบเริ่มต้น ถ้ายังไม่มีผู้ใช้เลย
  const usersSheet = getSheet_(SHEET_NAMES.USERS, ['id', 'username', 'pin', 'role', 'active']);
  if (usersSheet.getLastRow() < 2) {
    usersSheet.appendRow(['U' + new Date().getTime(), 'admin', '1234', 'admin', true]);
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function checkToken_(token) {
  if (!API_TOKEN) return true;
  return token === API_TOKEN;
}

function sheetRowsToObjects_(sheet) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  const headers = rows[0];
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i].join('') === '') continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = rows[i][idx]; });
    out.push(obj);
  }
  return out;
}

// ---------- Entry points ----------
function doGet(e) {
  ensureSheets_();
  const action = e.parameter.action;
  const token = e.parameter.token;

  if (action === 'ping') return jsonOut_({ ok: true, message: 'pong' });
  if (!checkToken_(token)) return jsonOut_({ ok: false, error: 'Unauthorized: token ไม่ถูกต้อง' });

  try {
    switch (action) {
      case 'getConfig':
        return jsonOut_({ ok: true, data: getConfig_() });
      case 'getProducts':
        return jsonOut_({ ok: true, data: getProducts_() });
      case 'getUsers':
        return jsonOut_({ ok: true, data: getUsers_() });
      case 'getSales':
        return jsonOut_({ ok: true, data: getSales_(e.parameter.from, e.parameter.to) });
      default:
        return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  ensureSheets_();
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Invalid JSON body' });
  }

  const action = body.action;
  if (action !== 'login' && !checkToken_(body.token)) {
    return jsonOut_({ ok: false, error: 'Unauthorized: token ไม่ถูกต้อง' });
  }
  if (action === 'login' && !checkToken_(body.token)) {
    return jsonOut_({ ok: false, error: 'Unauthorized: token ไม่ถูกต้อง' });
  }

  try {
    switch (action) {
      case 'saveConfig':
        saveConfig_(body.data || {});
        return jsonOut_({ ok: true });
      case 'saveProduct':
        return jsonOut_({ ok: true, data: saveProduct_(body.data || {}) });
      case 'deleteProduct':
        deleteProduct_(body.id);
        return jsonOut_({ ok: true });
      case 'saveUser':
        return jsonOut_({ ok: true, data: saveUser_(body.data || {}) });
      case 'deleteUser':
        deleteUser_(body.id);
        return jsonOut_({ ok: true });
      case 'login':
        const user = login_(body.username, body.pin);
        if (!user) return jsonOut_({ ok: false, error: 'ชื่อผู้ใช้หรือ PIN ไม่ถูกต้อง' });
        return jsonOut_({ ok: true, data: user });
      case 'addSale':
        return jsonOut_({ ok: true, data: addSale_(body.data || {}) });
      default:
        return jsonOut_({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ---------- Config ----------
function getConfig_() {
  const sheet = getSheet_(SHEET_NAMES.CONFIG, ['key', 'value']);
  const rows = sheet.getDataRange().getValues();
  const obj = {};
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0]) obj[rows[i][0]] = rows[i][1];
  }
  return obj;
}

function saveConfig_(data) {
  const sheet = getSheet_(SHEET_NAMES.CONFIG, ['key', 'value']);
  const rows = sheet.getDataRange().getValues();
  const keyToRow = {};
  for (let i = 1; i < rows.length; i++) keyToRow[rows[i][0]] = i + 1;
  Object.keys(data).forEach(k => {
    const val = data[k];
    if (keyToRow[k]) {
      sheet.getRange(keyToRow[k], 2).setValue(val);
    } else {
      sheet.appendRow([k, val]);
    }
  });
}

// ---------- Products ----------
function getProducts_() {
  const sheet = getSheet_(SHEET_NAMES.PRODUCTS, ['id', 'name', 'category', 'price', 'cost', 'stock', 'sku', 'active']);
  return sheetRowsToObjects_(sheet);
}

function saveProduct_(p) {
  const sheet = getSheet_(SHEET_NAMES.PRODUCTS, ['id', 'name', 'category', 'price', 'cost', 'stock', 'sku', 'active']);
  const rows = sheet.getDataRange().getValues();
  const rowData = [
    p.id || '', p.name || '', p.category || '', Number(p.price) || 0,
    Number(p.cost) || 0, Number(p.stock) || 0, p.sku || '', p.active !== false
  ];
  if (!p.id) {
    p.id = 'P' + new Date().getTime();
    rowData[0] = p.id;
    sheet.appendRow(rowData);
  } else {
    let rowIdx = -1;
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] == p.id) { rowIdx = i + 1; break; } }
    if (rowIdx === -1) {
      sheet.appendRow(rowData);
    } else {
      sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
    }
  }
  return p;
}

function deleteProduct_(id) {
  const sheet = getSheet_(SHEET_NAMES.PRODUCTS, ['id', 'name', 'category', 'price', 'cost', 'stock', 'sku', 'active']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == id) { sheet.deleteRow(i + 1); break; }
  }
}

// ---------- Users ----------
function getUsers_() {
  const sheet = getSheet_(SHEET_NAMES.USERS, ['id', 'username', 'pin', 'role', 'active']);
  const list = sheetRowsToObjects_(sheet);
  // ไม่ส่ง PIN ออกไปตอนดึงรายชื่อผู้ใช้ทั้งหมด (เพื่อความปลอดภัย)
  return list.map(u => ({ id: u.id, username: u.username, role: u.role, active: u.active }));
}

function saveUser_(u) {
  const sheet = getSheet_(SHEET_NAMES.USERS, ['id', 'username', 'pin', 'role', 'active']);
  const rows = sheet.getDataRange().getValues();
  if (!u.id) {
    u.id = 'U' + new Date().getTime();
    sheet.appendRow([u.id, u.username || '', u.pin || '1234', u.role || 'cashier', u.active !== false]);
  } else {
    let rowIdx = -1;
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] == u.id) { rowIdx = i + 1; break; } }
    const existingPin = rowIdx > -1 ? rows[rowIdx - 1][2] : '1234';
    const pinToUse = u.pin ? u.pin : existingPin;
    const rowData = [u.id, u.username || '', pinToUse, u.role || 'cashier', u.active !== false];
    if (rowIdx === -1) {
      sheet.appendRow(rowData);
    } else {
      sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
    }
  }
  return { id: u.id, username: u.username, role: u.role, active: u.active };
}

function deleteUser_(id) {
  const sheet = getSheet_(SHEET_NAMES.USERS, ['id', 'username', 'pin', 'role', 'active']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == id) { sheet.deleteRow(i + 1); break; }
  }
}

function login_(username, pin) {
  const sheet = getSheet_(SHEET_NAMES.USERS, ['id', 'username', 'pin', 'role', 'active']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === username && String(rows[i][2]) === String(pin) && rows[i][4] !== false) {
      return { id: rows[i][0], username: rows[i][1], role: rows[i][3] };
    }
  }
  return null;
}

// ---------- Sales ----------
function getSales_(from, to) {
  const sheet = getSheet_(SHEET_NAMES.SALES, ['id', 'datetime', 'cashier', 'items_json', 'subtotal', 'discount', 'tax', 'total', 'payment_method']);
  let list = sheetRowsToObjects_(sheet);
  if (from) list = list.filter(s => s.datetime >= from);
  if (to) list = list.filter(s => s.datetime <= to);
  return list;
}

function addSale_(s) {
  const sheet = getSheet_(SHEET_NAMES.SALES, ['id', 'datetime', 'cashier', 'items_json', 'subtotal', 'discount', 'tax', 'total', 'payment_method']);
  const id = 'S' + new Date().getTime();
  const dt = new Date();
  sheet.appendRow([
    id, dt.toISOString(), s.cashier || '', JSON.stringify(s.items || []),
    Number(s.subtotal) || 0, Number(s.discount) || 0, Number(s.tax) || 0,
    Number(s.total) || 0, s.payment_method || 'cash'
  ]);

  // ตัดสต๊อกสินค้าอัตโนมัติ
  if (s.items && s.items.length) {
    const psheet = getSheet_(SHEET_NAMES.PRODUCTS, ['id', 'name', 'category', 'price', 'cost', 'stock', 'sku', 'active']);
    const prows = psheet.getDataRange().getValues();
    s.items.forEach(item => {
      for (let i = 1; i < prows.length; i++) {
        if (prows[i][0] == item.id) {
          const newStock = (Number(prows[i][5]) || 0) - (Number(item.qty) || 0);
          psheet.getRange(i + 1, 6).setValue(newStock);
          break;
        }
      }
    });
  }
  return { id: id, datetime: dt.toISOString() };
}
