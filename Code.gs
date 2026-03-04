/**
 * =========================
 * Patties Bill - Code.gs
 * =========================
 * ต้องตั้งค่า SPREADSHEET_ID ก่อนใช้งาน
 */

const CONFIG = {
  SPREADSHEET_ID: 'PUT_YOUR_SHEET_ID_HERE', // <-- ใส่ Spreadsheet ID
  SHEET_INVOICES: 'Invoices',
  SHEET_SETTINGS: 'Settings',
  DRIVE_FOLDER_NAME: 'Patties Bill Documents',
};

// ---------- Web App Entry ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Patties Bill')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // เพื่อให้ iframe ได้
}

// ---------- Router ----------
function apiRequest(action, payload) {
  try {
    payload = payload || {};
    switch (action) {
      case 'checkAuth': return json_(checkAuth_());
      case 'read': return json_(readInvoices_());
      case 'create': return json_(createInvoice_(payload));
      case 'update': return json_(updateInvoice_(payload));
      case 'delete': return json_(deleteInvoice_(payload));
      case 'saveSettings': return json_(saveSettings_(payload));
      case 'listDriveFiles': return json_(listDriveFiles_(payload));
      case 'deleteDriveFile': return json_(deleteDriveFile_(payload));
      case 'savePdf': return json_(savePdf_(payload));
      default:
        return json_({ status: 'error', message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return json_({ status: 'error', message: err && err.message ? err.message : String(err) });
  }
}

function json_(obj) {
  return JSON.stringify(obj);
}

// ---------- Auth & Settings ----------
function checkAuth_() {
  const ownerEmail = Session.getEffectiveUser().getEmail(); // เจ้าของสคริปต์
  const userEmail = (Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';

  const settings = getSettings_(); // {folderId, allowedEmails}
  const allowed = parseAllowedEmails_(settings.allowedEmails);

  // owner เข้าได้เสมอ
  if (userEmail && userEmail.toLowerCase() === ownerEmail.toLowerCase()) {
    ensureDriveFolder_(settings);
    return { status: 'success', userEmail, isOwner: true, settings };
  }

  // ถ้าไม่มีอีเมล userEmail แปลว่าดีพลอยตั้ง Execute as ไม่ใช่ "User accessing..."
  // ยังให้เข้าได้ แต่จะตรวจสิทธิ์แบบอีเมลไม่ได้ (แจ้งไปที่หน้า UI)
  if (!userEmail) {
    ensureDriveFolder_(settings);
    return {
      status: 'success',
      userEmail: '',
      isOwner: false,
      settings,
      message: 'ไม่พบอีเมลผู้ใช้ (โปรดตั้ง Deploy เป็น "User accessing the web app")'
    };
  }

  // ถ้า allowedEmails ว่าง = public (ใครมีลิงก์และล็อกอินก็เข้าได้)
  if (allowed.length === 0) {
    ensureDriveFolder_(settings);
    return { status: 'success', userEmail, isOwner: false, settings };
  }

  const ok = allowed.includes(userEmail.toLowerCase());
  if (!ok) {
    return { status: 'error', userEmail, message: `บัญชี ${userEmail} ไม่ได้รับอนุญาต` };
  }

  ensureDriveFolder_(settings);
  return { status: 'success', userEmail, isOwner: false, settings };
}

function parseAllowedEmails_(str) {
  if (!str) return [];
  return str.split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean);
}

function getSettings_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_SETTINGS, ['Key', 'Value']);

  const values = sh.getDataRange().getValues();
  const map = {};
  for (let r = 1; r < values.length; r++) {
    const k = String(values[r][0] || '').trim();
    const v = String(values[r][1] || '').trim();
    if (k) map[k] = v;
  }

  return {
    folderId: map.folderId || '',
    allowedEmails: map.allowedEmails || ''
  };
}

function saveSettings_(p) {
  const ownerEmail = Session.getEffectiveUser().getEmail();
  const userEmail = (Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';

  if (!userEmail || userEmail.toLowerCase() !== ownerEmail.toLowerCase()) {
    return { status: 'error', message: 'คุณไม่ใช่เจ้าของระบบ' };
  }

  const folderId = (p.folderId || '').trim();
  const allowedEmails = (p.allowedEmails || '').trim();

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_SETTINGS, ['Key', 'Value']);

  upsertSetting_(sh, 'folderId', folderId);
  upsertSetting_(sh, 'allowedEmails', allowedEmails);

  const settings = getSettings_();
  ensureDriveFolder_(settings);

  return { status: 'success', message: 'บันทึกการตั้งค่าเรียบร้อย', settings };
}

function upsertSetting_(sh, key, value) {
  const rng = sh.getDataRange();
  const values = rng.getValues();
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim() === key) {
      sh.getRange(r + 1, 2).setValue(value);
      return;
    }
  }
  sh.appendRow([key, value]);
}

function ensureDriveFolder_(settings) {
  if (settings.folderId) {
    // ตรวจว่า folderId ใช้งานได้
    try {
      DriveApp.getFolderById(settings.folderId).getName();
      return;
    } catch (e) {
      // ถ้า folderId พัง ให้สร้างใหม่ทับ
    }
  }

  // สร้างโฟลเดอร์ใหม่ใน My Drive ของ owner/user ที่รัน
  const folder = DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
  settings.folderId = folder.getId();

  // บันทึกกลับ settings sheet
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_SETTINGS, ['Key', 'Value']);
  upsertSetting_(sh, 'folderId', settings.folderId);
}

// ---------- Invoices CRUD ----------
function readInvoices_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_INVOICES, invoiceHeaders_());

  const settings = getSettings_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { status: 'success', data: [], settings };

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const data = rows.map((row, idx) => {
    const obj = { _rowIndex: idx + 2 }; // sheet row number
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse(); // ใหม่อยู่บนสุด

  return { status: 'success', data, settings };
}

function createInvoice_(p) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_INVOICES, invoiceHeaders_());
  const headers = invoiceHeaders_();

  // ทำให้ TotalAmt เป็นตัวเลข/ข้อความสวย
  const row = headers.map(h => (p[h] !== undefined ? p[h] : ''));
  sh.appendRow(row);

  return { status: 'success', message: 'สร้างบิลใหม่เรียบร้อย' };
}

function updateInvoice_(p) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_INVOICES, invoiceHeaders_());
  const headers = invoiceHeaders_();

  const rowIndex = Number(p._rowIndex);
  if (!rowIndex || rowIndex < 2) return { status: 'error', message: 'ไม่พบ Row index สำหรับแก้ไข' };

  const row = headers.map(h => (p[h] !== undefined ? p[h] : ''));
  sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);

  return { status: 'success', message: 'แก้ไขบิลเรียบร้อย' };
}

function deleteInvoice_(p) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_INVOICES, invoiceHeaders_());

  const rowIndex = Number(p._rowIndex);
  if (!rowIndex || rowIndex < 2) return { status: 'error', message: 'ไม่พบ Row index สำหรับลบ' };

  sh.deleteRow(rowIndex);
  return { status: 'success', message: 'ลบบิลเรียบร้อย' };
}

function invoiceHeaders_() {
  return [
    'LogNo', 'Date', 'CustomerName', 'CustomerPhone', 'Room',
    'Item1Desc', 'Item1Amt',
    'Item2Desc', 'Item2Amt',
    'Item3Desc', 'Item3Amt',
    'Item4Desc', 'Item4Amt',
    'TotalAmt', 'TotalText',
    'BankName', 'AccName', 'AccNo', 'QRPicUrl'
  ];
}

// ---------- Drive: PDF & File Manager ----------
function savePdf_(p) {
  const folderId = (p.folderId || '').trim();
  const htmlContent = p.htmlContent || '';
  const fileName = (p.fileName || 'Invoice.pdf').trim();

  if (!folderId) return { status: 'error', message: 'folderId ว่าง' };
  if (!htmlContent) return { status: 'error', message: 'htmlContent ว่าง' };

  const folder = DriveApp.getFolderById(folderId);

  // แปลง HTML -> PDF
  const blob = HtmlService.createHtmlOutput(htmlContent).getBlob().getAs(MimeType.PDF).setName(fileName);
  const file = folder.createFile(blob);

  return { status: 'success', message: 'สร้าง PDF และบันทึกลง Drive เรียบร้อย', fileUrl: file.getUrl(), fileId: file.getId() };
}

function listDriveFiles_(p) {
  const folderId = (p.folderId || '').trim();
  if (!folderId) return { status: 'error', message: 'ยังไม่ได้ตั้งค่า folderId' };

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  const out = [];
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() !== MimeType.PDF) continue;
    out.push({
      id: f.getId(),
      name: f.getName(),
      url: f.getUrl(),
      dateCreated: Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    });
  }

  // ใหม่สุดก่อน
  out.sort((a, b) => (a.dateCreated < b.dateCreated ? 1 : -1));

  return { status: 'success', data: out };
}

function deleteDriveFile_(p) {
  const fileId = (p.fileId || '').trim();
  if (!fileId) return { status: 'error', message: 'fileId ว่าง' };

  const file = DriveApp.getFileById(fileId);
  file.setTrashed(true);

  return { status: 'success', message: 'ลบไฟล์แล้ว (ย้ายไปถังขยะ) เรียบร้อย' };
}

// ---------- Utilities ----------
function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}