// Paste this entire file into Google Apps Script (Extensions -> Apps Script
// from your Sheet). Deploy once as a Web App with "Anyone" access. The
// generated /exec URL goes into config.js as `apiUrl`.
//
// You only ever come back here if you change SHARED_SECRET. Frontend changes
// happen in the repo and ship with `git push` — the script never needs
// re-deploying for those.

// Optional shared secret. If non-empty, the frontend must send the same
// value in config.js so requests match. Leave empty to disable.
const SHARED_SECRET = '';

const SHEET_CHECKS = 'Checks';
const SHEET_SUPPLIERS = 'Suppliers';
const SHEET_BANKS = 'Banks';
const CHECK_HEADERS = ['Timestamp', 'Supplier', 'Check Date', 'Check Number', 'Bank', 'Amount', 'Notes'];

function doGet() {
  return json({ ok: true, message: 'BKM Checks API. POST JSON to use.' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (SHARED_SECRET && req.secret !== SHARED_SECRET) {
      return json({ error: 'unauthorized' });
    }
    switch (req.action) {
      case 'getLists':     return json(getLists_());
      case 'saveCheck':    return json(saveCheck_(req.payload || {}));
      case 'findChecks':   return json(findChecks_(req.payload || {}));
      case 'updateCheck':  return json(updateCheck_(req.payload || {}));
      case 'searchChecks': return json(searchChecks_(req.payload || {}));
      default:             return json({ error: 'unknown action: ' + req.action });
    }
  } catch (err) {
    return json({ error: (err && err.message) || String(err) });
  }
}

function getLists_() {
  ensureSheets_();
  return {
    suppliers: getColumn_(SHEET_SUPPLIERS),
    banks: getColumn_(SHEET_BANKS)
  };
}

function saveCheck_(p) {
  ensureSheets_();

  const supplier = String(p.supplier || '').trim();
  const bank = String(p.bank || '').trim();
  const checkNumber = String(p.checkNumber || '').trim();
  const checkDate = String(p.checkDate || '').trim();
  const notes = String(p.notes || '').trim();
  const amount = Number(String(p.amount || '').replace(/,/g, ''));

  if (!supplier) throw new Error('Supplier is required.');
  if (!bank) throw new Error('Bank is required.');
  if (!checkNumber) throw new Error('Check number is required.');
  if (!checkDate) throw new Error('Check date is required.');
  if (!isFinite(amount) || amount <= 0) throw new Error('Amount must be a positive number.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = ss.getSheetByName(SHEET_CHECKS);

  addIfNew_(SHEET_SUPPLIERS, supplier);
  addIfNew_(SHEET_BANKS, bank);

  const row = checks.getLastRow() + 1;
  checks.getRange(row, 1, 1, CHECK_HEADERS.length).setValues([[
    new Date(), supplier, checkDate, checkNumber, bank, amount, notes
  ]]);
  checks.getRange(row, 3).setNumberFormat('yyyy-mm-dd');
  checks.getRange(row, 6).setNumberFormat('#,##0.00');

  return { ok: true };
}

function findChecks_(p) {
  ensureSheets_();
  const checkNumber = String(p.checkNumber || '').trim();
  if (!checkNumber) throw new Error('Check number is required.');

  const rows = readChecks_();
  const matches = rows
    .filter(r => String(r.row[3] || '').trim() === checkNumber)
    .map(r => rowToCheck_(r.row, r.rowIndex));
  return { matches };
}

function updateCheck_(p) {
  ensureSheets_();
  const rowIndex = Number(p.rowIndex);
  if (!rowIndex || rowIndex < 2) throw new Error('Invalid row index.');

  const supplier = String(p.supplier || '').trim();
  const bank = String(p.bank || '').trim();
  const checkNumber = String(p.checkNumber || '').trim();
  const checkDate = String(p.checkDate || '').trim();
  const notes = String(p.notes || '').trim();
  const amount = Number(String(p.amount || '').replace(/,/g, ''));

  if (!supplier) throw new Error('Supplier is required.');
  if (!bank) throw new Error('Bank is required.');
  if (!checkNumber) throw new Error('Check number is required.');
  if (!checkDate) throw new Error('Check date is required.');
  if (!isFinite(amount) || amount <= 0) throw new Error('Amount must be a positive number.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = ss.getSheetByName(SHEET_CHECKS);
  if (rowIndex > checks.getLastRow()) throw new Error('Row no longer exists.');

  addIfNew_(SHEET_SUPPLIERS, supplier);
  addIfNew_(SHEET_BANKS, bank);

  const originalTimestamp = checks.getRange(rowIndex, 1).getValue();
  checks.getRange(rowIndex, 1, 1, CHECK_HEADERS.length).setValues([[
    originalTimestamp || new Date(), supplier, checkDate, checkNumber, bank, amount, notes
  ]]);
  checks.getRange(rowIndex, 3).setNumberFormat('yyyy-mm-dd');
  checks.getRange(rowIndex, 6).setNumberFormat('#,##0.00');

  return { ok: true };
}

function searchChecks_(p) {
  ensureSheets_();
  const supplier = String(p.supplier || '').trim().toLowerCase();
  const fromDate = String(p.from || '').trim();
  const toDate = String(p.to || '').trim();

  const rows = readChecks_();
  const checks = rows.map(r => rowToCheck_(r.row, r.rowIndex)).filter(c => {
    if (supplier && c.supplier.toLowerCase().indexOf(supplier) === -1) return false;
    if (fromDate && c.checkDate < fromDate) return false;
    if (toDate && c.checkDate > toDate) return false;
    return true;
  });
  return { checks };
}

function readChecks_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CHECKS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, CHECK_HEADERS.length).getValues();
  return values.map((row, idx) => ({ row, rowIndex: idx + 2 }));
}

function rowToCheck_(row, rowIndex) {
  return {
    rowIndex,
    timestamp: formatDate_(row[0]),
    supplier: String(row[1] || ''),
    checkDate: formatDate_(row[2]),
    checkNumber: String(row[3] || ''),
    bank: String(row[4] || ''),
    amount: Number(row[5]) || 0,
    notes: String(row[6] || '')
  };
}

function formatDate_(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(v || '');
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let checks = ss.getSheetByName(SHEET_CHECKS);
  if (!checks) {
    checks = ss.insertSheet(SHEET_CHECKS);
    checks.getRange(1, 1, 1, CHECK_HEADERS.length).setValues([CHECK_HEADERS]).setFontWeight('bold');
    checks.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_SUPPLIERS)) {
    const s = ss.insertSheet(SHEET_SUPPLIERS);
    s.getRange(1, 1).setValue('Supplier').setFontWeight('bold');
    s.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET_BANKS)) {
    const b = ss.insertSheet(SHEET_BANKS);
    b.getRange(1, 1).setValue('Bank').setFontWeight('bold');
    b.setFrozenRows(1);
  }
}

function getColumn_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const seen = new Set();
  const out = [];
  values.forEach(r => {
    const v = String(r[0] || '').trim();
    if (!v) return;
    const k = v.toLowerCase();
    if (seen.has(k)) return;
    seen.add(k);
    out.push(v);
  });
  out.sort((a, b) => a.localeCompare(b));
  return out;
}

function addIfNew_(sheetName, value) {
  const v = String(value || '').trim();
  if (!v) return;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const existing = getColumn_(sheetName).map(s => s.toLowerCase());
  if (existing.includes(v.toLowerCase())) return;
  sheet.appendRow([v]);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
