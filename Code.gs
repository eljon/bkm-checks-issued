const SHEET_CHECKS = 'Checks';
const SHEET_SUPPLIERS = 'Suppliers';
const SHEET_BANKS = 'Banks';

const CHECK_HEADERS = ['Timestamp', 'Supplier', 'Check Date', 'Check Number', 'Bank', 'Amount', 'Notes'];

function doGet() {
  ensureSheets_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('BKM Checks Issued')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getLists() {
  ensureSheets_();
  return {
    suppliers: getColumn_(SHEET_SUPPLIERS),
    banks: getColumn_(SHEET_BANKS)
  };
}

function saveCheck(payload) {
  ensureSheets_();

  const supplier = String(payload.supplier || '').trim();
  const bank = String(payload.bank || '').trim();
  const checkNumber = String(payload.checkNumber || '').trim();
  const checkDate = String(payload.checkDate || '').trim();
  const amountRaw = payload.amount;
  const notes = String(payload.notes || '').trim();

  if (!supplier) throw new Error('Supplier is required.');
  if (!bank) throw new Error('Bank is required.');
  if (!checkNumber) throw new Error('Check number is required.');
  if (!checkDate) throw new Error('Check date is required.');

  const amount = Number(String(amountRaw).replace(/,/g, ''));
  if (!isFinite(amount) || amount <= 0) throw new Error('Amount must be a positive number.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = ss.getSheetByName(SHEET_CHECKS);

  addIfNew_(SHEET_SUPPLIERS, supplier);
  addIfNew_(SHEET_BANKS, bank);

  const row = checks.getLastRow() + 1;
  checks.getRange(row, 1, 1, CHECK_HEADERS.length).setValues([[
    new Date(),
    supplier,
    checkDate,
    checkNumber,
    bank,
    amount,
    notes
  ]]);

  checks.getRange(row, 3).setNumberFormat('yyyy-mm-dd');
  checks.getRange(row, 6).setNumberFormat('#,##0.00');

  return { ok: true };
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let checks = ss.getSheetByName(SHEET_CHECKS);
  if (!checks) {
    checks = ss.insertSheet(SHEET_CHECKS);
    checks.getRange(1, 1, 1, CHECK_HEADERS.length).setValues([CHECK_HEADERS]).setFontWeight('bold');
    checks.setFrozenRows(1);
  }

  const suppliers = ss.getSheetByName(SHEET_SUPPLIERS);
  if (!suppliers) {
    const s = ss.insertSheet(SHEET_SUPPLIERS);
    s.getRange(1, 1).setValue('Supplier').setFontWeight('bold');
    s.setFrozenRows(1);
  }

  const banks = ss.getSheetByName(SHEET_BANKS);
  if (!banks) {
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
