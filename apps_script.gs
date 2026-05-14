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
const CHECK_HEADERS = ['Timestamp', 'Supplier', 'Order Month', 'Check Date', 'Check Number', 'Bank', 'Amount', 'Notes'];

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
  const orderMonth = String(p.orderMonth || '').trim();
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
    new Date(), supplier, orderMonth, checkDate, checkNumber, bank, amount, notes
  ]]);

  return { ok: true };
}

function findChecks_(p) {
  ensureSheets_();
  const checkNumber = String(p.checkNumber || '').trim();
  if (!checkNumber) throw new Error('Check number is required.');

  const rows = readChecks_();
  const matches = rows
    .filter(r => String(r.row[4] || '').trim() === checkNumber)
    .map(r => rowToCheck_(r.row, r.rowIndex));
  return { matches };
}

function updateCheck_(p) {
  ensureSheets_();
  const rowIndex = Number(p.rowIndex);
  if (!rowIndex || rowIndex < 2) throw new Error('Invalid row index.');

  const supplier = String(p.supplier || '').trim();
  const orderMonth = String(p.orderMonth || '').trim();
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
    originalTimestamp || new Date(), supplier, orderMonth, checkDate, checkNumber, bank, amount, notes
  ]]);

  return { ok: true };
}

function searchChecks_(p) {
  ensureSheets_();
  const supplier = String(p.supplier || '').trim().toLowerCase();
  const bank = String(p.bank || '').trim().toLowerCase();
  const fromDate = String(p.from || '').trim();
  const toDate = String(p.to || '').trim();
  const orderMonth = String(p.orderMonth || '').trim();

  const rows = readChecks_();
  const checks = rows.map(r => rowToCheck_(r.row, r.rowIndex)).filter(c => {
    if (supplier && c.supplier.toLowerCase().indexOf(supplier) === -1) return false;
    if (bank && c.bank.toLowerCase().indexOf(bank) === -1) return false;
    if (fromDate && c.checkDate < fromDate) return false;
    if (toDate && c.checkDate > toDate) return false;
    if (orderMonth && c.orderMonth !== orderMonth) return false;
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
    orderMonth: formatYearMonth_(row[2]),
    checkDate: formatDate_(row[3]),
    checkNumber: String(row[4] || ''),
    bank: String(row[5] || ''),
    amount: Number(row[6]) || 0,
    notes: String(row[7] || '')
  };
}

function formatYearMonth_(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}`;
  }
  return String(v || '').trim();
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

// Bump SCHEMA_VERSION whenever migrateChecksHeader_ needs to re-run on
// existing spreadsheets. Each version runs the migration once per Sheet
// (tracked in document properties) and then short-circuits on every
// subsequent call so doPost stays fast.
const SCHEMA_VERSION = 'v4';
const CACHE_TTL_SEC = 600; // 10 minutes

function ensureSheets_() {
  // Hot path: most requests just hit the cache and skip the Sheets API.
  const cache = CacheService.getScriptCache();
  if (cache.get('schema_ok') === SCHEMA_VERSION) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let checks = ss.getSheetByName(SHEET_CHECKS);
  if (!checks) {
    checks = ss.insertSheet(SHEET_CHECKS);
    checks.getRange(1, 1, 1, CHECK_HEADERS.length).setValues([CHECK_HEADERS]).setFontWeight('bold');
    checks.setFrozenRows(1);
    applyChecksColumnFormats_(checks);
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

  runMigrationsIfNeeded_(checks);

  cache.put('schema_ok', SCHEMA_VERSION, CACHE_TTL_SEC);
}

function applyChecksColumnFormats_(checks) {
  // Column-level formats so saveCheck_/updateCheck_ never need per-row formatting.
  checks.getRange('C:C').setNumberFormat('@');           // Order Month as text
  checks.getRange('D:D').setNumberFormat('yyyy-mm-dd');  // Check Date
  checks.getRange('G:G').setNumberFormat('#,##0.00');    // Amount
}

function runMigrationsIfNeeded_(checks) {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('schema_version') === SCHEMA_VERSION) return;
  migrateChecksHeader_(checks);
  applyChecksColumnFormats_(checks);
  backfillListFromChecks_(checks, SHEET_SUPPLIERS, 2); // column B: Supplier
  backfillListFromChecks_(checks, SHEET_BANKS, 6);     // column F: Bank
  props.setProperty('schema_version', SCHEMA_VERSION);
}

// ============================================================================
// ONE-TIME UTILITY — delete this function after running it.
//
// How to run:
//   1. In the Apps Script editor, paste this file's contents and save.
//   2. In the function dropdown at the top, select `mergeSupplierAliases`.
//   3. Click Run. Grant permissions on first run.
//   4. Open the Sheet to verify the Suppliers tab and the Checks rows look
//      right.
//   5. Delete this function (and this comment block) from apps_script.gs
//      and save again. No redeploy needed — this runs only from the editor.
// ============================================================================
function mergeSupplierAliases() {
  const merges = [
    { canonical: "ERICK'S MOTOR SALES",                       aliases: ['ERICKS MOTOR SALES'] },
    { canonical: "ROBERT'S AIPMC",                            aliases: ['ROBERTS AIPMC'] },
    { canonical: 'CSS AUTO PARTS INC.',                       aliases: ['CSS AUTO PARTS INC'] },
    { canonical: "FLY DRAGON INTERNATIONAL HOLDING'S INC",    aliases: ['FLY DRAGON INTERNATIONAL HOLDINGS INC'] },
    { canonical: 'XIAOHAI CHEN',                              aliases: ['XIAHAI CHEN'] },
    { canonical: 'TRIPLE P INCLUSIVE SALES CORP',             aliases: ['TRIPLE P INCLUSNE SALES CORP'] },
    { canonical: 'DREAMVOLTS MARKETING',                      aliases: ['DREAMVOLTS MARKETINGB'] },
    { canonical: 'AUTOPHIL ZONE SALES CORPORATION',           aliases: ['AUTOPHIL ZONE SALES CORP'] },
    { canonical: 'KINGSWEALTH INTERNATIONAL TRADING',         aliases: ['KINGSWEALTH INTERNATIONAL'] },
    { canonical: 'LRY BRAKE CLUTCH LINING TRADING',           aliases: ['LRY BRAKE CLUTCH LINING'] },
    { canonical: 'MAODUO TRADING CORP',                       aliases: ['MAODUO TRADING'] },
    { canonical: 'Southeast Marketing Corporation',           aliases: ['Southeast'] },
    { canonical: 'NUVO AUTO PARTS CORP',                      aliases: ['NUVO PARTS PARTS SUPPLY'] }
  ];

  const aliasMap = new Map();
  merges.forEach(m => {
    m.aliases.forEach(a => aliasMap.set(a.toLowerCase(), m.canonical));
  });

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checks = ss.getSheetByName(SHEET_CHECKS);
  if (!checks) throw new Error('Checks sheet not found.');

  // Rewrite supplier column in Checks.
  let rewritten = 0;
  const lastRow = checks.getLastRow();
  if (lastRow >= 2) {
    const range = checks.getRange(2, 2, lastRow - 1, 1); // column B
    const values = range.getValues();
    for (let i = 0; i < values.length; i++) {
      const v = String(values[i][0] || '').trim();
      if (!v) continue;
      const replacement = aliasMap.get(v.toLowerCase());
      if (replacement && replacement !== v) {
        values[i][0] = replacement;
        rewritten++;
      }
    }
    if (rewritten) range.setValues(values);
  }

  // Rebuild Suppliers list: drop aliases, keep everything else (deduped),
  // make sure every canonical is present.
  const suppliers = ss.getSheetByName(SHEET_SUPPLIERS);
  if (!suppliers) throw new Error('Suppliers sheet not found.');
  const sLastRow = suppliers.getLastRow();
  const kept = [];
  const keptLower = new Set();
  if (sLastRow >= 2) {
    const sValues = suppliers.getRange(2, 1, sLastRow - 1, 1).getValues();
    sValues.forEach(r => {
      const v = String(r[0] || '').trim();
      if (!v) return;
      const key = v.toLowerCase();
      if (aliasMap.has(key)) return;
      if (keptLower.has(key)) return;
      kept.push(v);
      keptLower.add(key);
    });
  }
  merges.forEach(m => {
    const k = m.canonical.toLowerCase();
    if (!keptLower.has(k)) {
      kept.push(m.canonical);
      keptLower.add(k);
    }
  });
  kept.sort((a, b) => a.localeCompare(b));

  if (sLastRow >= 2) {
    suppliers.getRange(2, 1, sLastRow - 1, 1).clearContent();
  }
  if (kept.length) {
    suppliers.getRange(2, 1, kept.length, 1).setValues(kept.map(v => [v]));
  }

  CacheService.getScriptCache().remove('col_' + SHEET_SUPPLIERS);

  Logger.log('Rewrote ' + rewritten + ' Checks row(s). Suppliers list now has ' +
             kept.length + ' entries.');
}

function backfillListFromChecks_(checks, listSheetName, columnIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName(listSheetName);
  if (!listSheet) return;

  const lastRow = checks.getLastRow();
  if (lastRow < 2) return;

  // Unique non-empty values currently in the Checks column.
  const seenInChecks = new Map(); // lower -> original casing (first occurrence)
  const checkValues = checks.getRange(2, columnIndex, lastRow - 1, 1).getValues();
  for (let i = 0; i < checkValues.length; i++) {
    const v = String(checkValues[i][0] || '').trim();
    if (!v) continue;
    const k = v.toLowerCase();
    if (!seenInChecks.has(k)) seenInChecks.set(k, v);
  }

  // Whatever already exists in the list tab.
  const existingLower = new Set();
  const listLastRow = listSheet.getLastRow();
  if (listLastRow >= 2) {
    const listValues = listSheet.getRange(2, 1, listLastRow - 1, 1).getValues();
    for (let i = 0; i < listValues.length; i++) {
      const v = String(listValues[i][0] || '').trim();
      if (v) existingLower.add(v.toLowerCase());
    }
  }

  // Append rows for anything missing.
  const missing = [];
  seenInChecks.forEach((original, lower) => {
    if (!existingLower.has(lower)) missing.push([original]);
  });
  if (!missing.length) return;

  const startRow = listSheet.getLastRow() + 1;
  listSheet.getRange(startRow, 1, missing.length, 1).setValues(missing);
  CacheService.getScriptCache().remove('col_' + listSheetName);
}

function migrateChecksHeader_(checks) {
  const lastCol = Math.max(checks.getLastColumn(), 1);
  const header = checks.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('Order Month') === -1) {
    // Insert empty Order Month column right after Supplier (column 2).
    checks.insertColumnAfter(2);
    checks.getRange(1, 3).setValue('Order Month').setFontWeight('bold');
  }
  // Sync the header row with the canonical headers in case anything drifted.
  checks.getRange(1, 1, 1, CHECK_HEADERS.length)
    .setValues([CHECK_HEADERS])
    .setFontWeight('bold');

  // Rewrite any Order Month cells that Sheets auto-parsed into Date values
  // back as plain "YYYY-MM" text so equality filtering works.
  const lastRow = checks.getLastRow();
  if (lastRow >= 2) {
    const range = checks.getRange(2, 3, lastRow - 1, 1);
    const values = range.getValues();
    let dirty = false;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] instanceof Date) {
        values[i][0] = formatYearMonth_(values[i][0]);
        dirty = true;
      }
    }
    if (dirty) {
      range.setNumberFormat('@');
      range.setValues(values);
    }
  }
}

function getColumn_(sheetName) {
  const cacheKey = 'col_' + sheetName;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached !== null) return JSON.parse(cached);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const lastRow = sheet.getLastRow();
  const out = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const seen = new Set();
    values.forEach(r => {
      const v = String(r[0] || '').trim();
      if (!v) return;
      const k = v.toLowerCase();
      if (seen.has(k)) return;
      seen.add(k);
      out.push(v);
    });
    out.sort((a, b) => a.localeCompare(b));
  }
  cache.put(cacheKey, JSON.stringify(out), CACHE_TTL_SEC);
  return out;
}

function addIfNew_(sheetName, value) {
  const v = String(value || '').trim();
  if (!v) return;
  const existing = getColumn_(sheetName);
  if (existing.some(s => s.toLowerCase() === v.toLowerCase())) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.appendRow([v]);

  // Keep the cache warm so the next save doesn't re-read the column.
  const updated = existing.concat(v).sort((a, b) => a.localeCompare(b));
  CacheService.getScriptCache().put('col_' + sheetName, JSON.stringify(updated), CACHE_TTL_SEC);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
