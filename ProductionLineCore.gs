/**
 * =================================================================
 * PRODUCTION LINE CORE — Consolidated Apps Script
 * =================================================================
 * Replaces 4 separate Apps Scripts (wiring, assembly, dispatch, hub)
 * with ONE script, ONE deployment URL.
 *
 * Functions exposed via Web App:
 *   GET  ?action=readMachines            → returns JSON list of machines
 *   POST ?action=logWiringScan           → scanner app writes wiring scan
 *   POST ?action=logAssemblyScan         → scanner app writes assembly scan
 *   POST ?action=logDispatchScan         → scanner app writes dispatch scan
 *   POST ?action=manualStockEntry        → data-entry app adds to frame stock
 *   POST ?action=addMachine              → admin adds new machine (atomic)
 *   POST (legacy, no action)             → backwards-compat with old Hub calls
 *
 * Routing: each POST/GET includes `?action=<name>` in URL.
 * Scanner apps use mode:"no-cors" (cannot read response).
 * Data-entry app uses Content-Type:"text/plain" (can read JSON response).
 * =================================================================
 */

// ==================== CONFIGURATION ====================
// Paste your actual Sheet IDs here (from the URLs of each Google Sheet)
// These are the IDs we extracted from your existing HTML files.

const SHEET_IDS = {
  wiring:    '19y16Nu7GdYPN0DdTZId2bfoBn4jSmV5ZVHXSkdDdcRY',
  assembly:  '1vYAuJxLY3DXoTf_5Pbvin1zwA6L8N4pNOhy180ag4p0',
  dispatch:  '15Kj3Dw073HjuK54uYOEq4kn-04SeSOUqHydgZetDphI',
  stock:     '1nsLwfGbA_rVtdZxn9fSzW9_2o9JAT0OGVWwrO-vv2sQ',
  registry:  '1nsLwfGbA_rVtdZxn9fSzW9_2o9JAT0OGVWwrO-vv2sQ'  // same as stock, different tab
};

// Tab names inside each spreadsheet
const SHEET_NAMES = {
  wiring:    'Sheet1',
  assembly:  'Sheet1',
  dispatch:  'Sheet1',
  stock:     'Sheet1',
  registry:  'MachineRegistry'  // new tab to create in the stock spreadsheet
};

// ==================== JUNK COLUMN SKIP LOGIC ====================
// PRESERVED VERBATIM from your existing scripts (wiring/assembly/dispatch)
// Do NOT modify these — they reflect what junk columns exist in your sheets.

const JUNK_SKIP_EXACT = ['id', 's.no', 'sno', 'no', '#'];
const JUNK_SKIP_SUBSTRINGS = ['date', 'time', 'shift', 'operator', 'status',
                              'remark', 'note', 'total', 'count', 'day',
                              'name', 'model', 'serial', 'input', 'scan'];

function isJunkColumn(header) {
  if (!header) return true;
  const h = header.toString().toLowerCase().trim();
  if (JUNK_SKIP_EXACT.includes(h)) return true;
  if (JUNK_SKIP_SUBSTRINGS.some(k => h.includes(k))) return true;
  return false;
}

// ==================== ROUTER (doGet / doPost / doOptions) ====================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action || '').toLowerCase();
  switch (action) {
    case 'readmachines':
      return readMachines();
    default:
      return jsonResponse({ status: 'error', message: 'Unknown GET action: ' + action });
  }
}

function doPost(e) {
  // Parse body JSON (works for both no-cors and text/plain POSTs)
  let bodyData = {};
  try {
    if (e && e.postData && e.postData.contents) {
      bodyData = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return textResponse('Error: Invalid JSON body');
  }

  // Action can come from URL query OR body
  const action = (
    (e && e.parameter && e.parameter.action) ||
    bodyData.action ||
    ''
  ).toLowerCase();

  switch (action) {
    case 'logwiring':
      return logWiringScan(bodyData);
    case 'logassembly':
      return logAssemblyScan(bodyData);
    case 'logdispatch':
      return logDispatchScan(bodyData);
    case 'manualstockentry':
      return manualStockEntry(bodyData);
    case 'addmachine':
      return addMachine(bodyData);
    case 'readmachines':
      return readMachines();
    default:
      // LEGACY HUB COMPATIBILITY
      // Old scanner APKs still call the old Hub URL with no action parameter,
      // sending {machineName, action, amount}. Support that pattern.
      if (bodyData.machineName && bodyData.action) {
        return legacyHubUpdate(bodyData);
      }
      return textResponse('Error: Unknown action: ' + action);
  }
}

function doOptions(e) {
  // CORS preflight handler (defensive — most calls use no-cors which doesn't preflight)
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ==================== MACHINE REGISTRY ====================

/**
 * readMachines() — GET endpoint
 * Returns the list of all ACTIVE machines from MachineRegistry sheet.
 * Called by all 5 HTML apps on page load to populate dropdowns dynamically.
 *
 * Response: { machines: [ {machine_id, model_name, material_type, serial_prefix,
 *                          serial_pattern, category, added_date, status, min_stock}, ... ] }
 */
function readMachines() {
  try {
    const sheet = getRegistrySheet();
    if (!sheet) {
      return jsonResponse({ status: 'error', message: 'MachineRegistry sheet not found. Create it first.' });
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return jsonResponse({ machines: [] });
    }
    const machines = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;  // skip empty rows
      machines.push({
        machine_id:      row[0],
        model_name:      row[1],
        material_type:  row[2],
        serial_prefix:   row[3],
        serial_pattern:  row[4],
        category:        row[5],
        added_date:      row[6] ? formatDate(row[6]) : '',
        status:          row[7] || 'active',
        min_stock:       Number(row[8]) || 0
      });
    }
    return jsonResponse({ machines: machines });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * addMachine() — POST endpoint
 * Atomically adds a new machine type to the system:
 *   1. Insert row in MachineRegistry
 *   2. Append column to wiring, assembly, dispatch sheets
 *   3. Append row to Stock sheet (col A = "SS ModelName" or "MS ModelName")
 *   4. Verify: column header exists in wiring sheet
 *   5. If any step fails, ROLL BACK all prior changes
 *
 * Body: { model_name, material_type, serial_prefix, serial_pattern, category, min_stock }
 */
function addMachine(data) {
  // Validate input
  if (!data.model_name || !data.material_type || !data.serial_prefix) {
    return jsonResponse({
      status: 'error',
      message: 'Missing required fields. Need: model_name, material_type, serial_prefix'
    });
  }

  // Generate machine_id (e.g. "CIDM-SS")
  const machineId = `${data.model_name}-${data.material_type}`;
  const fullMachineName = `${data.material_type} ${data.model_name}`;

  // Open registry
  const regSheet = getRegistrySheet();
  if (!regSheet) {
    return jsonResponse({ status: 'error', message: 'MachineRegistry sheet not found. Create it first.' });
  }

  // Check if machine_id already exists
  const regData = regSheet.getDataRange().getValues();
  for (let i = 1; i < regData.length; i++) {
    if (regData[i][0] === machineId) {
      return jsonResponse({ status: 'error', message: `Machine '${machineId}' already exists` });
    }
  }

  // Check serial pattern collision (skip if pattern is .* = wildcard)
  if (data.serial_pattern && data.serial_pattern !== '.*') {
    for (let i = 1; i < regData.length; i++) {
      const existingPattern = regData[i][4];
      if (existingPattern && existingPattern !== '.*' && patternsOverlap(data.serial_pattern, existingPattern)) {
        return jsonResponse({
          status: 'error',
          message: `Serial pattern overlaps with existing machine: ${regData[i][1]}-${regData[i][2]} (pattern: ${existingPattern})`
        });
      }
    }
  }

  // === ATOMIC OPERATIONS WITH ROLLBACK TRACKING ===
  const addedColumns = [];  // [{sheet, col}] for rollback
  const addedRows = [];     // [{sheet, rowIndex}] for rollback

  try {
    // 1. Add row to MachineRegistry
    const today = new Date();
    regSheet.appendRow([
      machineId,
      data.model_name,
      data.material_type,
      data.serial_prefix,
      data.serial_pattern || '.*',
      data.category || data.material_type,
      today,
      'active',
      Number(data.min_stock) || 5
    ]);
    // Find the row we just added (last row)
    const regLastRow = regSheet.getLastRow();
    addedRows.push({ sheet: regSheet, rowIndex: regLastRow, type: 'registry' });

    // 2. Append column to each stage sheet (wiring, assembly, dispatch)
    // Stage sheets have columns: Date | CIDM | CSEM | ... | (junk cols)
    // We append the new model_name as a new column at the end.
    // Junk columns will still be skipped by the junk-skip logic during reverse-lookup.
    for (const stage of ['wiring', 'assembly', 'dispatch']) {
      const stageSheet = getSheet(stage);
      if (stageSheet) {
        const lastCol = stageSheet.getLastColumn();
        stageSheet.getRange(1, lastCol + 1).setValue(data.model_name);
        addedColumns.push({ sheet: stageSheet, col: lastCol + 1 });
      }
    }

    // 3. Append row to Stock sheet
    // Stock columns: A=Machine Name | B=Ready | C=Min | D=Prepared
    const stockSheet = getSheet('stock');
    if (stockSheet) {
      stockSheet.appendRow([fullMachineName, 0, Number(data.min_stock) || 5, 0]);
      const stockLastRow = stockSheet.getLastRow();
      addedRows.push({ sheet: stockSheet, rowIndex: stockLastRow, type: 'stock' });
    }

    // 4. Post-add verification: confirm new column header is in wiring sheet
    const wireSheet = getSheet('wiring');
    if (wireSheet) {
      const headers = wireSheet.getRange(1, 1, 1, wireSheet.getLastColumn()).getValues()[0];
      if (!headers.includes(data.model_name)) {
        throw new Error('Verification failed: new column not found in WiringLog after append');
      }
    }

    // === SUCCESS ===
    return jsonResponse({
      status: 'success',
      machine_id: machineId,
      machine_name: fullMachineName,
      message: `Machine '${fullMachineName}' added successfully. ` +
               `Column '${data.model_name}' added to wiring/assembly/dispatch sheets. ` +
               `Row '${fullMachineName}' added to Stock sheet.`
    });

  } catch (err) {
    // === ROLLBACK ALL CHANGES ===
    // Delete added columns (in reverse order to preserve indices)
    addedColumns.reverse().forEach(c => {
      try { c.sheet.deleteColumn(c.col); } catch (e) {}
    });
    // Delete added rows (in reverse order)
    addedRows.reverse().forEach(r => {
      try { r.sheet.deleteRow(r.rowIndex); } catch (e) {}
    });
    return jsonResponse({
      status: 'error',
      message: `Rollback after error: ${err.toString()}. All changes were undone.`
    });
  }
}

// ==================== SCAN LOGGING (3 scanners) ====================

/**
 * logWiringScan() — POST endpoint (no-cors)
 * Body: { serial, machineType, date }
 * Writes the scan to WiringLog under the matching machine column.
 * Decreases Stock.Prepared (col D) by 1 for that machine.
 */
function logWiringScan(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!data.serial || !data.machineType) {
      lock.releaseLock();
      return textResponse('Error: Missing Data');
    }

    const sheet = getSheet('wiring');
    if (!sheet) {
      lock.releaseLock();
      return textResponse('Error: Wiring sheet not found');
    }

    // DYNAMIC column lookup — find column by header match
    const colIndex = findColumnByHeader(sheet, data.machineType);
    if (colIndex === -1) {
      lock.releaseLock();
      return textResponse('Error: Invalid Machine Type: ' + data.machineType);
    }

    // Build row: [date, "", "", ...] with serial at colIndex-1
    const lastCol = sheet.getLastColumn();
    const dateVal = data.date || new Date().toISOString().split('T')[0];
    const rowValues = new Array(lastCol).fill("");
    rowValues[0] = dateVal;
    rowValues[colIndex - 1] = data.serial;
    sheet.appendRow(rowValues);

    // Update stock internally (no HTTP call needed since we're consolidated)
    const prefix = data.serial.startsWith("SSKM") ? "SS" : "MS";
    const fullMachineName = `${prefix} ${data.machineType}`;
    updateStockInternal(fullMachineName, 'wiring', -1);

    lock.releaseLock();
    return textResponse('Success');
  } catch (err) {
    lock.releaseLock();
    return textResponse('Error: ' + err.toString());
  }
}

/**
 * logAssemblyScan() — POST endpoint (no-cors)
 * Body: { serial, machineType, date }
 * Writes the scan to AssemblyLog under the matching machine column.
 * Increases Stock.Ready (col B) by 1 for that machine.
 */
function logAssemblyScan(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!data.serial || !data.machineType) {
      lock.releaseLock();
      return textResponse('Error: Missing Data');
    }

    const sheet = getSheet('assembly');
    if (!sheet) {
      lock.releaseLock();
      return textResponse('Error: Assembly sheet not found');
    }

    // DYNAMIC column lookup
    const colIndex = findColumnByHeader(sheet, data.machineType);
    if (colIndex === -1) {
      lock.releaseLock();
      return textResponse('Error: Invalid Machine Type: ' + data.machineType);
    }

    const lastCol = sheet.getLastColumn();
    const dateVal = data.date || new Date().toISOString().split('T')[0];
    const rowValues = new Array(lastCol).fill("");
    rowValues[0] = dateVal;
    rowValues[colIndex - 1] = data.serial;
    sheet.appendRow(rowValues);

    const prefix = data.serial.startsWith("SSKM") ? "SS" : "MS";
    const fullMachineName = `${prefix} ${data.machineType}`;
    updateStockInternal(fullMachineName, 'assembly', 1);

    lock.releaseLock();
    return textResponse('Success');
  } catch (err) {
    lock.releaseLock();
    return textResponse('Error: ' + err.toString());
  }
}

/**
 * logDispatchScan() — POST endpoint (no-cors)
 * Body: { serial, customer, invoice, date }
 * Dispatch sheet shape: Serial | Invoice | Customer | Date
 * Determines machineType via reverse-lookup in wiring → assembly sheets.
 * Decreases Stock.Ready (col B) by 1 for that machine.
 */
function logDispatchScan(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!data.serial) {
      lock.releaseLock();
      return textResponse('Error: Missing Serial');
    }

    const serial = data.serial.toUpperCase();
    const sheet = getSheet('dispatch');
    if (!sheet) {
      lock.releaseLock();
      return textResponse('Error: Dispatch sheet not found');
    }

    // Write row first (Serial | Invoice | Customer | Date)
    const date = data.date || new Date().toLocaleDateString('en-GB');
    sheet.appendRow([serial, data.invoice || "", data.customer || "", date]);
    SpreadsheetApp.flush();

    // Reverse-lookup: search in wiring sheet, fallback to assembly sheet
    let machineName = lookupMachineInStageSheet('wiring', serial);
    if (!machineName) {
      machineName = lookupMachineInStageSheet('assembly', serial);
    }

    if (machineName) {
      updateStockInternal(machineName, 'dispatch', -1);
    }

    lock.releaseLock();
    return textResponse('Success');
  } catch (err) {
    lock.releaseLock();
    return textResponse('Error: ' + err.toString());
  }
}

/**
 * manualStockEntry() — POST endpoint (Content-Type: text/plain, returns JSON)
 * Body: { machineName, amount }
 * Called by data-entry HTML when worker taps +1 or enters bulk quantity.
 * Updates Stock.Prepared (col D) by +amount for that machine.
 */
function manualStockEntry(data) {
  try {
    if (!data.machineName) {
      return jsonResponse({ status: 'error', message: 'Missing machineName' });
    }
    const amount = parseInt(data.amount) || 0;
    if (amount === 0) {
      return jsonResponse({ status: 'error', message: 'Invalid amount' });
    }
    updateStockInternal(data.machineName, 'worker', amount);
    return jsonResponse({ status: 'success', machine: data.machineName, amount: amount });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * legacyHubUpdate() — backwards-compat for old scanner APKs
 * Old scanners POST {machineName, action, amount} to the old Hub URL.
 * If old scanner APKs are still pointed at the consolidated URL (during migration),
 * this handles their calls without breaking.
 */
function legacyHubUpdate(data) {
  try {
    updateStockInternal(data.machineName, data.action || 'worker', parseInt(data.amount) || 1);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== INTERNAL STOCK UPDATE ====================

/**
 * updateStockInternal() — replaces the old sendToHub() HTTP call.
 * Since the script is consolidated, no HTTP roundtrip needed.
 *
 * Stock sheet columns: A=Machine Name | B=Ready | C=Min | D=Prepared
 * Action rules (preserved from old Hub script):
 *   "wiring"     → update col D (Prepared), amount (typically -1, frame used)
 *   "assembly"   → update col B (Ready),    amount (typically +1, finished unit)
 *   "dispatch"   → update col B (Ready),    amount (typically -1, unit shipped)
 *   "worker"     → update col D (Prepared), amount (manual entry from data-entry app)
 *   default      → update col D (Prepared)
 */
function updateStockInternal(machineName, action, amount) {
  const sheet = getSheet('stock');
  if (!sheet) return;

  const values = sheet.getDataRange().getValues();
  var targetCol = 4;  // Default: Prepared (D)
  if (action === 'assembly' || action === 'dispatch') {
    targetCol = 2;  // Switch to Ready (B)
  }

  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == machineName) {
      var currentVal = Number(values[i][targetCol - 1]) || 0;
      var newVal = currentVal + amount;
      if (newVal < 0) newVal = 0;
      sheet.getRange(i + 1, targetCol).setValue(newVal);
      return;  // Done
    }
  }
  // Machine not found — silently fail (preserves old behavior)
  // Could optionally log to a "StockLog" sheet for debugging
}

// ==================== REVERSE-LOOKUP HELPER ====================

/**
 * lookupMachineInStageSheet() — pattern-agnostic exact-match reverse-lookup.
 * Uses TextFinder for FAST native search (Google's backend) instead of reading
 * the entire sheet. Returns "SS CIDM" or "MS CSEM" etc., or null if not found.
 *
 * TextFinder is much faster than reading 1700+ rows × 13 cols of data into
 * a 2D array and iterating in JS — typical lookup goes from 2-5s down to <500ms.
 */
function lookupMachineInStageSheet(stage, serialToFind) {
  try {
    const sheet = getSheet(stage);
    if (!sheet) return null;

    const lastCol = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2 || lastCol < 2) return null;

    // Read all headers ONCE (small read, fast)
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Use TextFinder for FAST native search — Google's backend does the search
    // instead of us reading 22,000+ cells into memory and iterating.
    const textFinder = sheet.createTextFinder(serialToFind);
    textFinder.matchCase(true);  // serial is uppercase, match exactly
    const matches = textFinder.findAll();

    for (const match of matches) {
      const col = match.getColumn();
      const row = match.getRow();
      if (row === 1) continue;  // skip header row (defensive)

      const header = headers[col - 1];  // 0-indexed
      if (isJunkColumn(header)) continue;

      // Found the serial in a non-junk machine column — return the machine name
      const prefix = serialToFind.startsWith("SSKM") ? "SS" : "MS";
      return `${prefix} ${header}`;
    }
    return null;  // serial not found in any active machine column
  } catch (e) {
    return null;
  }
}

// ==================== UTILITY FUNCTIONS ====================

/**
 * findColumnByHeader() — returns 1-indexed column number matching the header.
 * Returns -1 if not found. Replaces the old hardcoded colMap.
 */
function findColumnByHeader(sheet, headerToFind) {
  if (!sheet) return -1;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === headerToFind) {
      return i + 1;  // 1-indexed
    }
  }
  return -1;
}

/**
 * patternsOverlap() — checks if two regex patterns could match the same string.
 * Simple heuristic: extracts literal prefixes from patterns and checks overlap.
 * Returns true if overlap is likely, false otherwise.
 * Conservative: if can't determine, returns false (allows the addition).
 */
function patternsOverlap(p1, p2) {
  if (p1 === '.*' || p2 === '.*') return true;  // wildcard matches anything
  try {
    // Extract prefix literals (very simplified)
    const getPrefix = (pattern) => {
      const match = pattern.match(/^\^?([A-Za-z]+)/);
      return match ? match[1] : '';
    };
    const prefix1 = getPrefix(p1);
    const prefix2 = getPrefix(p2);
    if (prefix1 && prefix2) {
      return prefix1.startsWith(prefix2) || prefix2.startsWith(prefix1);
    }
    return false;  // Conservative: don't flag as overlap if can't determine
  } catch (e) {
    return false;
  }
}

/**
 * getSheet() — opens a sheet by stage name ('wiring', 'assembly', etc.)
 * CACHES open spreadsheets per request to avoid repeated opens.
 */
const _sheetCache = {};
function getSheet(stage) {
  if (_sheetCache[stage]) return _sheetCache[stage];
  const sheetId = SHEET_IDS[stage];
  const sheetName = SHEET_NAMES[stage];
  if (!sheetId || !sheetName) return null;
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) _sheetCache[stage] = sheet;
    return sheet;
  } catch (e) {
    return null;
  }
}

/**
 * getRegistrySheet() — opens the MachineRegistry sheet
 */
function getRegistrySheet() {
  return getSheet('registry');
}

/**
 * formatDate() — formats a Date object as YYYY-MM-DD string
 */
function formatDate(dateObj) {
  if (!dateObj) return '';
  if (typeof dateObj === 'string') return dateObj;
  const d = new Date(dateObj);
  if (isNaN(d.getTime())) return '';
  return d.toISOString().split('T')[0];
}

/**
 * textResponse() — for no-cors POST endpoints (scanners)
 * Scanners can't read the response anyway, but we return plain text
 * so testing via browser or curl shows the result.
 */
function textResponse(text) {
  return ContentService.createTextOutput(text);
}

/**
 * jsonResponse() — for text/plain POST endpoints (data-entry app)
 * Returns JSON with proper mime type so the data-entry app can parse it.
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== ONE-TIME SETUP HELPER ====================

/**
 * setupMachineRegistry() — run this ONCE from the Apps Script editor
 * to create the MachineRegistry sheet (tab) inside the stock spreadsheet
 * and pre-populate it with the current 10 machines (5 models × 2 materials).
 *
 * To use: open this script in the Apps Script editor, select this function
 * from the function dropdown at the top, click "Run". Authorize when prompted.
 * After running, the MachineRegistry tab will exist with 10 rows pre-filled.
 */
function setupMachineRegistry() {
  const ss = SpreadsheetApp.openById(SHEET_IDS.registry);
  let sheet = ss.getSheetByName(SHEET_NAMES.registry);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.registry);
  } else {
    // Clear existing data (be careful — only run this on first setup)
    sheet.clearContents();
  }

  // Headers
  const headers = [
    'machine_id', 'model_name', 'material_type', 'serial_prefix',
    'serial_pattern', 'category', 'added_date', 'status', 'min_stock'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');

  // Pre-populate with current 10 machines (5 models × 2 materials)
  const models = ['CIDM', 'CSEM', 'CSIM', 'CCBM', 'CIDM110'];
  const materials = ['SS', 'MS'];
  const today = new Date();
  const rows = [];
  for (const model of models) {
    for (const material of materials) {
      const machineId = `${model}-${material}`;
      const serialPrefix = material === 'SS' ? 'SSKM' : 'KM';
      const serialPattern = material === 'SS' ? '^SSKM\\d+$' : '^KM\\d+$';
      rows.push([
        machineId,    // A: machine_id
        model,        // B: model_name
        material,     // C: material_type
        serialPrefix, // D: serial_prefix
        serialPattern,// E: serial_pattern
        material,     // F: category (= material for now)
        today,        // G: added_date
        'active',     // H: status
        5             // I: min_stock (default, adjust per machine)
      ]);
    }
  }
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  // Format the date column
  sheet.getRange(2, 7, rows.length, 1).setNumberFormat('yyyy-MM-dd');

  // Auto-resize columns for readability
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }

  Logger.log(`MachineRegistry sheet created with ${rows.length} machines`);
  Logger.log('Sheet URL: ' + ss.getUrl());
}
