/**
 * =================================================================
 * DEMO APP CORE — Customer Demo Visit Tracking System
 * =================================================================
 * Tracks walk-in customer demos at the company demo room.
 * Staff checks in customers, runs demo timers, checks out with outcome.
 * Owner views real-time dashboard.
 *
 * Tabs:
 *   1. DemoLog — one row per customer visit
 *   2. AttachmentRegistry — list of attachments (dynamically addable)
 *   3. MachineList — list of model names (dynamically addable)
 *   4. DailyNotes — end-of-day other work notes
 * =================================================================
 */

const SHEET_ID = '1hEaCA46B7ichGeHgOvtWpS7RtHPuGaHba4C8rRVVOIA';

const TABS = {
  DEMO_LOG: 'DemoLog',
  ATTACHMENTS: 'AttachmentRegistry',
  MACHINES: 'MachineList',
  DAILY_NOTES: 'DailyNotes'
};

// ==================== ROUTER ====================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action || '').toLowerCase();
  switch (action) {
    case 'getactivevisits': return getActiveVisits();
    case 'gettodaysummary': return getTodaySummary();
    case 'getrecentvisits': return getRecentVisits(parseInt(e.parameter.limit) || 50);
    case 'getmachines': return getMachines();
    case 'getattachments': return getAttachments();
    case 'getdailynotes': return getDailyNotes();
    case 'getweeklytrend': return getWeeklyTrend();
    default: return jsonResponse({ status: 'error', message: 'Unknown action: ' + action });
  }
}

function doPost(e) {
  let bodyData = {};
  try {
    if (e && e.postData && e.postData.contents) {
      bodyData = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return textResponse('Error: Invalid JSON');
  }
  const action = ((e && e.parameter && e.parameter.action) || bodyData.action || '').toLowerCase();
  switch (action) {
    case 'checkin': return checkIn(bodyData);
    case 'checkout': return checkOut(bodyData);
    case 'addmachine': return addMachine(bodyData);
    case 'addattachment': return addAttachment(bodyData);
    case 'savedailynotes': return saveDailyNotes(bodyData);
    case 'updateredoemo': return updateReDemo(bodyData);
    default: return textResponse('Error: Unknown action: ' + action);
  }
}

function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

// ==================== CHECK-IN ====================

function checkIn(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ status: 'error', message: 'DemoLog sheet not found' });

    const today = new Date();
    const dateStr = formatDate(today);

    // Find next customer number for today
    const allData = sheet.getDataRange().getValues();
    let maxNum = 0;
    for (let i = 1; i < allData.length; i++) {
      if (formatDate(allData[i][0]) === dateStr) {
        const num = parseInt(allData[i][1]) || 0;
        if (num > maxNum) maxNum = num;
      }
    }
    const customerNum = maxNum + 1;
    const startTime = today.toISOString();

    // Build machines demoed string
    const machinesDemoed = Array.isArray(data.machinesDemoed) ? data.machinesDemoed.join(', ') : (data.machinesDemoed || '');
    const attachmentsDemoed = Array.isArray(data.attachmentsDemoed) ? data.attachmentsDemoed.join(', ') : (data.attachmentsDemoed || '');

    sheet.appendRow([
      dateStr,                    // A: Date
      customerNum,               // B: CustomerNum
      data.fromLocation || '',    // C: FromLocation
      machinesDemoed,            // D: MachinesDemoed
      attachmentsDemoed,         // E: AttachmentsDemoed
      startTime,                 // F: StartTime
      '',                        // G: EndTime (empty = active)
      '',                        // H: DemoMinutes (calculated at checkout)
      '',                        // I: Outcome
      '',                        // J: OutcomeNotes
      '',                        // K: MachinesPurchased
      '',                        // L: AttachmentsPurchased
      data.returningCustomer ? 'Yes' : 'No',  // M: ReturningCustomer
      ''                         // N: ReDemoAfterBilling
    ]);

    const rowId = sheet.getLastRow();
    lock.releaseLock();

    return jsonResponse({
      status: 'success',
      customerNum: customerNum,
      rowId: rowId,
      startTime: startTime
    });
  } catch (err) {
    lock.releaseLock();
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== CHECK-OUT ====================

function checkOut(data) {
  try {
    if (!data.rowId) return jsonResponse({ status: 'error', message: 'Missing rowId' });
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ status: 'error', message: 'DemoLog sheet not found' });

    const rowId = parseInt(data.rowId);
    const endTime = new Date().toISOString();

    // Get start time to calculate demo minutes
    const startStr = sheet.getRange(rowId, 6).getValue();
    let demoMinutes = 0;
    if (startStr) {
      const start = new Date(startStr);
      const end = new Date(endTime);
      demoMinutes = Math.round((end - start) / 60000);
    }

    // Build purchase strings
    const machinesPurchased = Array.isArray(data.machinesPurchased) ? data.machinesPurchased.join(', ') : (data.machinesPurchased || '');
    const attachmentsPurchased = Array.isArray(data.attachmentsPurchased) ? data.attachmentsPurchased.join(', ') : (data.attachmentsPurchased || '');

    // Update row
    sheet.getRange(rowId, 7).setValue(endTime);                    // G: EndTime
    sheet.getRange(rowId, 8).setValue(demoMinutes);                // H: DemoMinutes
    sheet.getRange(rowId, 9).setValue(data.outcome || '');         // I: Outcome
    sheet.getRange(rowId, 10).setValue(data.outcomeNotes || '');   // J: OutcomeNotes
    sheet.getRange(rowId, 11).setValue(machinesPurchased);         // K: MachinesPurchased
    sheet.getRange(rowId, 12).setValue(attachmentsPurchased);       // L: AttachmentsPurchased

    // Also update attachments demoed if changed
    if (data.attachmentsDemoed) {
      const attDemoed = Array.isArray(data.attachmentsDemoed) ? data.attachmentsDemoed.join(', ') : data.attachmentsDemoed;
      sheet.getRange(rowId, 5).setValue(attDemoed);
    }

    return jsonResponse({
      status: 'success',
      demoMinutes: demoMinutes
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== UPDATE RE-DEMO ====================

function updateReDemo(data) {
  try {
    if (!data.rowId) return jsonResponse({ status: 'error', message: 'Missing rowId' });
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ status: 'error', message: 'DemoLog sheet not found' });

    const rowId = parseInt(data.rowId);
    sheet.getRange(rowId, 14).setValue(data.reDemoNotes || 'Re-demo requested at ' + new Date().toLocaleTimeString());

    return jsonResponse({ status: 'success' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== GET ACTIVE VISITS ====================

function getActiveVisits() {
  try {
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ visits: [] });

    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const visits = [];

    for (let i = 1; i < data.length; i++) {
      const endTime = data[i][6]; // column G
      if (!endTime || endTime === '') {
        // Active visit (no end time)
        const startTime = data[i][5] ? new Date(data[i][5]) : null;
        let elapsedMinutes = 0;
        if (startTime) {
          elapsedMinutes = Math.round((now - startTime) / 60000);
        }
        visits.push({
          rowId: i + 1,
          date: formatDate(data[i][0]),
          customerNum: data[i][1],
          fromLocation: data[i][2],
          machinesDemoed: data[i][3],
          attachmentsDemoed: data[i][4],
          startTime: data[i][5],
          elapsedMinutes: elapsedMinutes,
          returningCustomer: data[i][12],
          reDemo: data[i][13]
        });
      }
    }
    return jsonResponse({ visits: visits });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== GET TODAY SUMMARY ====================

function getTodaySummary() {
  try {
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ summary: {} });

    const data = sheet.getDataRange().getValues();
    const today = formatDate(new Date());
    const now = new Date();

    let totalVisits = 0, activeCount = 0, purchased = 0, buyingLater = 0, quotation = 0;
    let courier = 0, justVisiting = 0, wrongMachine = 0, empty = 0;
    let totalDemoMinutes = 0;
    let activeElapsedMinutes = 0;

    for (let i = 1; i < data.length; i++) {
      if (formatDate(data[i][0]) !== today) continue;
      totalVisits++;

      const endTime = data[i][6];
      const startTime = data[i][5] ? new Date(data[i][5]) : null;

      if (!endTime || endTime === '') {
        // Active
        activeCount++;
        if (startTime) {
          activeElapsedMinutes += Math.round((now - startTime) / 60000);
        }
      } else {
        // Completed
        const demoMin = parseInt(data[i][7]) || 0;
        totalDemoMinutes += demoMin;
      }

      const outcome = data[i][8] || '';
      switch (outcome.toLowerCase()) {
        case 'purchased': purchased++; break;
        case 'buying later': buyingLater++; break;
        case 'quotation': quotation++; break;
        case 'courier to home': courier++; break;
        case 'just visiting': justVisiting++; break;
        case 'wrong machine': wrongMachine++; break;
        default: if (!outcome) empty++; break;
      }
    }

    // Hourly breakdown
    const hourly = {};
    for (let h = 9; h <= 18; h++) hourly[h + ':00'] = 0;
    for (let i = 1; i < data.length; i++) {
      if (formatDate(data[i][0]) !== today) continue;
      const startTime = data[i][5] ? new Date(data[i][5]) : null;
      if (startTime) {
        const h = startTime.getHours();
        const key = h + ':00';
        if (hourly[key] !== undefined) hourly[key]++;
      }
    }

    return jsonResponse({
      summary: {
        totalVisits: totalVisits,
        activeCount: activeCount,
        purchased: purchased,
        buyingLater: buyingLater,
        quotation: quotation,
        courier: courier,
        justVisiting: justVisiting,
        wrongMachine: wrongMachine,
        empty: empty,
        conversionRate: totalVisits > 0 ? Math.round((purchased / totalVisits) * 100) : 0,
        totalDemoMinutes: totalDemoMinutes,
        activeElapsedMinutes: activeElapsedMinutes,
        hourlyBreakdown: hourly
      }
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== GET RECENT VISITS ====================

function getRecentVisits(limit) {
  try {
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ visits: [] });

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return jsonResponse({ visits: [] });

    const visits = [];
    const maxRows = Math.min(limit, data.length - 1);

    for (let i = 0; i < maxRows; i++) {
      const rowIdx = data.length - 1 - i;
      const row = data[rowIdx];
      visits.push({
        rowId: rowIdx + 1,
        date: formatDate(row[0]),
        customerNum: row[1],
        fromLocation: row[2],
        machinesDemoed: row[3],
        attachmentsDemoed: row[4],
        startTime: row[5],
        endTime: row[6],
        demoMinutes: row[7],
        outcome: row[8],
        outcomeNotes: row[9],
        machinesPurchased: row[10],
        attachmentsPurchased: row[11],
        returningCustomer: row[12],
        reDemo: row[13]
      });
    }
    return jsonResponse({ visits: visits });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== GET WEEKLY TREND ====================

function getWeeklyTrend() {
  try {
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ trend: [] });

    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const trend = [];

    for (let d = 6; d >= 0; d--) {
      const day = new Date(now);
      day.setDate(now.getDate() - d);
      const dateStr = formatDate(day);
      let visits = 0, purchased = 0, totalMin = 0;

      for (let i = 1; i < data.length; i++) {
        if (formatDate(data[i][0]) === dateStr) {
          visits++;
          if ((data[i][8] || '').toLowerCase() === 'purchased') purchased++;
          totalMin += parseInt(data[i][7]) || 0;
        }
      }
      trend.push({
        date: dateStr,
        dayName: day.toLocaleDateString('en-US', { weekday: 'short' }),
        visits: visits,
        purchased: purchased,
        conversionRate: visits > 0 ? Math.round((purchased / visits) * 100) : 0,
        totalDemoMinutes: totalMin
      });
    }
    return jsonResponse({ trend: trend });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== MACHINE LIST ====================

function getMachines() {
  try {
    const sheet = getSheet(TABS.MACHINES);
    if (!sheet) return jsonResponse({ machines: [] });
    const data = sheet.getDataRange().getValues();
    const machines = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && (data[i][1] || 'active') === 'active') {
        machines.push(data[i][0]);
      }
    }
    return jsonResponse({ machines: machines });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function addMachine(data) {
  try {
    if (!data.modelName) return jsonResponse({ status: 'error', message: 'Missing modelName' });
    const sheet = getSheet(TABS.MACHINES);
    if (!sheet) return jsonResponse({ status: 'error', message: 'MachineList sheet not found' });
    
    // Check if already exists
    const existing = sheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if (existing[i][0] === data.modelName) {
        return jsonResponse({ status: 'error', message: 'Machine already exists' });
      }
    }
    sheet.appendRow([data.modelName, 'active']);
    return jsonResponse({ status: 'success', modelName: data.modelName });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== ATTACHMENT REGISTRY ====================

function getAttachments() {
  try {
    const sheet = getSheet(TABS.ATTACHMENTS);
    if (!sheet) return jsonResponse({ attachments: [] });
    const data = sheet.getDataRange().getValues();
    const attachments = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && (data[i][2] || 'active') === 'active') {
        attachments.push(data[i][1] || data[i][0]);
      }
    }
    return jsonResponse({ attachments: attachments });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function addAttachment(data) {
  try {
    if (!data.attachmentName) return jsonResponse({ status: 'error', message: 'Missing attachmentName' });
    const sheet = getSheet(TABS.ATTACHMENTS);
    if (!sheet) return jsonResponse({ status: 'error', message: 'AttachmentRegistry sheet not found' });
    
    const existing = sheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if ((existing[i][1] || existing[i][0]) === data.attachmentName) {
        return jsonResponse({ status: 'error', message: 'Attachment already exists' });
      }
    }
    const id = 'ATT-' + String(existing.length).padStart(3, '0');
    sheet.appendRow([id, data.attachmentName, 'active']);
    return jsonResponse({ status: 'success', attachmentName: data.attachmentName });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== DAILY NOTES ====================

function getDailyNotes() {
  try {
    const sheet = getSheet(TABS.DAILY_NOTES);
    if (!sheet) return jsonResponse({ notes: '' });
    const data = sheet.getDataRange().getValues();
    const today = formatDate(new Date());
    for (let i = 1; i < data.length; i++) {
      if (formatDate(data[i][0]) === today) {
        return jsonResponse({ notes: data[i][1] || '', date: today });
      }
    }
    return jsonResponse({ notes: '', date: today });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function saveDailyNotes(data) {
  try {
    if (!data.notes && data.notes !== '') return jsonResponse({ status: 'error', message: 'Missing notes' });
    const sheet = getSheet(TABS.DAILY_NOTES);
    if (!sheet) return jsonResponse({ status: 'error', message: 'DailyNotes sheet not found' });

    const today = formatDate(new Date());
    const existing = sheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if (formatDate(existing[i][0]) === today) {
        sheet.getRange(i + 1, 2).setValue(data.notes);
        return jsonResponse({ status: 'success' });
      }
    }
    sheet.appendRow([today, data.notes]);
    return jsonResponse({ status: 'success' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== DIAGNOSTIC (RUN FIRST) ====================
// Run testConnection() BEFORE setupSheet() to verify:
//   1) The script has been authorized
//   2) The SHEET_ID is correct and accessible
// Returns a clear message you can read in the Execution Log.

function testConnection() {
  Logger.log('--- testConnection START ---');
  Logger.log('SHEET_ID = ' + SHEET_ID);

  let ss;
  try {
    ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('✓ openById OK — sheet name: "' + ss.getName() + '"');
    Logger.log('  URL: ' + ss.getUrl());
  } catch (e) {
    Logger.log('✗ openById FAILED: ' + e.toString());
    Logger.log('  → Check SHEET_ID, or that your account owns the sheet.');
    return 'FAIL: cannot open spreadsheet — ' + e.toString();
  }

  let sheets;
  try {
    sheets = ss.getSheets();
    Logger.log('✓ getSheets OK — count: ' + sheets.length);
    sheets.forEach(function (s) {
      Logger.log('  - ' + s.getName() + ' (rows: ' + s.getLastRow() + ')');
    });
  } catch (e) {
    Logger.log('✗ getSheets FAILED: ' + e.toString());
    return 'FAIL: cannot list sheets — ' + e.toString();
  }

  Logger.log('--- testConnection OK ---');
  return 'OK: spreadsheet accessible. ' + sheets.length + ' sheets present. Now run setupSheet().';
}

// ==================== ONE-TIME SETUP ====================
// Defensive version — each step wrapped in try-catch with explicit logging,
// so if one step fails you see exactly which one in the Execution Log
// instead of Google's generic "unknown error".

function setupSheet() {
  Logger.log('=== setupSheet START ===');

  // ---- Step 0: open the spreadsheet ----
  let ss;
  try {
    ss = SpreadsheetApp.openById(SHEET_ID);
    Logger.log('[0] Opened spreadsheet: ' + ss.getName());
  } catch (e) {
    Logger.log('[0] FAILED to open spreadsheet: ' + e.toString());
    throw new Error('Cannot open spreadsheet. Run testConnection() first. Details: ' + e.toString());
  }

  // Helper: ensure a tab exists; if it exists, just clear it
  function ensureTab(name, headers, nCols) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      Logger.log('  [+] Created new tab: ' + name);
    } else {
      sh.clearContents();
      Logger.log('  [~] Cleared existing tab: ' + name);
    }
    sh.getRange(1, 1, 1, nCols).setValues([headers]);
    sh.getRange(1, 1, 1, nCols).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }

  // ---- Step 1: DemoLog ----
  try {
    Logger.log('[1] Setting up DemoLog...');
    const logHeaders = ['Date', 'CustomerNum', 'FromLocation', 'MachinesDemoed',
      'AttachmentsDemoed', 'StartTime', 'EndTime', 'DemoMinutes', 'Outcome',
      'OutcomeNotes', 'MachinesPurchased', 'AttachmentsPurchased',
      'ReturningCustomer', 'ReDemoAfterBilling'];
    const logSheet = ensureTab(TABS.DEMO_LOG, logHeaders, logHeaders.length);
    for (let i = 1; i <= logHeaders.length; i++) logSheet.autoResizeColumn(i);
    Logger.log('[1] OK — DemoLog ready');
  } catch (e) {
    Logger.log('[1] FAILED — DemoLog: ' + e.toString());
    throw e;
  }

  // ---- Step 2: MachineList ----
  try {
    Logger.log('[2] Setting up MachineList...');
    const machHeaders = ['model_name', 'status'];
    const machSheet = ensureTab(TABS.MACHINES, machHeaders, 2);
    const defaultMachines = ['CIDM', 'CSEM', 'CSIM', 'CCBM', 'CIDM110', 'AIDM'];
    defaultMachines.forEach(function (m) { machSheet.appendRow([m, 'active']); });
    machSheet.autoResizeColumn(1);
    machSheet.autoResizeColumn(2);
    Logger.log('[2] OK — MachineList ready (6 default machines)');
  } catch (e) {
    Logger.log('[2] FAILED — MachineList: ' + e.toString());
    throw e;
  }

  // ---- Step 3: AttachmentRegistry ----
  try {
    Logger.log('[3] Setting up AttachmentRegistry...');
    const attHeaders = ['attachment_id', 'attachment_name', 'status'];
    const attSheet = ensureTab(TABS.ATTACHMENTS, attHeaders, 3);
    attSheet.appendRow(['ATT-001', 'Muruku', 'active']);
    attSheet.appendRow(['ATT-002', 'Porota', 'active']);
    attSheet.autoResizeColumn(1);
    attSheet.autoResizeColumn(2);
    attSheet.autoResizeColumn(3);
    Logger.log('[3] OK — AttachmentRegistry ready (2 default attachments)');
  } catch (e) {
    Logger.log('[3] FAILED — AttachmentRegistry: ' + e.toString());
    throw e;
  }

  // ---- Step 4: DailyNotes ----
  try {
    Logger.log('[4] Setting up DailyNotes...');
    const notesHeaders = ['Date', 'OtherWorkNotes'];
    const notesSheet = ensureTab(TABS.DAILY_NOTES, notesHeaders, 2);
    notesSheet.autoResizeColumn(1);
    notesSheet.autoResizeColumn(2);
    Logger.log('[4] OK — DailyNotes ready');
  } catch (e) {
    Logger.log('[4] FAILED — DailyNotes: ' + e.toString());
    throw e;
  }

  // ---- Step 5: optionally remove the default empty "Sheet1" ----
  // SAFE: only delete if Sheet1 exists, is empty, AND is NOT the only sheet
  // AND is NOT the active sheet (Google rejects deleting the active sheet).
  try {
    Logger.log('[5] Checking for default Sheet1...');
    const sheet1 = ss.getSheetByName('Sheet1');
    const allSheets = ss.getSheets();
    if (sheet1) {
      const isEmpty = sheet1.getLastRow() <= 1 && sheet1.getLastColumn() <= 1;
      const isOnlySheet = allSheets.length <= 1;
      const isActive = ss.getActiveSheet().getSheetId() === sheet1.getSheetId();
      Logger.log('  Sheet1 found. isEmpty=' + isEmpty + ' isOnlySheet=' + isOnlySheet + ' isActive=' + isActive);
      if (isEmpty && !isOnlySheet && !isActive) {
        ss.deleteSheet(sheet1);
        Logger.log('  [−] Deleted empty default Sheet1');
      } else {
        Logger.log('  [skip] Left Sheet1 in place (not safe to delete)');
      }
    } else {
      Logger.log('  No Sheet1 present — nothing to clean up');
    }
  } catch (e) {
    Logger.log('[5] Non-fatal — Sheet1 cleanup skipped: ' + e.toString());
  }

  Logger.log('=== setupSheet COMPLETE ===');
  Logger.log('Sheet URL: ' + ss.getUrl());
  Logger.log('Tabs created: ' + TABS.DEMO_LOG + ', ' + TABS.MACHINES + ', ' + TABS.ATTACHMENTS + ', ' + TABS.DAILY_NOTES);
  Logger.log('');
  Logger.log('NEXT: Deploy → New deployment → Web app');
  Logger.log('  Execute as: Me');
  Logger.log('  Who has access: Anyone');

  return 'Setup complete. 4 tabs ready. Check the Execution Log for details.';
}

// ==================== UTILITIES ====================

function getSheet(name) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    return ss.getSheetByName(name);
  } catch (e) {
    return null;
  }
}

function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'string') return date.substring(0, 10);
  const d = new Date(date);
  if (isNaN(d.getTime())) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + day;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function textResponse(text) {
  return ContentService.createTextOutput(text);
}
