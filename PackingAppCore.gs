/**
 * =================================================================
 * PACKING APP CORE — Apps Script for packing verification app
 * =================================================================
 * Worker takes photo of packed box (with courier label visible).
 * App OCRs the label, extracts fields (customer, address, items, etc.),
 * worker confirms, photo is sent to Telegram bot for storage.
 *
 * Sheet structure (12 columns in "PackingLog" tab):
 *   A: Timestamp | B: InvoiceNumber | C: CustomerName | D: CustomerAddress
 *   E: CustomerPhone | F: Pincode | G: LabelDate | H: ItemsPacked (JSON)
 *   I: TelegramMessageId | J: OCRConfidence | K: Notes | L: CourierCompany
 *
 * Also uses a "Settings" tab for retention period.
 * =================================================================
 */

// ==================== CONFIGURATION ====================

const SHEET_ID = '1fILq8yK7n-S7NavTWDKZswyQAToKnCp_eGQcF_fTlVY';
const SHEET_NAME = 'PackingLog';
const SETTINGS_SHEET_NAME = 'Settings';

// Telegram bot credentials (reused from wiring HTML)
const TG_TOKEN = '7801890257:AAGnOw2gbplPToKKJLLtwzDz6SHj2GLkCb4';
const TG_CHAT = '8341470050';

// ==================== ROUTER ====================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action || '').toLowerCase();
  switch (action) {
    case 'getrecentpackings':
      return getRecentPackings(parseInt(e.parameter.limit) || 20);
    case 'getpacking':
      return getPacking(e.parameter.rowId);
    case 'getcouriersuggestions':
      return getCourierSuggestions();
    case 'getsettings':
      return getSettings();
    default:
      return jsonResponse({ status: 'error', message: 'Unknown GET action: ' + action });
  }
}

function doPost(e) {
  let bodyData = {};
  try {
    if (e && e.postData && e.postData.contents) {
      bodyData = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return textResponse('Error: Invalid JSON body');
  }

  const action = ((e && e.parameter && e.parameter.action) || bodyData.action || '').toLowerCase();

  switch (action) {
    case 'logpacking':
      return logPacking(bodyData);
    case 'updatepacking':
      return updatePacking(bodyData);
    case 'deletepacking':
      return deletePacking(bodyData);
    case 'updatesettings':
      return updateSettings(bodyData);
    default:
      return textResponse('Error: Unknown action: ' + action);
  }
}

function doOptions(e) {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

// ==================== LOG PACKING (main endpoint) ====================

/**
 * Receives packing data + photo from the app.
 * Sends photo to Telegram bot, stores message_id + all data in sheet.
 *
 * Body: {
 *   invoiceNumber, customerName, customerAddress, customerPhone,
 *   pincode, labelDate, itemsPacked (JSON array), ocrConfidence,
 *   notes, courierCompany, photoBase64
 * }
 */
function logPacking(data) {
  try {
    if (!data.invoiceNumber || !data.photoBase64) {
      return jsonResponse({ status: 'error', message: 'Missing required fields: invoiceNumber, photoBase64' });
    }

    // Decode photo from base64
    const photoBytes = Utilities.base64Decode(data.photoBase64);
    const photoBlob = Utilities.newBlob(photoBytes, 'image/jpeg', 'packing.jpg');

    // Send photo to Telegram bot
    const formData = {
      'chat_id': TG_CHAT,
      'caption': '📦 Packing: ' + data.invoiceNumber + (data.courierCompany ? ' | ' + data.courierCompany : '')
    };

    const boundary = '----FormBoundary' + Utilities.getUuid();
    let body = '';
    for (const key in formData) {
      body += '--' + boundary + '\r\n';
      body += 'Content-Disposition: form-data; name="' + key + '"\r\n\r\n';
      body += formData[key] + '\r\n';
    }
    body += '--' + boundary + '\r\n';
    body += 'Content-Disposition: form-data; name="photo"; filename="packing.jpg"\r\n';
    body += 'Content-Type: image/jpeg\r\n\r\n';

    const photoBinary = Utilities.newBlob(body).getBytes().concat(photoBytes);
    const closingBoundary = Utilities.newBlob('\r\n--' + boundary + '--\r\n').getBytes();
    const fullBody = photoBinary.concat(closingBoundary);

    const tgResponse = UrlFetchApp.fetch(
      'https://api.telegram.org/bot' + TG_TOKEN + '/sendPhoto',
      {
        method: 'post',
        contentType: 'multipart/form-data; boundary=' + boundary,
        payload: Utilities.newBlob(fullBody).getBytes(),
        muteHttpExceptions: true
      }
    );

    const tgData = JSON.parse(tgResponse.getContentText());
    if (!tgData.ok) {
      return jsonResponse({ status: 'error', message: 'Telegram upload failed: ' + (tgData.description || 'Unknown') });
    }

    const messageId = tgData.result.message_id;
    const photoFileId = tgData.result.photo ? tgData.result.photo[tgData.result.photo.length - 1].file_id : '';

    // Log to sheet
    const sheet = getSheet(SHEET_NAME);
    if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found: ' + SHEET_NAME });

    const now = new Date();
    const itemsJson = typeof data.itemsPacked === 'string' ? data.itemsPacked : JSON.stringify(data.itemsPacked || []);

    sheet.appendRow([
      now,                                          // A: Timestamp
      data.invoiceNumber || '',                     // B: InvoiceNumber
      data.customerName || '',                      // C: CustomerName
      data.customerAddress || '',                   // D: CustomerAddress
      data.customerPhone || '',                     // E: CustomerPhone
      data.pincode || '',                           // F: Pincode
      data.labelDate || '',                         // G: LabelDate
      itemsJson,                                    // H: ItemsPacked
      String(messageId),                            // I: TelegramMessageId
      data.ocrConfidence || 'unknown',              // J: OCRConfidence
      data.notes || '',                             // K: Notes
      data.courierCompany || ''                     // L: CourierCompany
    ]);

    const rowId = sheet.getLastRow();

    return jsonResponse({
      status: 'success',
      rowId: rowId,
      telegramMessageId: String(messageId),
      photoFileId: photoFileId,
      timestamp: now.toISOString()
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== GET RECENT PACKINGS (for history view) ====================

function getRecentPackings(limit) {
  try {
    const sheet = getSheet(SHEET_NAME);
    if (!sheet) return jsonResponse({ packings: [] });

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return jsonResponse({ packings: [] });

    const packings = [];
    const maxRows = Math.min(limit, data.length - 1);

    for (let i = 0; i < maxRows; i++) {
      const rowIdx = data.length - 1 - i;
      const row = data[rowIdx];
      packings.push({
        rowId: rowIdx + 1,
        timestamp: row[0] ? new Date(row[0]).toISOString() : '',
        invoiceNumber: row[1],
        customerName: row[2],
        customerAddress: row[3],
        customerPhone: row[4],
        pincode: row[5],
        labelDate: row[6],
        itemsPacked: row[7] ? parseJsonSafe(row[7]) : [],
        telegramMessageId: row[8],
        ocrConfidence: row[9],
        notes: row[10],
        courierCompany: row[11]
      });
    }
    return jsonResponse({ packings: packings });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function getPacking(rowId) {
  try {
    const sheet = getSheet(SHEET_NAME);
    if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found' });

    const rowIdx = parseInt(rowId);
    if (isNaN(rowIdx) || rowIdx < 2) return jsonResponse({ status: 'error', message: 'Invalid rowId' });

    const row = sheet.getRange(rowIdx, 1, 1, 12).getValues()[0];
    return jsonResponse({
      packing: {
        rowId: rowId,
        timestamp: row[0] ? new Date(row[0]).toISOString() : '',
        invoiceNumber: row[1],
        customerName: row[2],
        customerAddress: row[3],
        customerPhone: row[4],
        pincode: row[5],
        labelDate: row[6],
        itemsPacked: row[7] ? parseJsonSafe(row[7]) : [],
        telegramMessageId: row[8],
        ocrConfidence: row[9],
        notes: row[10],
        courierCompany: row[11]
      }
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== UPDATE PACKING (edit existing record) ====================

function updatePacking(data) {
  try {
    if (!data.rowId) return jsonResponse({ status: 'error', message: 'Missing rowId' });

    const sheet = getSheet(SHEET_NAME);
    if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found' });

    const rowIdx = parseInt(data.rowId);
    if (isNaN(rowIdx) || rowIdx < 2) return jsonResponse({ status: 'error', message: 'Invalid rowId' });

    const itemsJson = typeof data.itemsPacked === 'string' ? data.itemsPacked : JSON.stringify(data.itemsPacked || []);

    sheet.getRange(rowIdx, 1, 1, 12).setValues([[
      new Date(),                                  // A: Timestamp (update on edit)
      data.invoiceNumber || '',                    // B
      data.customerName || '',                     // C
      data.customerAddress || '',                  // D
      data.customerPhone || '',                    // E
      data.pincode || '',                          // F
      data.labelDate || '',                        // G
      itemsJson,                                   // H
      data.telegramMessageId || '',                // I (preserve existing)
      data.ocrConfidence || 'edited',              // J
      data.notes || '',                            // K
      data.courierCompany || ''                    // L
    ]]);

    return jsonResponse({ status: 'success', rowId: rowIdx });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== DELETE PACKING (delete record + Telegram message) ====================

function deletePacking(data) {
  try {
    if (!data.rowId) return jsonResponse({ status: 'error', message: 'Missing rowId' });

    const sheet = getSheet(SHEET_NAME);
    if (!sheet) return jsonResponse({ status: 'error', message: 'Sheet not found' });

    const rowIdx = parseInt(data.rowId);
    if (isNaN(rowIdx) || rowIdx < 2) return jsonResponse({ status: 'error', message: 'Invalid rowId' });

    // Get telegram message_id before deleting
    const row = sheet.getRange(rowIdx, 1, 1, 12).getValues()[0];
    const telegramMessageId = row[8];

    // Delete Telegram message if exists
    if (telegramMessageId) {
      try {
        UrlFetchApp.fetch(
          'https://api.telegram.org/bot' + TG_TOKEN + '/deleteMessage?chat_id=' + TG_CHAT + '&message_id=' + telegramMessageId,
          { method: 'post', muteHttpExceptions: true }
        );
      } catch (e) {
        // Telegram delete might fail if message already deleted — continue anyway
      }
    }

    // Delete the row from sheet
    sheet.deleteRow(rowIdx);

    return jsonResponse({ status: 'success', rowId: rowIdx });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== COURIER SUGGESTIONS (autocomplete) ====================

function getCourierSuggestions() {
  try {
    const sheet = getSheet(SHEET_NAME);
    if (!sheet) return jsonResponse({ suggestions: [] });

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return jsonResponse({ suggestions: [] });

    const counts = {};
    for (let i = 1; i < data.length; i++) {
      const courier = data[i][11]; // Column L
      if (courier && courier.toString().trim()) {
        const name = courier.toString().trim().toUpperCase();
        counts[name] = (counts[name] || 0) + 1;
      }
    }

    // Sort by frequency (most used first)
    const suggestions = Object.keys(counts)
      .map(name => ({ name: name, count: counts[name] }))
      .sort((a, b) => b.count - a.count);

    return jsonResponse({ suggestions: suggestions });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== SETTINGS (retention period for auto-delete) ====================

function getSettings() {
  try {
    let sheet = getSheet(SETTINGS_SHEET_NAME);
    if (!sheet) {
      // Auto-create Settings tab if missing (defensive)
      sheet = ensureSettingsSheet();
      if (!sheet) {
        return jsonResponse({
          settings: { retentionDays: 21, autoDeleteEnabled: true, lastCleanup: null }
        });
      }
    }

    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 0; i < data.length; i++) {
      if (data[i][0]) settings[data[i][0]] = data[i][1];
    }

    return jsonResponse({
      settings: {
        retentionDays: parseInt(settings.retentionDays) || 21,
        autoDeleteEnabled: settings.autoDeleteEnabled !== 'false',
        lastCleanup: settings.lastCleanup || null
      }
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function updateSettings(data) {
  try {
    let sheet = getSheet(SETTINGS_SHEET_NAME);
    if (!sheet) {
      // Auto-create Settings tab if missing (defensive — fixes "Settings sheet not found" error)
      sheet = ensureSettingsSheet();
      if (!sheet) return jsonResponse({ status: 'error', message: 'Could not create Settings tab' });
    }

    // Update or insert settings
    const retentionDays = parseInt(data.retentionDays);
    if (isNaN(retentionDays)) return jsonResponse({ status: 'error', message: 'Invalid retentionDays' });

    const existingData = sheet.getDataRange().getValues();
    let foundRetention = false;
    let foundAutoDelete = false;

    for (let i = 0; i < existingData.length; i++) {
      if (existingData[i][0] === 'retentionDays') {
        sheet.getRange(i + 1, 2).setValue(retentionDays);
        foundRetention = true;
      }
      if (existingData[i][0] === 'autoDeleteEnabled') {
        sheet.getRange(i + 1, 2).setValue(retentionDays === 0 ? 'false' : 'true');
        foundAutoDelete = true;
      }
    }

    if (!foundRetention) {
      sheet.appendRow(['retentionDays', retentionDays]);
    }
    if (!foundAutoDelete) {
      sheet.appendRow(['autoDeleteEnabled', retentionDays === 0 ? 'false' : 'true']);
    }

    return jsonResponse({ status: 'success', settings: { retentionDays: retentionDays, autoDeleteEnabled: retentionDays !== 0 } });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * ensureSettingsSheet() — creates the Settings tab if it doesn't exist.
 * Called defensively by getSettings() and updateSettings().
 * Returns the sheet object, or null on failure.
 */
function ensureSettingsSheet() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (sheet) return sheet;

    // Create the Settings tab with default values
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    sheet.appendRow(['retentionDays', 21]);
    sheet.appendRow(['autoDeleteEnabled', 'true']);
    sheet.appendRow(['lastCleanup', '']);

    // Auto-resize
    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);

    Logger.log('Created Settings tab: ' + sheet.getName());
    return sheet;
  } catch (e) {
    Logger.log('Failed to create Settings tab: ' + e.toString());
    return null;
  }
}

// ==================== AUTO-DELETE TRIGGER (runs daily) ====================

/**
 * Runs daily via time-driven trigger.
 * Reads retention period from Settings tab.
 * Deletes packings older than retention period (calls Telegram deleteMessage + removes sheet row).
 */
function cleanupOldPackings() {
  try {
    const settingsResp = getSettings();
    const settingsData = JSON.parse(settingsResp.getContentText());
    const settings = settingsData.settings || {};

    if (!settings.autoDeleteEnabled || settings.retentionDays === 0) {
      return; // Auto-delete disabled
    }

    const retentionDays = settings.retentionDays;
    const cutoffDate = new Date(Date.now() - retentionDays * 24 * 60 * 60 * 1000);

    const sheet = getSheet(SHEET_NAME);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      const timestamp = data[i][0] ? new Date(data[i][0]) : null;
      if (timestamp && timestamp < cutoffDate) {
        rowsToDelete.push({
          rowIndex: i + 1, // 1-indexed
          telegramMessageId: data[i][8]
        });
      }
    }

    // Delete from bottom to top (to preserve row indices)
    rowsToDelete.sort((a, b) => b.rowIndex - a.rowIndex);

    for (const item of rowsToDelete) {
      // Delete Telegram message
      if (item.telegramMessageId) {
        try {
          UrlFetchApp.fetch(
            'https://api.telegram.org/bot' + TG_TOKEN + '/deleteMessage?chat_id=' + TG_CHAT + '&message_id=' + item.telegramMessageId,
            { method: 'post', muteHttpExceptions: true }
          );
        } catch (e) {}
      }
      // Delete sheet row
      sheet.deleteRow(item.rowIndex);
    }

    // Update last cleanup timestamp
    const settingsSheet = getSheet(SETTINGS_SHEET_NAME);
    if (settingsSheet) {
      const settingsData = settingsSheet.getDataRange().getValues();
      let foundLastCleanup = false;
      for (let i = 0; i < settingsData.length; i++) {
        if (settingsData[i][0] === 'lastCleanup') {
          settingsSheet.getRange(i + 1, 2).setValue(new Date().toISOString());
          foundLastCleanup = true;
          break;
        }
      }
      if (!foundLastCleanup) {
        settingsSheet.appendRow(['lastCleanup', new Date().toISOString()]);
      }
    }

    Logger.log('Cleanup: deleted ' + rowsToDelete.length + ' old packings (older than ' + retentionDays + ' days)');
  } catch (err) {
    Logger.log('Cleanup error: ' + err.toString());
  }
}

// ==================== ONE-TIME SETUP HELPER ====================

/**
 * Run this ONCE from the Apps Script editor to:
 * 1. Ensure the PackingLog tab has correct headers
 * 2. Create the Settings tab with default values
 * 3. Create the daily cleanup trigger
 *
 * After running, deploy as Web App (Deploy → New deployment → Web app → "Anyone" access).
 */
function setupSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // === Setup PackingLog tab ===
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Check if headers are already set (skip if already correct)
  const existingHeaders = sheet.getRange(1, 1, 1, 12).getValues()[0];
  const expectedHeaders = ['Timestamp', 'InvoiceNumber', 'CustomerName', 'CustomerAddress', 'CustomerPhone', 'Pincode', 'LabelDate', 'ItemsPacked', 'TelegramMessageId', 'OCRConfidence', 'Notes', 'CourierCompany'];

  const headersMatch = existingHeaders.length === 12 && expectedHeaders.every((h, i) => existingHeaders[i] === h);

  if (!headersMatch) {
    sheet.getRange(1, 1, 1, 12).setValues([expectedHeaders]);
    sheet.getRange(1, 1, 1, 12).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Auto-resize columns
  for (let i = 1; i <= 12; i++) {
    sheet.autoResizeColumn(i);
  }

  // === Setup Settings tab ===
  let settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
  }

  const settingsData = settingsSheet.getDataRange().getValues();
  if (settingsData.length < 2 || !settingsData[0][0]) {
    settingsSheet.clearContents();
    settingsSheet.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    settingsSheet.appendRow(['retentionDays', 21]);
    settingsSheet.appendRow(['autoDeleteEnabled', 'true']);
    settingsSheet.appendRow(['lastCleanup', '']);
  }

  // === Setup daily cleanup trigger (runs at 2 AM every day) ===
  const triggers = ScriptApp.getProjectTriggers();
  let hasCleanupTrigger = false;
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'cleanupOldPackings') {
      hasCleanupTrigger = true;
      break;
    }
  }
  if (!hasCleanupTrigger) {
    ScriptApp.newTrigger('cleanupOldPackings')
      .timeBased()
      .everyDays(1)
      .atHour(2)
      .create();
    Logger.log('Created daily cleanup trigger (runs at 2 AM)');
  } else {
    Logger.log('Cleanup trigger already exists');
  }

  Logger.log('=== Setup complete ===');
  Logger.log('Sheet: ' + ss.getUrl());
  Logger.log('PackingLog tab: ' + sheet.getName());
  Logger.log('Settings tab: ' + settingsSheet.getName());
  Logger.log('');
  Logger.log('Now deploy as Web App:');
  Logger.log('  Deploy → New deployment → Web app');
  Logger.log('  Execute as: Me');
  Logger.log('  Who has access: Anyone');
}

// ==================== UTILITY FUNCTIONS ====================

function getSheet(name) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    return ss.getSheetByName(name);
  } catch (e) {
    return null;
  }
}

function parseJsonSafe(str) {
  try {
    return JSON.parse(str);
  } catch (e) {
    return [];
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function textResponse(text) {
  return ContentService.createTextOutput(text);
}
