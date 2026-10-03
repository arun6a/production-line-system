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
  DAILY_NOTES: 'DailyNotes',
  FCM_TOKENS: 'FCMTokens'
};

// ==================== FCM (Firebase Cloud Messaging) ====================
// Firebase credentials stored in Script Properties (NOT in code, for security).
// Apps Script editor -> Project Settings (gear icon) -> Script Properties -> Add:
//   FIREBASE_PROJECT_ID    = km-demo-tracker
//   FIREBASE_CLIENT_EMAIL  = firebase-adminsdk-fbsvc@km-demo-tracker.iam.gserviceaccount.com
//   FIREBASE_PRIVATE_KEY   = <paste the entire private_key value from the service-account JSON, including the PEM BEGIN/END markers and \n newlines>
//
// To get the JSON: Firebase console -> Project Settings -> Service Accounts -> Generate new private key.
// The downloaded JSON's private_key field has the un-redacted key (with real PEM markers).
function getFirebaseConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    PROJECT_ID: props.getProperty('FIREBASE_PROJECT_ID') || 'km-demo-tracker',
    CLIENT_EMAIL: props.getProperty('FIREBASE_CLIENT_EMAIL') || '',
    PRIVATE_KEY: props.getProperty('FIREBASE_PRIVATE_KEY') || ''
  };
}

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
    case 'startdemo': return startDemo(bodyData);
    case 'checkout': return checkOut(bodyData);
    case 'updatevisit': return updateVisit(bodyData);
    case 'addmachine': return addMachine(bodyData);
    case 'addattachment': return addAttachment(bodyData);
    case 'savedailynotes': return saveDailyNotes(bodyData);
    case 'updateredoemo': return updateReDemo(bodyData);
    case 'registerfcmtoken': return registerFCMToken(bodyData);
    case 'unregisterfcmtoken': return unregisterFCMToken(bodyData);
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
    const waitStartTime = today.toISOString();

    // Build machines demoed string (planned at check-in)
    const machinesDemoed = Array.isArray(data.machinesDemoed) ? data.machinesDemoed.join(', ') : (data.machinesDemoed || '');
    const attachmentsDemoed = Array.isArray(data.attachmentsDemoed) ? data.attachmentsDemoed.join(', ') : (data.attachmentsDemoed || '');

    // Customer enters WAITING state — StartTime (F) and EndTime (G) left empty.
    // WaitStartTime (O) is set so the wait timer starts immediately.
    // OriginalRowId (Q) — if this is a re-demo, links to the original visit's row.
    sheet.appendRow([
      dateStr,                                   // A: Date
      customerNum,                               // B: CustomerNum
      data.fromLocation || '',                    // C: FromLocation (optional)
      machinesDemoed,                             // D: MachinesDemoed (planned)
      attachmentsDemoed,                          // E: AttachmentsDemoed (planned)
      '',                                        // F: StartTime (demo start) — empty = waiting
      '',                                        // G: EndTime
      '',                                        // H: DemoMinutes
      '',                                        // I: Outcome
      '',                                        // J: OutcomeNotes
      '',                                        // K: MachinesPurchased
      '',                                        // L: AttachmentsPurchased
      data.returningCustomer ? 'Yes' : 'No',     // M: ReturningCustomer
      '',                                        // N: ReDemoAfterBilling
      waitStartTime,                             // O: WaitStartTime
      '',                                        // P: WaitMinutes
      data.originalRowId ? String(data.originalRowId) : ''  // Q: OriginalRowId
    ]);

    const rowId = sheet.getLastRow();
    lock.releaseLock();

    // If this is a re-demo, update the original visit's ReDemoAfterBilling field
    // to mark it as "Re-demoed as Customer N on <date>"
    if (data.originalRowId) {
      try {
        const originalRowId = parseInt(data.originalRowId);
        const reDemoStatus = 'Yes - Re-demoed as Customer ' + customerNum + ' on ' + dateStr;
        sheet.getRange(originalRowId, 14).setValue(reDemoStatus);   // N: ReDemoAfterBilling
        Logger.log('Re-demo link: original row ' + originalRowId + ' → new Customer ' + customerNum);
      } catch (linkErr) {
        Logger.log('Non-fatal: could not update original row re-demo status: ' + linkErr.toString());
      }
    }

    // Send FCM notification: "Customer N checked in (Loc). X in waiting. Y customer demo running."
    try {
      // Compute current waiting + active counts for today
      const allData2 = sheet.getDataRange().getValues();
      let waitCount = 0, activeCount = 0;
      for (let i = 1; i < allData2.length; i++) {
        if (formatDate(allData2[i][0]) !== dateStr) continue;
        const st = allData2[i][5];        // F: StartTime (demo start)
        const wt = allData2[i][14];       // O: WaitStartTime
        const et = allData2[i][6];        // G: EndTime
        if ((!et || et === '') && st && st !== '') activeCount++;
        else if ((!et || et === '') && wt && wt !== '') waitCount++;
      }
      const locStr = data.fromLocation ? `(${data.fromLocation})` : '';
      const title = `Customer ${customerNum} checked in`;
      const body = `Customer ${customerNum} checked in ${locStr}. ${waitCount} in waiting. ${activeCount} customer demo running.`;
      sendFCMNotification(title, body, { type: 'checkin', customerNum: customerNum });
    } catch (fcmErr) {
      Logger.log('FCM notification failed (non-fatal): ' + fcmErr.toString());
    }

    return jsonResponse({
      status: 'success',
      customerNum: customerNum,
      rowId: rowId,
      waitStartTime: waitStartTime,
      originalRowId: data.originalRowId || null
    });
  } catch (err) {
    lock.releaseLock();
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== START DEMO ====================
// Moves a customer from WAITING to ACTIVE.
// Sets StartTime (F) = now, calculates WaitMinutes (P) = now - WaitStartTime.

function startDemo(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (!data.rowId) return jsonResponse({ status: 'error', message: 'Missing rowId' });
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ status: 'error', message: 'DemoLog sheet not found' });

    const rowId = parseInt(data.rowId);
    const now = new Date();
    const startTime = now.toISOString();

    // Read existing wait start time to calculate wait minutes
    const waitStartStr = sheet.getRange(rowId, 15).getValue();   // O: WaitStartTime
    let waitMinutes = 0;
    if (waitStartStr) {
      const ws = new Date(waitStartStr);
      waitMinutes = Math.round((now - ws) / 60000);
    }

    // Update StartTime (F) and WaitMinutes (P)
    sheet.getRange(rowId, 6).setValue(startTime);          // F: StartTime (demo start)
    sheet.getRange(rowId, 16).setValue(waitMinutes);       // P: WaitMinutes
    lock.releaseLock();

    return jsonResponse({
      status: 'success',
      startTime: startTime,
      waitMinutes: waitMinutes
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

    // Update purchase strings
    const machinesPurchased = Array.isArray(data.machinesPurchased) ? data.machinesPurchased.join(', ') : (data.machinesPurchased || '');
    const attachmentsPurchased = Array.isArray(data.attachmentsPurchased) ? data.attachmentsPurchased.join(', ') : (data.attachmentsPurchased || '');

    // Update row
    sheet.getRange(rowId, 7).setValue(endTime);                    // G: EndTime
    sheet.getRange(rowId, 8).setValue(demoMinutes);                // H: DemoMinutes
    if (data.outcome) sheet.getRange(rowId, 9).setValue(data.outcome);              // I: Outcome (optional)
    if (data.outcomeNotes !== undefined) sheet.getRange(rowId, 10).setValue(data.outcomeNotes);  // J: OutcomeNotes
    sheet.getRange(rowId, 11).setValue(machinesPurchased);         // K: MachinesPurchased
    sheet.getRange(rowId, 12).setValue(attachmentsPurchased);      // L: AttachmentsPurchased

    // Merge "additional demoed" with existing (NOT overwrite — preserves check-in selections)
    // D: MachinesDemoed
    if (data.machinesDemoed) {
      const existingStr = (sheet.getRange(rowId, 4).getValue() || '').toString();
      const existingList = existingStr.split(',').map(function(s){return s.trim();}).filter(Boolean);
      const additionalList = Array.isArray(data.machinesDemoed) ? data.machinesDemoed : [data.machinesDemoed];
      const merged = [];
      const seen = {};
      [].concat(existingList, additionalList).forEach(function(m) {
        if (m && !seen[m]) { seen[m] = true; merged.push(m); }
      });
      sheet.getRange(rowId, 4).setValue(merged.join(', '));
    }
    // E: AttachmentsDemoed
    if (data.attachmentsDemoed) {
      const existingStr = (sheet.getRange(rowId, 5).getValue() || '').toString();
      const existingList = existingStr.split(',').map(function(s){return s.trim();}).filter(Boolean);
      const additionalList = Array.isArray(data.attachmentsDemoed) ? data.attachmentsDemoed : [data.attachmentsDemoed];
      const merged = [];
      const seen = {};
      [].concat(existingList, additionalList).forEach(function(a) {
        if (a && !seen[a]) { seen[a] = true; merged.push(a); }
      });
      sheet.getRange(rowId, 5).setValue(merged.join(', '));
    }

    // Send FCM notification: "Customer N complete demo, need to talk to owner and billing"
    try {
      const custNum = sheet.getRange(rowId, 2).getValue();
      const title = `Customer ${custNum} complete demo`;
      const body = `Customer ${custNum} complete demo, need to talk to owner and billing. (${fmtElapsed(demoMinutes)} demo)`;
      sendFCMNotification(title, body, { type: 'checkout', customerNum: custNum });
    } catch (fcmErr) {
      Logger.log('FCM checkout notification failed (non-fatal): ' + fcmErr.toString());
    }

    return jsonResponse({
      status: 'success',
      demoMinutes: demoMinutes
    });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// Helper for formatting minutes as "Xm" or "Xh Ym"
function fmtElapsed(min) {
  if (min < 60) return min + 'm';
  const h = Math.floor(min/60), m = min%60;
  return h + 'h' + (m > 0 ? ' ' + m + 'm' : '');
}

// ==================== UPDATE VISIT (edit from dashboard) ====================
// Allows owner to edit any field of a completed (or active) visit.
// Only fields provided in `data` will be updated; others are left as-is.

function updateVisit(data) {
  try {
    if (!data.rowId) return jsonResponse({ status: 'error', message: 'Missing rowId' });
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ status: 'error', message: 'DemoLog sheet not found' });

    const rowId = parseInt(data.rowId);

    // Update each field only if provided
    // C: FromLocation
    if (data.fromLocation !== undefined) {
      sheet.getRange(rowId, 3).setValue(data.fromLocation);
    }
    // D: MachinesDemoed
    if (data.machinesDemoed !== undefined) {
      const v = Array.isArray(data.machinesDemoed) ? data.machinesDemoed.join(', ') : (data.machinesDemoed || '');
      sheet.getRange(rowId, 4).setValue(v);
    }
    // E: AttachmentsDemoed
    if (data.attachmentsDemoed !== undefined) {
      const v = Array.isArray(data.attachmentsDemoed) ? data.attachmentsDemoed.join(', ') : (data.attachmentsDemoed || '');
      sheet.getRange(rowId, 5).setValue(v);
    }
    // I: Outcome (optional — allow setting to empty string)
    if (data.outcome !== undefined) {
      sheet.getRange(rowId, 9).setValue(data.outcome);
    }
    // J: OutcomeNotes
    if (data.outcomeNotes !== undefined) {
      sheet.getRange(rowId, 10).setValue(data.outcomeNotes);
    }
    // K: MachinesPurchased
    if (data.machinesPurchased !== undefined) {
      const v = Array.isArray(data.machinesPurchased) ? data.machinesPurchased.join(', ') : (data.machinesPurchased || '');
      sheet.getRange(rowId, 11).setValue(v);
    }
    // L: AttachmentsPurchased
    if (data.attachmentsPurchased !== undefined) {
      const v = Array.isArray(data.attachmentsPurchased) ? data.attachmentsPurchased.join(', ') : (data.attachmentsPurchased || '');
      sheet.getRange(rowId, 12).setValue(v);
    }
    // M: ReturningCustomer
    if (data.returningCustomer !== undefined) {
      const v = data.returningCustomer === true || data.returningCustomer === 'true' || data.returningCustomer === 'Yes' ? 'Yes' : 'No';
      sheet.getRange(rowId, 13).setValue(v);
    }
    // N: ReDemoAfterBilling
    if (data.reDemoAfterBilling !== undefined) {
      const v = data.reDemoAfterBilling === true || data.reDemoAfterBilling === 'true' || data.reDemoAfterBilling === 'Yes' ? 'Yes' : 'No';
      sheet.getRange(rowId, 14).setValue(v);
    }

    return jsonResponse({ status: 'success', rowId: rowId });
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

// ==================== GET ACTIVE + WAITING VISITS ====================
// Returns both arrays: active (demo in progress) and waiting (checked in, demo not yet started).

function getActiveVisits() {
  try {
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return jsonResponse({ active: [], waiting: [] });

    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const active = [];
    const waiting = [];

    for (let i = 1; i < data.length; i++) {
      const endTime = data[i][6];        // G
      const startTime = data[i][5];      // F (demo start)
      const waitStartTime = data[i][14]; // O (check-in time)

      if (!endTime || endTime === '') {
        if (startTime && startTime !== '') {
          // ACTIVE: demo started, not yet checked out
          const start = new Date(startTime);
          const elapsed = Math.round((now - start) / 60000);
          active.push({
            rowId: i + 1,
            date: formatDate(data[i][0]),
            customerNum: data[i][1],
            fromLocation: data[i][2],
            machinesDemoed: data[i][3],
            attachmentsDemoed: data[i][4],
            startTime: startTime,
            elapsedMinutes: elapsed,
            returningCustomer: data[i][12],
            reDemo: data[i][13],
            waitMinutes: data[i][15],
            originalRowId: data[i][16]
          });
        } else if (waitStartTime && waitStartTime !== '') {
          // WAITING: checked in but demo not yet started
          const ws = new Date(waitStartTime);
          const waited = Math.round((now - ws) / 60000);
          waiting.push({
            rowId: i + 1,
            date: formatDate(data[i][0]),
            customerNum: data[i][1],
            fromLocation: data[i][2],
            machinesDemoed: data[i][3],
            attachmentsDemoed: data[i][4],
            waitStartTime: waitStartTime,
            waitedMinutes: waited,
            returningCustomer: data[i][12],
            originalRowId: data[i][16]
          });
        }
      }
    }
    return jsonResponse({ active: active, waiting: waiting });
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

    let totalVisits = 0, activeCount = 0, waitingCount = 0, purchased = 0, buyingLater = 0, quotation = 0;
    let courier = 0, justVisiting = 0, wrongMachine = 0, empty = 0;
    let totalDemoMinutes = 0, totalWaitMinutes = 0;
    let activeElapsedMinutes = 0, waitingElapsedMinutes = 0;

    for (let i = 1; i < data.length; i++) {
      if (formatDate(data[i][0]) !== today) continue;
      totalVisits++;

      const endTime = data[i][6];
      const startTime = data[i][5] ? new Date(data[i][5]) : null;
      const waitStart = data[i][14] ? new Date(data[i][14]) : null;

      if (!endTime || endTime === '') {
        // Not yet checked out
        if (startTime) {
          // Active demo
          activeCount++;
          if (startTime) {
            activeElapsedMinutes += Math.round((now - startTime) / 60000);
          }
        } else if (waitStart) {
          // Waiting to start demo
          waitingCount++;
          waitingElapsedMinutes += Math.round((now - waitStart) / 60000);
        }
      } else {
        // Completed
        const demoMin = parseInt(data[i][7]) || 0;
        totalDemoMinutes += demoMin;
      }

      // Always count wait minutes (set at startDemo time)
      const wMin = parseInt(data[i][15]) || 0;
      totalWaitMinutes += wMin;

      const outcome = (data[i][8] || '').toLowerCase();
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

    // Hourly breakdown (based on demo start time, falling back to wait start if demo not yet started)
    const hourly = {};
    for (let h = 9; h <= 18; h++) hourly[h + ':00'] = 0;
    for (let i = 1; i < data.length; i++) {
      if (formatDate(data[i][0]) !== today) continue;
      let t = data[i][5]; // F: StartTime (demo start)
      if (!t) t = data[i][14]; // O: WaitStartTime (check-in time) if demo not yet started
      if (t) {
        const dt = new Date(t);
        const h = dt.getHours();
        const key = h + ':00';
        if (hourly[key] !== undefined) hourly[key]++;
      }
    }

    // Weekly counter: Monday to Saturday of CURRENT week (Sunday rolls back to last Mon-Sat).
    // dayOfWeek: 0=Sunday, 1=Monday, ..., 6=Saturday
    // Monthly counter: visits from 1st of current calendar month to today.
    const dow = now.getDay();
    let weekStart, weekEnd;
    if (dow === 0) {
      // Sunday — show last Monday to last Saturday (i.e., the just-finished week)
      weekStart = new Date(now);
      weekStart.setDate(now.getDate() - 6);   // last Monday
      weekEnd = new Date(now);
      weekEnd.setDate(now.getDate() - 1);     // last Saturday
    } else {
      // Monday–Saturday — weekStart = Monday of this week, weekEnd = today (cap)
      weekStart = new Date(now);
      weekStart.setDate(now.getDate() - (dow - 1));  // Monday
      weekEnd = new Date(now);                       // today (don't count future days)
    }
    const weekStartStr = formatDate(weekStart);
    const weekEndStr = formatDate(weekEnd);
    const monthPrefix = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

    let weeklyVisits = 0, weeklyPurchased = 0, monthlyVisits = 0, monthlyPurchased = 0;
    for (let i = 1; i < data.length; i++) {
      const rowDate = formatDate(data[i][0]);
      if (rowDate >= weekStartStr && rowDate <= weekEndStr) {
        weeklyVisits++;
        if ((data[i][8] || '').toLowerCase() === 'purchased') weeklyPurchased++;
      }
      if (rowDate.substring(0, 7) === monthPrefix && rowDate <= today) {
        monthlyVisits++;
        if ((data[i][8] || '').toLowerCase() === 'purchased') monthlyPurchased++;
      }
    }

    return jsonResponse({
      summary: {
        today: today,
        totalVisits: totalVisits,
        activeCount: activeCount,
        waitingCount: waitingCount,
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
        totalWaitMinutes: totalWaitMinutes,
        waitingElapsedMinutes: waitingElapsedMinutes,
        hourlyBreakdown: hourly,
        weeklyVisits: weeklyVisits,
        weeklyPurchased: weeklyPurchased,
        monthlyVisits: monthlyVisits,
        monthlyPurchased: monthlyPurchased
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
        startTime: row[5],         // F: demo start time
        endTime: row[6],
        demoMinutes: row[7],
        outcome: row[8],
        outcomeNotes: row[9],
        machinesPurchased: row[10],
        attachmentsPurchased: row[11],
        returningCustomer: row[12],
        reDemo: row[13],
        waitStartTime: row[14],   // O: check-in time
        waitMinutes: row[15],     // P: wait duration
        originalRowId: row[16]   // Q: links to original visit (for re-demo)
      });
    }
    return jsonResponse({
      today: formatDate(new Date()),
      visits: visits
    });
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
      'ReturningCustomer', 'ReDemoAfterBilling', 'WaitStartTime', 'WaitMinutes', 'OriginalRowId'];
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

  // ---- Step 4b: FCMTokens (for push notifications) ----
  try {
    Logger.log('[4b] Setting up FCMTokens...');
    const fcmHeaders = ['device_id', 'fcm_token', 'registered_at', 'last_seen_at', 'device_name', 'platform', 'status'];
    const fcmSheet = ensureTab(TABS.FCM_TOKENS, fcmHeaders, fcmHeaders.length);
    for (let i = 1; i <= fcmHeaders.length; i++) fcmSheet.autoResizeColumn(i);
    Logger.log('[4b] OK — FCMTokens ready');
  } catch (e) {
    Logger.log('[4b] FAILED — FCMTokens: ' + e.toString());
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

// ==================== FCM REGISTRATION ====================

function registerFCMToken(data) {
  try {
    if (!data.deviceId || !data.fcmToken) {
      return jsonResponse({ status: 'error', message: 'Missing deviceId or fcmToken' });
    }
    const sheet = getSheet(TABS.FCM_TOKENS);
    if (!sheet) return jsonResponse({ status: 'error', message: 'FCMTokens sheet not found' });

    const nowStr = new Date().toISOString();
    const allData = sheet.getDataRange().getValues();

    // Check if device_id already exists
    let existingRow = -1;
    let existingToken = '';
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.deviceId) {
        existingRow = i + 1;
        existingToken = allData[i][1] || '';
        break;
      }
    }

    if (existingRow > 0) {
      // Update existing row (also check if token changed)
      sheet.getRange(existingRow, 2).setValue(data.fcmToken);            // B: fcm_token
      sheet.getRange(existingRow, 3).setValue(nowStr);                    // C: registered_at
      sheet.getRange(existingRow, 4).setValue(nowStr);                    // D: last_seen_at
      sheet.getRange(existingRow, 5).setValue(data.deviceName || '');    // E: device_name
      sheet.getRange(existingRow, 6).setValue(data.platform || '');     // F: platform
      sheet.getRange(existingRow, 7).setValue('active');                 // G: status
      Logger.log('FCM: Updated existing device row ' + existingRow + ' (token ' + (existingToken === data.fcmToken ? 'unchanged' : 'CHANGED') + ')');
      return jsonResponse({ status: 'success', action: 'updated', tokenChanged: existingToken !== data.fcmToken });
    } else {
      // Insert new row
      sheet.appendRow([
        data.deviceId,         // A
        data.fcmToken,         // B
        nowStr,                // C: registered_at
        nowStr,                // D: last_seen_at
        data.deviceName || '', // E
        data.platform || '',  // F
        'active'               // G
      ]);
      Logger.log('FCM: Registered new device ' + data.deviceId);
      return jsonResponse({ status: 'success', action: 'created' });
    }
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function unregisterFCMToken(data) {
  try {
    if (!data.deviceId) return jsonResponse({ status: 'error', message: 'Missing deviceId' });
    const sheet = getSheet(TABS.FCM_TOKENS);
    if (!sheet) return jsonResponse({ status: 'error', message: 'FCMTokens sheet not found' });
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === data.deviceId) {
        sheet.getRange(i + 1, 7).setValue('unregistered');  // G: status
        sheet.getRange(i + 1, 4).setValue(new Date().toISOString());  // D: last_seen_at
        return jsonResponse({ status: 'success' });
      }
    }
    return jsonResponse({ status: 'error', message: 'Device not found' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

// ==================== FCM NOTIFICATION SENDING ====================
// Uses FCM HTTP v1 API with OAuth2 JWT (service account) auth.

function getAllFCMTokens() {
  const sheet = getSheet(TABS.FCM_TOKENS);
  if (!sheet) return [];
  const allData = sheet.getDataRange().getValues();
  const tokens = [];
  for (let i = 1; i < allData.length; i++) {
    const status = allData[i][6] || 'active';
    if (status === 'active' && allData[i][1]) {
      tokens.push({
        deviceId: allData[i][0],
        token: allData[i][1],
        rowIdx: i + 1
      });
    }
  }
  return tokens;
}

// Get OAuth2 access token by signing a JWT with the Firebase service account private key
function getFCMAccessToken() {
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: 'RS256', typ: 'JWT' };
  const payload = {
    iss: getFirebaseConfig().CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  };

  const encHeader = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=+$/, '');
  const encPayload = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/, '');
  const toSign = encHeader + '.' + encPayload;

  // Sign with RSA-SHA256 using the private key
  const signature = Utilities.computeRsaSha256Signature(toSign, getFirebaseConfig().PRIVATE_KEY.replace(/\\n/g, '\n'));
  const encSignature = Utilities.base64EncodeWebSafe(signature).replace(/=+$/, '');

  const jwt = toSign + '.' + encSignature;

  // Exchange JWT for access token
  const tokenResp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });
  const tokenJson = JSON.parse(tokenResp.getContentText());
  if (!tokenJson.access_token) {
    throw new Error('Failed to get FCM access token: ' + tokenResp.getContentText());
  }
  return tokenJson.access_token;
}

// Send FCM push notification to ALL registered devices
function sendFCMNotification(title, body, data) {
  try {
    const tokens = getAllFCMTokens();
    if (tokens.length === 0) {
      Logger.log('FCM: No registered devices, skipping notification');
      return;
    }
    const accessToken = getFCMAccessToken();
    const projectId = getFirebaseConfig().PROJECT_ID;
    const url = 'https://fcm.googleapis.com/v1/projects/' + projectId + '/messages:send';

    let sentCount = 0, failCount = 0;
    for (let i = 0; i < tokens.length; i++) {
      const t = tokens[i];
      const message = {
        message: {
          token: t.token,
          notification: { title: title, body: body },
          data: (function() {
            const d = {};
            if (data) for (const k in data) d[k] = String(data[k]);
            return d;
          })(),
          android: {
            priority: 'high',
            notification: { channel_id: 'km_demo_default', sound: 'default' }
          }
        }
      };
      try {
        const resp = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: { Authorization: 'Bearer ' + accessToken },
          payload: JSON.stringify(message),
          muteHttpExceptions: true
        });
        if (resp.getResponseCode() === 200) {
          sentCount++;
        } else {
          // 404 = token not registered (unregister on server)
          if (resp.getResponseCode() === 404 || resp.getContentText().indexOf('UNREGISTERED') >= 0) {
            const fcmSheet = getSheet(TABS.FCM_TOKENS);
            if (fcmSheet) fcmSheet.getRange(t.rowIdx, 7).setValue('unregistered');
          }
          Logger.log('FCM send failed for device ' + t.deviceId + ': ' + resp.getContentText());
          failCount++;
        }
      } catch (e) {
        failCount++;
        Logger.log('FCM send error for device ' + t.deviceId + ': ' + e.toString());
      }
    }
    Logger.log('FCM: sent=' + sentCount + ' failed=' + failCount + ' total=' + tokens.length);
  } catch (e) {
    Logger.log('sendFCMNotification error: ' + e.toString());
  }
}

// ==================== DAILY 7PM SUMMARY (time-driven trigger) ====================
// Set up trigger: Apps Script editor → Triggers → Add Trigger
//   Function: sendDailySummary
//   Event source: Time-driven
//   Type: Day timer
//   Time of day: 7pm to 8pm

function sendDailySummary() {
  try {
    const sheet = getSheet(TABS.DEMO_LOG);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const today = formatDate(new Date());
    let visitedCount = 0;
    let purchasedCount = 0;
    let totalDemoMin = 0;
    for (let i = 1; i < data.length; i++) {
      if (formatDate(data[i][0]) !== today) continue;
      visitedCount++;
      const demoMin = parseInt(data[i][7]) || 0;
      totalDemoMin += demoMin;
      if ((data[i][8] || '').toLowerCase() === 'purchased') purchasedCount++;
    }
    const title = 'Daily Summary — ' + visitedCount + ' customer' + (visitedCount === 1 ? '' : 's') + ' visited';
    const body = visitedCount + ' customer' + (visitedCount === 1 ? '' : 's') + ' visited today. ' +
      purchasedCount + ' purchased. Total demo time: ' + fmtElapsed(totalDemoMin) + '.';
    sendFCMNotification(title, body, { type: 'daily_summary', date: today });
    Logger.log('Daily summary sent: ' + title + ' — ' + body);
  } catch (e) {
    Logger.log('sendDailySummary error: ' + e.toString());
  }
}


// ==================== TEST FCM (diagnostic) ====================
function testFCM() {
  const props = PropertiesService.getScriptProperties();
  Logger.log('1. FIREBASE_PROJECT_ID: ' + props.getProperty('FIREBASE_PROJECT_ID'));
  Logger.log('2. FIREBASE_CLIENT_EMAIL: ' + props.getProperty('FIREBASE_CLIENT_EMAIL'));
  const pk = props.getProperty('FIREBASE_PRIVATE_KEY');
  Logger.log('3. FIREBASE_PRIVATE_KEY: ' + (pk ? 'set, length=' + pk.length + ', starts with: ' + pk.substring(0, 30) : 'NOT SET'));
  if (pk && pk.indexOf('BEGIN PRIVATE KEY') < 0) {
    Logger.log('ERROR: Private key missing BEGIN marker. Re-download JSON from Firebase.');
    return;
  }
  const tokens = getAllFCMTokens();
  Logger.log('4. Registered device tokens: ' + tokens.length);
  if (tokens.length === 0) {
    Logger.log('NOTE: Open the APK on your phone, then re-run testFCM.');
    return;
  }
  try {
    const at = getFCMAccessToken();
    Logger.log('5. FCM access token: ' + at.substring(0, 30) + '...');
  } catch (e) {
    Logger.log('5. FAILED: ' + e.toString());
    return;
  }
  sendFCMNotification('KM Test', 'If you see this, FCM works!', { type: 'test' });
  Logger.log('6. Test sent. Check phone.');
}
