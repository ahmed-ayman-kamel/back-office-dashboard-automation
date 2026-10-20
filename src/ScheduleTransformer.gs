/**
 * ============================================================
 * Agent Activity Tracker — Schedule Transformer  v2
 * ============================================================
 * Reads unstructured schedule data from the 'Raw_Schedule' tab
 * and writes a clean, flat table to the 'CS-Schedule' tab.
 *
 * NEW in v2
 * ─────────
 *  • Hourly time-based trigger — install once, runs forever.
 *  • Safe UI wrapper — script works both from the menu AND
 *    from the Apps Script editor ▶ Run button.
 *  • Execution log tab ('Run_Log') records every run with
 *    timestamp, rows written, and any error message.
 *  • Trigger management helpers (install / remove / status)
 *    exposed as menu items so no code changes are needed.
 *
 * Deployment
 * ──────────
 *  1. Open the Google Sheet → Extensions → Apps Script.
 *  2. Paste this entire file (replace any existing code) → Save.
 *  3. Reload the Sheet — the "⚙ Schedule Tools" menu appears.
 *  4. Click  ⚙ Schedule Tools → Install Hourly Trigger.
 *  5. Grant permissions when prompted (first time only).
 *  6. Done — the script now auto-runs every hour.
 *
 *  To stop auto-runs : ⚙ Schedule Tools → Remove Hourly Trigger.
 *  To check status   : ⚙ Schedule Tools → Trigger Status.
 * ============================================================
 */


// ══════════════════════════════════════════════════════════════
// CONFIGURATION
// ══════════════════════════════════════════════════════════════

var CONFIG = {
  SOURCE_TAB:       'Raw_Schedule',
  DEST_TAB:         'CS-Schedule',
  LOG_TAB:          'Run_Log',          // NEW: execution history tab
  MAX_DATA_COLS:    14,                 // Columns A–N only; ignore O and P
  AGENT_INFO_COLS:  7,                  // Metadata in columns A–G (indices 0–6)
  BLOCK_SIZE:       67,                 // Rows per repeating block
  TRIGGER_FUNC:     'transformSchedule', // Function name the trigger calls
  HOUR_INTERVAL:    1,                  // How often to run (in hours)
  OUTPUT_HEADERS: [
    'Agent Name', 'Email', 'Gender', 'Batch',
    'Skill', 'Seniority', 'Reporting To', 'Date', 'Shift'
  ],
  LOG_HEADERS: [
    'Timestamp', 'Triggered By', 'Blocks Processed',
    'Rows Written', 'Status', 'Message'
  ]
};


// ══════════════════════════════════════════════════════════════
// MENU
// ══════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙ Schedule Tools')
    .addItem('▶ Transform Now',            'transformSchedule')
    .addSeparator()
    .addItem('⏱ Install Hourly Trigger',   'installHourlyTrigger')
    .addItem('🛑 Remove Hourly Trigger',   'removeHourlyTrigger')
    .addItem('ℹ Trigger Status',           'showTriggerStatus')
    .addSeparator()
    .addItem('📋 View Run Log',            'openRunLog')
    .addItem('About',                      'showAbout')
    .addToUi();
}

function showAbout() {
  showMessage(
    'Agent Schedule Transformer v2',
    'Reads Raw_Schedule → writes CS-Schedule.\n\n' +
    '• Blocks of 67 rows are auto-detected.\n' +
    '• Shift ranges ("10:00 ~ 17:00") become start-time only.\n' +
    '• Every run is logged to the Run_Log tab.\n' +
    '• Use "Install Hourly Trigger" to run automatically.'
  );
}

function openRunLog() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var log  = ss.getSheetByName(CONFIG.LOG_TAB);
  if (!log) {
    showMessage('Run Log', 'No runs recorded yet. Run the transform first.');
    return;
  }
  ss.setActiveSheet(log);
}


// ══════════════════════════════════════════════════════════════
// TRIGGER MANAGEMENT
// ══════════════════════════════════════════════════════════════

/**
 * Installs a time-based trigger that fires transformSchedule()
 * every CONFIG.HOUR_INTERVAL hours.
 * Safe to call multiple times — removes any duplicate triggers first.
 */
function installHourlyTrigger() {
  // Remove existing triggers for the same function to avoid duplicates
  _removeTriggers(CONFIG.TRIGGER_FUNC);

  ScriptApp.newTrigger(CONFIG.TRIGGER_FUNC)
    .timeBased()
    .everyHours(CONFIG.HOUR_INTERVAL)
    .create();

  showMessage(
    '⏱ Trigger Installed',
    '"' + CONFIG.TRIGGER_FUNC + '" will now run automatically ' +
    'every ' + CONFIG.HOUR_INTERVAL + ' hour(s).\n\n' +
    'Use "Trigger Status" to confirm.'
  );
}

/**
 * Removes all time-based triggers for transformSchedule().
 */
function removeHourlyTrigger() {
  var removed = _removeTriggers(CONFIG.TRIGGER_FUNC);
  showMessage(
    '🛑 Trigger Removed',
    removed > 0
      ? removed + ' trigger(s) removed. Auto-runs are now OFF.'
      : 'No active triggers found for "' + CONFIG.TRIGGER_FUNC + '".'
  );
}

/**
 * Shows how many triggers are currently installed and when the
 * next run is roughly expected.
 */
function showTriggerStatus() {
  var triggers = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === CONFIG.TRIGGER_FUNC;
  });

  if (triggers.length === 0) {
    showMessage(
      'ℹ Trigger Status',
      'No active triggers for "' + CONFIG.TRIGGER_FUNC + '".\n' +
      'Use "Install Hourly Trigger" to enable auto-runs.'
    );
    return;
  }

  var lines = triggers.map(function (t, i) {
    return (i + 1) + '. Type: ' + t.getTriggerSource() +
           ' | Event: ' + t.getEventType();
  });

  showMessage(
    'ℹ Trigger Status',
    triggers.length + ' active trigger(s) found:\n\n' +
    lines.join('\n') + '\n\n' +
    'The script runs every ' + CONFIG.HOUR_INTERVAL + ' hour(s).'
  );
}

/**
 * Internal helper — deletes all project triggers for a given function.
 * Returns the number of triggers removed.
 * @param {string} funcName
 * @returns {number}
 */
function _removeTriggers(funcName) {
  var triggers = ScriptApp.getProjectTriggers();
  var count = 0;
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  });
  return count;
}


// ══════════════════════════════════════════════════════════════
// MAIN FUNCTION
// ══════════════════════════════════════════════════════════════

/**
 * Reads all data from Raw_Schedule, parses each 67-row block,
 * and writes the flat output to CS-Schedule.
 * Called both by the custom menu and by the hourly trigger.
 */
function transformSchedule() {

  // Detect whether we were called by a trigger (no UI) or by a human
  var calledByTrigger = !_hasUiContext();
  var triggeredBy     = calledByTrigger ? 'Hourly Trigger' : 'Manual Run';

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Validate source tab ──────────────────────────────────────
  var rawSheet = ss.getSheetByName(CONFIG.SOURCE_TAB);
  if (!rawSheet) {
    var errMsg = 'Tab "' + CONFIG.SOURCE_TAB + '" not found.';
    _writeLog(ss, triggeredBy, 0, 0, 'ERROR', errMsg);
    throw new Error(errMsg);
  }

  // ── Get or create destination tab ───────────────────────────
  var csSheet = ss.getSheetByName(CONFIG.DEST_TAB);
  if (!csSheet) {
    csSheet = ss.insertSheet(CONFIG.DEST_TAB);
    Logger.log('Created new tab: ' + CONFIG.DEST_TAB);
  }

  // ── Read raw data ────────────────────────────────────────────
  var lastRow = rawSheet.getLastRow();
  if (lastRow < 2) {
    var emptyMsg = CONFIG.SOURCE_TAB + ' is empty. Nothing to transform.';
    _writeLog(ss, triggeredBy, 0, 0, 'SKIPPED', emptyMsg);
    showMessage('Skipped', emptyMsg);
    return;
  }

  var rawData = rawSheet
    .getRange(1, 1, lastRow, CONFIG.MAX_DATA_COLS)
    .getDisplayValues();

  Logger.log('Rows read: ' + rawData.length + ' | Triggered by: ' + triggeredBy);

  // ── Locate block header rows ─────────────────────────────────
  var blockStarts = [];
  for (var i = 0; i < rawData.length; i++) {
    if (rawData[i][0].trim() === 'Agent Name') {
      blockStarts.push(i);
    }
  }

  if (blockStarts.length === 0) {
    var noBlockMsg = 'No "Agent Name" header rows found in ' + CONFIG.SOURCE_TAB + '.';
    _writeLog(ss, triggeredBy, 0, 0, 'ERROR', noBlockMsg);
    showMessage('Error', noBlockMsg);
    return;
  }

  Logger.log('Blocks found: ' + blockStarts.length);

  // ── Process every block ──────────────────────────────────────
  var outputRows = [];

  for (var b = 0; b < blockStarts.length; b++) {
    var blockStart = blockStarts[b];
    var blockEnd   = (b + 1 < blockStarts.length)
      ? blockStarts[b + 1]
      : rawData.length;

    var headerRow = rawData[blockStart];

    // Identify date columns (pattern "D-Mon" / "DD-Mon")
    var dateCols = [];
    for (var c = CONFIG.AGENT_INFO_COLS; c < CONFIG.MAX_DATA_COLS; c++) {
      var hVal = headerRow[c].trim();
      if (isDateHeader(hVal)) {
        dateCols.push({ index: c, label: hVal });
      }
    }

    if (dateCols.length === 0) {
      Logger.log('Block ' + (b + 1) + ' row ' + (blockStart + 1) + ': no dates — skipped.');
      continue;
    }

    // Iterate data rows inside this block
    for (var row = blockStart + 1; row < blockEnd; row++) {
      var r         = rawData[row];
      var agentName = r[0].trim();

      if (!agentName || agentName === 'Agent Name' || agentName.indexOf('#REF') !== -1) {
        continue;
      }

      var email     = r[1].trim();
      var gender    = r[2].trim();
      var batch     = r[3].trim();
      var skill     = r[4].trim();
      var seniority = r[5].trim();
      var reportTo  = r[6].trim();

      for (var d = 0; d < dateCols.length; d++) {
        var shift = processShift(r[dateCols[d].index]);
        if (shift === '') continue;

        outputRows.push([
          agentName, email, gender, batch,
          skill, seniority, reportTo,
          dateCols[d].label, shift
        ]);
      }
    }
  }

  Logger.log('Output rows: ' + outputRows.length);

  // ── Write to CS-Schedule ─────────────────────────────────────
  csSheet.clearContents();
  csSheet.clearFormats();

  var headRange = csSheet.getRange(1, 1, 1, CONFIG.OUTPUT_HEADERS.length);
  headRange.setValues([CONFIG.OUTPUT_HEADERS]);
  headRange.setFontWeight('bold');
  headRange.setBackground('#4A86E8');
  headRange.setFontColor('#FFFFFF');

  if (outputRows.length > 0) {
    csSheet
      .getRange(2, 1, outputRows.length, CONFIG.OUTPUT_HEADERS.length)
      .setValues(outputRows);
  }

  csSheet.autoResizeColumns(1, CONFIG.OUTPUT_HEADERS.length);

  // ── Write to Run_Log ─────────────────────────────────────────
  _writeLog(
    ss, triggeredBy,
    blockStarts.length, outputRows.length,
    'SUCCESS',
    outputRows.length + ' rows written to ' + CONFIG.DEST_TAB
  );

  // ── Show result (only visible for manual runs) ───────────────
  var summary =
    'Triggered by     : ' + triggeredBy      + '\n' +
    'Blocks processed : ' + blockStarts.length + '\n' +
    'Rows written     : ' + outputRows.length  + '\n' +
    'Destination tab  : ' + CONFIG.DEST_TAB;

  Logger.log('✅ Done.\n' + summary);
  showMessage('✅ Transform Complete', summary);
}


// ══════════════════════════════════════════════════════════════
// RUN LOG
// ══════════════════════════════════════════════════════════════

/**
 * Appends one row to the Run_Log tab.
 * Creates the tab with headers if it doesn't exist yet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} triggeredBy   - "Manual Run" or "Hourly Trigger"
 * @param {number} blocks        - Number of blocks processed
 * @param {number} rows          - Number of rows written
 * @param {string} status        - "SUCCESS" | "ERROR" | "SKIPPED"
 * @param {string} message       - Human-readable detail
 */
function _writeLog(ss, triggeredBy, blocks, rows, status, message) {
  var logSheet = ss.getSheetByName(CONFIG.LOG_TAB);

  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_TAB);
    // Write header row
    var hRange = logSheet.getRange(1, 1, 1, CONFIG.LOG_HEADERS.length);
    hRange.setValues([CONFIG.LOG_HEADERS]);
    hRange.setFontWeight('bold');
    hRange.setBackground('#34A853');
    hRange.setFontColor('#FFFFFF');
    logSheet.setFrozenRows(1);
    Logger.log('Created Run_Log tab.');
  }

  var now = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );

  logSheet.appendRow([now, triggeredBy, blocks, rows, status, message]);

  // Colour the Status cell: green = SUCCESS, red = ERROR, yellow = SKIPPED
  var lastRow   = logSheet.getLastRow();
  var statusCol = 5;   // Column E = "Status"
  var cell      = logSheet.getRange(lastRow, statusCol);
  var colours   = { SUCCESS: '#B7E1CD', ERROR: '#F4C7C3', SKIPPED: '#FCE8B2' };
  cell.setBackground(colours[status] || '#FFFFFF');

  logSheet.autoResizeColumns(1, CONFIG.LOG_HEADERS.length);
}


// ══════════════════════════════════════════════════════════════
// SAFE UI WRAPPER
// ══════════════════════════════════════════════════════════════

/**
 * Shows an alert when a UI context is available (menu click),
 * falls back to Logger.log when running headlessly (trigger / editor).
 *
 * @param {string} title
 * @param {string} message
 */
function showMessage(title, message) {
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    // No UI available — log instead (visible in Apps Script → Executions)
    Logger.log('[' + title + '] ' + message);
  }
}

/**
 * Returns true when the script is running inside a browser session
 * (menu click, onOpen, onEdit), false when running headlessly
 * (time-based trigger, editor ▶ Run).
 *
 * @returns {boolean}
 */
function _hasUiContext() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}


// ══════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ══════════════════════════════════════════════════════════════

/**
 * Returns true if a string looks like a "D-Mon" / "DD-Mon" date label.
 * e.g. "1-Jan", "12-Feb", "31-Dec"
 */
function isDateHeader(value) {
  return /^\d{1,2}-[A-Za-z]{3}$/.test(value);
}

/**
 * Cleans a raw shift cell value.
 *
 *  "10:00 ~ 17:00"  →  "10:00"   (range: keep start time only)
 *  "15:00\t15:00"   →  "15:00"   (tab-duplicate: keep start time only)
 *  "OFF" / "Task"   →  unchanged
 *  blank / null     →  ""
 */
function processShift(rawValue) {
  if (rawValue === null || rawValue === undefined) return '';
  var val = String(rawValue).trim();
  if (!val) return '';
  if (val.indexOf(' ~ ') !== -1) return val.substring(0, 5);
  if (val.indexOf('\t')   !== -1) return val.substring(0, 5);
  return val;
}