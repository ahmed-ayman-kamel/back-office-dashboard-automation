// ==========================================
// SECTION 1: CONFIGURATION
// ==========================================

const PHONE_CONFIG = {
  // --- Spreadsheet Target ---
  SHEET_ID: 'YOUR_SLA_DASHBOARD_SHEET_ID',
  TAB_NAME: 'Phone_API',            // CONFIRMED: Your phone data tab name

  // --- Maqsam API ---
  // base_url, access_key_id, access_secret → stored in Script Properties
  // Do NOT hardcode credentials here.

  // --- Sync Behavior ---
  BACKFILL_CHUNK_DAYS: 2,         // Days per backfill chunk (stay under 6-min limit)
  INCREMENTAL_LOOKBACK_MIN: 5,    // Overlap buffer to catch straggler records
  HOURLY_TRIGGER_NAME: 'runHourlyPhoneSync',

  // --- Safety Limits ---
  MAX_ROWS_PER_RUN: 5000,         // Hard cap on rows processed per execution
  MAX_PAGES_PER_RUN: 50,          // Hard cap on API pages fetched per execution
  REQUEST_DELAY_MS: 1200,         // 1.2s between API calls (rate limit courtesy)
  SCRIPT_TIMEOUT_BUFFER_MS: 30000,// Exit 30s before 6-min GAS hard limit
};

// E.164 Prefix → Country Name (expanded as needed)
const COUNTRY_PREFIX_MAP = {
  '966': 'Saudi Arabia',
  '971': 'United Arab Emirates',
  '974': 'Qatar',
  '962': 'Jordan',
  '965': 'Kuwait',
  '968': 'Oman',
  '91':  'India',
  '1':   'United States',
  '20':  'Egypt',
  '973': 'Bahrain',
  '964': 'Iraq',
  '963': 'Syria',
  '961': 'Lebanon',
  '970': 'Palestine',
  '212': 'Morocco',
  '213': 'Algeria',
  '216': 'Tunisia',
  '218': 'Libya',
  '249': 'Sudan',
  '252': 'Somalia',
  '254': 'Kenya',
  '255': 'Tanzania',
  '256': 'Uganda',
  '27':  'South Africa',
  '44':  'United Kingdom',
  '49':  'Germany',
  '33':  'France',
  '39':  'Italy',
  '34':  'Spain',
  '90':  'Turkey',
  '86':  'China',
  '81':  'Japan',
  '82':  'South Korea',
  '7':   'Russia',
  '380': 'Ukraine',
  '48':  'Poland',
  '31':  'Netherlands',
  '32':  'Belgium',
  '41':  'Switzerland',
  '46':  'Sweden',
  '47':  'Norway',
  '45':  'Denmark',
  '358': 'Finland',
  '43':  'Austria',
  '30':  'Greece',
  '351': 'Portugal',
  '353': 'Ireland',
  '36':  'Hungary',
  '420': 'Czech Republic',
  '421': 'Slovakia',
  '40':  'Romania',
  '359': 'Bulgaria',
  '386': 'Slovenia',
  '385': 'Croatia',
  '372': 'Estonia',
  '371': 'Latvia',
  '370': 'Lithuania',
};

// Sheet headers — final column order
const SHEET_HEADERS = [
  'Call ID',        // id
  'Date Time',      // timestamp
  'direction',      // direction ("inbound" / "outbound")
  'Caller',         // callerNumber
  'Callee',         // calleeNumber
  'Country',        // derived: callerNumber if inbound, calleeNumber if outbound
  'Call duration',  // handlingTime
  'Hold Time',      // holdTime
  'Wrapup Time',    // acwTime
  'Wait Duration',  // waitingTime
  'Call Status',    // state
  'Agent name',     // agents[].name
  'Agent email'     // agents[].email
];

// ==========================================
// SECTION 2: MAIN ENTRY POINTS
// ==========================================

/**
 * ENTRY POINT 1: One-time setup
 * Run this ONCE manually to create sheet headers and the hourly trigger.
 * Safe to run multiple times — it's idempotent (won't create duplicates).
 */
function setupPhoneSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('Another instance is running. Skipping setup.');
    return;
  }

  try {
    _ensureHeaders();
    _createHourlyTrigger();
    _logAudit('setupPhoneSync', 'SUCCESS', 'Headers ensured, trigger created');
    Logger.log('=== Setup complete. Headers ready. Hourly trigger active. ===');
    Logger.log('Next: Run runHistoricalBackfill() once to load data from Jan 1, 2026.');
  } catch (e) {
    _logAudit('setupPhoneSync', 'FAILED', e.toString());
    Logger.log('Setup FAILED: ' + e.toString());
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * ENTRY POINT 2: Historical backfill
 * Run this ONCE manually. It fetches all calls from Jan 1, 2026 through now,
 * processing in small date chunks to stay within GAS execution limits.
 * Progress is saved automatically — if it times out, just run it again.
 */
function runHistoricalBackfill() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('Another instance is running. Skipping backfill.');
    return;
  }

  const startTime = Date.now();
  const props = PropertiesService.getScriptProperties();

  try {
    // Backfill configuration
    const BACKFILL_START = new Date('2026-01-01T00:00:00+04:00'); // GST timezone
    const now = new Date();

    // Read resume state (if previous run timed out)
    let chunkStart = _getBackfillState() || BACKFILL_START;
    let totalRows = 0;
    let totalPages = 0;
    let chunksProcessed = 0;

    Logger.log(`=== BACKFILL START === From: ${chunkStart.toISOString()}`);

    while (chunkStart < now) {
      // Check time budget
      if (_isApproachingTimeout(startTime)) {
        Logger.log(`Approaching timeout. Pausing at ${chunkStart.toISOString()}. Run again to resume.`);
        _saveBackfillState(chunkStart);
        _logAudit('runHistoricalBackfill', 'PAUSED', `Processed ${chunksProcessed} chunks, ${totalRows} rows. Resumes at ${chunkStart.toISOString()}`);
        return;
      }

      // Define chunk window
      const chunkEnd = new Date(chunkStart.getTime() + PHONE_CONFIG.BACKFILL_CHUNK_DAYS * 86400000);
      const effectiveEnd = chunkEnd > now ? now : chunkEnd;

      const startUnix = Math.floor(chunkStart.getTime() / 1000);
      const endUnix = Math.floor(effectiveEnd.getTime() / 1000);

      Logger.log(`Chunk: ${chunkStart.toISOString()} → ${effectiveEnd.toISOString()}`);

      // Fetch this chunk
      const result = _fetchAndWriteChunk(startUnix, endUnix, startTime);

      totalRows += result.rows;
      totalPages += result.pages;
      chunksProcessed++;

      if (result.hitRowLimit) {
        Logger.log(`Hit row limit (${PHONE_CONFIG.MAX_ROWS_PER_RUN}). Pausing. Run again to resume.`);
        _saveBackfillState(chunkStart);
        _logAudit('runHistoricalBackfill', 'PAUSED_ROW_LIMIT', `Processed ${chunksProcessed} chunks, ${totalRows} rows`);
        return;
      }

      // Advance to next chunk
      chunkStart = new Date(effectiveEnd.getTime() + 1000); // +1s to avoid overlap
    }

    // Backfill complete
    props.deleteProperty('BACKFILL_RESUME_AT');

    // Set watermark for incremental sync to now
    const watermark = Math.floor(Date.now() / 1000);
    props.setProperty('PHONE_LAST_SYNC_UNIX', String(watermark));

    Logger.log(`=== BACKFILL COMPLETE === ${chunksProcessed} chunks, ${totalPages} pages, ${totalRows} rows`);
    _logAudit('runHistoricalBackfill', 'COMPLETE', `Chunks: ${chunksProcessed}, Pages: ${totalPages}, Inserted: ${totalRows}`);

  } catch (e) {
    _logAudit('runHistoricalBackfill', 'FAILED', e.toString());
    Logger.log('Backfill FAILED: ' + e.toString());
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * ENTRY POINT 3: Hourly incremental sync (called by trigger)
 * Fetches only new/updated records since the last successful sync.
 * The watermark is updated ONLY after successful write.
 */
function runHourlyPhoneSync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('Another instance is running. Skipping hourly sync.');
    return;
  }

  const startTime = Date.now();
  const props = PropertiesService.getScriptProperties();

  try {
    // If backfill hasn't completed, don't run incremental
    const backfillState = props.getProperty('BACKFILL_RESUME_AT');
    if (backfillState) {
      Logger.log('Backfill not complete. Run runHistoricalBackfill() first.');
      return;
    }

    const nowUnix = Math.floor(Date.now() / 1000);

    // Always look back to the START OF TODAY (GST, UTC+4).
    // This ensures calls that were "in_progress" earlier today get re-fetched
    // and their rows updated when they finish — not just newly created calls.
    const todayGST     = new Date();
    const gstOffset    = 4 * 60; // UTC+4 in minutes
    const utcMs        = todayGST.getTime() + (todayGST.getTimezoneOffset() * 60000);
    const gstMs        = utcMs + (gstOffset * 60000);
    const gstDate      = new Date(gstMs);
    gstDate.setHours(0, 0, 0, 0);
    const startOfTodayUnix = Math.floor((gstDate.getTime() - (gstOffset * 60000)) / 1000);

    Logger.log(`=== INCREMENTAL SYNC === From start-of-day: ${new Date(startOfTodayUnix * 1000).toISOString()}`);

    // _fetchAndWriteChunk now does UPSERT: new rows are inserted, existing rows
    // (matched by Call ID) are overwritten in place — this is what updates
    // in_progress calls once they reach a final state.
    const result = _fetchAndWriteChunk(startOfTodayUnix, nowUnix, startTime);

    // Always update watermark after a clean run
    props.setProperty('PHONE_LAST_SYNC_UNIX', String(nowUnix));
    Logger.log(`Watermark updated to: ${new Date(nowUnix * 1000).toISOString()}`);

    Logger.log(`=== INCREMENTAL COMPLETE === Pages: ${result.pages}, New: ${result.inserted}, Updated: ${result.updated}`);
    _logAudit('runHourlyPhoneSync', 'SUCCESS', `Pages: ${result.pages}, New rows: ${result.inserted}, Updated rows: ${result.updated}`);

  } catch (e) {
    _logAudit('runHourlyPhoneSync', 'FAILED', e.toString());
    Logger.log('Hourly sync FAILED: ' + e.toString());
    // Do NOT throw — triggers should not be disabled on error
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// SECTION 3: CORE SYNC ENGINE
// ==========================================

/**
 * Fetches a date-range chunk from Maqsam and writes to the sheet.
 * Returns { rows, pages, hitRowLimit } for progress tracking.
 */
function _fetchAndWriteChunk(startUnix, endUnix, scriptStartTime) {
  const allRows = [];
  let page = 1;
  let hasMore = true;
  let totalPages = 0;
  let hitRowLimit = false;

  while (hasMore && totalPages < PHONE_CONFIG.MAX_PAGES_PER_RUN) {
    // Time check
    if (_isApproachingTimeout(scriptStartTime)) {
      Logger.log('Approaching script timeout. Stopping pagination.');
      break;
    }

    // Row limit check
    if (allRows.length >= PHONE_CONFIG.MAX_ROWS_PER_RUN) {
      hitRowLimit = true;
      break;
    }

    // Fetch page
    const response = _fetchCallsPage(page, startUnix, endUnix);

    if (!response || !response.calls || response.calls.length === 0) {
      hasMore = false;
      break;
    }

    // Transform calls to sheet rows
    const rows = response.calls.map(_mapCallToRow).filter(r => r !== null);
    allRows.push(...rows);

    totalPages++;

    // Check if this was the last page
    if (response.calls.length < 100) {
      hasMore = false;
    } else {
      page++;
      Utilities.sleep(PHONE_CONFIG.REQUEST_DELAY_MS);
    }
  }

  // Upsert: insert new rows, overwrite rows whose Call ID already exists
  let inserted = 0;
  let updated  = 0;
  if (allRows.length > 0) {
    const result = _upsertToSheet(allRows);
    inserted = result.inserted;
    updated  = result.updated;
    SpreadsheetApp.flush();
  }

  return {
    rows:        allRows.length,
    inserted:    inserted,
    updated:     updated,
    pages:       totalPages,
    hitRowLimit: hitRowLimit
  };
}

/**
 * Fetches a single page from Maqsam Calls API v2.
 * Returns { calls: [...] } or null on failure.
 */
function _fetchCallsPage(page, startUnix, endUnix) {
  const props = PropertiesService.getScriptProperties();
  const baseUrl = props.getProperty('MAQSAM_BASE_URL');
  const key = props.getProperty('MAQSAM_ACCESS_KEY_ID');
  const secret = props.getProperty('MAQSAM_ACCESS_SECRET');

  if (!baseUrl || !key || !secret) {
    throw new Error('Missing Maqsam credentials in Script Properties. See setup guide.');
  }

  // Build URL
  let url = `https://api.${baseUrl}/v2/calls/page/${page}`;
  const params = [];
  if (startUnix) params.push(`start_time=${startUnix}`);
  if (endUnix)   params.push(`end_time=${endUnix}`);
  if (params.length > 0) url += '?' + params.join('&');

  // HTTP Basic Auth
  const credentials = Utilities.base64Encode(key + ':' + secret);
  const options = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  };

  // Execute with retry
  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const body = response.getContentText();

      if (code === 200) {
        const json = JSON.parse(body);
        // Handle different possible response structures
        const calls = json.message || json.calls || json.data || json;
        if (Array.isArray(calls)) {
          return { calls: calls };
        }
        if (calls && Array.isArray(calls.result)) {
          return { calls: calls.result };
        }
        // If we can't find an array, log and return empty
        Logger.log('Unexpected response structure. Keys: ' + Object.keys(json).join(', '));
        return { calls: [] };
      }

      if (code === 401) {
        throw new Error(`Authentication failed (401). Verify MAQSAM_ACCESS_KEY_ID and MAQSAM_ACCESS_SECRET.`);
      }

      if (code === 429 || code >= 500) {
        // Rate limit or server error — retry with backoff
        const delay = Math.pow(2, attempt) * 1000;
        Logger.log(`HTTP ${code} on attempt ${attempt}. Retrying in ${delay}ms...`);
        Utilities.sleep(delay);
        lastError = `HTTP ${code}`;
        continue;
      }

      throw new Error(`HTTP ${code}: ${body.substring(0, 200)}`);

    } catch (e) {
      lastError = e.toString();
      if (attempt < 3) {
        Utilities.sleep(Math.pow(2, attempt) * 1000);
      }
    }
  }

  throw new Error(`Failed after 3 attempts. Last error: ${lastError}`);
}

// ==========================================
// SECTION 4: DATA TRANSFORMATION
// ==========================================

/**
 * Maps a single Maqsam API call record to a sheet row array.
 * Returns array matching SHEET_HEADERS order, or null if invalid.
 *
 * FIELD MAPPINGS:
 *   id              → Call ID
 *   timestamp       → Date Time
 *   direction       → direction        ("inbound" / "outbound")
 *   callerNumber    → Caller
 *   calleeNumber    → Callee
 *   (derived)       → Country          callerNumber if inbound, calleeNumber if outbound
 *   handlingTime    → Call duration
 *   holdTime        → Hold Time
 *   acwTime         → Wrapup Time
 *   waitingTime     → Wait Duration
 *   state           → Call Status
 *   agents[].name   → Agent name
 *   agents[].email  → Agent email
 */
function _mapCallToRow(call) {
  try {
    const id = _safeGet(call, 'id', '');
    if (!id) return null;

    // Date Time
    const timestampRaw = _safeGet(call, 'timestamp', 0);
    const dateTimeStr  = _unixToDateTimeStr(timestampRaw);

    // Direction — "inbound" or "outbound"
    const direction = _safeGet(call, 'direction', '');

    // Caller / Callee numbers
    const callerNumber = _safeGet(call, 'callerNumber', '');
    const calleeNumber = _safeGet(call, 'calleeNumber', '');

    // Country: use callerNumber for inbound calls, calleeNumber for outbound
    const phoneForCountry = (direction.toLowerCase() === 'outbound') ? calleeNumber : callerNumber;
    const country = _getCountryFromPhone(phoneForCountry);

    // Duration fields (seconds from API)
    const callDurationSec = _safeGet(call, 'handlingTime', 0);
    const holdTimeSec     = _safeGet(call, 'holdTime',     0);
    const wrapupTimeSec   = _safeGet(call, 'acwTime',      0);
    const waitTimeSec     = _safeGet(call, 'waitingTime',  0);

    // Call Status
    const callState = _safeGet(call, 'state', '');

    // Agents array [{name, email}, ...]
    const agentsArr = _safeGet(call, 'agents', []);
    const agentList = Array.isArray(agentsArr) ? agentsArr : (agentsArr ? [agentsArr] : []);

    const agentNames = agentList
      .map(function(a) {
        const n = _safeGet(a, 'name', '');
        return n || _extractNameFromEmail(_safeGet(a, 'email', ''));
      })
      .filter(Boolean)
      .join(' | ');

    const agentEmails = agentList
      .map(function(a) { return _safeGet(a, 'email', ''); })
      .filter(Boolean)
      .join(' | ');

    // Row order MUST match SHEET_HEADERS exactly
    return [
      id,                              // Call ID
      dateTimeStr,                     // Date Time
      direction,                       // direction
      callerNumber,                    // Caller
      calleeNumber,                    // Callee
      country,                         // Country
      _secondsToHms(callDurationSec),  // Call duration
      _secondsToHms(holdTimeSec),      // Hold Time
      _secondsToHms(wrapupTimeSec),    // Wrapup Time
      _secondsToHms(waitTimeSec),      // Wait Duration
      callState,                       // Call Status
      agentNames,                      // Agent name(s)
      agentEmails                      // Agent email(s)
    ];

  } catch (e) {
    Logger.log('Error mapping call ' + _safeGet(call, 'id', '?') + ': ' + e.toString());
    return null;
  }
}

/**
 * Safe nested property getter with default value.
 * Supports dot-notation paths like 'agent.email'.
 */
function _safeGet(obj, path, defaultValue) {
  try {
    const keys = path.split('.');
    let current = obj;
    for (let i = 0; i < keys.length; i++) {
      if (current === null || current === undefined) return defaultValue;
      current = current[keys[i]];
    }
    return current !== undefined && current !== null ? current : defaultValue;
  } catch (e) {
    return defaultValue;
  }
}

/**
 * Converts UNIX timestamp (seconds) to formatted datetime string.
 * Format: "M/D/YYYY H:mm" (matching your CSV style)
 * Timezone: Asia/Dubai (GST, UTC+4)
 */
function _unixToDateTimeStr(unixSec) {
  if (!unixSec || isNaN(unixSec)) return '';

  // Handle both seconds and milliseconds
  let ms = unixSec;
  if (unixSec < 10000000000) ms = unixSec * 1000; // seconds → ms

  const date = new Date(ms);
  return Utilities.formatDate(date, 'Asia/Dubai', 'M/d/yyyy H:mm');
}

/**
 * Converts seconds to HH:MM:SS string format.
 * Returns empty string for zero/invalid values (matching CSV style).
 */
function _secondsToHms(totalSeconds) {
  if (!totalSeconds || isNaN(totalSeconds) || totalSeconds <= 0) return '';

  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = Math.floor(totalSeconds % 60);

  const pad = (n) => n.toString().padStart(2, '0');
  return `${pad(hours)}:${pad(minutes)}:${pad(seconds)}`;
}

/**
 * Derives country name from E.164 phone number prefix.
 */
function _getCountryFromPhone(phone) {
  if (!phone) return '';

  // Remove leading + and any non-digits
  const digits = String(phone).replace(/\D/g, '');

  // Try 3-digit prefix first, then 2-digit, then 1-digit
  for (let len = 3; len >= 1; len--) {
    const prefix = digits.substring(0, len);
    if (COUNTRY_PREFIX_MAP[prefix]) {
      return COUNTRY_PREFIX_MAP[prefix];
    }
  }

  return '';
}

/**
 * Extracts a display name from an email address.
 * e.g., "jane.doe.ext@example.com" → "Jane Doe"
 */
function _extractNameFromEmail(email) {
  if (!email) return '';

  const local = email.split('@')[0];
  if (!local) return '';

  // Remove common suffixes like .ext
  const clean = local.replace(/\.ext$/, '').replace(/\.[0-9]+$/, '');

  // Split on dots and capitalize
  const parts = clean.split('.');
  return parts
    .filter(p => p.length > 0)
    .map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase())
    .join(' ');
}

// ==========================================
// SECTION 5: SHEET OPERATIONS
// ==========================================

/**
 * Ensures the Phone_API tab exists with correct headers.
 * Creates the tab if it doesn't exist.
 */
function _ensureHeaders() {
  const ss = SpreadsheetApp.openById(PHONE_CONFIG.SHEET_ID);
  let sheet = ss.getSheetByName(PHONE_CONFIG.TAB_NAME);

  // Create tab if missing
  if (!sheet) {
    sheet = ss.insertSheet(PHONE_CONFIG.TAB_NAME);
    Logger.log(`Created new tab: ${PHONE_CONFIG.TAB_NAME}`);
  }

  // Check/update headers
  const existingHeaders = sheet.getRange(1, 1, 1, SHEET_HEADERS.length).getValues()[0];
  const needsHeaders = existingHeaders.every(h => !h);

  if (needsHeaders) {
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);

    // Format header row
    const headerRange = sheet.getRange(1, 1, 1, SHEET_HEADERS.length);
    headerRange.setFontWeight('bold')
               .setBackground('#E8EAED')
               .setBorder(true, true, true, true, true, true);

    Logger.log('Headers written to row 1.');
  }

  return sheet;
}

/**
 * UPSERT rows into the Phone_API sheet.
 *
 * For each incoming row (keyed by Call ID in column 1):
 *   - If the Call ID does NOT exist → append as a new row at the bottom
 *   - If the Call ID DOES exist     → overwrite that row in place
 *
 * This is what allows in_progress calls to be updated to their final state
 * (answered, missed, etc.) on subsequent sync runs.
 *
 * Returns { inserted, updated } counts for audit logging.
 */
function _upsertToSheet(rows) {
  const result = { inserted: 0, updated: 0 };
  if (!rows || rows.length === 0) return result;

  const ss    = SpreadsheetApp.openById(PHONE_CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(PHONE_CONFIG.TAB_NAME);

  if (!sheet) {
    throw new Error('Tab "' + PHONE_CONFIG.TAB_NAME + '" not found. Run setupPhoneSync() first.');
  }

  // ── Build a Map of existing Call ID → sheet row number (1-indexed) ─────────
  const lastRow       = sheet.getLastRow();
  const existingIdMap = {};   // { callId: rowNumber }

  if (lastRow > 1) {
    // Read all Call IDs in one batch — fast, avoids per-row reads
    const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    idValues.forEach(function(r, i) {
      if (r[0]) existingIdMap[String(r[0])] = i + 2; // +2: 1-indexed + header offset
    });
  }

  // ── Separate rows into inserts vs updates ──────────────────────────────────
  const toInsert = [];   // rows with Call IDs not in the sheet
  const toUpdate = [];   // { rowNum, rowData } for rows already in the sheet

  rows.forEach(function(row) {
    if (!row || !row[0]) return;
    const callId = String(row[0]);
    if (existingIdMap.hasOwnProperty(callId)) {
      toUpdate.push({ rowNum: existingIdMap[callId], data: row });
    } else {
      toInsert.push(row);
    }
  });

  // ── Apply updates in place ─────────────────────────────────────────────────
  // Batch by contiguous row blocks where possible to minimise setValues() calls.
  // For simplicity (and because update sets are usually small) we write each
  // updated row individually — still one API call per row but safe and correct.
  toUpdate.forEach(function(item) {
    sheet.getRange(item.rowNum, 1, 1, SHEET_HEADERS.length).setValues([item.data]);
  });
  result.updated = toUpdate.length;

  if (toUpdate.length > 0) {
    Logger.log('Updated ' + toUpdate.length + ' existing row(s) in place.');
  }

  // ── Append new rows at the bottom ──────────────────────────────────────────
  if (toInsert.length > 0) {
    // Ensure sheet has enough physical rows
    const currentMax = sheet.getMaxRows();
    const needed     = sheet.getLastRow() + toInsert.length;
    if (needed > currentMax) {
      sheet.insertRowsAfter(currentMax, needed - currentMax + 50);
    }

    const targetRow   = sheet.getLastRow() + 1;
    const targetRange = sheet.getRange(targetRow, 1, toInsert.length, SHEET_HEADERS.length);
    targetRange.setValues(toInsert);
    Logger.log('Inserted ' + toInsert.length + ' new row(s) starting at row ' + targetRow + '.');
  }
  result.inserted = toInsert.length;

  return result;
}

// ==========================================
// SECTION 6: TRIGGER MANAGEMENT
// ==========================================

/**
 * Creates the hourly time-driven trigger (idempotent — won't duplicate).
 */
function _createHourlyTrigger() {
  const fnName = PHONE_CONFIG.HOURLY_TRIGGER_NAME;

  // Check if trigger already exists
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === fnName);

  if (exists) {
    Logger.log('Hourly trigger already exists. Skipping creation.');
    return;
  }

  ScriptApp.newTrigger(fnName)
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('Hourly trigger created successfully.');
}

/**
 * Removes all Phone sync triggers (useful if you need to reconfigure).
 */
function removePhoneTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const fnName = PHONE_CONFIG.HOURLY_TRIGGER_NAME;
  let removed = 0;

  triggers.forEach(t => {
    if (t.getHandlerFunction() === fnName) {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });

  Logger.log(`Removed ${removed} trigger(s).`);
}

// ==========================================
// SECTION 7: STATE MANAGEMENT
// ==========================================

/**
 * Saves backfill resume position for chunked processing.
 */
function _saveBackfillState(date) {
  if (!date) return;
  PropertiesService.getScriptProperties().setProperty('BACKFILL_RESUME_AT', date.toISOString());
}

/**
 * Reads backfill resume position.
 */
function _getBackfillState() {
  const iso = PropertiesService.getScriptProperties().getProperty('BACKFILL_RESUME_AT');
  if (!iso) return null;
  return new Date(iso);
}

/**
 * Checks if we're approaching the GAS 6-minute timeout.
 */
function _isApproachingTimeout(scriptStartTime) {
  const elapsed = Date.now() - scriptStartTime;
  return elapsed > (360000 - PHONE_CONFIG.SCRIPT_TIMEOUT_BUFFER_MS); // 6 min = 360,000ms
}

// ==========================================
// SECTION 8: AUDIT & LOGGING
// ==========================================

/**
 * Logs audit events to a dedicated _AUDIT sheet for monitoring.
 * Creates the sheet if it doesn't exist.
 */
function _logAudit(functionName, status, details) {
  try {
    const ss = SpreadsheetApp.openById(PHONE_CONFIG.SHEET_ID);
    const auditSheetName = '_AUDIT';
    let sheet = ss.getSheetByName(auditSheetName);

    if (!sheet) {
      sheet = ss.insertSheet(auditSheetName);
      sheet.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Function', 'Status', 'Details', 'User']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#E8EAED');
    }

    const row = [
      Utilities.formatDate(new Date(), 'Asia/Dubai', 'yyyy-MM-dd HH:mm:ss'),
      functionName,
      status,
      details,
      Session.getActiveUser().getEmail() || 'trigger'
    ];

    sheet.appendRow(row);
  } catch (e) {
    // Silent fail — don't let audit logging break the main flow
    Logger.log('Audit log failed: ' + e.toString());
  }
}

// ==========================================
// SECTION 9: DIAGNOSTIC & TESTING
// ==========================================

/**
 * TEST FUNCTION: Verify connectivity and inspect API response structure.
 * Run this after setting up credentials. It fetches 1 page (max 100 records)
 * and logs the field structure so you can verify field names.
 *
 * EXPECTED OUTPUT: You should see values for:
 *   direction, callerNumber, calleeNumber, state, handlingTime, acwTime, holdTime, waitingTime, agents[]
 */
function testMaqsamConnection() {
  try {
    Logger.log('=== MAQSAM CONNECTION TEST ===');

    // Check credentials exist
    const props = PropertiesService.getScriptProperties();
    const baseUrl = props.getProperty('MAQSAM_BASE_URL');
    const key = props.getProperty('MAQSAM_ACCESS_KEY_ID');
    const secret = props.getProperty('MAQSAM_ACCESS_SECRET');

    Logger.log(`BASE_URL set: ${baseUrl ? 'YES (' + baseUrl + ')' : 'NO — MISSING'}`);
    Logger.log(`ACCESS_KEY_ID set: ${key ? 'YES (' + key.substring(0, 4) + '...)' : 'NO — MISSING'}`);
    Logger.log(`ACCESS_SECRET set: ${secret ? 'YES (hidden)' : 'NO — MISSING'}`);

    if (!baseUrl || !key || !secret) {
      Logger.log('ERROR: Missing credentials. Add them to Script Properties first.');
      Logger.log('See Step 2 in the deployment guide.');
      return;
    }

    // Fetch most recent page (no time filter = get everything, API returns newest first)
    const response = _fetchCallsPage(1, null, null);

    if (!response || !response.calls || response.calls.length === 0) {
      Logger.log('No call records returned. Response structure:');
      Logger.log(JSON.stringify(response, null, 2));
      return;
    }

    Logger.log(`SUCCESS: Fetched ${response.calls.length} records.`);
    Logger.log('');
    Logger.log('=== FIRST RECORD FIELD CHECK ===');
    Logger.log('(Verify these match the expected fields below)');
    Logger.log('');

    const first = response.calls[0];

    // Check every field the sync engine relies on
    const agentsArr  = _safeGet(first, 'agents', []);
    const firstAgent = Array.isArray(agentsArr) && agentsArr.length > 0 ? agentsArr[0] : {};
    const checks = [
      ['id',              _safeGet(first, 'id',            'MISSING')],
      ['timestamp',       _safeGet(first, 'timestamp',     'MISSING')],
      ['direction',       _safeGet(first, 'direction',     'MISSING')],
      ['callerNumber',    _safeGet(first, 'callerNumber',  'MISSING')],
      ['calleeNumber',    _safeGet(first, 'calleeNumber',  'MISSING')],
      ['state',           _safeGet(first, 'state',         'MISSING')],
      ['handlingTime',    _safeGet(first, 'handlingTime',  'MISSING')],
      ['acwTime',         _safeGet(first, 'acwTime',       'MISSING')],
      ['holdTime',        _safeGet(first, 'holdTime',      'MISSING')],
      ['waitingTime',     _safeGet(first, 'waitingTime',   'MISSING')],
      ['agents (array)',  Array.isArray(agentsArr) ? 'YES (' + agentsArr.length + ' agent(s))' : 'MISSING'],
      ['agents[0].name',  _safeGet(firstAgent, 'name',     'MISSING')],
      ['agents[0].email', _safeGet(firstAgent, 'email',    'MISSING')],
    ];

    checks.forEach(([field, value]) => {
      const status = (value !== 'MISSING' && value !== '') ? 'OK' : 'MISSING';
      Logger.log(`[${status}] ${field}: ${value}`);
    });

    Logger.log('');
    Logger.log('=== FULL FIRST RECORD (for reference) ===');
    Logger.log(JSON.stringify(first, null, 2));

  } catch (e) {
    Logger.log('TEST FAILED: ' + e.toString());
  }
}

/**
 * TEST FUNCTION: Verify country mapping for known phone numbers.
 */
function testCountryMapping() {
  const tests = [
    ['+966501234567', 'Saudi Arabia'],
    ['+971501234567', 'United Arab Emirates'],
    ['+97450123456', 'Qatar'],
    ['+962791234567', 'Jordan'],
  ];

  Logger.log('=== COUNTRY MAPPING TEST ===');
  tests.forEach(([phone, expected]) => {
    const result = _getCountryFromPhone(phone);
    const status = result === expected ? 'PASS' : 'FAIL';
    Logger.log(`${status}: ${phone} → ${result} (expected: ${expected})`);
  });
}

/**
 * UTILITY: Full reset — clears sync watermark AND backfill resume pointer.
 *
 * Run this if:
 *   - The backfill started from the wrong date (e.g. a stale BACKFILL_RESUME_AT
 *     from a previous run caused it to skip January data)
 *   - You want to re-sync everything from scratch
 *
 * After running this:
 *   1. Delete the Phone_API tab contents (or the whole tab) from the sheet
 *   2. Run setupPhoneSync()
 *   3. Run runHistoricalBackfill()  ← will restart cleanly from Jan 1, 2026
 */
function resetSyncWatermark() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('PHONE_LAST_SYNC_UNIX');
  props.deleteProperty('BACKFILL_RESUME_AT');
  Logger.log('=== RESET COMPLETE ===');
  Logger.log('Watermark and backfill resume pointer cleared.');
  Logger.log('Next steps:');
  Logger.log('  1. Clear or delete the Phone_API sheet tab');
  Logger.log('  2. Run setupPhoneSync()');
  Logger.log('  3. Run runHistoricalBackfill()  ← will start from 2026-01-01');
}

// ==============================================================================
// END OF FILE
// ==============================================================================