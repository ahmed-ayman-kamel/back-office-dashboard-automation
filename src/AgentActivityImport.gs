/**
 * CONFIGURATION - Scorecard Raw Data (Hybrid Version)
 */
const SCORECARD_CONFIG = {
  SHEET_ID: 'YOUR_SCORECARD_SHEET_ID',
  TAB_NAME: 'db - Freshdesk',
  GMAIL_QUERY: 'from:support@freshdesk.com subject:"Scorecard Raw Data - Hourly Widget"',
  KEYWORD: 'scorecard_raw_data_total_agent_activity',
  CLEAR_RANGE: 'A2:D' 
};

function syncScorecardData() {
  try {
    Logger.log('Starting Hybrid Sync (Checking Attachments & Links)...');

    const threads = GmailApp.search(SCORECARD_CONFIG.GMAIL_QUERY, 0, 1);
    if (threads.length === 0) return;

    const latestMessage = threads[0].getMessages().pop();
    const attachments = latestMessage.getAttachments();
    const body = latestMessage.getBody();
    
    let csvContent = null;

    // --- STEP 1: CHECK FOR ATTACHMENT ---
    for (let i = 0; i < attachments.length; i++) {
      const fileName = attachments[i].getName().toLowerCase();
      if (fileName.includes(SCORECARD_CONFIG.KEYWORD) && fileName.endsWith('.csv')) {
        csvContent = attachments[i].getDataAsString();
        Logger.log('Data source: ATTACHMENT found.');
        break;
      }
    }

    // --- STEP 2: FALLBACK TO LINK (IF NO ATTACHMENT) ---
    if (!csvContent) {
      Logger.log('No attachment found. Searching for link in body...');
      const regexPattern = new RegExp(`https:\\/\\/[^\\s"'<>]+${SCORECARD_CONFIG.KEYWORD}[^\\s"'<>]*`, 'gi');
      const match = body.match(regexPattern);

      if (match) {
        const downloadUrl = match[0].replace(/&amp;/g, '&');
        csvContent = UrlFetchApp.fetch(downloadUrl).getContentText();
        Logger.log('Data source: LINK found in body.');
      }
    }

    // --- STEP 3: PROCESS DATA ---
    if (!csvContent) {
      Logger.log('FAILED: No attachment and no link found in the email.');
      return;
    }

    const csvData = Utilities.parseCsv(csvContent);
    if (csvData.length <= 1) return;

    const dataToValues = csvData.slice(1).map(row => row.slice(0, 4)); 

    // --- STEP 4: TURBO IMPORT (Sheets API) ---
    const ssId = SCORECARD_CONFIG.SHEET_ID;
    const tabName = SCORECARD_CONFIG.TAB_NAME;
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(tabName);

    // Dynamic Expansion
    const requiredRows = dataToValues.length + 1;
    if (requiredRows > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    }

    // Clear and Update
    Sheets.Spreadsheets.Values.clear({}, ssId, `${tabName}!A2:D${sheet.getMaxRows()}`);
    
    const valueRange = Sheets.newValueRange();
    valueRange.values = dataToValues;
    Sheets.Spreadsheets.Values.update(valueRange, ssId, `${tabName}!A2`, {
      valueInputOption: 'USER_ENTERED'
    });

    Logger.log(`Successfully synced ${dataToValues.length} rows.`);

  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}