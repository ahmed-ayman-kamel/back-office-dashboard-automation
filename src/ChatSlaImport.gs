/**
 * CONFIGURATION - SLA Dashboard Backend (Hybrid Version)
 */
const SLA_CONFIG = {
  SHEET_ID: 'YOUR_SLA_DASHBOARD_SHEET_ID',
  TAB_NAME: 'Chat_RD',
  GMAIL_QUERY: 'from:support@freshdesk.com subject:"Scorecard Raw Data - Hourly Widget"',
  // The unique keyword to grab only the Chat Conversations file out of the 3 attached/linked
  KEYWORD: 'scorecard_raw_data_total_chat_conversations'
};

function syncChatSlaData() {
  try {
    Logger.log('Starting SLA Backend Sync (Checking for Attachments & Links)...');

    const threads = GmailApp.search(SLA_CONFIG.GMAIL_QUERY, 0, 1);
    if (threads.length === 0) {
      Logger.log('No matching email found.');
      return;
    }

    const latestMessage = threads[0].getMessages().pop();
    const attachments = latestMessage.getAttachments();
    const body = latestMessage.getBody();
    
    let csvContent = null;

    // --- STEP 1: CHECK FOR ATTACHMENT ---
    for (let i = 0; i < attachments.length; i++) {
      const fileName = attachments[i].getName().toLowerCase();
      if (fileName.includes(SLA_CONFIG.KEYWORD) && fileName.endsWith('.csv')) {
        csvContent = attachments[i].getDataAsString();
        Logger.log('Data source: ATTACHMENT found matching keyword.');
        break; // Stop looking once we find the exact SLA file
      }
    }

    // --- STEP 2: FALLBACK TO LINK (IF NO ATTACHMENT) ---
    if (!csvContent) {
      Logger.log('No attachment found. Searching for link in body...');
      // Regex looks specifically for the AWS link containing the exact keyword
      const regexPattern = new RegExp(`https:\\/\\/[^\\s"'<>]+${SLA_CONFIG.KEYWORD}[^\\s"'<>]*`, 'gi');
      const match = body.match(regexPattern);

      if (match) {
        const downloadUrl = match[0].replace(/&amp;/g, '&');
        csvContent = UrlFetchApp.fetch(downloadUrl).getContentText();
        Logger.log('Data source: LINK found in body matching keyword.');
      }
    }

    // --- STEP 3: PROCESS DATA ---
    if (!csvContent) {
      Logger.log('FAILED: No attachment or link found containing the keyword.');
      return;
    }

    const csvData = Utilities.parseCsv(csvContent);
    if (csvData.length <= 1) return;

    // Remove headers. We map the entire row since we are not restricting to 4 columns this time.
    const dataToValues = csvData.slice(1); 

    // --- STEP 4: PREPARE SPREADSHEET ---
    const ssId = SLA_CONFIG.SHEET_ID;
    const tabName = SLA_CONFIG.TAB_NAME;
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      Logger.log(`ERROR: Tab "${tabName}" not found.`);
      return;
    }

    // Dynamic Expansion: Ensure sheet has enough rows to avoid API crash
    const requiredRows = dataToValues.length + 1;
    if (requiredRows > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
    }

    // --- STEP 5: TURBO IMPORT WITH STRICT DATA TYPING ---
    // Clear old data first (Leaves Row 1 intact). We clear a wide range (A2:ZZ) to catch everything.
    Sheets.Spreadsheets.Values.clear({}, ssId, `${tabName}!A2:ZZ${sheet.getMaxRows()}`);
    
    const valueRange = Sheets.newValueRange();
    valueRange.values = dataToValues;

    // CRITICAL SLA COMPONENT: 'USER_ENTERED'
    // This tells Google Sheets to actively parse Date/Time text strings into mathematical Date/Time serial numbers.
    Sheets.Spreadsheets.Values.update(valueRange, ssId, `${tabName}!A2`, {
      valueInputOption: 'USER_ENTERED'
    });

    Logger.log(`Successfully synced ${dataToValues.length} rows for SLA analysis.`);

  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}