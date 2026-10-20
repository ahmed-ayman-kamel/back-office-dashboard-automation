/**
 * CONFIGURATION
 * Ensure these match your environment exactly.
 */
const CONFIG = {
  SHEET_ID: 'YOUR_BACK_OFFICE_SHEET_ID',
  TAB_NAME: 'BO RD',
  GMAIL_SEARCH_QUERY: 'from:support@freshdesk.com subject:"Ticket Performance - Hourly Widget"',
  FILENAME_KEYWORDS: 'total_tickets' 
};

/**
 * Main function to sync the Ticket Performance report.
 * Uses Google Sheets API for maximum speed and data integrity.
 */
function syncFreshdeskViaAttachment() {
  try {
    Logger.log('Starting Turbo Sync via Attachment...');

    // 1. Search for the latest email thread
    const threads = GmailApp.search(CONFIG.GMAIL_SEARCH_QUERY, 0, 1);
    if (threads.length === 0) {
      Logger.log('No matching email thread found.');
      return;
    }

    const messages = threads[0].getMessages();
    const latestMessage = messages[messages.length - 1];
    const attachments = latestMessage.getAttachments();
    
    let csvBlob = null;
    // 2. Identify the correct CSV attachment
    for (let i = 0; i < attachments.length; i++) {
      const name = attachments[i].getName().toLowerCase();
      if (name.includes(CONFIG.FILENAME_KEYWORDS) && name.endsWith('.csv')) {
        csvBlob = attachments[i];
        break;
      }
    }

    if (!csvBlob) {
      Logger.log('Email thread found, but no matching CSV attachment was detected.');
      return;
    }

    Logger.log('Found CSV: ' + csvBlob.getName());

    // 3. Fetch and Parse CSV Content
    const csvContent = csvBlob.getDataAsString();
    const csvData = Utilities.parseCsv(csvContent);
    
    if (csvData.length <= 1) {
      Logger.log('CSV is empty or contains only headers. Exiting.');
      return;
    }

    // Prepare data (Remove headers to preserve Sheet Row 1)
    const dataToValues = csvData.slice(1);
    const ssId = CONFIG.SHEET_ID;
    const tabName = CONFIG.TAB_NAME;

    // 4. Access Sheet and Expand if necessary
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(tabName);
    
    if (!sheet) {
      Logger.log('ERROR: Tab "' + tabName + '" not found in the spreadsheet.');
      return;
    }

    const requiredRows = dataToValues.length + 1;
    const currentMaxRows = sheet.getMaxRows();

    if (requiredRows > currentMaxRows) {
      const rowsToAdd = requiredRows - currentMaxRows;
      sheet.insertRowsAfter(currentMaxRows, rowsToAdd);
      Logger.log(`Sheet expanded: Added ${rowsToAdd} new rows.`);
    }

    // 5. TURBO UPDATE VIA SHEETS API
    // Clear old data from Row 2 down to the current limit of the sheet
    const clearRange = `${tabName}!A2:Z${sheet.getMaxRows()}`;
    Sheets.Spreadsheets.Values.clear({}, ssId, clearRange);
    Logger.log('Old data cleared.');

    // Prepare the update payload
    const valueRange = Sheets.newValueRange();
    valueRange.values = dataToValues;

    // Write starting at A2
    const updateRange = `${tabName}!A2`;
    
    /**
     * 'USER_ENTERED' is critical: It ensures Google Sheets parses numbers 
     * and dates correctly so they work in Looker Studio metrics.
     */
    Sheets.Spreadsheets.Values.update(valueRange, ssId, updateRange, {
      valueInputOption: 'USER_ENTERED'
    });

    Logger.log('Turbo Sync Complete! Successfully imported ' + dataToValues.length + ' rows.');

  } catch (e) {
    Logger.log('FATAL ERROR: ' + e.toString());
  }
}