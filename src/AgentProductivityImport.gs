/**
 * CONFIGURATION FOR AGENT VIEW (Turbo Version)
 */
const AGENT_CONFIG = {
  SHEET_ID: 'YOUR_BACK_OFFICE_SHEET_ID',
  TAB_NAME: 'Agent View', 
  // Updated subject and added 'has:attachment' for faster Gmail searching
  GMAIL_SEARCH_QUERY: 'from:support@freshdesk.com subject:"BO Productivity Per Agent - Hourly Widget" has:attachment',
};

function syncAgentViewAttachment() {
  try {
    Logger.log('Searching for the Agent Productivity email...');

    // 1. Fetch Email
    const threads = GmailApp.search(AGENT_CONFIG.GMAIL_SEARCH_QUERY, 0, 1);
    if (threads.length === 0) { 
      Logger.log('No email found with query.'); 
      return; 
    }

    // 2. Extract Attachment (Grabbing from the latest message in the thread)
    const messages = threads[0].getMessages();
    const latestMessage = messages[messages.length - 1];
    const attachments = latestMessage.getAttachments();
    
    let csvBlob = null;
    for (let i = 0; i < attachments.length; i++) {
      if (attachments[i].getName().toLowerCase().endsWith('.csv')) {
        csvBlob = attachments[i];
        break;
      }
    }

    if (!csvBlob) { 
      Logger.log('No CSV attachment found in the email.'); 
      return; 
    }
    
    Logger.log('SUCCESS: Found CSV: ' + csvBlob.getName());

    // 3. Parse Data
    const csvData = Utilities.parseCsv(csvBlob.getDataAsString());
    if (csvData.length <= 1) {
      Logger.log('CSV is empty or only contains headers.');
      return;
    }

    // Remove headers and strictly limit to 7 columns to fit perfectly into B through H
    const finalData = csvData.slice(1).map(row => row.slice(0, 7));

    // 4. Prepare Spreadsheet
    const ssId = AGENT_CONFIG.SHEET_ID;
    const tabName = AGENT_CONFIG.TAB_NAME;
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(tabName);

    if (!sheet) {
      Logger.log(`ERROR: Tab "${tabName}" not found.`);
      return;
    }

    // Dynamic Expansion: Ensure sheet has enough rows so the API doesn't fail
    const requiredRows = finalData.length + 1;
    const currentMaxRows = sheet.getMaxRows();
    if (requiredRows > currentMaxRows) {
      sheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
      Logger.log(`Expanded sheet by adding ${requiredRows - currentMaxRows} rows.`);
    }

    // 5. TURBO IMPORT (Sheets API)
    // Clear old data strictly in columns B through H (Leaving Col A intact)
    const clearRange = `${tabName}!B2:H${sheet.getMaxRows()}`;
    Sheets.Spreadsheets.Values.clear({}, ssId, clearRange);

    // Update with new data starting at B2
    const valueRange = Sheets.newValueRange();
    valueRange.values = finalData;

    // USER_ENTERED ensures numbers and dates format correctly for BI tools
    Sheets.Spreadsheets.Values.update(valueRange, ssId, `${tabName}!B2`, {
      valueInputOption: 'USER_ENTERED' 
    });

    Logger.log(`Success! Turbo Imported ${finalData.length} rows into B2:H on "${tabName}".`);

  } catch (e) {
    Logger.log('Error: ' + e.toString());
  }
}