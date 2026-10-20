/**
 * TEMPORARY ONE-TIME CLEANUP SCRIPT
 * Run this once to fix historical names and emails, then delete it.
 */

function runOneTimeCleanup() {
  const SHEET_NAME = 'D'; 

  
  // ⚠️ GOOGLE SHEETS COLUMNS START AT 1 (A=1, B=2, C=3... H=8, I=9, J=10)
  // Double check your sheet and change these numbers if your columns are different!
  const NAME_COLUMN_NUMBER = 1;  // Assuming Agent Name is Column I
  const EMAIL_COLUMN_NUMBER = 10; // Assuming Email is Column J

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Could not find the sheet named " + SHEET_NAME);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No data to clean.");
    return;
  }

  // 1. Clean Agent Names
  const nameRange = sheet.getRange(2, NAME_COLUMN_NUMBER, lastRow - 1, 1);
  const rawNames = nameRange.getValues();
  const cleanNames = rawNames.map(row => {
    return [String(row[0]).split(/,| and /i)[0].trim()];
  });
  nameRange.setValues(cleanNames);

  // 2. Clean Emails
  const emailRange = sheet.getRange(2, EMAIL_COLUMN_NUMBER, lastRow - 1, 1);
  const rawEmails = emailRange.getValues();
  const cleanEmails = rawEmails.map(row => {
    return [String(row[0]).split(/,| and /i)[0].trim()];
  });
  emailRange.setValues(cleanEmails);

  SpreadsheetApp.getUi().alert(`Success! Cleaned ${lastRow - 1} rows. You can now delete this script!`);
}