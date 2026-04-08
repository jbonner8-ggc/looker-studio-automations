/**
 * Copies data from 'Live Feed' to 'History'
 * and appends it to the bottom of the History sheet.
 */
function archiveLiveFeed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let liveSheet = ss.getSheetByName('Live Feed');
  let historySheet = ss.getSheetByName('History');

  // Create 'Live Feed' sheet if it doesn't exist and initialize with IMPORTFEED
  if (!liveSheet) {
    liveSheet = ss.insertSheet('Live Feed');
    liveSheet.getRange('A1').setFormula('=IMPORTFEED("https://www.google.com/alerts/feeds/09040407867941873068/5296938900162927815","items",TRUE,20)');
  }

  // Create 'History' sheet if it doesn't exist
  if (!historySheet) {
    historySheet = ss.insertSheet('History');
  }
  
  // Get all data from Live Feed (excluding header if you want)
  // This assumes your data starts on row 2 to skip headers
  const lastRow = liveSheet.getLastRow();
  if (lastRow < 2) return; // Nothing to copy
  
  const range = liveSheet.getRange(2, 1, lastRow - 1, liveSheet.getLastColumn());
  const values = range.getValues();
  
  // Append data to the next available row in History
  historySheet.getRange(historySheet.getLastRow() + 1, 1, values.length, values[0].length)
              .setValues(values);
              
  // Optional: Clear the Live Feed after copying
  // liveSheet.getRange(2, 1, lastRow - 1, liveSheet.getLastColumn()).clearContent();
}

/**
 * Creates a trigger to run the archive function every 12 hours.
 * Run this function manually once from the editor.
 */
function createTwelveHourTrigger() {
  // Delete existing triggers for this function to avoid duplicates
  Browser.msgBox('Existing triggers for archiveLiveFeed will be removed. Continue?');
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'archiveLiveFeed') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create the new 12-hour trigger
  ScriptApp.newTrigger('archiveLiveFeed')
    .timeBased()
    .everyHours(12)
    .create();
}