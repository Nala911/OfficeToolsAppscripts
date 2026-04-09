/**
 * @file FormattingTools.js
 * @description Sheet formatting utilities for rapid cleanup.
 */

/**
 * Automatically resizes all populated columns to fit their content.
 */
function autofitColumns() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getDataRange();
  var numCols = dataRange.getNumColumns();
  if (numCols > 0) {
    sheet.autoResizeColumns(1, numCols);
  }
}

/**
 * Formats the first row (header) with a blue background and white bold text.
 * Also freezes the first row for better readability.
 */
function formatHeaderBlue() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var maxCols = sheet.getMaxColumns();
  
  // Select the first row as the header
  var headerRange = sheet.getRange(1, 1, 1, maxCols);
  
  // Apply blue background, white bold text
  headerRange.setBackground('#4a86e8') // Royal Blue shade
             .setFontColor('#ffffff')
             .setFontWeight('bold');
  
  // Freeze the top row
  sheet.setFrozenRows(1);
}
