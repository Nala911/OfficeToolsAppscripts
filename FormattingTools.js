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

