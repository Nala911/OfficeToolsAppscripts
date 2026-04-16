/**
 * @file FormattingTools.js
 * @description Sheet formatting utilities for rapid cleanup.
 */

var FormattingTools = {
  /**
   * Automatically resizes all populated columns to fit their content.
   * @return {object} Result object {success, message, error}
   */
  autofitColumns: function() {
    try {
      var sheet = SpreadsheetApp.getActiveSheet();
      var dataRange = sheet.getDataRange();
      var numCols = dataRange.getNumColumns();
      if (numCols > 0) {
        sheet.autoResizeColumns(1, numCols);
        return { success: true, message: 'All columns autofitted successfully.' };
      }
      return { success: false, error: 'Sheet is empty.' };
    } catch (e) {
      return { success: false, error: e.toString() };
    }
  },

  /**
   * Locates a Transaction ID in Column B and highlights Columns A through G.
   * @param {string} transactionId
   * @return {object} Result object {success, message, error}
   */
  highlightTransaction: function(transactionId) {
    try {
      if (!transactionId) return { success: false, error: 'Transaction ID is required.' };
      
      var sheet = SpreadsheetApp.getActiveSheet();
      var searchId = transactionId.trim();
      
      // Auto-prepend 'TXN-' if the user just typed the number
      if (!searchId.toUpperCase().startsWith('TXN-')) {
          searchId = 'TXN-' + searchId;
      }
      
      var lastRow = sheet.getLastRow();
      if (lastRow < 1) return { success: false, error: 'Sheet is empty.' };
      
      var colBData = sheet.getRange(1, 2, lastRow, 1).getValues();
      var foundRow = -1;

      for (var i = 0; i < colBData.length; i++) {
          if (String(colBData[i][0]).trim() === searchId) {
              foundRow = i + 1; 
              break;
          }
      }
      
      if (foundRow !== -1) {
          var targetRange = sheet.getRange(foundRow, 1, 1, 7);
          targetRange.setBackground('#FFFF00'); // Yellow
          sheet.setActiveRange(targetRange); // Scroll to it
          return { success: true, message: 'Transaction ' + searchId + ' highlighted at row ' + foundRow + '.' };
      } else {
          return { success: false, error: 'Transaction ID ' + searchId + ' not found in Column B.' };
      }
    } catch (e) {
      return { success: false, error: e.toString() };
    }
  }
};

/**
 * Global wrappers for backward compatibility and API Dispatcher.
 */
function apiAutofitColumns() {
  return FormattingTools.autofitColumns();
}

function apiHighlightTransaction(transactionId) {
  return FormattingTools.highlightTransaction(transactionId);
}

/**
 * Legacy wrapper for the top menu.
 */
function autofitColumns() {
  const result = FormattingTools.autofitColumns();
  if (!result.success) {
    SpreadsheetApp.getUi().alert(result.error);
  }
}

/**
 * Legacy wrapper for the top menu.
 */
function highlightTransactionRow() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Highlight Transaction', 'Enter Transaction ID (e.g., 1005):', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const result = FormattingTools.highlightTransaction(response.getResponseText());
    if (!result.success) {
      ui.alert(result.error);
    }
  }
}

