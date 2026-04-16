/**
 * @file pdf_maker.js
 * @description Logic for mapping row data to Template.html and generating PDFs.
 */

var TransactionTools = {
  /**
   * Generates PDF and returns base64.
   * @param {string} txnId
   * @return {object} Result object {success, base64, fileName, message, error}
   */
  generateLetter: function(txnId) {
    try {
      if (!txnId) return { success: false, error: 'Transaction ID is required.' };

      const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      let rowData = null;

      // Search for the ID in Column B (index 1)
      for (let i = 1; i < data.length; i++) {
        if (data[i][1].toString() === txnId) {
          rowData = data[i];
          break;
        }
      }

      if (!rowData) {
        return { success: false, error: 'Transaction ID ' + txnId + ' not found in Column B.' };
      }

      // Create a template from the HTML file
      const htmlTemplate = HtmlService.createTemplateFromFile('Template');

      // Dynamically map headers to row values for the template
      headers.forEach((header, index) => {
        htmlTemplate[header] = rowData[index];
      });

      const htmlOutput = htmlTemplate.evaluate().getContent();
      const blob = Utilities.newBlob(htmlOutput, 'text/html').getAs('application/pdf');
      const fileName = "Letter_" + txnId + ".pdf";

      return {
        success: true,
        base64: Utilities.base64Encode(blob.getBytes()),
        fileName: fileName,
        message: 'PDF generated successfully.'
      };
    } catch (e) {
      return { success: false, error: e.toString() };
    }
  },

  /**
   * Removes all rows above the given transaction ID in 'Table1'.
   * @param {string} txnId 
   * @return {object} Result of the operation.
   */
  removeOldTransactions: function(txnId) {
    try {
      if (!txnId) return { success: false, error: 'Transaction ID is required.' };

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getActiveSheet();

      const data = sheet.getDataRange().getValues();
      let targetRowIndex = -1;

      // Find the row with the transaction ID in Column B (index 1)
      for (let i = 1; i < data.length; i++) {
        const cellValue = data[i][1].toString().trim();
        // Match either the exact input, or the input prefixed with "TXN-"
        if (cellValue === txnId || cellValue === 'TXN-' + txnId) {
          targetRowIndex = i + 1; // GAS is 1-indexed
          break;
        }
      }

      if (targetRowIndex === -1) {
        return { success: false, error: 'Transaction ID ' + txnId + ' not found in Column B.' };
      }

      if (targetRowIndex <= 2) {
        return { success: true, message: 'No rows above the specified transaction to remove.' };
      }

      const numRowsToDelete = targetRowIndex - 2;
      sheet.deleteRows(2, numRowsToDelete);

      return {
        success: true,
        message: 'Successfully removed ' + numRowsToDelete + ' old transaction row(s).'
      };
    } catch (e) {
      return { success: false, error: e.toString() };
    }
  }
};