/**
 * @file pdf_maker.js
 * @description Logic for mapping row data to Template.html and generating PDFs.
 */

const TransactionTools = {
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
  }
};

/**
 * Global wrapper for backward compatibility and API Dispatcher.
 */
function apiGenerateTransactionLetter(txnId) {
  return TransactionTools.generateLetter(txnId);
}

/**
 * Legacy wrapper for the top menu.
 */
function promptForTransactionId() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Generate PDF', 'Please enter the Transaction ID (e.g., TXN-1001):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const txnId = response.getResponseText();
    processTransaction(txnId);
  }
}

/**
 * Finds the row and triggers PDF generation.
 */
function processTransaction(txnId) {
  const result = TransactionTools.generateLetter(txnId);

  if (!result.success) {
    SpreadsheetApp.getUi().alert(result.error);
    return;
  }

  // To download from browser, we must serve a small HTML dialog with a download link
  const downloadHtml = HtmlService.createHtmlOutput(
    '<script>' +
    '  var a = document.createElement("a");' +
    '  a.href = "data:application/pdf;base64,' + result.base64 + '";' +
    '  a.download = "' + result.fileName + '";' +
    '  a.click();' +
    '  google.script.host.close();' +
    '</script>' +
    '<p>Downloading...</p>'
  ).setWidth(50).setHeight(50);

  SpreadsheetApp.getUi().showModalDialog(downloadHtml, 'Processing...');
}