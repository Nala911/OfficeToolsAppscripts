/**
 * @file API.js
 * @description Central dispatcher for sidebar requests.
 */

/**
 * Dispatches calls from the sidebar to the appropriate backend function.
 * @param {string} toolName
 * @param {any} inputVal
 * @return {object} Result of the operation.
 */
function apiDispatcher(toolName, inputVal) {
  switch (toolName) {
    case 'autofitColumns':
      return apiAutofitColumns();
    case 'highlightTransaction':
      return apiHighlightTransaction(inputVal);
    case 'generateLetter':
      return apiGenerateTransactionLetter(inputVal);
    case 'deductAmount':
      return apiDeductAmounts(inputVal);
    case 'exportDepts':
      return apiExportDepartments(inputVal);
    case 'updateStatus':
      return apiUpdateDepartmentStatus(inputVal);
    case 'getUniqueDepartments':
      return getUniqueDepartments();
    default:
      return { success: false, error: 'Unknown tool: ' + toolName };
  }
}

/**
 * Opens the Office Tools sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('🧰 Office Tools')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}
