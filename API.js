/**
 * @file API.js
 * @description Central dispatcher for sidebar requests.
 */

/**
 * Dispatches calls from the sidebar to the appropriate backend function.
 * Supports both legacy string-based tool names and the new namespace/method objects.
 * 
 * @param {string|object} request Either a toolName string (legacy) or {namespace, method} object.
 * @param {any} payload The data to pass to the tool.
 * @return {object} Result of the operation.
 */
function apiDispatcher(request, payload) {
  // 1. Handle Legacy String Requests
  if (typeof request === 'string') {
    return handleLegacyRequest(request, payload);
  }

  // 2. Handle New Namespace/Method Requests
  const { namespace, method } = request;
  
  try {
    // In GAS, top-level objects are accessible via the global scope (this)
    const targetNamespace = this[namespace];
    
    if (targetNamespace && typeof targetNamespace[method] === 'function') {
      return targetNamespace[method](payload);
    }
    
    return { success: false, error: `Tool not found: ${namespace}.${method}` };
  } catch (e) {
    console.error(`Dispatcher Error: ${e.toString()}`);
    return { success: false, error: `Execution error: ${e.toString()}` };
  }
}

/**
 * Map of legacy tool names to their new namespace implementations.
 */
function handleLegacyRequest(toolName, inputVal) {
  const legacyMap = {
    'autofitColumns': () => FormattingTools.autofitColumns(),
    'highlightTransaction': (val) => FormattingTools.highlightTransaction(val),
    'generateLetter': (val) => TransactionTools.generateLetter(val),
    'deductAmount': (val) => FinanceTools.deductAmounts(val),
    'exportDepts': (val) => ReportTools.exportDepartments(val),
    'updateStatus': (val) => ReportTools.updateDepartmentStatus(val),
    'getUniqueDepartments': () => ReportTools.getUniqueDepartments()
  };

  if (legacyMap[toolName]) {
    return legacyMap[toolName](inputVal);
  }

  return { success: false, error: 'Unknown legacy tool: ' + toolName };
}

/**
 * Opens the Office Tools sidebar.
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('🧰 Office Tools')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Helper to include HTML files in templates.
 * @param {string} filename 
 * @return {string} HTML content.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
