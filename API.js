/**
 * @file API.js
 * @description Central dispatcher for sidebar requests.
 */

/**
 * Dispatches calls from the sidebar to the appropriate backend function.
 * 
 * @param {object} request Provide {namespace, method} object string locating the target.
 * @param {any} payload The data to pass to the tool.
 * @return {object} Standardized result of the operation.
 */
function apiDispatcher(request, payload) {
  if (typeof request === 'string') {
    return { success: false, error: 'Legacy string requests are no longer supported. Please update to the new namespace architecture.' };
  }

  const { namespace, method } = request;
  
  try {
    // In GAS, top-level objects are accessible via the global scope (this)
    const targetNamespace = this[namespace];
    
    if (targetNamespace && typeof targetNamespace[method] === 'function') {
      const result = targetNamespace[method](payload);
      return result;
    }
    
    return { success: false, error: `Tool not found: ${namespace}.${method}` };
  } catch (e) {
    console.error(`Dispatcher Error [${namespace}.${method}]: ${e.toString()}`);
    return { success: false, error: `Execution error in backend: ${e.toString()}` };
  }
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
