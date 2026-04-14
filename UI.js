/**
 * @file UI.js
 * @description Creates custom menus on document load.
 */

/**
 * Triggered automatically when the spreadsheet is opened.
 * Adds custom menus to the UI.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // Office Tools Menu
  ui.createMenu('🧰 Office Tools')
    .addItem('🚀 Open Sidebar', 'showSidebar')
    .addToUi();

  // Formatting Tools Menu
  ui.createMenu('✨ Formatting tools')
    .addItem('📏 Autofit Columns', 'autofitColumns')
    .addToUi();
}
