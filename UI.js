/**
 * @file UI.js
 * @description Creates custom menus on document load.
 */

/**
 * Triggered automatically when the spreadsheet is opened.
 * Adds a custom menu "Office Tools" to the UI.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🧰 Office Tools')
    .addItem('📏 Autofit Columns', 'autofitColumns')
    .addItem('💰 Deduct Amounts', 'deductAmountsFromColumnF')
    .addItem('📋 Export Departments to PDF', 'showDeptSelector')
    .addToUi();
}
