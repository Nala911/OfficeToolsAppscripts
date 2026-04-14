/**
 * @file UI.js
 * @description Dynamically creates custom menus based on user roles and configuration.
 */

/**
 * Triggered automatically when the spreadsheet is opened.
 * Dynamically builds menus for the authorized role.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const role = getCurrentUserRole();

  APP_CONFIG.MENU_CONFIG.forEach(menu => {
    // Check if at least one item in the menu is authorized for the user
    const authorizedItems = menu.items.filter(item => 
      item.roles.includes(role) || item.roles.includes('all')
    );

    if (authorizedItems.length > 0) {
      const customMenu = ui.createMenu(menu.title);
      
      authorizedItems.forEach(item => {
        customMenu.addItem(item.label, item.functionName);
      });
      
      customMenu.addToUi();
    }
  });
}
