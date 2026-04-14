/**
 * @file Config.js
 * @description Central configuration for menus, roles, and permissions.
 */

const APP_CONFIG = {
  /**
   * Role Definitions:
   * - admin: All tools
   * - finance: Finance and Transactions
   * - hr: Reports and Status updates
   * - default: Basic viewing tools
   */
  ROLES: {
    ADMIN: 'admin',
    FINANCE: 'finance',
    HR: 'hr',
    DEFAULT: 'default'
  },

  /**
   * Map of email addresses to roles.
   * In a real production app, this could be fetched from a dedicated 'Users' sheet.
   */
  USER_ROLES: {
    'admin@example.com': 'admin',
    'finance_user@example.com': 'finance',
    'hr_user@example.com': 'hr'
  },

  /**
   * Menu structure configuration.
   */
  /**
   * Menu structure configuration.
   * Each object in this array represents a separate top-level custom menu in Google Sheets.
   */
  MENU_CONFIG: [
    {
      title: '🧰 Office Tools',
      items: [
        { label: '🚀 Open Sidebar', functionName: 'showSidebar', roles: ['admin', 'finance', 'hr', 'default'] }
      ]
    },
    {
      title: '✨ Formatting Tools',
      items: [
        { label: '📏 Autofit Columns', functionName: 'autofitColumns', roles: ['admin', 'finance', 'default'] }
      ]
    }
  ]
};

/**
 * Gets the role for the current user.
 * @return {string}
 */
function getCurrentUserRole() {
  const email = Session.getActiveUser().getEmail();
  return APP_CONFIG.USER_ROLES[email] || APP_CONFIG.ROLES.DEFAULT;
}
