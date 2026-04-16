# 🧰 Office Apps Scripts - GEMINI.md

## 🎯 Project Purpose & User Requirements
The primary goal of this project is to consolidate multiple Google Apps Script (GAS) tools into a unified, scalable platform within Google Sheets.

### Key Requirements:
- **Flexible UI Architecture**: The project supports combining multiple tools into a unified, vertically stacked sidebar, OR keeping standalone tools in custom top-level Google Sheets menus, depending on user discretion.
- **Multi-Menu Integration**: Support for multiple custom top-level menus in the Spreadsheet UI, each containing specialized tool items.
- **Scalability & Robustness**: A decoupled architecture designed to handle hundreds of tools across various categories (Finance, HR, Formatting, etc.).
- **Role-Based Access Control (RBAC)**: Dynamic visibility of menus and sidebar tools based on the active user's role.

## 🏛️ Architectural Blueprint: Scalable "API-First" Framework

### 1. Dynamic Configuration (`Config.js`)
- **Central Source of Truth**: All menus, sub-items, and role-based permissions are defined in the `APP_CONFIG` object.
- **Menu Generation**: The spreadsheet UI is built dynamically on `onOpen` by iterating through the configuration.

### 2. Frontend (Single Page Application) - `Sidebar.html`
- **Modular Components**: The sidebar UI is composed of smaller HTML fragments (e.g., `Sidebar_Transactions.html`) included dynamically via a backend `include()` helper.
- **Unified Interface**: All tools are vertically stacked for easy access, with automatic role-based visibility filtering.
- **Client JS**: A universal `runTool({namespace, method}, inputIds)` function that facilitates communication with the backend.

### 3. Namespace-Based Backend (`*.js`)
- **Logical Grouping**: Tools are grouped into global namespace objects (e.g., `FinanceTools`, `ReportTools`) to prevent global variable pollution and ease maintainability.
- **Standardized Response**: All functions return a consistent JSON object: `{success: true/false, message: "...", base64: "...", fileName: "..."}`.

### 4. Communication Hub (`API.js`)
- **The Dispatcher**: A central `apiDispatcher(request, payload)` function that routes client requests to the appropriate namespace and method dynamically.
- **Secure Invocation**: Resolves target functions using the global scope (`this[namespace][method]`), ensuring a robust bridge between frontend and backend.

## 🛠️ Current Tools & Instructions

| Tool Name | Input(s) | Functionality |
| :--- | :--- | :--- |
| **Autofit Columns** | None | Automatically resizes all sheet columns to fit content. (Located in '✨ Formatting tools' menu). |
| **Highlight Transaction** | Transaction ID | Finds an ID in Column B and highlights the row in yellow. |
| **Generate Letter PDF** | Transaction ID | Maps row data to `Template.html` and generates a downloadable PDF. |
| **Deduct Amounts** | Amount | Cumulatively deducts a total amount from Column F, row by row. |
| **Export Reports** | Dept Name(s) | Filters the sheet by department and exports a PDF report. |
| **Update Status** | Dept Name + Status | Updates Column G for all rows matching the Dept in Column D to 'Paid' or 'Pending'. |

## 🚀 Development Instructions

### Syncing Changes
Always use **clasp** to push your local changes to the cloud:
```sh
clasp push
```

### Adding a New Tool (Step-by-Step)
1. **Create Backend**: Add the logic in a new `.js` file or append to an existing namespace object (e.g., `FinanceTools.myNewTool`).
2. **Standardize Return**: Ensure the namespace function returns the `{success, message}` object.
3. **Register via Interface**: 
   - *If Sidebar*: Add a new `<div class="section">` in a `Sidebar_*.html` snippet with inputs. Call `runTool({namespace: 'Obj', method: 'func'}, 'inputId')` from the button.
   - *If Custom Menu*: Create a global wrapper function in the backend (e.g., `function menuMyNewTool() { ... }`) that calls your namespace tool and uses `SpreadsheetApp.getActive().toast()` or `ui.alert()` for feedback. Then register this function in `Config.js` MENU_CONFIG.

## 🎨 UI/UX Guidelines
- **Typography**: Use the **Inter** font (Google Fonts import).
- **Themes**: Primary color `#2563eb`, background `#f8fafc`.
- **Transitions**: Use `all 0.2s ease` for buttons and input focus.
- **Feedback**: Never use `ui.alert()` in new tools; always use the `#status` toast area in the sidebar.

---
*This document serves as the primary instructional context for the Office Apps Scripts project.*
