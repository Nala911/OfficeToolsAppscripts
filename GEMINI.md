# 🧰 Office Apps Scripts - GEMINI.md

## 🎯 Project Purpose & User Requirements
The primary goal of this project is to consolidate multiple, previously disjointed Google Apps Script (GAS) tools into a single, unified, and **elegant** interface within Google Sheets.

### Key Requirements:
- **Unified Interface**: All tools must be accessible from a single, persistent **Sidebar**.
- **Scalability**: The system must be robust enough to handle many tools without becoming cluttered or unmanageable.
- **Elegant UX/UI**: High emphasis on a modern, professional look using the **Inter** font family, a card-based layout, and a clean slate/blue color palette.
- **Integrated Inputs**: Tools should use the sidebar for user input (text fields, dropdowns) instead of intrusive modal dialog boxes or prompts whenever possible.
- **Automated Deployment**: Use `clasp` for local development and reliable syncing with the Apps Script environment.

## 🏛️ Architectural Blueprint: The "API-First" Dispatcher
To achieve robustness and scalability, the project follows a strict decoupled architecture.

### 1. Frontend (The "Face") - `Sidebar.html`
- **UI Components**: Cards/Sections with clear titles and integrated form elements.
- **Client JS**: A universal `runTool(toolName, inputIds)` function that:
  - Collects values from one or multiple input fields.
  - Sends a "Payload" (data object) to the backend.
  - Handles the "State" (disabling buttons, showing a loading spinner).
  - Displays a "Toast" message (success/error) upon completion.

### 2. Communication Hub - `API.js`
- **The Dispatcher**: A central `apiDispatcher` function that acts as a router. Every client request passes through here before reaching the specific tool logic.
- **Benefits**: Simplifies the frontend-backend bridge and allows for easy auditing or logging of all tool calls in one place.

### 3. Backend (The "Brain") - Specialized JS Files
- **API Version Functions**: Every tool logic starts with an `api...` function (e.g., `apiUpdateDepartmentStatus`).
- **Standardized Response**: All backend functions MUST return a JSON object:
  ```json
  {
    "success": true, 
    "message": "Operation description",
    "base64": "...", // Optional: For file downloads
    "fileName": "..." // Optional: For file downloads
  }
  ```

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
1.  **Create Backend**: Add the logic in a new `.js` file or append to an existing one. Use the `api...` prefix.
2.  **Standardize Return**: Ensure it returns the `{success, message}` object.
3.  **Register Dispatcher**: Add a new `case` to `apiDispatcher` in `API.js`.
4.  **Update UI**: Add a new `<div class="section">` in `Sidebar.html` with necessary inputs.
5.  **Connect**: Call `runTool('yourNewTool', 'inputId')` from the button's `onclick` event.

## 🎨 UI/UX Guidelines
- **Typography**: Use the **Inter** font (Google Fonts import).
- **Themes**: Primary color `#2563eb`, background `#f8fafc`.
- **Transitions**: Use `all 0.2s ease` for buttons and input focus.
- **Feedback**: Never use `ui.alert()` in new tools; always use the `#status` toast area in the sidebar.

---
*This document serves as the primary instructional context for the Office Apps Scripts project.*
