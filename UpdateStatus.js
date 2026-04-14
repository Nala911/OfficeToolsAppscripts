/**
 * @file UpdateStatus.js
 * @description Updates column G (Status) for a specific department in column D.
 */

var ReportTools = ReportTools || {};

/**
 * Updates the status for a given department.
 * @param {object} input Object containing {statusDept, statusVal}
 * @return {object} Result object {success, message, error}
 */
ReportTools.updateDepartmentStatus = function(input) {
  try {
    const deptName = (input.statusDept || '').trim();
    const statusValue = (input.statusVal || '').trim();

    if (!deptName) return { success: false, error: 'Department name is required.' };
    if (!statusValue) return { success: false, error: 'Status value is required.' };

    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) return { success: false, error: 'The sheet is empty.' };

    // Column D is index 3 (4th col), Column G is index 6 (7th col)
    const range = sheet.getRange(2, 1, lastRow - 1, 7); 
    const values = range.getValues();
    let updatedCount = 0;

    for (let i = 0; i < values.length; i++) {
      if (String(values[i][3]).trim().toLowerCase() === deptName.toLowerCase()) {
        values[i][6] = statusValue;
        updatedCount++;
      }
    }

    if (updatedCount > 0) {
      // Write back the updated values
      range.setValues(values);
      return { success: true, message: `Updated ${updatedCount} row(s) to '${statusValue}' for department '${deptName}'.` };
    } else {
      return { success: false, error: `No rows found for department '${deptName}'.` };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
};

/**
 * Fetches unique department names from Column D of the "Table1" sheet.
 * @return {string[]} Array of unique department names.
 */
ReportTools.getUniqueDepartments = function() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Table1');
    
    // Fallback to active sheet if 'Table1' doesn't exist
    if (!sheet) {
      sheet = ss.getActiveSheet();
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Column D is 4th column
    const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const uniqueDepts = new Set();
    
    values.forEach(row => {
      const dept = String(row[0]).trim();
      if (dept) {
        uniqueDepts.add(dept);
      }
    });

    return Array.from(uniqueDepts).sort();
  } catch (e) {
    console.error('Error in getUniqueDepartments:', e.toString());
    return [];
  }
};

/**
 * Global wrappers for backward compatibility and API Dispatcher.
 */
function apiUpdateDepartmentStatus(input) {
  return ReportTools.updateDepartmentStatus(input);
}

function getUniqueDepartments() {
  return ReportTools.getUniqueDepartments();
}
