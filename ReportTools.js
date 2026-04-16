/**
 * @file FilterAndsave.js
 * @description Logic for filtering the sheet by department and exporting to PDF.
 */

var ReportTools = ReportTools || {};

/**
 * Exports selected departments to PDF.
 * @param {string} input String of comma-separated departments.
 * @return {object} Result object {success, base64, fileName, message, error}
 */
ReportTools.exportDepartments = function(input) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) throw new Error("The sheet is empty.");

    const data = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const uniqueDepts = [...new Set(data.flat())].filter(String).sort();

    let selectedDepts = [];
    const trimmedInput = (input || '').trim();
    
    if (trimmedInput === '') {
      selectedDepts = uniqueDepts;
    } else {
      selectedDepts = trimmedInput.split(',').map(d => d.trim()).filter(String);
      const invalid = selectedDepts.filter(d => !uniqueDepts.includes(d));
      if (invalid.length > 0) {
        return { success: false, error: 'Invalid departments: ' + invalid.join(', ') };
      }
    }

    const result = processSelection(selectedDepts);
    if (result.success) {
      return {
        success: true,
        base64: result.bytes,
        fileName: result.fileName,
        message: 'PDF exported successfully for ' + selectedDepts.length + ' departments.'
      };
    } else {
      return { success: false, error: result.error };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
};


function processSelection(selectedDepts) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    // 1. Get ALL unique values to determine what to HIDE
    const allData = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const allDepts = [...new Set(allData.flat())].filter(String);
    
    // Calculate hidden values: Any dept in the sheet NOT in our selected list
    const hiddenDepts = allDepts.filter(dept => !selectedDepts.includes(dept));
    
    // 2. Apply Filter using setHiddenValues instead
    let filter = sheet.getFilter();
    if (filter) filter.remove();
    
    const range = sheet.getDataRange();
    filter = range.createFilter();
    
    if (hiddenDepts.length > 0) {
      const criteria = SpreadsheetApp.newFilterCriteria()
        .setHiddenValues(hiddenDepts)
        .build();
      filter.setColumnFilterCriteria(4, criteria); 
    }
    
    SpreadsheetApp.flush(); 
    Utilities.sleep(1000); // Allow time for filter UI update for PDF engine to catch up

    // 3. Construct Export URL
    const ssId = ss.getId();
    const sheetId = sheet.getSheetId();
    const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?format=pdf&size=A4" +
                "&portrait=true&fitw=true&sheetnames=false&printtitle=false" +
                "&pagenumbers=false&gridlines=false&fzr=true&gid=" + sheetId;

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("PDF Engine Error: " + response.getContentText());
    }

    const blob = response.getBlob();
    sheet.getFilter().remove(); // Cleanup

    return {
      success: true,
      bytes: Utilities.base64Encode(blob.getBytes()),
      fileName: "Export_" + new Date().getTime() + ".pdf"
    };

  } catch (e) {
    try { SpreadsheetApp.getActiveSheet().getFilter().remove(); } catch(i) {}
    return { success: false, error: e.toString() };
  }
}


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
    const sheet = ss.getActiveSheet();
    
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


