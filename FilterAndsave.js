

/**
 * API Version: Exports selected departments to PDF.
 * @param {string} input String of comma-separated departments.
 */
function apiExportDepartments(input) {
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
}

/**
 * Legacy wrapper for the top menu.
 */
function showDeptSelector() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error("The sheet appears to be empty.");

    const data = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const uniqueDepts = [...new Set(data.flat())].filter(String).sort();

    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Filter and Export to PDF',
      'Available Departments:\n' + uniqueDepts.join(', ') + 
      '\n\nEnter departments to export (comma-separated) or leave blank for all:',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    const result = apiExportDepartments(response.getResponseText());
    
    if (result.success) {
      const htmlOutput = HtmlService.createHtmlOutput(`
        <script>
          const link = document.createElement('a');
          link.href = 'data:application/pdf;base64,${result.base64}';
          link.download = '${result.fileName}';
          link.click();
          setTimeout(() => google.script.host.close(), 1000);
        </script>
      `).setWidth(350).setHeight(100);
      ui.showModalDialog(htmlOutput, 'Exporting...');
    } else {
      ui.alert('Export Failed', result.error, ui.ButtonSet.OK);
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: " + e.message);
  }
}

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