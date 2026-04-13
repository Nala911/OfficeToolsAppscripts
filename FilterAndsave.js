

function showDeptSelector() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      throw new Error("The sheet appears to be empty or only has a header.");
    }

    const data = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const uniqueDepts = [...new Set(data.flat())].filter(String).sort();

    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Filter and Export to PDF',
      'Available Departments:\n' + uniqueDepts.join(', ') + 
      '\n\nEnter departments to export (comma-separated) or leave blank to export all:',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    let selectedDepts = [];
    const input = response.getResponseText().trim();
    
    if (input === '') {
      selectedDepts = uniqueDepts;
    } else {
      selectedDepts = input.split(',').map(d => d.trim()).filter(String);
      const invalid = selectedDepts.filter(d => !uniqueDepts.includes(d));
      if (invalid.length > 0) {
        ui.alert('Invalid departments entered: ' + invalid.join(', '));
        return;
      }
    }

    // Give some immediate feedback that it's running
    SpreadsheetApp.getActiveSpreadsheet().toast('Processing PDF export...', 'Please Wait', 5);

    const result = processSelection(selectedDepts);
    
    if (result.success) {
      ui.alert('Export Successful', `Saved to Drive File:\n${result.fileName}\n\nURL:\n${result.url}`, ui.ButtonSet.OK);
    } else {
      ui.alert('Export Failed', "Error: " + result.error, ui.ButtonSet.OK);
    }

  } catch (e) {
    SpreadsheetApp.getUi().alert("Error loading departments: " + e.message);
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

    // Save directly to Drive instead of returning base64
    const blob = response.getBlob().setName("Export_" + new Date().getTime() + ".pdf");
    const file = DriveApp.createFile(blob);

    sheet.getFilter().remove(); // Cleanup

    return {
      success: true,
      fileName: file.getName(),
      url: file.getUrl()
    };

  } catch (e) {
    try { SpreadsheetApp.getActiveSheet().getFilter().remove(); } catch(i) {}
    return { success: false, error: e.toString() };
  }
}