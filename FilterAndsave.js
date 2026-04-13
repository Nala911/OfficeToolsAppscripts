

function showDeptSelector() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      throw new Error("The sheet appears to be empty or only has a header.");
    }

    const data = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
    const uniqueDepts = [...new Set(data.flat())].filter(String).sort();

    const htmlTemplate = HtmlService.createTemplateFromFile('DownloadDialog');
    htmlTemplate.departments = uniqueDepts;
    
    const html = htmlTemplate.evaluate()
      .setWidth(400)
      .setHeight(450);
      
    SpreadsheetApp.getUi().showModalDialog(html, 'Select Departments');
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
    
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(hiddenDepts) // This is the fix
      .build();
    
    filter.setColumnFilterCriteria(4, criteria); 
    SpreadsheetApp.flush(); 

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
      throw new Error("PDF Engine Error: " + response.getResponseCode());
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