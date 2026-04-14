/**
 * @file DeductAmounts.js
 * @description Tool to deduct a specific amount from column F cumulatively.
 */

/**
 * API Version: Deducts the amount from row 2 downwards.
 * @param {number|string} amount The total amount to deduct.
 */
function apiDeductAmounts(amount) {
  try {
    let remainingDeduction = parseFloat(amount);

    if (isNaN(remainingDeduction) || remainingDeduction <= 0) {
      return { success: false, error: 'Please enter a valid positive number for deduction.' };
    }

    const sheet = SpreadsheetApp.getActiveSheet();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return { success: false, error: 'The sheet is empty or only contains a header.' };
    }

    const range = sheet.getRange(2, 6, lastRow - 1, 1);
    const values = range.getValues();
    let totalDeducted = 0;

    for (let i = 0; i < values.length; i++) {
      let cellValue = parseFloat(values[i][0]);
      if (!isNaN(cellValue) && cellValue > 0) {
        if (remainingDeduction >= cellValue) {
          remainingDeduction -= cellValue;
          totalDeducted += cellValue;
          values[i][0] = 0;
        } else {
          values[i][0] = cellValue - remainingDeduction;
          totalDeducted += remainingDeduction;
          remainingDeduction = 0;
        }
      }
      if (remainingDeduction <= 0) break;
    }

    range.setValues(values);
    return { 
      success: true, 
      message: 'Successfully deducted ' + totalDeducted.toFixed(2) + '. Remaining target: ' + remainingDeduction.toFixed(2)
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Legacy wrapper for the top menu.
 */
function deductAmountsFromColumnF() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Deduct Amounts', 'Enter the total amount to deduct from column F:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    const result = apiDeductAmounts(response.getResponseText());
    if (!result.success) {
      ui.alert('Error', result.error, ui.ButtonSet.OK);
    }
  }
}
