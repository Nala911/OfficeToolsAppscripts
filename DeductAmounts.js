/**
 * @file DeductAmounts.js
 * @description Tool to deduct a specific amount from column F cumulatively.
 */

/**
 * Prompts the user for a total amount to deduct from column F.
 * Deducts the amount from row 2 downwards until the target is reached or cells are exhausted.
 */
function deductAmountsFromColumnF() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Deduct Amounts', 'Enter the total amount to deduct from column F:', ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  if (response.getSelectedButton() !== ui.Button.OK) {
    return; // User cancelled
  }

  const text = response.getResponseText();
  let remainingDeduction = parseFloat(text);

  // Validate the input amount
  if (isNaN(remainingDeduction) || remainingDeduction <= 0) {
    ui.alert('Invalid Input', 'Please enter a valid positive number.', ui.ButtonSet.OK);
    return;
  }

  const originalDeduction = remainingDeduction;
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('No Data', 'The sheet is empty or only contains a header.', ui.ButtonSet.OK);
    return;
  }

  // Column F is the 6th column
  const range = sheet.getRange(2, 6, lastRow - 1, 1);
  const values = range.getValues();
  let totalDeducted = 0;

  // Iterate through the values in column F
  for (let i = 0; i < values.length; i++) {
    let cellValue = parseFloat(values[i][0]);

    if (!isNaN(cellValue) && cellValue > 0) {
      if (remainingDeduction >= cellValue) {
        // Entire cell value is deducted
        remainingDeduction -= cellValue;
        totalDeducted += cellValue;
        values[i][0] = 0;
      } else {
        // Only part of the cell value is deducted
        values[i][0] = cellValue - remainingDeduction;
        totalDeducted += remainingDeduction;
        remainingDeduction = 0;
      }
    }

    // Stop if we've deducted the target amount
    if (remainingDeduction <= 0) {
      break;
    }
  }

  // Update the sheet with new values
  range.setValues(values);

  // Final summary alert
  if (remainingDeduction > 0) {
    ui.alert('Deduction Complete', 
             'Successfully deducted ' + totalDeducted.toFixed(2) + '.\n' +
             'Remaining amount of ' + remainingDeduction.toFixed(2) + ' could not be deducted as column F was exhausted.', 
             ui.ButtonSet.OK);
  } else {
    ui.alert('Deduction Complete', 'Successfully deducted the full amount of ' + originalDeduction.toFixed(2) + '.', ui.ButtonSet.OK);
  }
}
