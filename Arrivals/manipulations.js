/**
 * Deletes trailing empty rows based on a control column, inserts one new row,
 * and returns the inserted row number.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet where the new row should be created.
 * @param {number} controlColumn One-based column number used to detect trailing empty rows.
 * @return {number} Inserted row number.
 */
function emptyRowsDel(sheet, controlColumn) {
  const maxRows = sheet.getMaxRows();
  const controlRange = sheet.getRange(1, controlColumn, maxRows, 1);
  const controlValues = controlRange.getValues();
  const controlFormulas = controlRange.getFormulas();

  let lastFilledRow = 1;
  for (let rowIndex = controlValues.length - 1; rowIndex >= 0; rowIndex--) {
    if (controlValues[rowIndex][0] !== '' || controlFormulas[rowIndex][0] !== '') {
      lastFilledRow = rowIndex + 1;
      break;
    }
  }

  const firstTrailingRow = lastFilledRow + 1;
  if (firstTrailingRow <= maxRows) {
    sheet.deleteRows(firstTrailingRow, maxRows - lastFilledRow);
  }

  sheet.insertRowAfter(lastFilledRow);

  return lastFilledRow + 1;
}

/**
 * Copies formatting from the previous row to the target row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet that contains the target row.
 * @param {number} lastRow Target row number.
 * @return {void}
 */
function format4LastRow(sheet, lastRow) {
  const sourceRange = sheet.getRange(lastRow - 1, 1, 1, sheet.getLastColumn());
  sourceRange.copyTo(sheet.getRange(lastRow, 1), {formatOnly: true});
}
