/**
 * Deletes safe trailing empty rows, inserts one new row,
 * and returns the inserted row number.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet where the new row should be created.
 * @param {number} controlColumn One-based column number used as the primary key column.
 * @return {number} Inserted row number.
 */
function emptyRowsDel(sheet, controlColumn) {
  const maxRows = sheet.getMaxRows();
  const frozenRows = sheet.getFrozenRows();
  const firstDataRow = frozenRows + 1;
  const lastColumn = sheet.getLastColumn();

  if (maxRows < firstDataRow) {
    sheet.insertRowsAfter(maxRows, firstDataRow - maxRows);
  }

  const scanRowCount = Math.max(sheet.getMaxRows() - firstDataRow + 1, 1);
  const controlRange = sheet.getRange(firstDataRow, controlColumn, scanRowCount, 1);
  const controlValues = controlRange.getValues();
  const controlFormulas = controlRange.getFormulas();
  const rowRange = sheet.getRange(firstDataRow, 1, scanRowCount, lastColumn);
  const rowValues = rowRange.getValues();
  const rowFormulas = rowRange.getFormulas();

  const lastControlRow = getLastFilledRow(controlValues, controlFormulas, firstDataRow);
  const lastContentRow = getLastFilledRow(rowValues, rowFormulas, firstDataRow);
  const hasDataRows = lastControlRow >= firstDataRow || lastContentRow >= firstDataRow;
  const insertAfterRow = hasDataRows ? Math.max(lastControlRow, lastContentRow) : Math.max(frozenRows, 1);

  const firstDeleteRow = hasDataRows ? insertAfterRow + 1 : firstDataRow + 1;
  const deleteCount = sheet.getMaxRows() - firstDeleteRow + 1;
  if (deleteCount > 0) {
    sheet.deleteRows(firstDeleteRow, deleteCount);
  }

  sheet.insertRowAfter(insertAfterRow);

  return insertAfterRow + 1;
}

/**
 * Gets the last row number that contains at least one value or formula.
 *
 * @param {*[][]} values Range values.
 * @param {string[][]} formulas Range formulas.
 * @param {number} firstRow One-based row number of the first range row.
 * @return {number} Last filled one-based row number, or the row above the range when empty.
 */
function getLastFilledRow(values, formulas, firstRow) {
  for (let rowIndex = values.length - 1; rowIndex >= 0; rowIndex--) {
    if (isFilledRow(values[rowIndex], formulas[rowIndex])) {
      return firstRow + rowIndex;
    }
  }

  return firstRow - 1;
}

/**
 * Checks whether at least one value or formula is present in a row.
 *
 * @param {*[]} rowValues Row values.
 * @param {string[]} rowFormulas Row formulas.
 * @return {boolean} True when the row contains a value or formula.
 */
function isFilledRow(rowValues, rowFormulas) {
  return rowValues.some((value, index) => value !== '' || rowFormulas[index] !== '');
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
