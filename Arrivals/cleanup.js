/**
 * Runs a one-time manual cleanup for a specific test person ID.
 * Update the hardcoded ID manually before running this function from Apps Script.
 * Parameters: none.
 *
 * @return {Object} Cleanup summary grouped by target.
 */
function cleanupTestArrivalManually() {
  return cleanupArrivalByPersonID('98618951');
}

/**
 * Deletes rows created by the Arrivals flow for a specific person ID.
 *
 * @param {string} personID Person ID to clean up.
 * @return {Object} Cleanup summary grouped by target.
 */
function cleanupArrivalByPersonID(personID) {
  const normalizedPersonID = normalizeCleanupID(personID);
  if (!normalizedPersonID) {
    throw new Error('Person ID is required.');
  }

  const forceIDs = collectArrivalForceIDs(normalizedPersonID);
  const summary = [];

  getPersonCleanupTargets().forEach(target => {
    summary.push(deleteMatchingRows(target, normalizedPersonID));
  });

  getSatisfactionCleanupTargets().forEach(target => {
    summary.push(deleteMatchingRowsFromAllSheets(target, normalizedPersonID));
  });

  forceIDs.forEach(forceID => {
    getForceCleanupTargets().forEach(target => {
      summary.push(deleteMatchingRows(target, forceID));
    });
  });

  getCoworkingCleanupTargets().forEach(target => {
    summary.push(deleteMatchingRows(target, normalizedPersonID));
  });

  Logger.log(JSON.stringify({
    personID: normalizedPersonID,
    forceIDs: forceIDs,
    summary: summary
  }, null, 2));

  return {
    personID: normalizedPersonID,
    forceIDs: forceIDs,
    summary: summary
  };
}

/**
 * Finds force IDs linked to a person ID in Arrivals target sheets.
 *
 * @param {string} personID Person ID to search.
 * @return {string[]} Unique force IDs linked to the person ID.
 */
function collectArrivalForceIDs(personID) {
  const forceIDs = []
    .concat(findLinkedValues({
      spreadsheetId: CONFIG.FORCES_PROJECTS_SPREADSHEET_ID,
      sheetName: 'AllData',
      matchColumn: 1,
      valueColumn: 2,
      headerProfile: HEADER_PROFILES.FORCES_PROJECTS
    }, personID))
    .concat(findLinkedValues({
      spreadsheetId: CONFIG.LEAVE_TRACKER_SPREADSHEET_ID,
      sheetName: 'Balance',
      matchColumn: 1,
      valueColumn: 2
    }, personID));

  return Array.from(new Set(forceIDs.filter(Boolean)));
}

/**
 * Gets row deletion targets keyed by person ID.
 *
 * @return {Object[]} Cleanup targets.
 */
function getPersonCleanupTargets() {
  return [
    {name: 'WorkForce', spreadsheetId: CONFIG.WORKFORCE_SPREADSHEET_ID, sheetName: 'Registry', matchColumn: 2, headerProfile: HEADER_PROFILES.WORKFORCE},
    {name: 'Forces-Projects', spreadsheetId: CONFIG.FORCES_PROJECTS_SPREADSHEET_ID, sheetName: 'AllData', matchColumn: 1, headerProfile: HEADER_PROFILES.FORCES_PROJECTS},
    {name: 'Forces-Personals', spreadsheetId: CONFIG.FORCES_PERSONALS_SPREADSHEET_ID, sheetName: 'AllData', matchColumn: 1},
    {name: 'Forces-Locations', spreadsheetId: CONFIG.FORCES_LOCATIONS_SPREADSHEET_ID, sheetName: 'AllData', matchColumn: 1},
    {name: 'Forces-CorpContacts', spreadsheetId: CONFIG.FORCES_CORP_CONTACTS_SPREADSHEET_ID, sheetName: 'AllData', matchColumn: 1},
    {name: 'Insurances', spreadsheetId: CONFIG.INSURANCES_SPREADSHEET_ID, sheetName: 'Registry', matchColumn: 1},
    {name: 'Leave Tracker', spreadsheetId: CONFIG.LEAVE_TRACKER_SPREADSHEET_ID, sheetName: 'Balance', matchColumn: 1},
    {name: 'Forces-BackgroundChecks', spreadsheetId: CONFIG.FORCES_BACKGROUND_CHECKS_SPREADSHEET_ID, sheetName: 'Records', matchColumn: 1},
    {name: 'Forces-Legal', spreadsheetId: CONFIG.FORCES_LEGAL_SPREADSHEET_ID, sheetName: 'AllData', matchColumn: 1},
    {name: 'Forces-Finance', spreadsheetId: CONFIG.FORCES_FINANCE_SPREADSHEET_ID, sheetName: 'AllData', matchColumn: 1},
    {name: 'Finance Basement | PL', spreadsheetId: CONFIG.FINANCE_BASEMENT_PL_SPREADSHEET_ID, sheetName: 'Finance', matchColumn: 1, headerProfile: HEADER_PROFILES.FINANCE_BASEMENT_PL},
    {name: 'Agreements', spreadsheetId: CONFIG.AGREEMENTS_SPREADSHEET_ID, sheetName: 'Contracts', matchColumn: 1, headerProfile: HEADER_PROFILES.AGREEMENTS}
  ];
}

/**
 * Gets row deletion targets keyed by force ID.
 *
 * @return {Object[]} Cleanup targets.
 */
function getForceCleanupTargets() {
  return [
    {name: 'Telegram bot database', spreadsheetId: CONFIG.TELEGRAM_BOT_DATABASE_SPREADSHEET_ID, sheetName: 'chatsDB', matchColumn: 1},
    {name: 'Forces-Inventory', spreadsheetId: CONFIG.FORCES_INVENTORY_SPREADSHEET_ID, sheetName: 'Teams', matchColumn: 1}
  ];
}

/**
 * Gets all-sheets deletion targets keyed by person ID.
 *
 * @return {Object[]} Cleanup targets.
 */
function getSatisfactionCleanupTargets() {
  return [
    {name: 'Forces-Satisfactions', spreadsheetId: CONFIG.FORCES_SATISFACTIONS_SPREADSHEET_ID, matchColumn: 1}
  ];
}

/**
 * Gets coworking deletion targets keyed by person ID.
 *
 * @return {Object[]} Cleanup targets.
 */
function getCoworkingCleanupTargets() {
  return (CONFIG.COWORKING_INTEGRATIONS || []).reduce((targets, integration) => {
    targets.push({
      name: `${integration.displayName} member register`,
      spreadsheetId: integration.memberRegisterSpreadsheetId,
      sheetName: 'List',
      matchColumn: 1,
      headerProfile: HEADER_PROFILES.COWORKING_MEMBER_REGISTER
    });
    targets.push({
      name: `${integration.displayName} resident badge register`,
      spreadsheetId: integration.residentBadgeRegisterSpreadsheetId,
      sheetName: 'List',
      matchColumn: 1
    });
    targets.push({
      name: `${integration.displayName} coworker finance register`,
      spreadsheetId: integration.financeSpreadsheetId,
      sheetName: 'Coworkers',
      matchColumn: 1
    });
    return targets;
  }, []);
}

/**
 * Finds values from one column in rows where another column matches the target ID.
 *
 * @param {Object} target Lookup target.
 * @param {string} target.spreadsheetId Spreadsheet ID.
 * @param {string} target.sheetName Sheet name.
 * @param {number} target.matchColumn Column used to match the target ID.
 * @param {number} target.valueColumn Column with linked values.
 * @param {string} matchValue Target ID to match.
 * @return {string[]} Linked values.
 */
function findLinkedValues(target, matchValue) {
  const sheet = SpreadsheetApp.openById(target.spreadsheetId).getSheetByName(target.sheetName);
  if (!sheet) return [];

  const firstDataRow = getCleanupFirstDataRow(sheet, target.headerProfile);
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) return [];

  const rowCount = lastRow - firstDataRow + 1;
  const matchValues = sheet.getRange(firstDataRow, target.matchColumn, rowCount, 1).getValues();
  const linkedValues = sheet.getRange(firstDataRow, target.valueColumn, rowCount, 1).getValues();

  return matchValues.reduce((result, row, index) => {
    if (normalizeCleanupID(row[0]) === matchValue) {
      result.push(normalizeCleanupID(linkedValues[index][0]));
    }
    return result;
  }, []);
}

/**
 * Deletes matching rows from one sheet.
 *
 * @param {Object} target Cleanup target.
 * @param {string} target.name Human-readable target name.
 * @param {string} target.spreadsheetId Spreadsheet ID.
 * @param {string} target.sheetName Sheet name.
 * @param {number} target.matchColumn One-based column to match.
 * @param {string} matchValue Target ID to match.
 * @return {Object} Target cleanup summary.
 */
function deleteMatchingRows(target, matchValue) {
  const sheet = SpreadsheetApp.openById(target.spreadsheetId).getSheetByName(target.sheetName);
  if (!sheet) {
    return {target: target.name, sheetName: target.sheetName, deletedRows: [], error: 'Sheet not found'};
  }

  removeFiltersIfAny(sheet);
  const deletedRows = deleteRowsByColumnValue(sheet, target.matchColumn, matchValue, target.headerProfile);
  return {target: target.name, sheetName: target.sheetName, deletedRows: deletedRows};
}

/**
 * Deletes matching rows from every sheet in a spreadsheet.
 *
 * @param {Object} target Cleanup target.
 * @param {string} target.name Human-readable target name.
 * @param {string} target.spreadsheetId Spreadsheet ID.
 * @param {number} target.matchColumn One-based column to match.
 * @param {string} matchValue Target ID to match.
 * @return {Object} Target cleanup summary.
 */
function deleteMatchingRowsFromAllSheets(target, matchValue) {
  const spreadsheet = SpreadsheetApp.openById(target.spreadsheetId);
  const sheetSummaries = spreadsheet.getSheets().map(sheet => {
    removeFiltersIfAny(sheet);
    return {
      sheetName: sheet.getName(),
      deletedRows: deleteRowsByColumnValue(sheet, target.matchColumn, matchValue, target.headerProfile)
    };
  });

  return {target: target.name, sheets: sheetSummaries};
}

/**
 * Deletes rows where a column value matches the target ID.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to clean.
 * @param {number} matchColumn One-based column to match.
 * @param {string} matchValue Target ID to match.
 * @param {string[]=} headerProfile Required headers used to detect the correct header row.
 * @return {number[]} Deleted one-based row numbers.
 */
function deleteRowsByColumnValue(sheet, matchColumn, matchValue, headerProfile) {
  const firstDataRow = getCleanupFirstDataRow(sheet, headerProfile);
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) return [];

  const values = sheet.getRange(firstDataRow, matchColumn, lastRow - firstDataRow + 1, 1).getValues();
  const rowsToDelete = values.reduce((rows, row, index) => {
    if (normalizeCleanupID(row[0]) === matchValue) {
      rows.push(firstDataRow + index);
    }
    return rows;
  }, []);

  rowsToDelete.slice().reverse().forEach(row => sheet.deleteRow(row));
  return rowsToDelete;
}

/**
 * Gets the first data row while protecting header rows from deletion.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to inspect.
 * @param {string[]=} headerProfile Required headers used to detect the correct header row.
 * @return {number} First one-based data row.
 */
function getCleanupFirstDataRow(sheet, headerProfile) {
  const headerRow = findHeaderRow(sheet, headerProfile);
  if (!headerRow && typeof logMissingHeaderRow === 'function') {
    logMissingHeaderRow(sheet, headerProfile);
  }
  return Math.max((headerRow || 1) + 1, 2);
}

/**
 * Normalizes an ID value for exact string matching.
 *
 * @param {*} value ID value.
 * @return {string} Normalized ID string.
 */
function normalizeCleanupID(value) {
  return String(value || '').trim();
}
