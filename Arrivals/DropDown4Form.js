const LISTS_SHEET_NAME = 'Lists';
const LISTS_FIRST_DATA_ROW = 2;

const DROPDOWN_BINDINGS = [
  {
    itemId: CONFIG.PROJECT_FORM_ITEM_ID,
    sourceColumn: 1,
    title: 'Project',
    label: 'Projects'
  },
  {
    itemId: CONFIG.RECRUITER_FORM_ITEM_ID,
    sourceColumn: 3,
    title: 'Recruiter',
    label: 'Recruiters'
  },
  {
    itemId: CONFIG.RESEARCHER_FORM_ITEM_ID,
    sourceColumn: 5,
    title: 'Researcher',
    label: 'Researchers'
  },
  {
    itemId: CONFIG.TEAM_FORM_ITEM_ID,
    sourceColumn: 7,
    title: 'Team',
    label: 'Team'
  }
];

/**
 * Updates configured Google Form dropdown items from the Lists sheet.
 * Parameters: none.
 *
 * @return {void}
 */
function updateForm() {
  const form = FormApp.openById(CONFIG.ARRIVALS_FORM_ID);
  const listsSheet = SpreadsheetApp.openById(CONFIG.LISTS_SPREADSHEET_ID).getSheetByName(LISTS_SHEET_NAME);

  DROPDOWN_BINDINGS.forEach(binding => {
    const choices = getColumnChoices(listsSheet, binding.sourceColumn);
    if (!choices.length) {
      Logger.log(`No choices found for ${binding.label}.`);
      return;
    }

    const item = getFormItem(form, binding);
    setItemChoices(item, choices, binding.label);
    Logger.log(`${binding.label}: updated ${choices.length} choices for "${item.getTitle()}" (${item.getType()}, ${item.getId()}).`);
  });
}

/**
 * Logs all form item titles, types, and IDs.
 * Parameters: none.
 *
 * @return {void}
 */
function logFormItems() {
  const form = FormApp.openById(CONFIG.ARRIVALS_FORM_ID);

  form.getItems().forEach(item => {
    Logger.log(`${item.getTitle()} | ${item.getType()} | ${item.getId()}`);
  });
}

/**
 * Gets a form item by configured ID and falls back to title lookup when needed.
 *
 * @param {GoogleAppsScript.Forms.Form} form Form that contains the item.
 * @param {Object} binding Dropdown binding configuration.
 * @return {GoogleAppsScript.Forms.Item} Matching form item.
 */
function getFormItem(form, binding) {
  try {
    const item = form.getItemById(binding.itemId);
    if (item.getTitle() === binding.title) return item;

    Logger.log(`${binding.label}: item ID ${binding.itemId} points to "${item.getTitle()}", expected "${binding.title}". Falling back to title lookup.`);
  } catch (err) {
    Logger.log(`${binding.label}: item ID ${binding.itemId} was not found. Falling back to title lookup.`);
  }

  const itemByTitle = form.getItems()
    .find(item => item.getTitle() === binding.title);

  if (!itemByTitle) {
    throw new Error(`${binding.label} item was not found by title "${binding.title}".`);
  }

  return itemByTitle;
}

/**
 * Updates choices for a supported Google Form choice item.
 *
 * @param {GoogleAppsScript.Forms.Item} item Form item to update.
 * @param {string[]} choices Choice values to apply.
 * @param {string} label Human-readable binding label for logs and errors.
 * @return {void}
 */
function setItemChoices(item, choices, label) {
  const itemType = item.getType();

  if (itemType === FormApp.ItemType.LIST) {
    item.asListItem().setChoiceValues(choices);
    return;
  }

  if (itemType === FormApp.ItemType.MULTIPLE_CHOICE) {
    item.asMultipleChoiceItem().setChoiceValues(choices);
    return;
  }

  throw new Error(`${label} item type is not supported: ${itemType}`);
}

/**
 * Reads non-empty choices from a source column in the Lists sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet with dropdown source values.
 * @param {number} sourceColumn One-based source column number.
 * @return {string[]} Non-empty dropdown choices.
 */
function getColumnChoices(sheet, sourceColumn) {
  const lastRow = sheet.getLastRow();
  const rowCount = lastRow - LISTS_FIRST_DATA_ROW + 1;

  if (rowCount <= 0) return [];

  return sheet.getRange(LISTS_FIRST_DATA_ROW, sourceColumn, rowCount, 1)
    .getValues()
    .flat()
    .map(value => String(value).trim())
    .filter(Boolean);
}
