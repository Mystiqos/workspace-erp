const LAST_GENERATED_ID_PROPERTY = 'LAST_GENERATED_ID';
const GENERATED_ID_LENGTH = 8;
const GENERATED_ID_MAX_VALUE = 99999999;

/**
 * Generates a unique 8-digit numeric ID using a locked persistent counter.
 * Parameters: none.
 *
 * @return {string} Unique 8-digit ID.
 */
function generateUniqueID() {
  return generateUniqueIDs(1)[0];
}

/**
 * Generates unique 8-digit numeric IDs using a locked persistent counter.
 *
 * @param {number} count Number of IDs to generate.
 * @return {string[]} Unique 8-digit IDs.
 */
function generateUniqueIDs(count) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    return getNextGeneratedIDs(count);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Sets the persistent ID counter to a known value.
 * Use this once before production if existing tables already contain generated IDs.
 *
 * @param {number|string} value Last generated numeric ID value.
 * @return {void}
 */
function setLastGeneratedID(value) {
  const numericValue = Number(value);
  if (!Number.isInteger(numericValue) || numericValue < 0 || numericValue > GENERATED_ID_MAX_VALUE) {
    throw new Error(`Invalid last generated ID value: ${value}`);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    PropertiesService.getScriptProperties()
      .setProperty(LAST_GENERATED_ID_PROPERTY, String(numericValue));
    Logger.log(`Last generated ID was set to ${formatGeneratedID(numericValue)}.`);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Logs the current persistent ID counter value.
 * Parameters: none.
 *
 * @return {void}
 */
function logLastGeneratedID() {
  const value = PropertiesService.getScriptProperties()
    .getProperty(LAST_GENERATED_ID_PROPERTY);

  Logger.log(`Last generated ID: ${value ? formatGeneratedID(Number(value)) : 'not set'}.`);
}

/**
 * Generates unique IDs without acquiring a lock.
 * The caller must hold the script lock before using this helper directly.
 *
 * @param {number} count Number of IDs to generate.
 * @return {string[]} Unique 8-digit IDs.
 */
function getNextGeneratedIDs(count) {
  if (!Number.isInteger(count) || count <= 0) {
    throw new Error(`Invalid generated ID count: ${count}`);
  }

  const properties = PropertiesService.getScriptProperties();
  const currentValue = Number(properties.getProperty(LAST_GENERATED_ID_PROPERTY)) || getInitialIDValue();
  const lastGeneratedValue = currentValue + count;

  if (lastGeneratedValue > GENERATED_ID_MAX_VALUE) {
    throw new Error(`Generated ID limit reached: ${GENERATED_ID_MAX_VALUE}`);
  }

  properties.setProperty(LAST_GENERATED_ID_PROPERTY, String(lastGeneratedValue));

  return Array.from({length: count}, (_, index) => formatGeneratedID(currentValue + index + 1));
}

/**
 * Creates the initial ID counter value from the current timestamp.
 * Parameters: none.
 *
 * @return {number} Initial 8-digit numeric counter value.
 */
function getInitialIDValue() {
  return Number(String(Date.now()).slice(-GENERATED_ID_LENGTH));
}

/**
 * Formats a numeric ID value as an 8-digit string.
 *
 * @param {number} value Numeric ID value.
 * @return {string} Zero-padded 8-digit ID.
 */
function formatGeneratedID(value) {
  return String(value).padStart(GENERATED_ID_LENGTH, '0');
}
