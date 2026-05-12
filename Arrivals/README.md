# Arrivals Module

The Arrivals module automates the onboarding registration flow for new suppliers.

Recruiters submit all required data through a Google Form. The form writes the original submission into the connected Arrivals responses spreadsheet. These original response rows are treated as the source backup and should not be modified by scripts.

After each form submission, the Apps Script flow distributes the submitted data into department-specific spreadsheets. Each department receives only the fields needed for its workflow. The module also sends email notifications to the relevant teams so they can process the new arrival according to the standard internal flow.

**Core Flow**

1. A recruiter submits the Arrivals Google Form.
2. Google Forms stores the raw response in the connected Arrivals responses spreadsheet.
3. `onFormSubmit(e)` reads the submitted values.
4. The script generates internal IDs.
5. The script writes mapped data into HR, project, finance, legal, tech, coworking, and other target spreadsheets.
6. The script sends workflow email notifications.
7. Any processing errors are collected and sent as an error summary.

**Files**

`Code.js`  
Main Arrivals processing flow. Contains `onFormSubmit(e)`, spreadsheet writes, notifications, helper functions, and validation-safe value handling.

`DropDown4Form.js`  
Updates Google Form dropdown and multiple-choice options from the configured Lists spreadsheet.

`idsGeneration.js`  
Generates locked sequential 8-digit IDs using `PropertiesService` and `LockService`.

`manipulations.js`  
Shared spreadsheet row helpers, including safe row insertion and formatting.

`cleanup.js`  
Manual maintenance helpers for deleting test rows by `personID` across Arrivals target spreadsheets.

`config.example.js`  
GitHub-safe configuration template. It documents required config keys but contains no real IDs, emails, or business-specific names.

`config.local.js`  
Local real configuration used by Apps Script. This file is ignored by Git but is pushed to Apps Script by `clasp`.

`.clasp.json`  
Active local Apps Script binding. This file is ignored by Git.

`.clasp.prod.json` / `.clasp.test.json`  
Local production and test Apps Script bindings. Copy one of them into `.clasp.json` before running `clasp` commands.

`.claspignore`  
Controls which files are excluded from Apps Script deployments. For example, `config.example.js` is kept in GitHub but not pushed to Apps Script.

**Configuration**

All real resource IDs, email recipients, client names, coworking brands, and integration values must live in `config.local.js`.

Tracked code should use `CONFIG`, for example:

```js
SpreadsheetApp.openById(CONFIG.WORKFORCE_SPREADSHEET_ID);
```

Do not hardcode real IDs, emails, or business names in tracked files.

`config.example.js` should be updated whenever a new required config key is added.

**Deployment**

Run `clasp` commands from the module folder:

```bash
cd /Users/vadym_1/CodeSpace/GAS/Arrivals
```

Before deploying to production:

```bash
cp .clasp.prod.json .clasp.json
clasp status
clasp push
```

Before deploying to the test Apps Script project:

```bash
cp .clasp.test.json .clasp.json
clasp status
clasp push
```

Always check `clasp status` before `clasp push`.

Expected tracked files for Apps Script:

```text
appsscript.json
cleanup.js
Code.js
config.local.js
DropDown4Form.js
idsGeneration.js
manipulations.js
```

`config.example.js` and `.clasp*.json` should not be pushed to Apps Script.

**Testing**

For production smoke tests:

1. Use clearly marked test values in the form, such as `TEST` in first or last name.
2. Use valid dropdown values where target spreadsheets have data validation.
3. Submit one controlled test response.
4. Check Apps Script `Executions`.
5. Check whether an error summary email was sent.
6. Verify the affected target spreadsheets.
7. Clean up test rows with `cleanupArrivalByPersonID(personID)`.

**Cleanup Test Data**

Use `cleanup.js` for manual cleanup after test submissions.

For a one-time cleanup, update the hardcoded ID in:

```js
function cleanupTestArrivalManually() {
  return cleanupArrivalByPersonID('put-person-id-here');
}
```

Then run `cleanupTestArrivalManually()` manually from Apps Script.

The cleanup searches for related `forceID` values and deletes matching rows from the configured Arrivals target spreadsheets. It logs a JSON summary of deleted rows.

Cleanup does not undo side effects such as already-sent emails or moved Drive files.

**ID Generation**

IDs are generated sequentially through `idsGeneration.js`.

Before using the module on a dataset with existing IDs, seed the last generated value once:

```js
function seedLastGeneratedIDOnce() {
  setLastGeneratedID('put-last-existing-id-here');
}
```

Then verify:

```js
function checkLastGeneratedID() {
  logLastGeneratedID();
}
```

Do this only once for a production counter initialization or a deliberate reset.

**Form Dropdown Updates**

`DropDown4Form.js` updates configured Google Form items from the Lists spreadsheet.

Use:

```js
updateForm()
```

To inspect form item IDs:

```js
logFormItems()
```

When adding a new form dropdown:

1. Add the source values to the Lists spreadsheet.
2. Find the Google Form item ID with `logFormItems()`.
3. Add the item ID and source column to the config/binding list.
4. Run `updateForm()`.

**Operational Notes**

- The Arrivals responses spreadsheet is the original raw data backup.
- Target spreadsheets are derived operational views for departments.
- The script uses `LockService` to reduce concurrency risks during form submissions.
- Processing blocks are isolated so that one target failure does not stop the entire flow.
- Email sending is wrapped safely and recorded in the processing error summary when it fails.
- Some target spreadsheets use data validation. Values must be normalized before writing.
- Some target spreadsheets use array formulas. Avoid writing into formula output cells before formulas expand.
- `config.local.js` is ignored by Git but required in Apps Script.
- `config.example.js` is tracked in Git but ignored by `clasp`.
