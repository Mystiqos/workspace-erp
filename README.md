# Workspace ERP

Custom ERP automation built on top of Google Workspace: Google Sheets, Google Forms, Google Drive, Gmail, and Google Apps Script.

The project is organized as a set of independent Apps Script modules. Each module is connected to its own Google Workspace resources and can be deployed separately with `clasp`.

**Modules**

`Arrivals`  
Handles the onboarding arrival flow. It processes Google Form submissions, writes structured data into connected spreadsheets, generates internal IDs, updates form dropdowns, and sends workflow notifications.

More modules can be added as separate folders under the repository root after they are cleaned and moved to the same configuration pattern.

**Repository Structure**

```text
.
├── Arrivals/
│   ├── Code.js
│   ├── DropDown4Form.js
│   ├── idsGeneration.js
│   ├── manipulations.js
│   ├── config.example.js
│   └── appsscript.json
├── .gitignore
└── README.md
```

**Local Configuration**

Real spreadsheet IDs, form IDs, folder IDs, script IDs, email addresses, and brand/client names must not be committed.

Each module should keep sensitive values in local ignored files:

```text
config.local.js
.clasp.json
.clasp.test.json
.clasp.prod.json
```

Use `config.example.js` as a template for the required configuration shape.

Example:

```js
const CONFIG = {
  ARRIVALS_FORM_ID: 'put-form-id-here',
  LISTS_SPREADSHEET_ID: 'put-spreadsheet-id-here',
  ERROR_NOTIFICATION_EMAIL: 'put-email-here'
};
```

**Security Rules**

- Do not hardcode spreadsheet IDs, form IDs, folder IDs, script IDs, email addresses, client names, or business brand names in tracked code.
- Put real values only in ignored local config files.
- Keep `.clasp*.json` ignored because it contains Apps Script project bindings.
- Before pushing, check for accidental sensitive values.

Useful checks:

```bash
rg -n "@" .
rg -n "scriptId|spreadsheetId|formId|folderId" .
git status --short
```

**Working With Clasp**

Run `clasp` commands from the module folder, not from the repository root.

For example:

```bash
cd Arrivals
clasp status
clasp push
```

To switch between test and production Apps Script projects, copy the needed local clasp file:

```bash
cp .clasp.test.json .clasp.json
clasp push
```

or:

```bash
cp .clasp.prod.json .clasp.json
clasp push
```

For safer testing, use a test Apps Script project first. If the local config points to production spreadsheets, test submissions may still write test rows into production data.

**Development Flow**

Recommended flow for larger changes:

```bash
git switch main
git pull
git switch -c optimize-arrivals
```

Then:

1. Make and review changes locally.
2. Push to a test Apps Script project with `clasp`.
3. Run a small number of controlled test submissions.
4. Check all affected spreadsheets and notifications.
5. Merge back to `main` only after validation.
6. Push the production Apps Script project with `clasp`.

**Arrivals Notes**

The `Arrivals` module currently includes:

- form submission processing via `onFormSubmit(e)`;
- centralized config access through `CONFIG`;
- script locking for safer concurrent submissions;
- generated internal IDs;
- grouped writes into HR, project, finance, legal, tech, and coworking-related targets;
- safe notification sending with error collection;
- form dropdown updates from a source spreadsheet;
- local-only configuration for integrations and recipients.

**Git Hygiene**

Before committing:

```bash
git status --short
git diff
```

Only commit files that are safe for GitHub. Local configs and clasp bindings should remain ignored.
