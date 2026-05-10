/**
 * Handles a Google Form submission and distributes the submitted arrival data
 * to all connected workforce, finance, legal, tech, coworking, and notification targets.
 *
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e Form submit event with namedValues.
 * @return {void}
 */
function onFormSubmit(e) {
  const submissionLock = LockService.getScriptLock();
  submissionLock.waitLock(30000);

  try {
    const values = flattenNamedValues(e.namedValues);
    const timestamp = values['Timestamp'];

    const [personID, newForceID] = getNextGeneratedIDs(2);

    let spreadsheet, sheet, lastRow;
    let rate;
    let countryCode = '';
    const processingErrors = [];
    
    /* ============================== HR DATA ============================== */
    try {
      // WorkForce.
      spreadsheet = SpreadsheetApp.openById(CONFIG.WORKFORCE_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('Registry');
      removeFiltersIfAny(sheet);
      lastRow = emptyRowsDel(sheet, 1);

      setRowValues(sheet, lastRow, {
        1: timestamp,
        2: personID,
        4: values['First Name'],
        5: values['Last Name'],
        6: values['Start Date'],
        16: values['Telephone'],
        17: values['2nd Telephone'],
        18: values['Telegram'],
        19: values['WhatsApp'],
        20: values['Private Skype'],
        21: values['Private Email'],
        22: values['LinkedIn'],
        23: values['Full Name by Local Passport'],
        24: values['Gender'],
        25: values['Birthday'],
        26: '',
        27: values['Emergency Contact'],
        28: values['Present Location'],
        33: values['Official Address'],
        34: values['Coworking'],
        35: values['Workplace Now'],
        36: values['Workplace After'],
        37: values['Office Attendance After'],
        38: values['Tech Stacks'],
        39: values['Insurance'],
        40: values['Legal Processing'],
        41: values['PE Support'],
        42: values['CV'],
        43: values['Recruiter'],
        44: values['Researcher'],
        49: timestamp,
        50: values['Email Address']
      });
      ensureCheckboxInRow(sheet, lastRow, 3);

      format4LastRow(sheet, lastRow);

    } catch (err) {
      recordProcessingError(processingErrors, "WorkForce", err);
    }

    try {
      // Forces-Projects.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_PROJECTS_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('AllData');
      removeFiltersIfAny(sheet);
      lastRow = emptyRowsDel(sheet, 1);

      setRowValues(sheet, lastRow, {
        1: personID,
        2: newForceID,
        6: values['Project'],
        7: values['Team'],
        8: values['Role'],
        9: values['Start Date'],
        11: values['Capacity'],
        12: values['Seniority'],
        13: values['Project Email'],
        14: values['Reporting Manager'],
        15: values['Rep Manager\'s Email'],
        16: values['Equipment'],
        17: values['Schedule Specificity']
      });
      sheet.getRange(lastRow, 3).clear();
      format4LastRow(sheet, lastRow);

    } catch (err) {
      recordProcessingError(processingErrors, "Forces-Projects", err);
    }

    try {
      // Forces-Personals.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_PERSONALS_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('AllData');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {
        1: personID,
        10: values['Comments Personality'],
        11: values['Comments Compensation'],
        12: values['Comments Technical'],
        13: values['Offer']
      });
      sheet.getRange(lastRow, 2).clear();

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-Personals", err);
    }

    try {
      // Forces-Locations.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_LOCATIONS_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('AllData');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      const countryName = removeSpaces(values['Residence Country']);
      switch (countryName) {
        case "Ukraine":
          countryCode = "UA";
          break;
        case "Poland":
          countryCode = "PL";
          break;
        default:
          countryCode = "UA";
      }
      setRowValues(sheet, lastRow, {
        1: personID,
        7: values['Start Date'],
        8: countryName,
        9: removeSpaces(values['Residence City']),
        10: removeSpaces(values['Residence Address']),
        12: "initial",
        13: countryCode
      });
      sheet.getRange(lastRow, 2).clear();

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-Locations", err);
    }  

    try {
      // Forces-CorpContacts.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_CORP_CONTACTS_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('AllData');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {
        1: personID,
        11: values['Photo'],
        13: values['Corp Email']
      });
      sheet.getRange(lastRow, 2).clear();

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-CorpContacts", err);
    }

    try {
      // Insurances.
      spreadsheet = SpreadsheetApp.openById(CONFIG.INSURANCES_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('Registry');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {1: personID});

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Insurances", err);
    }

    // Leave Tracker.
    try {
      spreadsheet = SpreadsheetApp.openById(CONFIG.LEAVE_TRACKER_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('Balance');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {
        1: personID,
        2: newForceID,
        13: 0,
        33: 0
      });

    } catch (err) {
      recordProcessingError(processingErrors, "Leave Tracker", err);
    }

    try {
      // Forces-Satisfactions.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_SATISFACTIONS_SPREADSHEET_ID);

      if (!getGeneralSatisfactionExcludedProjects().includes(values['Project'])) {
        sheet = spreadsheet.getSheetByName('General');
      } else {
        sheet = spreadsheet.getSheetByName(values['Project']);
      }
      removeFiltersIfAny(sheet);
      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {1: personID});

      format4LastRow(sheet, lastRow);

      sheet = spreadsheet.getSheetByName('Registry');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {1: personID});

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-Satisfactions", err);
    }  

    try {
      // Forces-BackgroundChecks.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_BACKGROUND_CHECKS_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('Records');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {1: personID});

      format4LastRow(sheet, lastRow); 
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-BackgroundChecks", err);
    }  

    try {
      // Telegram bot database.
      spreadsheet = SpreadsheetApp.openById(CONFIG.TELEGRAM_BOT_DATABASE_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('chatsDB');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {1: newForceID});

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Telegram bot database", err);
    } 

    /* ========================== FINANCE AND LEGAL DATA ========================== */

    try {
      // Forces-Legal.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_LEGAL_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('AllData');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {
        1: personID,
        24: values['Legal Processing'],
        25: values['PE Support'],
        26: values['Individual Tax #'],
        27: values['Individual Tax Scan'],
        28: values['PE/JDG Status'],
        29: values['PE/JDG Start Date'],
        30: values['Bank USD Account #'],
        31: values['Beneficiary'],
        32: values['Beneficiary\’s Bank'],
        33: values['SWIFT Code'],
        34: values['Bank Account Certificate'],
        35: values['Local Passport #'],
        36: values['Local Passport Scan'],
        37: values['International Passport #'],
        38: values['Int Passport Scan'],
        39: values['Int Passport Validity']
      });
      
      sheet.getRange(lastRow, 2).clear();

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-Legal", err);
    } 

    try {
      // Forces-Finance.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_FINANCE_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('AllData');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {
        1: personID,
        8: values['Legal Processing'],
        9: values['Currency'],
        10: values['NET'],
        11: values['GROSS'],
        16: values['Tax Rate'],
        17: values['Final Cost'],
        18: values['Insurance'],
        19: values['Hiring Commission'],
        20: values['Hiring Commission Currency'],
        21: values['Recommender'],
        22: values['Sign-in Bonus'],
        23: values['Equipment']
      });
      sheet.getRange(lastRow, 2).clear();

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-Finance", err);
    }

    if (countryCode === 'PL') {
      try {
        // Finance Basement | PL.
        spreadsheet = SpreadsheetApp.openById(CONFIG.FINANCE_BASEMENT_PL_SPREADSHEET_ID);
        sheet = spreadsheet.getSheetByName('Finance');
        removeFiltersIfAny(sheet);

        lastRow = emptyRowsDel(sheet, 1);
        setRowValues(sheet, lastRow, {
          1: personID,
          2: `${values['First Name'] || ''} ${values['Last Name'] || ''}`.trim(),
          3: 'new',
          4: values['Role'],
          6: values['Start Date'],
          14: values['GROSS'],
          24: 0,
          27: 0
        });

        format4LastRow(sheet, lastRow);
      } catch (err) {
        recordProcessingError(processingErrors, "Finance Basement | PL", err);
      }
    }

    try {
      // Agreements.
      spreadsheet = SpreadsheetApp.openById(CONFIG.AGREEMENTS_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('Contracts');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      const startDate = values['Start Date'] ? new Date(values['Start Date']) : '';
      const grossDivisors = {
        'PL: JDG': 145,
        'UA: FOP': 168
      };
      const monthlyHours = grossDivisors[values['Legal Processing']];
      const gross   = Number(values['GROSS']);
      rate          = (isFinite(gross) && monthlyHours)
                      ? Math.ceil((gross / monthlyHours) * 100) / 100
                      : 'n/a';
      setRowValues(sheet, lastRow, {
        1: personID,
        2: 'in progress',
        3: `${values['First Name'] || ''} ${values['Last Name'] || ''}`.trim(),
        4: values['Project'] || '',
        5: values['Role'] || '',
        6: 'Service Start',
        7: countryCode,
        11: startDate,
        18: values['GROSS'],
        19: rate,
        26: 'Currency',
        42: 'Full Name by Local Passport',
        54: 'Official Address',
        56: 'Individual Tax #',
        62: 'Private Email'
      });
   
      format4LastRow(sheet, lastRow);

    } catch (err) {
      recordProcessingError(processingErrors, "Agreements", err);
    }  

    /* ============================== TECH DATA ============================== */

    try {
      // Forces-Inventory.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_INVENTORY_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('Teams');
      removeFiltersIfAny(sheet);

      lastRow = emptyRowsDel(sheet, 1);
      setRowValues(sheet, lastRow, {1: newForceID});
      sheet.getRange(lastRow, 2).clear();

      format4LastRow(sheet, lastRow);
    } catch (err) {
      recordProcessingError(processingErrors, "Forces-Inventory", err);
    }

    /* ============================ COWORKING DATA ============================ */

    const coworkingIntegration = getCoworkingIntegration(values['Coworking']);

    if (coworkingIntegration) {
      
      try {
        // Coworking member register.
        spreadsheet = SpreadsheetApp.openById(coworkingIntegration.memberRegisterSpreadsheetId);
        sheet = spreadsheet.getSheetByName('List');
        removeFiltersIfAny(sheet);

        lastRow = emptyRowsDel(sheet, 1);
        setRowValues(sheet, lastRow, {1: personID});
        sheet.getRange(lastRow, 2).clear();

        format4LastRow(sheet, lastRow);
      } catch (err) {
        recordProcessingError(processingErrors, "Coworking member register", err);
      }

      try {
        // Resident Badge Register.
        spreadsheet = SpreadsheetApp.openById(coworkingIntegration.residentBadgeRegisterSpreadsheetId);
        sheet = spreadsheet.getSheetByName('List');
        removeFiltersIfAny(sheet);

        lastRow = emptyRowsDel(sheet, 5);
        setRowValues(sheet, lastRow, {
          1: personID,
          5: values['First Name'],
          6: String(values['Last Name'] || '').toUpperCase(),
          8: 'Pending'
        });

        format4LastRow(sheet, lastRow);
      } catch (err) {
        recordProcessingError(processingErrors, "Resident Badge Register", err);
      }

      try {
        // Coworker finance register.
        spreadsheet = SpreadsheetApp.openById(coworkingIntegration.financeSpreadsheetId);
        sheet = spreadsheet.getSheetByName('Coworkers');
        removeFiltersIfAny(sheet);

        // Find the row with "Total" in column C.
        const colCValues = sheet.getRange(1, 3, sheet.getLastRow()).getValues().flat();
        const totalRowIndex = colCValues.findIndex(v => String(v).trim().toLowerCase() === 'total') + 1;
        if (totalRowIndex === 0) throw new Error(`Row with 'Total' in column C not found`);

        // Insert the new coworker row before the total row.
        sheet.insertRowBefore(totalRowIndex);
        const newRow = totalRowIndex;

        // Prepare values from the form submission.
        const fullName = `${values['First Name'] || ''} ${values['Last Name'] || ''}`.trim();
        const taxRate  = Number(values['Tax Rate']);
        const divisor  = 145;
        const gross    = Number(values['GROSS']);
        const grossPer = (isFinite(gross) && divisor) ? Math.ceil((gross / divisor) * 100) / 100 : '';
        const startDt  = values['Start Date'] ? new Date(values['Start Date']) : '';

        setRowValues(sheet, newRow, {
          1: personID,
          3: fullName,
          7: isFinite(taxRate) ? taxRate : '',
          8: divisor,
          9: grossPer,
          16: startDt,
          18: values['PE Support'] || '',
          20: values['Residence Address'] || '',
          21: values['Individual Tax #'] || '',
          22: values['Private Email'] || ''
        });

        format4LastRow(sheet, newRow);

      } catch (err) {
        recordProcessingError(processingErrors, "Coworker finance register", err);
      }


      if (coworkingIntegration.photoMainFolderId && values['Photo'] && values['Photo'].includes("drive.google.com")) {
        Utilities.sleep(2000);
        try {
          const fileId = extractFileIdFromUrl(values['Photo']);
          const file = DriveApp.getFileById(fileId);

          const mainFolder = DriveApp.getFolderById(coworkingIntegration.photoMainFolderId);
          const targetFolder = getSubfolderByName(mainFolder, coworkingIntegration.photoTargetFolderName);

          if (targetFolder) {
            targetFolder.addFile(file);
            mainFolder.removeFile(file);
          } else {
            Logger.log(`Subfolder '${coworkingIntegration.photoTargetFolderName}' not found.`);
          }
        } catch (err) {
          recordProcessingError(processingErrors, "Coworking photo handling", err);
        }
      }
    }
    
    /* ============================= NOTIFICATIONS ============================= */

    const subject = "New arrivals to the team!";
    let baseHtmlBody = '<p>Hey there!\n\n</br>A new team member has just been registered and needs to be proceeded by you according to the standard flow:</p>';
    baseHtmlBody += '<ul><li>' + 'Name : '+ values['First Name'] + ' ' + values['Last Name'] + '</li>';
    baseHtmlBody += '<li>' + 'Role : '+ values['Role'] + ', ' + values['Project'] + '</li>'; 
    baseHtmlBody += '<li>' + 'Start date : '+ values['Start Date'] + '</li>'; 
    baseHtmlBody += '</ul>';
    baseHtmlBody += '<p>A new registration made by ' + values['Recruiter'] + '.</p>';

    // Notify HR.
    const recipientsCc = getNotificationEmail('ARRIVALS_CC', processingErrors, 'Notification CC');
    let recipientsTo = getNotificationEmail('HR', processingErrors, 'HR notification');

    let linksHtml = '<p>Here are the links to the registered data:\n' +
                `WorkForce: ${getSpreadsheetUrl(CONFIG.WORKFORCE_SPREADSHEET_ID)} \n` +
                `<br>Forces-Projects: ${getSpreadsheetUrl(CONFIG.FORCES_PROJECTS_SPREADSHEET_ID)}</p>`;

    sendEmailSafely(processingErrors, "HR notification", {to: recipientsTo, cc: recipientsCc, subject: subject, htmlBody: baseHtmlBody + linksHtml, name: "Auto-generated message", noReply: true});
    
    // Notify Legal.
    recipientsTo = getNotificationEmail('LEGAL', processingErrors, 'Legal notification');
    linksHtml = '<p>Here are the links to the registered data:\n' +
                `Forces-Legal: ${getSpreadsheetUrl(CONFIG.FORCES_LEGAL_SPREADSHEET_ID)} \n` +
                `<br>Forces-Finance: ${getSpreadsheetUrl(CONFIG.FORCES_FINANCE_SPREADSHEET_ID)}</p>`;
    sendEmailSafely(processingErrors, "Legal notification", {to: recipientsTo, cc: recipientsCc, subject: subject, htmlBody: baseHtmlBody + linksHtml, name: "Auto-generated message", noReply: true});

    // Notify FinController.
    recipientsTo = getNotificationEmail('FINANCE_CONTROLLER', processingErrors, 'FinController notification');
    linksHtml = '<p>Here is the link to the registered data:\n' +
                `Forces-Finance: ${getSpreadsheetUrl(CONFIG.FORCES_FINANCE_SPREADSHEET_ID)}</p>`;
    sendEmailSafely(processingErrors, "FinController notification", {to: recipientsTo, cc: recipientsCc, subject: subject, htmlBody: baseHtmlBody + linksHtml, name: "Auto-generated message", noReply: true});

    // Notify HelpDesk.
    recipientsTo = getNotificationEmail('HELPDESK', processingErrors, 'HelpDesk notification');
    linksHtml = '<p>Here is the link to the registered data:\n' +
                `Forces-Inventory&Movement: ${getSpreadsheetUrl(CONFIG.FORCES_INVENTORY_SPREADSHEET_ID)}</p>`;
    sendEmailSafely(processingErrors, "HelpDesk notification", {to: recipientsTo, cc: recipientsCc, subject: subject, htmlBody: baseHtmlBody + linksHtml, name: "Auto-generated message", noReply: true});

    if (values['Project'] === CONFIG.CLIENT_OPERATIONS_PROJECT_NAME) {
      const recipientsToClientOperations = getNotificationEmail('CLIENT_OPERATIONS', processingErrors, 'Client operations notification');
      sendEmailSafely(processingErrors, "Client operations notification", {to: recipientsToClientOperations, cc: recipientsCc, subject: subject, htmlBody: baseHtmlBody, name: "Auto-generated message", noReply: true});
    }

    if (coworkingIntegration) {
      const coworkingSubject = `New member of ${coworkingIntegration.displayName} Coworking!`;
      let coworkingHtmlStr = '<p>Hey there!\n\n </br>A new coworking member has just been registered and needs to be proceeded by you according to the standard flow:</p>';
      coworkingHtmlStr += '<ul><li>' + 'Name : '+ values['First Name'] + ' ' + values['Last Name'] + '</li>';
      coworkingHtmlStr += '<li>' + 'Start date : '+ values['Start Date'] + '</li>'; 
      coworkingHtmlStr += '</ul>';
      linksHtml = '<p>Here are the links to the registered data:\n' +
                `Resident Access Log: ${getSpreadsheetUrl(coworkingIntegration.residentBadgeRegisterSpreadsheetId)} \n` +
                `<br>Coworking Member Registry: ${getSpreadsheetUrl(coworkingIntegration.memberRegisterSpreadsheetId)}</p>`;
      
      const coworkingNotificationName = `${coworkingIntegration.displayName} coworking notification`;
      const recipientsToCoworking = getNotificationEmail(coworkingIntegration.notificationEmailKey, processingErrors, coworkingNotificationName);
      sendEmailSafely(processingErrors, coworkingNotificationName, {to: recipientsToCoworking, cc: recipientsCc, subject: coworkingSubject, htmlBody: coworkingHtmlStr + linksHtml, name: "Auto-generated message", noReply: true});
    }

    sendProcessingErrorSummary(processingErrors, values, personID, newForceID);
    return;
  } finally {
    submissionLock.releaseLock();
  }
}

/**
 * Removes an active filter from a sheet and recreates a clean filter on the data range.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination sheet.
 * @return {void}
 */
function removeFiltersIfAny(sheet) {

  const filterRange = sheet.getFilter();
  
  if (filterRange) {
    filterRange.remove();
    sheet.getDataRange().createFilter();
  }
}

/**
 * Normalizes whitespace in a value by collapsing repeated spaces and trimming edges.
 *
 * @param {*} value Value to normalize.
 * @return {string} Normalized string value.
 */
function removeSpaces(value) {
  return String(value || '').replace(/\s+/g, " ").trim();
}

/**
 * Builds a Google Spreadsheet URL from a configured spreadsheet ID.
 *
 * @param {string} spreadsheetId Google Spreadsheet ID.
 * @return {string} Spreadsheet URL.
 */
function getSpreadsheetUrl(spreadsheetId) {
  return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/`;
}

/**
 * Gets project names that should write satisfaction data to the General sheet.
 *
 * @return {string[]} Project names excluded from direct satisfaction sheet lookup.
 */
function getGeneralSatisfactionExcludedProjects() {
  return (CONFIG.GENERAL_SATISFACTION_EXCLUDED_PROJECTS || []).concat(
    CONFIG.CLIENT_OPERATIONS_PROJECT_NAME ? [CONFIG.CLIENT_OPERATIONS_PROJECT_NAME] : []
  );
}

/**
 * Gets coworking integration settings by the form dropdown value.
 *
 * @param {string} formValue Coworking value submitted from the form.
 * @return {?Object} Matching coworking integration configuration, or null when not configured.
 */
function getCoworkingIntegration(formValue) {
  const integrations = CONFIG.COWORKING_INTEGRATIONS || [];
  return integrations.find(integration => integration.formValue === formValue) || null;
}

/**
 * Gets a notification email address from the local configuration.
 *
 * @param {string} emailKey Key from CONFIG.NOTIFICATION_EMAILS.
 * @param {Object[]} processingErrors Error collection for the current form submission.
 * @param {string} targetName Human-readable notification name.
 * @return {string} Configured email address or an empty string when missing.
 */
function getNotificationEmail(emailKey, processingErrors, targetName) {
  const notificationEmails = CONFIG.NOTIFICATION_EMAILS || {};
  const email = notificationEmails[emailKey];

  if (email) return email;

  recordProcessingError(
    processingErrors,
    `${targetName} configuration`,
    new Error(`CONFIG.NOTIFICATION_EMAILS.${emailKey} is not configured.`)
  );
  return '';
}

/**
 * Stores and logs an error from one processing block.
 *
 * @param {Object[]} processingErrors Error collection for the current form submission.
 * @param {string} targetName Human-readable target name.
 * @param {Error} err Error thrown by the processing block.
 * @return {void}
 */
function recordProcessingError(processingErrors, targetName, err) {
  const message = err && err.message ? err.message : String(err);
  processingErrors.push({
    targetName: targetName,
    message: message
  });
  Logger.log(`❌ ${targetName} error: ${message}`);
}

/**
 * Sends an email and records the error instead of stopping the full submission flow.
 *
 * @param {Object[]} processingErrors Error collection for the current form submission.
 * @param {string} targetName Human-readable notification name.
 * @param {Object} emailOptions MailApp.sendEmail options.
 * @return {void}
 */
function sendEmailSafely(processingErrors, targetName, emailOptions) {
  if (!emailOptions.to) {
    recordProcessingError(processingErrors, targetName, new Error('Email recipient is not configured.'));
    return;
  }

  try {
    MailApp.sendEmail(emailOptions);
  } catch (err) {
    recordProcessingError(processingErrors, targetName, err);
  }
}

/**
 * Sends a summary email when one or more processing blocks fail.
 *
 * @param {Object[]} processingErrors Error collection for the current form submission.
 * @param {Object.<string, string>} values Plain form values keyed by question title.
 * @param {string} personID Generated person ID.
 * @param {string} forceID Generated force ID.
 * @return {void}
 */
function sendProcessingErrorSummary(processingErrors, values, personID, forceID) {
  if (!processingErrors.length) return;
  if (!CONFIG.ERROR_NOTIFICATION_EMAIL) {
    Logger.log('❌ Error summary notification email is not configured.');
    return;
  }

  const fullName = `${values['First Name'] || ''} ${values['Last Name'] || ''}`.trim();
  const htmlBody = [
    '<p>One or more arrival processing blocks failed.</p>',
    '<ul>',
    `<li>Name: ${fullName || 'n/a'}</li>`,
    `<li>Project: ${values['Project'] || 'n/a'}</li>`,
    `<li>Person ID: ${personID}</li>`,
    `<li>Force ID: ${forceID}</li>`,
    '</ul>',
    '<p>Failed blocks:</p>',
    '<ul>',
    processingErrors.map(error => `<li>${error.targetName}: ${error.message}</li>`).join(''),
    '</ul>'
  ].join('');

  try {
    MailApp.sendEmail({
      to: CONFIG.ERROR_NOTIFICATION_EMAIL,
      subject: `Arrival processing errors: ${fullName || personID}`,
      htmlBody: htmlBody,
      name: "Auto-generated message",
      noReply: true
    });
  } catch (err) {
    Logger.log(`❌ Error summary notification failed: ${err.message}`);
  }
}

/**
 * Checks access to the first configured coworker finance spreadsheet and writes the result to the log.
 * Parameters: none.
 *
 * @return {void}
 */
function testAccessToSpreadsheet() {
  const integration = (CONFIG.COWORKING_INTEGRATIONS || [])[0];
  const id = integration && integration.financeSpreadsheetId;

  if (!id) {
    Logger.log("Coworker finance spreadsheet is not configured.");
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(id);
    const sheet = ss.getSheets()[0];
    Logger.log("Access granted. Sheet name: " + sheet.getName());
  } catch (e) {
    Logger.log("❌ Access denied or error: " + e.message);
  }
}

/**
 * Extracts a Google Drive file or folder ID from a supported Drive URL.
 *
 * @param {string} url Google Drive URL.
 * @return {?string} Extracted file or folder ID, or null when no ID is found.
 */
function extractFileIdFromUrl(url) {
  if (!url) return null;

  const patterns = [
    /\/d\/([a-zA-Z0-9_-]{25,})/,
    /id=([a-zA-Z0-9_-]{25,})/,
    /\/folders\/([a-zA-Z0-9_-]{25,})/,
    /\/uc\?export=download&id=([a-zA-Z0-9_-]{25,})/
  ];

  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) return match[1];
  }

  return null;
}

/**
 * Finds a direct child folder by name.
 *
 * @param {GoogleAppsScript.Drive.Folder} parentFolder Folder where the search should run.
 * @param {string} folderName Name of the child folder to find.
 * @return {?GoogleAppsScript.Drive.Folder} Matching child folder, or null when it is not found.
 */
function getSubfolderByName(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : null;
}

/**
 * Converts form namedValues from arrays to a plain object with string values.
 *
 * @param {Object.<string, string[]>} namedValues Google Form namedValues payload.
 * @return {Object.<string, string>} Plain object keyed by form question title.
 */
function flattenNamedValues(namedValues) {
  const result = {};
  for (const key in namedValues) {
    result[key] = namedValues[key]?.[0] || '';
  }
  return result;
}

/**
 * Writes sparse row values with the minimum number of setValues() calls.
 * Example: {1: id, 2: forceId, 6: project} writes A:B and F separately.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination sheet.
 * @param {number} row Destination row number.
 * @param {Object.<number, *>} columnValues Object keyed by 1-based column numbers.
 * @return {void}
 */
function setRowValues(sheet, row, columnValues) {
  const columns = Object.keys(columnValues)
    .map(Number)
    .filter(column => isFinite(column))
    .sort((a, b) => a - b);

  if (!columns.length) return;

  let startColumn = columns[0];
  let values = [columnValues[startColumn]];
  let previousColumn = startColumn;

  for (let i = 1; i < columns.length; i++) {
    const column = columns[i];
    if (column === previousColumn + 1) {
      values.push(columnValues[column]);
    } else {
      sheet.getRange(row, startColumn, 1, values.length).setValues([values]);
      startColumn = column;
      values = [columnValues[column]];
    }
    previousColumn = column;
  }

  sheet.getRange(row, startColumn, 1, values.length).setValues([values]);
}

/**
 * Applies checkbox data validation to a cell in the given row and column.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination sheet.
 * @param {number} row Destination row number.
 * @param {number} column Destination column number.
 * @return {void}
 */
function ensureCheckboxInRow(sheet, row, column) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(true)
    .build();
  sheet.getRange(row, column).setDataValidation(rule);
}
