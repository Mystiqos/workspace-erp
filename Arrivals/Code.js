const HEADER_PROFILES = {
  WORKFORCE: ['Person ID', 'First Name', 'Last Name', 'Start Date'],
  FORCES_PROJECTS: ['Person ID', 'Force ID', 'Project', 'Role'],
  FINANCE_BASEMENT_PL: ['Person ID', 'Role', 'Project Start Date', 'Offered'],
  AGREEMENTS: ['Person ID', 'Status', 'Project', 'Document Currency'],
  COWORKING_MEMBER_REGISTER: ['Person ID', 'Member Name']
};

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
      const workforceHeaderProfile = HEADER_PROFILES.WORKFORCE;
      lastRow = emptyRowsDel(sheet, 1, {headerProfile: workforceHeaderProfile});

      const workforceValues = {};
      const setWorkforceValue = createHeaderValueSetter(sheet, workforceValues, workforceHeaderProfile);
      setWorkforceValue('Timestamp', timestamp, 1);
      setWorkforceValue('Person ID', personID, 2);
      setWorkforceValue('First Name', values['First Name'], 4);
      setWorkforceValue('Last Name', values['Last Name'], 5);
      setWorkforceValue('Start Date', values['Start Date'], 6);
      setWorkforceValue('Telephone', values['Telephone'], 16);
      setWorkforceValue('2nd Telephone', values['2nd Telephone'], 17);
      setWorkforceValue('Telegram', values['Telegram'], 18);
      setWorkforceValue('WhatsApp', values['WhatsApp'], 19);
      setWorkforceValue('Private Skype', values['Private Skype'], 20);
      setWorkforceValue('Private Email', values['Private Email'], 21);
      setWorkforceValue('LinkedIn', values['LinkedIn'], 22);
      setWorkforceValue('Full Name by Local Passport', values['Full Name by Local Passport'], 23);
      setWorkforceValue('Gender', values['Gender'], 24);
      setWorkforceValue('Birthday', values['Birthday'], 25);
      setWorkforceValue('Emergency Contact', values['Emergency Contact'], 27);
      setWorkforceValue('Start Location', values['Present Location'], 28);
      setWorkforceValue('Official Address', values['Official Address'], 33);
      setWorkforceValue('Coworking', values['Coworking'], 34);
      setWorkforceValue('Workplace Now', values['Workplace Now'], 35);
      setWorkforceValue('Workplace After', values['Workplace After'], 36);
      setWorkforceValue('Office Attendance After', values['Office Attendance After'], 37);
      setWorkforceValue('Tech Stacks', values['Tech Stacks'], 38);
      setWorkforceValue('Insurance', values['Insurance'], 39);
      setWorkforceValue('Legal Processing', values['Legal Processing'], 40);
      setWorkforceValue('PE Support', values['PE Support'], 41);
      setWorkforceValue('CV', values['CV'], 42);
      setWorkforceValue('Recruiter', values['Recruiter'], 43);
      setWorkforceValue('Researcher', values['Researcher'], 44);
      setWorkforceValue('Last Modified', timestamp, 49);
      setWorkforceValue('Modified By', values['Email Address'], 50);
      setRowValues(sheet, lastRow, workforceValues);
      format4LastRow(sheet, lastRow);
      ensureCheckboxAfterFormulaUpdate(sheet, lastRow, 3);

    } catch (err) {
      recordProcessingError(processingErrors, "WorkForce", err);
    }

    try {
      // Forces-Projects.
      spreadsheet = SpreadsheetApp.openById(CONFIG.FORCES_PROJECTS_SPREADSHEET_ID);
      sheet = spreadsheet.getSheetByName('AllData');
      removeFiltersIfAny(sheet);
      const forceProjectsHeaderProfile = HEADER_PROFILES.FORCES_PROJECTS;
      lastRow = emptyRowsDel(sheet, 1, {headerProfile: forceProjectsHeaderProfile});

      const forceProjectValues = {};
      const setForceProjectValue = createHeaderValueSetter(sheet, forceProjectValues, forceProjectsHeaderProfile);
      setForceProjectValue('Person ID', personID, 1);
      setForceProjectValue('Force ID', newForceID, 2);
      setForceProjectValue('Project', values['Project'], 6);
      setForceProjectValue('Team', values['Team'], 7);
      setForceProjectValue('Role', values['Role'], 8);
      setForceProjectValue('Start Date', values['Start Date'], 9);
      setForceProjectValue('Workload', values['Capacity'], 11);
      setForceProjectValue('Seniority', values['Seniority'], 12);
      setForceProjectValue('Project Mail', values['Project Email'], 13);
      setForceProjectValue('Reporting Manager', values['Reporting Manager'], 14);
      setForceProjectValue('Rep Manager\'s Email', values['Rep Manager\'s Email'], 15);
      setForceProjectValue('Equipment', values['Equipment'], 16);
      setForceProjectValue('Schedule Specifics', values['Schedule Specifics'], 17);
      setRowValues(sheet, lastRow, forceProjectValues);
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

    const projectEndDate = getProjectEndDate(values['Start Date'], values['Project Term']);
    if (!projectEndDate) {
      recordProcessingError(processingErrors, "Project end date", new Error("Project Term is missing or invalid."));
    }

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

        const financeBasementHeaderProfile = HEADER_PROFILES.FINANCE_BASEMENT_PL;
        lastRow = emptyRowsDel(sheet, 1, {headerProfile: financeBasementHeaderProfile});
        const financeBasementValues = {};
        const setFinanceBasementValue = createHeaderValueSetter(sheet, financeBasementValues, financeBasementHeaderProfile);
        setFinanceBasementValue('Person ID', personID, 1);
        setFinanceBasementValue('Name', `${values['First Name'] || ''} ${values['Last Name'] || ''}`.trim(), 2);
        setFinanceBasementValue('Agreement', 'new', 3);
        setFinanceBasementValue('Role', getFinanceBasementRole(values['Role']), 4);
        setFinanceBasementValue('Project Start Date', values['Start Date'], 6);
        setFinanceBasementValue('Offered', values['GROSS'], 14);
        setFinanceBasementValue('Project End Date', projectEndDate || '', 0);
        setFinanceBasementValue('Carried-over', 0, 24);
        setFinanceBasementValue('Used Before', 0, 27);

        setRowValues(sheet, lastRow, financeBasementValues);

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

      const agreementsHeaderProfile = HEADER_PROFILES.AGREEMENTS;
      lastRow = emptyRowsDel(sheet, 1, {headerProfile: agreementsHeaderProfile});
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
      const agreementCurrency = getValidCurrency(values['Currency']);
      if (!agreementCurrency) {
        recordProcessingError(processingErrors, "Agreements currency", new Error("Currency is missing or invalid."));
      }

      const agreementValues = {};
      const setAgreementValue = createHeaderValueSetter(sheet, agreementValues, agreementsHeaderProfile);
      setAgreementValue('Person ID', personID, 1);
      setAgreementValue('Status', 'in progress', 2);
      setAgreementValue('Contractor Name', `${values['First Name'] || ''} ${values['Last Name'] || ''}`.trim(), 3);
      setAgreementValue('Project', values['Project'] || '', 4);
      setAgreementValue('Role', values['Role'] || '', 5);
      setAgreementValue('Purpose', 'Service Start', 6);
      setAgreementValue('Country Code', countryCode, 7);
      setAgreementValue('Effective DT', startDate, 11);
      setAgreementValue('Compensation Monthly Agreed', values['GROSS'], 18);
      setAgreementValue('Rate Offered', rate, 19);
      setAgreementValue('Contractor Legal Name Local', values['Legal Name']);
      setAgreementValue('Contractor Address Local', values['Official Address']);
      setAgreementValue('Contractor Tax ID (NIP/INN/PAN)', values['Individual Tax #']);
      setAgreementValue('Contractor REGON', values['Regon']);
      setAgreementValue('Contractor SWIFT', values['SWIFT Code']);
      setAgreementValue('Contractor IBAN', values['Bank USD Account #']);
      setAgreementValue('Contractor Email', values['Private Email']);

      if (agreementCurrency) {
        setAgreementValue('Document Currency', agreementCurrency, 26);
      }
      setAgreementValue('Expiry DT', projectEndDate || '', 0);

      setRowValues(sheet, lastRow, agreementValues);
   
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

        const memberHeaderProfile = HEADER_PROFILES.COWORKING_MEMBER_REGISTER;
        lastRow = emptyRowsDel(sheet, 1, {headerProfile: memberHeaderProfile});
        const memberValues = {};
        const setMemberValue = createHeaderValueSetter(sheet, memberValues, memberHeaderProfile);
        setMemberValue('Person ID', personID, 1);
        setMemberValue('Member Name', values['Legal Name'] || values['Full Name by Local Passport']);
        setMemberValue('Member Address', values['Official Address']);
        setMemberValue('Member PESEL / NIP', values['Individual Tax #']);

        setRowValues(sheet, lastRow, memberValues);
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

        const residentNameColumn = getColumnByHeader(sheet, 'RESIDENT NAME', 5);
        lastRow = findFirstEmptyRowInColumn(sheet, residentNameColumn);
        setRowValues(sheet, lastRow, {
          1: personID,
          [residentNameColumn]: values['First Name'],
          [residentNameColumn + 1]: String(values['Last Name'] || '').toUpperCase(),
          8: 'Pending'
        });
      } catch (err) {
        recordProcessingError(processingErrors, "Resident Badge Register", err);
      }

      try {
        // Coworker finance register.
        spreadsheet = SpreadsheetApp.openById(coworkingIntegration.financeSpreadsheetId);
        sheet = spreadsheet.getSheetByName('Coworkers');
        removeFiltersIfAny(sheet);

        // Find the row with "Total" in columns A:F.
        const totalRowIndex = findRowContainingText(sheet, 'Total', 1, 6);
        if (totalRowIndex === 0) throw new Error(`Row with 'Total' in columns A:F not found`);

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
 * Gets a single valid currency code from a form value.
 *
 * @param {string} value Currency value submitted from the form.
 * @return {string} Valid currency code, or an empty string when missing or invalid.
 */
function getValidCurrency(value) {
  const allowedCurrencies = ['USD', 'EUR', 'PLN', 'UAH'];
  const currency = String(value || '').trim().toUpperCase();
  return allowedCurrencies.includes(currency) ? currency : '';
}

/**
 * Gets a valid role code for the Finance Basement sheet.
 *
 * @param {string} role Role submitted from the form.
 * @return {string} Valid finance role code.
 */
function getFinanceBasementRole(role) {
  const allowedRoles = CONFIG.FINANCE_BASEMENT_PL_ALLOWED_ROLES || ['QA', 'AQA', 'DEV', 'BETA', 'OPR', 'ENG', 'EMP'];
  const defaultRole = CONFIG.FINANCE_BASEMENT_PL_DEFAULT_ROLE || 'EMP';
  const normalizedRole = String(role || '').trim().toUpperCase();

  return allowedRoles.includes(normalizedRole) ? normalizedRole : defaultRole;
}

/**
 * Calculates the project end date from a start date and project term.
 *
 * @param {string|Date} startDateValue Project start date.
 * @param {string} projectTerm Project term submitted from the form.
 * @return {?Date} Project end date, or null when inputs are invalid.
 */
function getProjectEndDate(startDateValue, projectTerm) {
  const startDate = parseDateValue(startDateValue);
  if (!startDate) return null;

  const normalizedTerm = String(projectTerm || '').trim().toLowerCase();
  if (normalizedTerm === 'until year-end') {
    return new Date(startDate.getFullYear(), 11, 31);
  }

  if (normalizedTerm === '1 year') {
    if (startDate.getDate() === 1) {
      return addMonthsAndSubtractOneDay(startDate, 12);
    }

    const annualEndMonthDate = addMonthsClamped(startDate, 12);
    if (isWithinDaysBeforeMonthEnd(startDate, 7)) {
      return getLastDayOfNextMonth(annualEndMonthDate);
    }

    return getLastDayOfMonth(annualEndMonthDate);
  }

  const monthTerms = {
    '1 month': 1,
    '2 months': 2,
    '3 months': 3,
    '6 months': 6,
    '9 months': 9,
    '2 years': 24
  };
  const months = monthTerms[normalizedTerm];
  return months ? addMonthsAndSubtractOneDay(startDate, months) : null;
}

/**
 * Parses a form date value.
 *
 * @param {string|Date} value Date value.
 * @return {?Date} Parsed date, or null when invalid.
 */
function parseDateValue(value) {
  if (!value) return null;
  const date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  return isNaN(date.getTime()) ? null : clearTime(date);
}

/**
 * Adds months to a date and subtracts one day from the result.
 *
 * @param {Date} date Source date.
 * @param {number} months Number of months to add.
 * @return {Date} Calculated end date.
 */
function addMonthsAndSubtractOneDay(date, months) {
  const result = addMonthsClamped(date, months);
  result.setDate(result.getDate() - 1);
  return result;
}

/**
 * Adds months while clamping the day to the target month's last day.
 *
 * @param {Date} date Source date.
 * @param {number} months Number of months to add.
 * @return {Date} Date in the target month.
 */
function addMonthsClamped(date, months) {
  const year = date.getFullYear();
  const month = date.getMonth() + months;
  const day = date.getDate();
  const lastTargetDay = new Date(year, month + 1, 0).getDate();
  return new Date(year, month, Math.min(day, lastTargetDay));
}

/**
 * Gets the last day of the month for a date.
 *
 * @param {Date} date Date inside the target month.
 * @return {Date} Last day of the date's month.
 */
function getLastDayOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 1, 0);
}

/**
 * Gets the last day of the month after the date's month.
 *
 * @param {Date} date Date inside the reference month.
 * @return {Date} Last day of the next month.
 */
function getLastDayOfNextMonth(date) {
  return new Date(date.getFullYear(), date.getMonth() + 2, 0);
}

/**
 * Checks whether a date is within the given number of days before month end, inclusive.
 *
 * @param {Date} date Date to check.
 * @param {number} days Number of final month days to include.
 * @return {boolean} True when the date is close to month end.
 */
function isWithinDaysBeforeMonthEnd(date, days) {
  const lastDay = getLastDayOfMonth(date).getDate();
  return lastDay - date.getDate() + 1 <= days;
}

/**
 * Clears time fields from a date.
 *
 * @param {Date} date Date to normalize.
 * @return {Date} Date at local midnight.
 */
function clearTime(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Finds the first row that contains the given text in a column range.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to search.
 * @param {string} text Text to find.
 * @param {number} startColumn One-based first column in the search range.
 * @param {number} endColumn One-based last column in the search range.
 * @return {number} Matching one-based row number, or 0 when not found.
 */
function findRowContainingText(sheet, text, startColumn, endColumn) {
  const lastRow = sheet.getLastRow();
  if (!lastRow) return 0;

  const columnCount = endColumn - startColumn + 1;
  const searchText = String(text).trim().toLowerCase();
  const values = sheet.getRange(1, startColumn, lastRow, columnCount).getValues();

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const hasMatch = values[rowIndex].some(value => String(value).trim().toLowerCase() === searchText);
    if (hasMatch) return rowIndex + 1;
  }

  return 0;
}

/**
 * Finds a column by header text in the detected header row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to inspect.
 * @param {string} headerText Header text to find.
 * @param {number} fallbackColumn One-based column returned when the header is not found.
 * @param {string[]=} headerProfile Required headers used to detect the correct header row.
 * @return {number} Matching one-based column number.
 */
function getColumnByHeader(sheet, headerText, fallbackColumn, headerProfile) {
  const headerInfo = getHeaderInfo(sheet, headerProfile);
  return getColumnFromHeaderInfo(sheet, headerInfo, headerText, fallbackColumn);
}

/**
 * Finds the most likely header row in the first top rows.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to inspect.
 * @param {string[]=} headerProfile Required headers used to detect the correct header row.
 * @return {number} One-based header row number, or 0 when not found.
 */
function findHeaderRow(sheet, headerProfile) {
  const rowsToScan = Math.min(10, sheet.getLastRow());
  const lastColumn = sheet.getLastColumn();
  if (!rowsToScan || !lastColumn) return 0;

  const values = sheet.getRange(1, 1, rowsToScan, lastColumn).getValues();
  const candidates = values.map((row, index) => ({
    rowNumber: index + 1,
    values: row,
    filledCells: countFilledCells(row)
  })).sort((first, second) => {
    if (second.filledCells !== first.filledCells) {
      return second.filledCells - first.filledCells;
    }
    return first.rowNumber - second.rowNumber;
  });

  if (headerProfile && headerProfile.length) {
    for (let i = 0; i < candidates.length; i++) {
      if (hasHeaderProfile(candidates[i].values, headerProfile)) {
        return candidates[i].rowNumber;
      }
    }
  }

  for (let i = 0; i < candidates.length; i++) {
    if (hasHeaderAnchor(candidates[i].values)) {
      return candidates[i].rowNumber;
    }
  }

  return 0;
}

/**
 * Gets detected header row metadata and a normalized header-to-column map.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to inspect.
 * @param {string[]=} headerProfile Required headers used to detect the correct header row.
 * @return {Object} Header metadata with row number and column map.
 */
function getHeaderInfo(sheet, headerProfile) {
  const headerRow = findHeaderRow(sheet, headerProfile);
  if (!headerRow) {
    logMissingHeaderRow(sheet, headerProfile);
    return {
      row: 0,
      columnsByHeader: {}
    };
  }

  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  return {
    row: headerRow,
    columnsByHeader: buildHeaderColumnMap(headers)
  };
}

/**
 * Builds a map from normalized header text to one-based column number.
 *
 * @param {*[]} headers Header row values.
 * @return {Object.<string, number>} Header-to-column map.
 */
function buildHeaderColumnMap(headers) {
  return headers.reduce((result, header, index) => {
    const normalizedHeader = normalizeHeaderText(header);
    if (normalizedHeader && !result[normalizedHeader]) {
      result[normalizedHeader] = index + 1;
    }
    return result;
  }, {});
}

/**
 * Finds a column in preloaded header metadata.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet being written.
 * @param {Object} headerInfo Header metadata from getHeaderInfo().
 * @param {string} headerText Header text to find.
 * @param {number} fallbackColumn One-based column returned when the header is not found.
 * @return {number} Matching one-based column number.
 */
function getColumnFromHeaderInfo(sheet, headerInfo, headerText, fallbackColumn) {
  if (!headerInfo.row) return fallbackColumn;

  const expectedHeaders = getHeaderAliases(headerText);
  for (let i = 0; i < expectedHeaders.length; i++) {
    const column = headerInfo.columnsByHeader[expectedHeaders[i]];
    if (column) return column;
  }

  logMissingHeaderColumn(sheet, headerText, fallbackColumn);
  return fallbackColumn;
}

/**
 * Counts non-empty cells in a row.
 *
 * @param {*[]} row Row values.
 * @return {number} Number of filled cells.
 */
function countFilledCells(row) {
  return row.filter(value => normalizeHeaderText(value) !== '').length;
}

/**
 * Checks whether a row contains all headers from a required profile.
 *
 * @param {*[]} row Row values.
 * @param {string[]} headerProfile Required headers.
 * @return {boolean} True when all profile headers are present.
 */
function hasHeaderProfile(row, headerProfile) {
  const normalizedHeaders = row.map(normalizeHeaderText);
  return headerProfile.every(headerText => {
    return getHeaderAliases(headerText).some(alias => normalizedHeaders.includes(alias));
  });
}

/**
 * Checks whether a row contains a generic ID header anchor.
 *
 * @param {*[]} row Row values.
 * @return {boolean} True when the row contains Person ID or Force ID.
 */
function hasHeaderAnchor(row) {
  return hasHeaderProfile(row, ['Person ID']) || hasHeaderProfile(row, ['Force ID']);
}

/**
 * Logs a missing header row warning once per sheet during one execution.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet where the header row was not found.
 * @param {string[]=} headerProfile Required headers used for matching.
 * @return {void}
 */
function logMissingHeaderRow(sheet, headerProfile) {
  const spreadsheetName = sheet.getParent().getName();
  const sheetName = sheet.getName();
  const expectedHeaders = headerProfile && headerProfile.length
    ? headerProfile.join(', ')
    : 'Person ID, Force ID, P ID, F ID';
  const logKey = `${spreadsheetName}::${sheetName}::${expectedHeaders}`;

  if (!logMissingHeaderRow.loggedSheets) {
    logMissingHeaderRow.loggedSheets = {};
  }
  if (logMissingHeaderRow.loggedSheets[logKey]) return;

  Logger.log(
    `Header row was not found in "${spreadsheetName}" / "${sheetName}". ` +
    `Expected headers: ${expectedHeaders}. Fallback columns will be used where available.`
  );
  logMissingHeaderRow.loggedSheets[logKey] = true;
}

/**
 * Logs a missing header column warning once per sheet and header during one execution.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet where the header was not found.
 * @param {string} headerText Expected header text.
 * @param {number} fallbackColumn Fallback column used when the header is missing.
 * @return {void}
 */
function logMissingHeaderColumn(sheet, headerText, fallbackColumn) {
  const spreadsheetName = sheet.getParent().getName();
  const sheetName = sheet.getName();
  const fallbackMessage = fallbackColumn
    ? `Fallback column ${fallbackColumn} will be used.`
    : 'No fallback column is configured, so this value will not be written.';
  const logKey = `${spreadsheetName}::${sheetName}::${normalizeHeaderText(headerText)}`;

  if (!logMissingHeaderColumn.loggedHeaders) {
    logMissingHeaderColumn.loggedHeaders = {};
  }
  if (logMissingHeaderColumn.loggedHeaders[logKey]) return;

  Logger.log(
    `Header "${headerText}" was not found in "${spreadsheetName}" / "${sheetName}". ` +
    fallbackMessage
  );
  logMissingHeaderColumn.loggedHeaders[logKey] = true;
}

/**
 * Gets normalized aliases for a header.
 *
 * @param {string} headerText Source header text.
 * @return {string[]} Header aliases.
 */
function getHeaderAliases(headerText) {
  const normalizedHeader = normalizeHeaderText(headerText);
  const aliases = {
    'person id': ['person id', 'p id'],
    'force id': ['force id', 'f id']
  };

  return aliases[normalizedHeader] || [normalizedHeader];
}

/**
 * Creates a header-based row value setter bound to a sheet and header profile.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet where values will be written.
 * @param {Object.<number, *>} columnValues Row values keyed by one-based column number.
 * @param {string[]} headerProfile Required headers used to detect the correct header row.
 * @return {Function} Setter accepting header text, value, and optional fallback column.
 */
function createHeaderValueSetter(sheet, columnValues, headerProfile) {
  const headerInfo = getHeaderInfo(sheet, headerProfile);
  return function(headerText, value, fallbackColumn) {
    const column = getColumnFromHeaderInfo(sheet, headerInfo, headerText, fallbackColumn || 0);
    if (column) {
      columnValues[column] = value;
    }
  };
}

/**
 * Adds a row value by finding the destination column by its header text.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet where the value will be written.
 * @param {Object.<number, *>} columnValues Row values keyed by one-based column number.
 * @param {string} headerText Header text that identifies the destination column.
 * @param {*} value Value to write when the header exists.
 * @param {number} [fallbackColumn] One-based column used when the header is not found.
 * @param {string[]=} headerProfile Required headers used to detect the correct header row.
 * @return {void}
 */
function setValueByHeader(sheet, columnValues, headerText, value, fallbackColumn, headerProfile) {
  const column = getColumnByHeader(sheet, headerText, fallbackColumn || 0, headerProfile);
  if (column) {
    columnValues[column] = value;
  }
}

/**
 * Finds the first empty cell row in a column after the detected header row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to inspect.
 * @param {number} column One-based column to scan.
 * @return {number} First one-based row number with no value or formula in the column.
 */
function findFirstEmptyRowInColumn(sheet, column) {
  const headerRow = findHeaderRow(sheet);
  const firstDataRow = Math.max((headerRow || 1) + 1, 2);
  const maxRows = sheet.getMaxRows();

  if (maxRows < firstDataRow) {
    sheet.insertRowsAfter(maxRows, firstDataRow - maxRows);
    return firstDataRow;
  }

  const rowCount = maxRows - firstDataRow + 1;
  const range = sheet.getRange(firstDataRow, column, rowCount, 1);
  const values = range.getValues();
  const formulas = range.getFormulas();

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    if (values[rowIndex][0] === '' && formulas[rowIndex][0] === '') {
      return firstDataRow + rowIndex;
    }
  }

  sheet.insertRowAfter(maxRows);
  return maxRows + 1;
}

/**
 * Normalizes header text for reliable comparisons.
 *
 * @param {*} value Header value.
 * @return {string} Normalized header text.
 */
function normalizeHeaderText(value) {
  return String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();
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
  let values = [normalizeValueForWrite(columnValues[startColumn])];
  let previousColumn = startColumn;

  for (let i = 1; i < columns.length; i++) {
    const column = columns[i];
    if (column === previousColumn + 1) {
      values.push(normalizeValueForWrite(columnValues[column]));
    } else {
      sheet.getRange(row, startColumn, 1, values.length).setValues([values]);
      startColumn = column;
      values = [normalizeValueForWrite(columnValues[column])];
    }
    previousColumn = column;
  }

  sheet.getRange(row, startColumn, 1, values.length).setValues([values]);
}

/**
 * Normalizes text values before writing to spreadsheets.
 *
 * @param {*} value Value to normalize.
 * @return {*} Normalized text value, or the original non-text value.
 */
function normalizeValueForWrite(value) {
  if (typeof value !== 'string') return value;
  return value.replace(/\s+/g, ' ').trim();
}

/**
 * Waits for sheet formulas to update and applies checkbox validation to a cell.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination sheet.
 * @param {number} row Destination row number.
 * @param {number} column Destination column number.
 * @return {void}
 */
function ensureCheckboxAfterFormulaUpdate(sheet, row, column) {
  sheet.getRange(row, column).clearContent().clearDataValidations();
  SpreadsheetApp.flush();
  const rule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .setAllowInvalid(true)
    .build();
  sheet.getRange(row, column).setDataValidation(rule);
}
