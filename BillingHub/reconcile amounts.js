/***************************************
 * Payroll vs Individual Timesheets reconciliation v3
 *
 * Does NOT change values.
 * Does NOT change formatting.
 *
 * Writes/clears NOTES only in Payroll:
 * - Primary Amount
 * - Secondary Amount
 * - Invoiced Amount
 *
 * ONE SERVICE:
 *   TS Invoiced Amount == Payroll Primary Amount
 *   Payroll Secondary Amount == 0
 *   TS Invoiced Amount == Payroll Invoiced Amount
 *
 * TWO SERVICES:
 *   TS Primary Service == Payroll Primary Amount
 *   TS Secondary Service == Payroll Secondary Amount
 *   TS Invoiced Total Amount == Payroll Invoiced Amount
 *
 * Internal Payroll check:
 *   Primary Amount + Secondary Amount == Invoiced Amount
 ***************************************/
function payroll_reconcileAmountsWithTimesheets_v3() {
  const CFG = {
    MASTER_SHEET: 'Master',
    MASTER_HEADER_ROW: 2,
    M_FORCE: 'Force ID',
    M_AGR: 'Agreement No',
    M_NAME: 'Name',
    M_TS_ID: 'Individual Timesheet',

    PAYROLL_SHEET: 'Payroll',
    PAYROLL_HEADER_ROW: 2,
    P_FORCE: 'Force ID',
    P_AGR: 'Agreement No',
    P_MONTH: 'Month',
    P_PRIMARY: 'Primary Amount',
    P_SECONDARY: 'Secondary Amount',
    P_INV: 'Invoiced Amount',

    TS_SHEET: 'Calculations',
    TS_MONTH: 'Month',

    // one-service model
    TS_ONE_INV: 'Invoiced Amount',

    // two-service model
    TS_TWO_PRIMARY: 'Primary Service',
    TS_TWO_SECONDARY: 'Secondary Service',
    TS_TWO_TOTAL: 'Invoiced Total Amount',

    LOG_SHEET: 'Payroll Reconciliation Log',

    HEADER_SCAN_ROWS: 60,
    HEADER_SCAN_COLS: 40,
    MAX_MONTH_ROWS: 13,

    TOLERANCE: 0.01,

    TS_TZ: 'Europe/Warsaw',
    TS_FMT: 'yyyy-MM-dd HH:mm:ss'
  };

  const ss = SpreadsheetApp.getActive();
  const shM = ss.getSheetByName(CFG.MASTER_SHEET);
  const shP = ss.getSheetByName(CFG.PAYROLL_SHEET);

  if (!shM) throw new Error(`Sheet "${CFG.MASTER_SHEET}" not found`);
  if (!shP) throw new Error(`Sheet "${CFG.PAYROLL_SHEET}" not found`);

  const runAt = Utilities.formatDate(new Date(), CFG.TS_TZ, CFG.TS_FMT);

  /******** MASTER ********/
  const mLastCol = shM.getLastColumn();
  const mHdr = shM.getRange(CFG.MASTER_HEADER_ROW, 1, 1, mLastCol)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const mIdx = recon_indexByHeader_(mHdr);

  const ixMF = recon_reqHeaderIndex_(mIdx, CFG.M_FORCE, 'Master');
  const ixMA = recon_reqHeaderIndex_(mIdx, CFG.M_AGR, 'Master');
  const ixMN = recon_reqHeaderIndex_(mIdx, CFG.M_NAME, 'Master');
  const ixMT = recon_reqHeaderIndex_(mIdx, CFG.M_TS_ID, 'Master');

  const mLastRow = recon_lastDataRowByColumn_(shM, ixMF + 1, CFG.MASTER_HEADER_ROW);
  if (mLastRow <= CFG.MASTER_HEADER_ROW) return 'No Master data';

  const mData = shM.getRange(
    CFG.MASTER_HEADER_ROW + 1,
    1,
    mLastRow - CFG.MASTER_HEADER_ROW,
    mLastCol
  ).getValues();

  /******** PAYROLL ********/
  const pLastRow = shP.getLastRow();
  const pLastCol = shP.getLastColumn();
  if (pLastRow <= CFG.PAYROLL_HEADER_ROW) return 'No Payroll data';

  const pHdr = shP.getRange(CFG.PAYROLL_HEADER_ROW, 1, 1, pLastCol)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const pIdx = recon_indexByHeader_(pHdr);

  const ixPF = recon_reqHeaderIndex_(pIdx, CFG.P_FORCE, 'Payroll');
  const ixPA = recon_reqHeaderIndex_(pIdx, CFG.P_AGR, 'Payroll');
  const ixPM = recon_reqHeaderIndex_(pIdx, CFG.P_MONTH, 'Payroll');
  const ixPPri = recon_reqHeaderIndex_(pIdx, CFG.P_PRIMARY, 'Payroll');
  const ixPSec = recon_reqHeaderIndex_(pIdx, CFG.P_SECONDARY, 'Payroll');
  const ixPInv = recon_reqHeaderIndex_(pIdx, CFG.P_INV, 'Payroll');

  const pData = shP.getRange(
    CFG.PAYROLL_HEADER_ROW + 1,
    1,
    pLastRow - CFG.PAYROLL_HEADER_ROW,
    pLastCol
  ).getValues();

  const payrollMap = new Map();

  for (let r = 0; r < pData.length; r++) {
    const key = recon_makeKey_(pData[r][ixPF], pData[r][ixPA], pData[r][ixPM]);
    if (!key) continue;

    if (payrollMap.has(key)) {
      const oldRow = payrollMap.get(key);
      throw new Error(
        'Duplicate Payroll key detected:\n' +
        key + '\n\n' +
        'Rows: ' +
        (CFG.PAYROLL_HEADER_ROW + 1 + oldRow) +
        ' and ' +
        (CFG.PAYROLL_HEADER_ROW + 1 + r)
      );
    }

    payrollMap.set(key, r);
  }

  /******** NOTES ONLY ********/
  const priRange = shP.getRange(CFG.PAYROLL_HEADER_ROW + 1, ixPPri + 1, pData.length, 1);
  const secRange = shP.getRange(CFG.PAYROLL_HEADER_ROW + 1, ixPSec + 1, pData.length, 1);
  const invRange = shP.getRange(CFG.PAYROLL_HEADER_ROW + 1, ixPInv + 1, pData.length, 1);

  const priNotes = priRange.getNotes();
  const secNotes = secRange.getNotes();
  const invNotes = invRange.getNotes();

  // Clear previous reconciliation notes in checked amount columns.
  // Values and formatting are untouched.
  for (let r = 0; r < pData.length; r++) {
    priNotes[r][0] = '';
    secNotes[r][0] = '';
    invNotes[r][0] = '';
  }

  /******** LOG ********/
  const logRows = [[
    'Run At',
    'Status',
    'Person',
    'Force ID',
    'Agreement No',
    'Month',
    'Timesheet Model',
    'Field',
    'Payroll Value',
    'Timesheet Value',
    'Difference',
    'Details'
  ]];

  let peopleProcessed = 0;
  let tsOpened = 0;
  let monthsChecked = 0;
  let mismatchCount = 0;
  let missingPayrollCount = 0;
  let tsErrorCount = 0;

  /******** ITERATE MASTER ********/
  for (const rowM of mData) {
    const fid = recon_normId_(rowM[ixMF]);
    const agr = recon_normAgreement_(rowM[ixMA]);
    const name = String(rowM[ixMN] || '').trim();
    const tsId = recon_extractSpreadsheetId_(rowM[ixMT]);

    if (!fid || !agr || !tsId) continue;

    peopleProcessed++;

    let ts;
    try {
      ts = SpreadsheetApp.openById(tsId);
      tsOpened++;
    } catch (e) {
      tsErrorCount++;
      logRows.push([
        runAt,
        'ERROR',
        name,
        fid,
        agr,
        '',
        '',
        'Timesheet',
        '',
        '',
        '',
        'Cannot open timesheet: ' + tsId
      ]);
      continue;
    }

    const shT = ts.getSheetByName(CFG.TS_SHEET);
    if (!shT) {
      tsErrorCount++;
      logRows.push([
        runAt,
        'ERROR',
        name,
        fid,
        agr,
        '',
        '',
        'Timesheet',
        '',
        '',
        '',
        `Sheet "${CFG.TS_SHEET}" not found`
      ]);
      continue;
    }

    const tsModel = recon_detectTimesheetAmountModel_(shT, CFG);
    if (!tsModel) {
      tsErrorCount++;
      logRows.push([
        runAt,
        'ERROR',
        name,
        fid,
        agr,
        '',
        '',
        'Timesheet',
        '',
        '',
        '',
        'Required amount headers not found'
      ]);
      continue;
    }

    const tsRows = recon_readTimesheetAmountsByModel_(shT, tsModel, CFG);

    for (const rec of tsRows) {
      const key = `${fid}|${agr}|${rec.monthKey}`;
      const pRowIndex = payrollMap.get(key);

      if (pRowIndex == null) {
        missingPayrollCount++;
        logRows.push([
          runAt,
          'MISSING_PAYROLL',
          name,
          fid,
          agr,
          rec.monthKey,
          rec.model,
          'Payroll row',
          '',
          '',
          '',
          'No matching Payroll row by Force ID + Agreement No + Month'
        ]);
        continue;
      }

      monthsChecked++;

      const pRow = pData[pRowIndex];

      const pPrimary = recon_toNum_(pRow[ixPPri]);
      const pSecondary = recon_toNum_(pRow[ixPSec]);
      const pInv = recon_toNum_(pRow[ixPInv]);

      /******** INTERNAL PAYROLL CONSISTENCY ********/
      const internalDiff = recon_round2_((pPrimary + pSecondary) - pInv);

      if (Math.abs(internalDiff) > CFG.TOLERANCE) {
        mismatchCount++;

        const note = [
          'INTERNAL PAYROLL ERROR',
          'Primary + Secondary must equal Invoiced',
          'Primary + Secondary: ' + recon_round2_(pPrimary + pSecondary),
          'Invoiced: ' + pInv,
          'Difference: ' + internalDiff,
          'Month: ' + recon_formatMonthLabel_(rec.monthKey),
          'Model: ' + rec.model
        ].join('\n');

        recon_setNote_(priNotes, pRowIndex, note);
        recon_setNote_(secNotes, pRowIndex, note);
        recon_setNote_(invNotes, pRowIndex, note);

        logRows.push([
          runAt,
          'INTERNAL_PAYROLL_INCONSISTENCY',
          name,
          fid,
          agr,
          recon_formatMonthLabel_(rec.monthKey),
          rec.model,
          'Primary + Secondary vs Invoiced',
          recon_round2_(pPrimary + pSecondary),
          pInv,
          internalDiff,
          'Primary + Secondary must equal Invoiced'
        ]);
      }

      /******** TS VS PAYROLL ********/
      if (rec.model === 'ONE_SERVICE') {
        mismatchCount += recon_compareAndNote_({
          runAt,
          logRows,
          name,
          fid,
          agr,
          monthKey: rec.monthKey,
          model: rec.model,
          field: 'Primary Amount',
          payrollValue: pPrimary,
          tsValue: rec.invoiced,
          notesArr: priNotes,
          rowIndex: pRowIndex,
          CFG
        });

        mismatchCount += recon_compareAndNote_({
          runAt,
          logRows,
          name,
          fid,
          agr,
          monthKey: rec.monthKey,
          model: rec.model,
          field: 'Secondary Amount',
          payrollValue: pSecondary,
          tsValue: 0,
          notesArr: secNotes,
          rowIndex: pRowIndex,
          CFG
        });

        mismatchCount += recon_compareAndNote_({
          runAt,
          logRows,
          name,
          fid,
          agr,
          monthKey: rec.monthKey,
          model: rec.model,
          field: 'Invoiced Amount',
          payrollValue: pInv,
          tsValue: rec.invoiced,
          notesArr: invNotes,
          rowIndex: pRowIndex,
          CFG
        });
      }

      if (rec.model === 'TWO_SERVICES') {
        mismatchCount += recon_compareAndNote_({
          runAt,
          logRows,
          name,
          fid,
          agr,
          monthKey: rec.monthKey,
          model: rec.model,
          field: 'Primary Amount',
          payrollValue: pPrimary,
          tsValue: rec.primary,
          notesArr: priNotes,
          rowIndex: pRowIndex,
          CFG
        });

        mismatchCount += recon_compareAndNote_({
          runAt,
          logRows,
          name,
          fid,
          agr,
          monthKey: rec.monthKey,
          model: rec.model,
          field: 'Secondary Amount',
          payrollValue: pSecondary,
          tsValue: rec.secondary,
          notesArr: secNotes,
          rowIndex: pRowIndex,
          CFG
        });

        mismatchCount += recon_compareAndNote_({
          runAt,
          logRows,
          name,
          fid,
          agr,
          monthKey: rec.monthKey,
          model: rec.model,
          field: 'Invoiced Amount',
          payrollValue: pInv,
          tsValue: rec.invoiced,
          notesArr: invNotes,
          rowIndex: pRowIndex,
          CFG
        });
      }
    }
  }

  /******** APPLY NOTES ONLY ********/
  priRange.setNotes(priNotes);
  secRange.setNotes(secNotes);
  invRange.setNotes(invNotes);

  /******** WRITE LOG ********/
  const shLog = recon_getOrCreateSheet_(ss, CFG.LOG_SHEET);
  shLog.clearContents();
  shLog.clearFormats();

  shLog.getRange(1, 1, logRows.length, logRows[0].length).setValues(logRows);
  shLog.getRange(1, 1, 1, logRows[0].length).setFontWeight('bold');
  shLog.setFrozenRows(1);
  shLog.autoResizeColumns(1, logRows[0].length);

  const summary =
    'Payroll reconciliation completed.\n' +
    'People processed: ' + peopleProcessed + '\n' +
    'Timesheets opened: ' + tsOpened + '\n' +
    'Months checked: ' + monthsChecked + '\n' +
    'Mismatches found: ' + mismatchCount + '\n' +
    'Missing Payroll rows: ' + missingPayrollCount + '\n' +
    'Timesheet errors: ' + tsErrorCount;

  Logger.log(summary);

  SpreadsheetApp.flush();

  try {
    SpreadsheetApp.getUi().alert(summary);
  } catch (e) {}

  return summary;
}

/* ===========================
 * Reconciliation helpers
 * =========================== */
function recon_detectTimesheetAmountModel_(shT, CFG) {
  const two = recon_findTableHeaders_(
    shT,
    [CFG.TS_MONTH, CFG.TS_TWO_PRIMARY, CFG.TS_TWO_SECONDARY, CFG.TS_TWO_TOTAL],
    CFG.HEADER_SCAN_ROWS,
    CFG.HEADER_SCAN_COLS
  );

  if (two) {
    return {
      model: 'TWO_SERVICES',
      headerRow: two.headerRow,
      colByName: two.colByName
    };
  }

  const one = recon_findTableHeaders_(
    shT,
    [CFG.TS_MONTH, CFG.TS_ONE_INV],
    CFG.HEADER_SCAN_ROWS,
    CFG.HEADER_SCAN_COLS
  );

  if (one) {
    return {
      model: 'ONE_SERVICE',
      headerRow: one.headerRow,
      colByName: one.colByName
    };
  }

  return null;
}

function recon_readTimesheetAmountsByModel_(shT, tsModel, CFG) {
  const startRow = tsModel.headerRow + 1;
  const colMonth = tsModel.colByName[CFG.TS_MONTH] + 1;

  const monthVals = shT.getRange(startRow, colMonth, CFG.MAX_MONTH_ROWS + 2, 1).getValues();

  const months = [];
  for (let i = 0; i < monthVals.length && months.length < CFG.MAX_MONTH_ROWS; i++) {
    const v = monthVals[i][0];

    if (!v) break;
    if (String(v).toUpperCase().includes('TOTAL')) break;

    const d = recon_coerceDate_(v);
    if (!d) break;

    months.push({
      row: startRow + i,
      monthKey: recon_monthKeyUTC_(recon_monthFirstLocal_(d))
    });
  }

  if (!months.length) return [];

  if (tsModel.model === 'ONE_SERVICE') {
    const colInv = tsModel.colByName[CFG.TS_ONE_INV] + 1;
    const invVals = shT.getRange(months[0].row, colInv, months.length, 1).getDisplayValues();

    return months.map((m, i) => ({
      model: 'ONE_SERVICE',
      monthKey: m.monthKey,
      invoiced: recon_toNum_(invVals[i][0])
    }));
  }

  const colPri = tsModel.colByName[CFG.TS_TWO_PRIMARY] + 1;
  const colSec = tsModel.colByName[CFG.TS_TWO_SECONDARY] + 1;
  const colInv = tsModel.colByName[CFG.TS_TWO_TOTAL] + 1;

  const priVals = shT.getRange(months[0].row, colPri, months.length, 1).getDisplayValues();
  const secVals = shT.getRange(months[0].row, colSec, months.length, 1).getDisplayValues();
  const invVals = shT.getRange(months[0].row, colInv, months.length, 1).getDisplayValues();

  return months.map((m, i) => ({
    model: 'TWO_SERVICES',
    monthKey: m.monthKey,
    primary: recon_toNum_(priVals[i][0]),
    secondary: recon_toNum_(secVals[i][0]),
    invoiced: recon_toNum_(invVals[i][0])
  }));
}

function recon_compareAndNote_(o) {
  const payroll = recon_toNum_(o.payrollValue);
  const ts = recon_toNum_(o.tsValue);
  const diff = recon_round2_(payroll - ts);

  if (Math.abs(diff) <= o.CFG.TOLERANCE) {
    // Note has already been cleared at the beginning.
    // If value is correct, it remains empty.
    return 0;
  }

  const note = [
    'PAYROLL RECONCILIATION MISMATCH',
    'Field: ' + o.field,
    'Payroll: ' + payroll,
    'Timesheet: ' + ts,
    'Difference: ' + diff,
    'Month: ' + recon_formatMonthLabel_(o.monthKey),
    'Model: ' + o.model
  ].join('\n');

  // Overwrite note, not append.
  recon_setNote_(o.notesArr, o.rowIndex, note);

  o.logRows.push([
    o.runAt,
    'MISMATCH',
    o.name,
    o.fid,
    o.agr,
    o.monthKey,
    o.model,
    o.field,
    payroll,
    ts,
    diff,
    ''
  ]);

  return 1;
}

function recon_setNote_(notesArr, rowIndex, text) {
  notesArr[rowIndex][0] = String(text || '');
}

function recon_getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function recon_makeKey_(forceId, agreementNo, monthDate) {
  const fid = recon_normId_(forceId);
  const ag = recon_normAgreement_(agreementNo);
  const d = recon_coerceDate_(monthDate);

  if (!fid || !ag || !d) return '';

  return fid + '|' + ag + '|' + recon_monthKeyUTC_(recon_monthFirstLocal_(d));
}

function recon_indexByHeader_(hdrRow) {
  const m = {};
  for (let i = 0; i < hdrRow.length; i++) {
    m[String(hdrRow[i] || '').trim()] = i;
  }
  return m;
}

function recon_reqHeaderIndex_(idxMap, name, sheetLabel) {
  const ix = idxMap[name];
  if (ix == null) throw new Error(sheetLabel + ' header not found: "' + name + '"');
  return ix;
}

function recon_findTableHeaders_(sh, requiredHeaders, maxRows, maxCols) {
  const rows = Math.min(maxRows || 60, sh.getMaxRows());
  const cols = Math.min(maxCols || 40, sh.getMaxColumns());

  const normalizeHeader = v => String(v || '')
    .replace(/\s+/g, ' ')
    .trim();

  const requiredNorm = requiredHeaders.map(normalizeHeader);
  const want = new Set(requiredNorm);

  const grid = sh.getRange(1, 1, rows, cols).getValues();

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(normalizeHeader);
    const colByName = {};

    for (let c = 0; c < row.length; c++) {
      const h = row[c];
      if (want.has(h)) colByName[h] = c;
    }

    const ok = requiredNorm.every(h => colByName[h] != null);

    if (ok) {
      const resultColByName = {};
      for (let i = 0; i < requiredHeaders.length; i++) {
        resultColByName[requiredHeaders[i]] = colByName[requiredNorm[i]];
      }

      return {
        headerRow: r + 1,
        colByName: resultColByName
      };
    }
  }

  return null;
}

function recon_lastDataRowByColumn_(sh, col, headerRow) {
  const lastRow = sh.getLastRow();
  if (lastRow <= headerRow) return headerRow;

  const vals = sh.getRange(headerRow + 1, col, lastRow - headerRow, 1).getValues();

  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || '').trim() !== '') {
      return headerRow + 1 + i;
    }
  }

  return headerRow;
}

function recon_extractSpreadsheetId_(v) {
  if (v == null) return '';

  const s = String(v).trim();
  if (!s) return '';

  if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return s;

  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return m && m[1] ? m[1] : '';
}

function recon_normId_(v) {
  if (v === null || v === undefined) return '';

  let s = String(v).trim();
  if (!s) return '';

  s = s
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, '-')
    .replace(/\s+/g, ' ');

  if (/^\d+$/.test(s)) {
    s = s.replace(/^0+/, '');
    return s === '' ? '0' : s;
  }

  return s.toUpperCase()
    .replace(/\s*-\s*/g, '-')
    .replace(/-+/g, '-')
    .replace(/\s+/g, ' ');
}

function recon_normAgreement_(v) {
  if (v === null || v === undefined) return '';

  return String(v).trim()
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, '-')
    .replace(/\s+/g, ' ');
}

function recon_coerceDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  if (typeof v === 'number') {
    const base = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(base.getTime() + Math.round(v) * 86400000);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  const s = String(v || '').trim();
  if (!s) return null;

  const d = new Date(s);
  if (isNaN(d.getTime())) return null;

  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function recon_monthFirstLocal_(d) {
  const x = recon_coerceDate_(d);
  if (!x) return null;
  return new Date(x.getFullYear(), x.getMonth(), 1);
}

function recon_monthKeyUTC_(d) {
  const x = recon_coerceDate_(d);
  if (!x) return '';

  const y = x.getFullYear();
  const m = x.getMonth() + 1;

  return y + '-' + (m < 10 ? '0' + m : m) + '-01';
}

function recon_toNum_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;

  let s = String(v).trim();
  if (!s) return 0;

  s = s.replace(/[^\d.,-]/g, '');

  const lastComma = s.lastIndexOf(',');
  const lastDot = s.lastIndexOf('.');

  if (lastComma !== -1 && lastDot !== -1) {
    if (lastDot > lastComma) {
      s = s.replace(/,/g, '');
    } else {
      s = s.replace(/\./g, '').replace(/,/g, '.');
    }
  } else if (lastComma !== -1) {
    s = s.replace(/,/g, '.');
  }

  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function recon_round2_(v) {
  const n = Number(v);
  if (!isFinite(n)) return 0;
  return Math.round(n * 100) / 100;
}

function recon_findAnyAmountTableHeaders_(sh, CFG, maxRows, maxCols) {
  const rows = Math.min(maxRows || 60, sh.getMaxRows());
  const cols = Math.min(maxCols || 40, sh.getMaxColumns());

  const grid = sh.getRange(1, 1, rows, cols).getValues();

  const headersToFind = [
    CFG.TS_MONTH,
    CFG.TS_ONE_INV,
    CFG.TS_TWO_PRIMARY,
    CFG.TS_TWO_SECONDARY
  ];

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(v => String(v || '').trim());
    const colByName = {};

    for (let c = 0; c < row.length; c++) {
      const h = row[c];
      if (headersToFind.includes(h)) {
        colByName[h] = c;
      }
    }

    const hasMonth = colByName[CFG.TS_MONTH] != null;
    const hasInvoiced = colByName[CFG.TS_ONE_INV] != null;

    if (hasMonth && hasInvoiced) {
      return {
        headerRow: r + 1,
        colByName
      };
    }
  }

  return null;
}

function recon_formatMonthLabel_(monthKey) {
  const d = recon_coerceDate_(monthKey);
  if (!d) return monthKey;

  return Utilities.formatDate(d, 'en-US', 'MMMM yyyy'); // October
}

function payroll_clearAllNotes() {
  const SHEET_NAME = 'Payroll';

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found`);

  const range = sh.getDataRange();
  range.clearNote();

  SpreadsheetApp.getUi().alert('All notes on Payroll were cleared.');
}