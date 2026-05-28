/***************************************
 * Payroll -> Individual Timesheets (Calculations)
 *
 * LOGIC:
 * - Current + Next month: ALWAYS overwrite Planned/Unplanned from Payroll
 * - Other months: overwrite ONLY if real absences detected
 * - Δ columns mirror Payroll 1:1 (including clearing)
 *
 * IMPORTANT:
 * - Supports merged first payroll period automatically
 * - Mapping is done by Payroll.Month <-> Timesheet.Month
 * - In merged case Payroll.Month already stores the full/second month,
 *   which matches the first visible milestone row in timesheet
 ***************************************/
function payroll_updateIndividualTimesheetsFromPayroll() {
  const CFG = {
    MASTER_SHEET: 'Master',
    MASTER_HEADER_ROW: 2,
    M_FORCE: 'Force ID',
    M_AGREEMENT: 'Agreement No',
    M_NAME: 'Name',
    M_TS_ID: 'Individual Timesheet',

    PAYROLL_SHEET: 'Payroll',
    PAYROLL_HEADER_ROW: 2,
    P_FORCE: 'Force ID',
    P_AGREEMENT: 'Agreement No',
    P_MONTH: 'Month',
    P_SERVICE_START: 'Service Start',
    P_SERVICE_END: 'Service End',
    P_MAX: 'Max Workload',
    P_PLAN: 'Planned Workload',
    P_UNPL: 'Unplanned Absence',
    P_D_IS: 'Δ Invoiced Scope',
    P_D_IA: 'Δ Invoiced Amount',

    TS_SHEET: 'Calculations',
    TS_MONTH: 'Month',
    TS_PLAN: 'Planned Workload',
    TS_UNPL: 'Unplanned Absence',
    TS_D_IS: 'Δ Invoiced Scope',
    TS_D_IA: 'Δ Invoiced Amount',

    MAX_MONTH_ROWS: 13,
    HEADER_SCAN_ROWS: 60,
    HEADER_SCAN_COLS: 26,
  };

  const ss = SpreadsheetApp.getActive();
  const shM = ss.getSheetByName(CFG.MASTER_SHEET);
  const shP = ss.getSheetByName(CFG.PAYROLL_SHEET);
  if (!shM || !shP) throw new Error('Master or Payroll sheet not found');

  /* ================= CURRENT + NEXT MONTH KEYS ================= */

  const today = new Date();
  const mkCur = _monthKeyUTC_(_monthFirstLocal_(today));
  const mkNext = _monthKeyUTC_(
    _monthFirstLocal_(new Date(today.getFullYear(), today.getMonth() + 1, 1))
  );

  /* ================= MASTER ================= */

  const mLastCol = shM.getLastColumn();
  const mHdr = shM.getRange(CFG.MASTER_HEADER_ROW, 1, 1, mLastCol)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const mIdx = Object.fromEntries(mHdr.map((h, i) => [h, i]));

  const ixMF = _reqHeaderIndex_(mIdx, CFG.M_FORCE, 'Master');
  const ixMA = _reqHeaderIndex_(mIdx, CFG.M_AGREEMENT, 'Master');
  const ixMN = _reqHeaderIndex_(mIdx, CFG.M_NAME, 'Master');
  const ixMT = _reqHeaderIndex_(mIdx, CFG.M_TS_ID, 'Master');

  const lastDataRow = _lastDataRowByColumn_(shM, ixMF + 1, CFG.MASTER_HEADER_ROW);
  if (lastDataRow <= CFG.MASTER_HEADER_ROW) return;

  const mData = shM.getRange(
    CFG.MASTER_HEADER_ROW + 1,
    1,
    lastDataRow - CFG.MASTER_HEADER_ROW,
    mLastCol
  ).getValues();

  /* ================= PAYROLL ================= */

  const pLastRow = shP.getLastRow();
  const pLastCol = shP.getLastColumn();
  if (pLastRow <= CFG.PAYROLL_HEADER_ROW) return;

  const pHdr = shP.getRange(CFG.PAYROLL_HEADER_ROW, 1, 1, pLastCol)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const pIdx = _indexByHeader_(pHdr);

  const ixPF   = _reqHeaderIndex_(pIdx, CFG.P_FORCE, 'Payroll');
  const ixPA   = _reqHeaderIndex_(pIdx, CFG.P_AGREEMENT, 'Payroll');
  const ixPM   = _reqHeaderIndex_(pIdx, CFG.P_MONTH, 'Payroll');
  const ixPSS  = _reqHeaderIndex_(pIdx, CFG.P_SERVICE_START, 'Payroll');
  const ixPSE  = _reqHeaderIndex_(pIdx, CFG.P_SERVICE_END, 'Payroll');
  const ixPMax = _reqHeaderIndex_(pIdx, CFG.P_MAX, 'Payroll');
  const ixPP   = _reqHeaderIndex_(pIdx, CFG.P_PLAN, 'Payroll');
  const ixPU   = _reqHeaderIndex_(pIdx, CFG.P_UNPL, 'Payroll');
  const ixPDS  = _reqHeaderIndex_(pIdx, CFG.P_D_IS, 'Payroll');
  const ixPDA  = _reqHeaderIndex_(pIdx, CFG.P_D_IA, 'Payroll');

  const pData = shP.getRange(
    CFG.PAYROLL_HEADER_ROW + 1,
    1,
    pLastRow - CFG.PAYROLL_HEADER_ROW,
    pLastCol
  ).getValues();

  /**
   * payrollMap:
   * key = ForceID|AgreementNo|MonthKey
   *
   * NOTE:
   * In merged-first-period case, Payroll.Month already contains the visible
   * timesheet milestone month, so this mapping remains correct.
   */
  const payrollMap = new Map();

  for (let i = 0; i < pData.length; i++) {
    const row = pData[i];

    const fid = _normId_(row[ixPF]);
    const ag  = _normAgreement_(row[ixPA]);
    if (!fid || !ag) continue;

    let monthDate = _coerceDate_(row[ixPM]);

    // fallback safety: if Month is broken, derive from Service End month
    if (!monthDate) {
      const serviceEnd = _coerceDate_(row[ixPSE]);
      if (serviceEnd) {
        monthDate = _monthFirstLocal_(serviceEnd);
      }
    }

    if (!monthDate) continue;

    const mk = _monthKeyUTC_(_monthFirstLocal_(monthDate));
    const key = fid + '|' + ag + '|' + mk;

    payrollMap.set(key, {
      max: row[ixPMax],
      plan: row[ixPP],
      unpl: row[ixPU],
      dIS: row[ixPDS],
      dIA: row[ixPDA],
      month: monthDate,
      serviceStart: row[ixPSS],
      serviceEnd: row[ixPSE],
    });
  }

  /* ================= ITERATE MASTER ================= */

  for (let r = 0; r < mData.length; r++) {
    const masterRow = mData[r];

    const fid = _normId_(masterRow[ixMF]);
    const ag = _normAgreement_(masterRow[ixMA]);
    const name = String(masterRow[ixMN] || '').trim();
    const tsId = _extractSpreadsheetId_(masterRow[ixMT]);

    if (!fid || !ag || !tsId) continue;

    let ts;
    try {
      ts = SpreadsheetApp.openById(tsId);
    } catch (e) {
      Logger.log(`Cannot open timesheet: ${name} | ${tsId}`);
      continue;
    }

    const shT = ts.getSheetByName(CFG.TS_SHEET);
    if (!shT) continue;

    const found = _findTableHeaders_(
      shT,
      [CFG.TS_MONTH, CFG.TS_PLAN, CFG.TS_UNPL, CFG.TS_D_IS, CFG.TS_D_IA],
      CFG.HEADER_SCAN_ROWS,
      CFG.HEADER_SCAN_COLS
    );
    if (!found) continue;

    const headerRow = found.headerRow;
    const colMonth = found.colByName[CFG.TS_MONTH] + 1;
    const colPlan  = found.colByName[CFG.TS_PLAN] + 1;
    const colUnpl  = found.colByName[CFG.TS_UNPL] + 1;
    const colDIS   = found.colByName[CFG.TS_D_IS] + 1;
    const colDIA   = found.colByName[CFG.TS_D_IA] + 1;

    const startRow = headerRow + 1;

    const monthVals = shT.getRange(startRow, colMonth, CFG.MAX_MONTH_ROWS + 2, 1).getValues();

    const months = [];
    for (let i = 0; i < monthVals.length && months.length < CFG.MAX_MONTH_ROWS; i++) {
      const v = monthVals[i][0];
      if (!v) break;
      if (String(v).toUpperCase().includes('TOTAL')) break;

      const d = _coerceDate_(v);
      if (!d) break;

      months.push({
        row: startRow + i,
        mk: _monthKeyUTC_(_monthFirstLocal_(d)),
        date: _monthFirstLocal_(d)
      });
    }

    if (!months.length) continue;

    const curPlanBlock = shT.getRange(months[0].row, colPlan, months.length, 2).getValues();
    const curDeltaBlock = shT.getRange(months[0].row, colDIS, months.length, 2).getValues();

    const newPlanBlock = [];
    const newDeltaBlock = [];

    for (let i = 0; i < months.length; i++) {

      const m = months[i];
      const rec = payrollMap.get(`${fid}|${ag}|${m.mk}`);

      const curPlan = curPlanBlock[i][0];
      const curUnpl = curPlanBlock[i][1];

      let nextPlan = curPlan;
      let nextUnpl = curUnpl;

      // always sync from Payroll if record exists
      if (rec) {
        nextPlan = rec.plan;
        nextUnpl = rec.unpl;
      }

      newPlanBlock.push([nextPlan, nextUnpl]);

      // Δ — mirror 1:1 as before
      let nextDIS = curDeltaBlock[i][0];
      let nextDIA = curDeltaBlock[i][1];

      if (rec) {
        nextDIS = (rec.dIS == null ? '' : rec.dIS);
        nextDIA = (rec.dIA == null ? '' : rec.dIA);
      }

      newDeltaBlock.push([nextDIS, nextDIA]);
    }

    if (!_same2D_(curPlanBlock, newPlanBlock)) {
      shT.getRange(months[0].row, colPlan, months.length, 2).setValues(newPlanBlock);
    }

    if (!_same2D_(curDeltaBlock, newDeltaBlock)) {
      shT.getRange(months[0].row, colDIS, months.length, 2).setValues(newDeltaBlock);
    }
  }
}