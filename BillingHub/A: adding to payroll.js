// BASELINE v5 (TOP insert, merged short first month): 2026-03-28
/***************************************
 * Master -> Payroll (billing periods) [TOP INSERT]
 *
 * Rules:
 * - Process ONLY Master rows where Status = "agrt signed"
 * - Create Payroll rows by billing periods, not blindly by calendar months
 * - Special first-period merge rule:
 *   If contract start is NOT the 1st day of month
 *   and workdays from contract start to end of that month <= 5,
 *   then that short first month is merged with the next full month
 *
 * For merged first period:
 * - Month         = first day of SECOND month
 * - Service Start = real contract start
 * - Service End   = end of second month (or contract end if earlier)
 * - Max Workload  = workdays across whole merged period
 *
 * Standard periods:
 * - Month         = first day of that month
 * - Service Start = max(contract start, month first day)
 * - Service End   = min(contract end, month end)
 * - Max Workload  = workdays(Service Start..Service End)
 *
 * - Insert NEW rows at TOP (right under headers)
 * - Do NOT overwrite existing Payroll rows
 * - Prevent duplicates by key: (Force ID, Agreement No, MonthKey)
 * - Hard-stop if Payroll already contains duplicates for that key
 * - After successful insert -> set Status = "payroll created" in Master
 ***************************************/
function payroll_recalcNewRows() {
  const CFG = {
    // Master
    SRC_SHEET: 'Master',
    SRC_HEADER_ROW: 2,
    STATUS_HEADER: 'Status',
    STATUS_REQUIRED: 'agrt signed',
    STATUS_DONE: 'payroll created',

    FORCE_ID_HEADER: 'Force ID',
    PERSON_NAME_HEADER: 'Name',
    AGREEMENT_NO_HEADER: 'Agreement No',

    CONTRACT_START_HEADER: 'Project Start Date',
    CONTRACT_END_HEADER: 'Project End Date',

    // Payroll
    OUT_SHEET: 'Payroll',
    OUT_HEADER_ROW: 2,

    // Calendar
    CALENDAR_SHEET: 'Calendar',
    HOLIDAYS_RANGE_A1: 'A2:A',

    // Merge rule
    SHORT_FIRST_MONTH_MAX_WORKDAYS: 5
  };

  const ss = SpreadsheetApp.getActive();

  const shSrc = ss.getSheetByName(CFG.SRC_SHEET);
  if (!shSrc) throw new Error(`Sheet "${CFG.SRC_SHEET}" not found`);

  const shOut = ss.getSheetByName(CFG.OUT_SHEET);
  if (!shOut) throw new Error(`Sheet "${CFG.OUT_SHEET}" not found`);

  const shCal = ss.getSheetByName(CFG.CALENDAR_SHEET);
  if (!shCal) throw new Error(`Sheet "${CFG.CALENDAR_SHEET}" not found`);

  // -----------------------
  // Payroll header map
  // -----------------------
  const pLastCol = shOut.getLastColumn();
  if (pLastCol < 1) throw new Error('Payroll has no columns');

  const pHdr = shOut.getRange(CFG.OUT_HEADER_ROW, 1, 1, pLastCol).getValues()[0]
    .map(h => String(h || '').trim());

  const pIdx = Object.fromEntries(pHdr.map((h, i) => [h, i]));

  const ixPFid   = _reqHeaderIndex_(pIdx, 'Force ID', 'Payroll');
  const ixPAgNo  = _reqHeaderIndex_(pIdx, 'Agreement No', 'Payroll');
  const ixPName  = _reqHeaderIndex_(pIdx, 'Name', 'Payroll');
  const ixPMonth = _reqHeaderIndex_(pIdx, 'Month', 'Payroll');

  const ixPMax   = pIdx['Max Workload'];
  const ixPMs    = pIdx['Invoice Item description'];
  const ixPSs    = pIdx['Service Start'];
  const ixPSe    = pIdx['Service End'];

  // -----------------------
  // Master headers
  // -----------------------
  const srcLastRow = shSrc.getLastRow();
  const srcLastCol = shSrc.getLastColumn();
  if (srcLastRow <= CFG.SRC_HEADER_ROW) return;

  const mHdr = shSrc.getRange(CFG.SRC_HEADER_ROW, 1, 1, srcLastCol).getValues()[0]
    .map(h => String(h || '').trim());
  const mIdx = Object.fromEntries(mHdr.map((h, i) => [h, i]));

  const ixMStatus = _reqHeaderIndex_(mIdx, CFG.STATUS_HEADER, 'Master');
  const ixMFid    = _reqHeaderIndex_(mIdx, CFG.FORCE_ID_HEADER, 'Master');
  const ixMName   = _reqHeaderIndex_(mIdx, CFG.PERSON_NAME_HEADER, 'Master');
  const ixMAgNo   = _reqHeaderIndex_(mIdx, CFG.AGREEMENT_NO_HEADER, 'Master');
  const ixMStart  = _reqHeaderIndex_(mIdx, CFG.CONTRACT_START_HEADER, 'Master');
  const ixMEnd    = _reqHeaderIndex_(mIdx, CFG.CONTRACT_END_HEADER, 'Master');

  // -----------------------
  // Holidays cache
  // -----------------------
  const holVals = shCal.getRange(CFG.HOLIDAYS_RANGE_A1).getValues().flat().filter(Boolean);
  const holidayTs = Array.from(new Set(
    holVals
      .map(_coerceDate_)
      .filter(Boolean)
      .map(d => _toMidnight_(d).getTime())
  )).sort((a, b) => a - b);

  // -----------------------
  // Existing Payroll keys
  // -----------------------
  const pDataRows = Math.max(0, shOut.getLastRow() - CFG.OUT_HEADER_ROW);
  const existingKeySet = new Set();
  const seenKeyCount = new Map();

  if (pDataRows > 0) {
    const block = shOut.getRange(CFG.OUT_HEADER_ROW + 1, 1, pDataRows, pLastCol).getValues();
    for (let r = 0; r < block.length; r++) {
      const row = block[r];

      const fid = _normId_(row[ixPFid]);
      const ag  = _normAgreement_(row[ixPAgNo]);
      const m   = _coerceDate_(row[ixPMonth]);
      if (!fid || !ag || !m) continue;

      const mk = _monthKeyUTC_(_monthFirstLocal_(m));
      const key = _makeKey_(fid, ag, mk);

      existingKeySet.add(key);
      seenKeyCount.set(key, (seenKeyCount.get(key) || 0) + 1);
    }
  }

  const dupKeys = [];
  for (const [k, cnt] of seenKeyCount.entries()) {
    if (cnt > 1) dupKeys.push(k);
  }
  if (dupKeys.length) {
    throw new Error(
      `Payroll has duplicate rows for key (Force ID, Agreement No, Month): ` +
      `${dupKeys.slice(0, 5).join(' | ')}${dupKeys.length > 5 ? ' ...' : ''}. ` +
      `Fix duplicates first; creation stopped.`
    );
  }

  // -----------------------
  // Read Master data
  // -----------------------
  const masterData = shSrc.getRange(
    CFG.SRC_HEADER_ROW + 1, 1,
    srcLastRow - CFG.SRC_HEADER_ROW, srcLastCol
  ).getValues();

  const outRows = [];
  const statusUpdates = [];

  for (let i = 0; i < masterData.length; i++) {
    const row = masterData[i];

    const st = String(row[ixMStatus] || '').trim().toLowerCase();
    if (st !== String(CFG.STATUS_REQUIRED).trim().toLowerCase()) continue;

    const fidRaw = row[ixMFid];
    const fidKey = _normId_(fidRaw);

    const name = String(row[ixMName] || '').trim();
    const agNo = _normAgreement_(row[ixMAgNo]);

    const start = _coerceDate_(row[ixMStart]);
    const end   = _coerceDate_(row[ixMEnd]);

    if (!fidKey || !agNo || !name || !start || !end) continue;

    const s = _toMidnight_(start);
    const e = _toMidnight_(end);
    if (e.getTime() < s.getTime()) continue;

    const periods = _buildPayrollPeriods_(s, e, holidayTs, CFG.SHORT_FIRST_MONTH_MAX_WORKDAYS);
    if (!periods.length) continue;

    let createdAny = false;
    let milestoneN = 0;

    for (let p = 0; p < periods.length; p++) {
      const period = periods[p];

      const mk = _monthKeyUTC_(period.monthFirst);
      const key = _makeKey_(fidKey, agNo, mk);
      if (existingKeySet.has(key)) continue;

      milestoneN++;
      const milestone = `Milestone ${milestoneN}`;

      const out = new Array(pLastCol).fill('');

      out[ixPFid]   = fidRaw;
      out[ixPAgNo]  = agNo;
      out[ixPName]  = name;
      out[ixPMonth] = new Date(period.monthFirst);

      if (ixPMax != null) out[ixPMax] = period.maxWorkload;
      if (ixPMs  != null) out[ixPMs]  = milestone;
      if (ixPSs  != null) out[ixPSs]  = new Date(period.serviceStart);
      if (ixPSe  != null) out[ixPSe]  = new Date(period.serviceEnd);

      outRows.push(out);

      existingKeySet.add(key);
      createdAny = true;
    }

    if (createdAny) {
      const sheetRow = CFG.SRC_HEADER_ROW + 1 + i;
      statusUpdates.push({ sheetRow, newStatus: CFG.STATUS_DONE });
    }
  }

  if (!outRows.length) return;

  // -----------------------
  // TOP INSERT
  // -----------------------
  const insertAt = CFG.OUT_HEADER_ROW + 1; // row 3
  shOut.insertRowsBefore(insertAt, outRows.length);
  shOut.getRange(insertAt, 1, outRows.length, pLastCol).setValues(outRows);

  // -----------------------
  // Update Master statuses
  // -----------------------
  const statusCol = ixMStatus + 1;
  for (let i = 0; i < statusUpdates.length; i++) {
    const u = statusUpdates[i];
    shSrc.getRange(u.sheetRow, statusCol).setValue(u.newStatus);
  }

  SpreadsheetApp.flush();

  // Recalculate Payroll fields immediately
  payroll_updateWorkloadsFromAbsences();
}

/* ===========================
 * A-SPECIFIC HELPERS
 * =========================== */

function _maxDate_(a, b) {
  return (a.getTime() >= b.getTime()) ? a : b;
}

function _minDate_(a, b) {
  return (a.getTime() <= b.getTime()) ? a : b;
}

function _buildPayrollPeriods_(start, end, holidayTs, shortFirstMonthMaxWorkdays) {
  const s = _toMidnight_(start);
  const e = _toMidnight_(end);
  if (e.getTime() < s.getTime()) return [];

  const periods = [];

  const firstMonth = new Date(s.getFullYear(), s.getMonth(), 1);
  firstMonth.setHours(0, 0, 0, 0);

  const firstMonthEnd = _endOfMonth_(firstMonth);
  const nextMonth = _addMonths_(firstMonth, 1);

  const shortFirstWorkdays =
    s.getDate() !== 1
      ? _workdaysFast_(s, firstMonthEnd, holidayTs)
      : null;

  const canMergeFirst =
    s.getDate() !== 1 &&
    shortFirstWorkdays <= shortFirstMonthMaxWorkdays &&
    nextMonth.getTime() <= e.getTime();

  let curMonth;

  if (canMergeFirst) {
    const mergedEnd = _minDate_(e, _endOfMonth_(nextMonth));

    periods.push({
      monthFirst: new Date(nextMonth),   // display / key month = second month
      serviceStart: new Date(s),
      serviceEnd: new Date(mergedEnd),
      maxWorkload: _workdaysFast_(s, mergedEnd, holidayTs)
    });

    curMonth = _addMonths_(firstMonth, 2);
  } else {
    curMonth = new Date(firstMonth);
  }

  while (curMonth.getTime() <= e.getTime()) {
    const monthEnd = _endOfMonth_(curMonth);
    const serviceStart = _maxDate_(s, curMonth);
    const serviceEnd = _minDate_(e, monthEnd);

    if (serviceStart.getTime() <= serviceEnd.getTime()) {
      periods.push({
        monthFirst: new Date(curMonth),
        serviceStart: new Date(serviceStart),
        serviceEnd: new Date(serviceEnd),
        maxWorkload: _workdaysFast_(serviceStart, serviceEnd, holidayTs)
      });
    }

    curMonth = _addMonths_(curMonth, 1);
  }

  return periods;
}

function _addMonths_(d, n) {
  const x = _coerceDate_(d);
  return new Date(x.getFullYear(), x.getMonth() + n, 1);
}

function _workdaysFast_(start, end, holidayTs) {
  var s = _coerceDate_(start);
  var e = _coerceDate_(end);
  if (!s || !e) return 0;
  if (e.getTime() < s.getTime()) return 0;

  var holidaySet = {};
  for (var i = 0; i < (holidayTs || []).length; i++) {
    var hd = _coerceDate_(new Date(holidayTs[i]));
    if (!hd) continue;
    holidaySet[_ymdKey_(hd)] = true;
  }

  var count = 0;
  var cur = new Date(s.getFullYear(), s.getMonth(), s.getDate());

  while (cur.getTime() <= e.getTime()) {
    var dow = cur.getDay(); // 0=Sun ... 6=Sat
    var key = _ymdKey_(cur);

    if (dow >= 1 && dow <= 5 && !holidaySet[key]) {
      count++;
    }

    cur.setDate(cur.getDate() + 1);
  }

  return count;
}

function _ymdKey_(d) {
  var x = _coerceDate_(d);
  if (!x) return '';
  var y = x.getFullYear();
  var m = x.getMonth() + 1;
  var day = x.getDate();
  return y + '-' +
    (m < 10 ? '0' + m : m) + '-' +
    (day < 10 ? '0' + day : day);
}