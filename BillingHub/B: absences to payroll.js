/***************************************
 * Absences -> Payroll (recalc)
 *
 * LOGIC:
 * - VL always reduces Invoiced Scope in the payroll period.
 * - VL yearly limit is analytics only:
 *     Cumulative VL
 *     Excess VL
 * - SL still uses yearly remaining-limit logic:
 *     only unpaidSL reduces Invoiced Scope
 *
 * FACT:
 *   Planned Workload  = Max Workload - VL_period
 *   Unplanned Absence = SL_period
 *   Actual Workload   = Max Workload - VL_period - SL_period
 *
 * PAYMENT:
 *   Invoiced Scope = Max Workload - VL_period - unpaidSL_period
 *
 * AMOUNTS:
 *   For each service (Primary / Secondary):
 *
 *   base_row =
 *     ROUND(maxW * total / projectWorkload)
 *     - INT( ROUND(maxW * total / projectWorkload) * (maxW - invScope) / maxW )
 *
 *   adjusted_total =
 *     ROUND(total * SUM(invScope) / projectWorkload, 0)
 *
 *   Last non-empty milestone:
 *     adjusted_total - sum(previous rows)
 *
 * IMPORTANT:
 * - Payroll row may represent merged first period
 *   (short first month + next full month)
 * - Therefore absences are aggregated by Service Start..Service End period,
 *   not just by Payroll.Month
 ***************************************/
function payroll_updateWorkloadsFromAbsences() {
  var CFG = {
    PAYROLL_SHEET: 'Payroll',
    PAYROLL_HEADER_ROW: 2,

    ABS_SHEET: 'Absences',
    ABS_HEADER_ROW: 1,

    MASTER_SHEET: 'Master',
    MASTER_HEADER_ROW: 2,

    // Payroll main
    P_FORCE: 'Force ID',
    P_AGR: 'Agreement No',
    P_NAME: 'Name',
    P_MONTH: 'Month',
    P_SERVICE_START: 'Service Start',
    P_SERVICE_END: 'Service End',
    P_MAX: 'Max Workload',
    P_PLAN: 'Planned Workload',
    P_UNPL: 'Unplanned Absence',
    P_ACT: 'Actual Workload',
    P_INV_SCOPE: 'Invoiced Scope',
    P_PRIMARY_AMT: 'Primary Amount',
    P_SECONDARY_AMT: 'Secondary Amount',
    P_INV_AMT: 'Invoiced Amount',
    P_FIN_CHECK: 'Finance Check',

    // Control columns
    P_CUM_VL: 'Cumulative VL',
    P_EXCESS_VL: 'Excess VL',
    P_CUM_SL: 'Cumulative SL',
    P_EXCESS_SL: 'Excess SL',

    // Meta
    P_LAST_RECALC: 'Last Recalculated at',

    // Deltas
    D_PLAN: 'Δ Planned Workload',
    D_UNPL: 'Δ Unplanned Absence',
    D_ACT: 'Δ Actual Workload',
    D_INV_SCOPE: 'Δ Invoiced Scope',
    D_PRIMARY_AMT: 'Δ Primary Amount',
    D_SECONDARY_AMT: 'Δ Secondary Amount',
    D_INV_AMT: 'Δ Invoiced Amount',

    // Absences
    A_FORCE: 'Force ID',
    A_MONTH: 'Month',
    A_TYPE: 'Leave Type',
    A_QTY: 'Quantity',

    // Master
    M_FORCE: 'Force ID',
    M_AGR: 'Agreement No',
    M_NAME: 'Name',
    M_START: 'Project Start Date',
    M_END: 'Project End Date',
    M_VL_LIMIT: 'Effective Suspension',
    M_SL_LIMIT: 'Tolerance Threshold',
    M_PROJECT_WL: 'Project Workload',
    M_PRIMARY_SERVICE: 'Primary Service',
    M_SECONDARY_SERVICE: 'Secondary Service',

    ROUND_DAYS_STEP: 0.5,
    TS_TZ: 'Europe/Warsaw',
    TS_FMT: 'yyyy-MM-dd HH:mm:ss',

    P_VL_RISK_STATUS: 'VL Risk Status',
    P_VL_REMAINING: 'VL Remaining',
    P_VL_REQUIRED_MONTHLY: 'VL Required / Month',
    P_VL_HARD_CAPACITY: 'VL Hard Capacity',
    P_VL_RISK_NOTE: 'VL Risk Note',

    TAIL_LIMIT: 0.7,

    VL_ALERT_EMAIL_TO: _getScriptProperty_('VL_ALERT_EMAIL_TO'),
    VL_ALERT_EMAIL_CC: _getScriptProperty_('VL_ALERT_EMAIL_CC'),
    VL_ALERT_EMAIL_STATUSES: ['WARNING', 'CRITICAL'],
    VL_ALERT_SUBJECT_PREFIX: '[VL Risk]',
  };

  var ss = SpreadsheetApp.getActive();

  var shP = ss.getSheetByName(CFG.PAYROLL_SHEET);
  var shA = ss.getSheetByName(CFG.ABS_SHEET);
  var shM = ss.getSheetByName(CFG.MASTER_SHEET);

  if (!shP) throw new Error('Sheet "' + CFG.PAYROLL_SHEET + '" not found');
  if (!shA) throw new Error('Sheet "' + CFG.ABS_SHEET + '" not found');
  if (!shM) throw new Error('Sheet "' + CFG.MASTER_SHEET + '" not found');

  /* =======================
   * PAYROLL HEADERS + DATA
   * ======================= */
  var pLastRow = shP.getLastRow();
  var pLastCol = shP.getLastColumn();
  if (pLastRow <= CFG.PAYROLL_HEADER_ROW) return;

  var ensureHeaders = [
    CFG.P_CUM_VL,
    CFG.P_EXCESS_VL,
    CFG.P_CUM_SL,
    CFG.P_EXCESS_SL,
    CFG.P_LAST_RECALC,
    CFG.D_PLAN,
    CFG.D_UNPL,
    CFG.D_ACT,
    CFG.D_INV_SCOPE,
    CFG.D_PRIMARY_AMT,
    CFG.D_SECONDARY_AMT,
    CFG.D_INV_AMT,

    CFG.P_VL_RISK_STATUS,
    CFG.P_VL_REMAINING,
    CFG.P_VL_REQUIRED_MONTHLY,
    CFG.P_VL_HARD_CAPACITY,
    CFG.P_VL_RISK_NOTE,
  ];
  pLastCol = _ensureHeaders_(shP, CFG.PAYROLL_HEADER_ROW, pLastCol, ensureHeaders);

  var pHeaders = shP.getRange(CFG.PAYROLL_HEADER_ROW, 1, 1, pLastCol).getValues()[0];
  var pIdx = _buildIndex_(pHeaders);

  var ixPForce        = _reqHeaderIndex_(pIdx, CFG.P_FORCE, 'Payroll');
  var ixPAgr          = _reqHeaderIndex_(pIdx, CFG.P_AGR, 'Payroll');
  var ixPMonth        = _reqHeaderIndex_(pIdx, CFG.P_MONTH, 'Payroll');
  var ixPServiceStart = _reqHeaderIndex_(pIdx, CFG.P_SERVICE_START, 'Payroll');
  var ixPServiceEnd   = _reqHeaderIndex_(pIdx, CFG.P_SERVICE_END, 'Payroll');
  var ixPMax          = _reqHeaderIndex_(pIdx, CFG.P_MAX, 'Payroll');
  var ixPPlan         = _reqHeaderIndex_(pIdx, CFG.P_PLAN, 'Payroll');
  var ixPUnpl         = _reqHeaderIndex_(pIdx, CFG.P_UNPL, 'Payroll');
  var ixPAct          = _reqHeaderIndex_(pIdx, CFG.P_ACT, 'Payroll');
  var ixPInvScope     = _reqHeaderIndex_(pIdx, CFG.P_INV_SCOPE, 'Payroll');
  var ixPPrimary      = _reqHeaderIndex_(pIdx, CFG.P_PRIMARY_AMT, 'Payroll');
  var ixPSecondary    = _reqHeaderIndex_(pIdx, CFG.P_SECONDARY_AMT, 'Payroll');
  var ixPInvAmt       = _reqHeaderIndex_(pIdx, CFG.P_INV_AMT, 'Payroll');
  var ixPFin          = _reqHeaderIndex_(pIdx, CFG.P_FIN_CHECK, 'Payroll');

  var ixPCumVL        = _reqHeaderIndex_(pIdx, CFG.P_CUM_VL, 'Payroll');
  var ixPExcessVL     = _reqHeaderIndex_(pIdx, CFG.P_EXCESS_VL, 'Payroll');
  var ixPVLRiskStatus = _reqHeaderIndex_(pIdx, CFG.P_VL_RISK_STATUS, 'Payroll');
  var ixPVLRemaining  = _reqHeaderIndex_(pIdx, CFG.P_VL_REMAINING, 'Payroll');
  var ixPVLRequiredMonthly = _reqHeaderIndex_(pIdx, CFG.P_VL_REQUIRED_MONTHLY, 'Payroll');
  var ixPVLHardCapacity = _reqHeaderIndex_(pIdx, CFG.P_VL_HARD_CAPACITY, 'Payroll');
  var ixPVLRiskNote   = _reqHeaderIndex_(pIdx, CFG.P_VL_RISK_NOTE, 'Payroll');

  var ixPCumSL        = _reqHeaderIndex_(pIdx, CFG.P_CUM_SL, 'Payroll');
  var ixPExcessSL     = _reqHeaderIndex_(pIdx, CFG.P_EXCESS_SL, 'Payroll');

  var ixPLast         = _reqHeaderIndex_(pIdx, CFG.P_LAST_RECALC, 'Payroll');

  var ixDPlan         = _reqHeaderIndex_(pIdx, CFG.D_PLAN, 'Payroll');
  var ixDUnpl         = _reqHeaderIndex_(pIdx, CFG.D_UNPL, 'Payroll');
  var ixDAct          = _reqHeaderIndex_(pIdx, CFG.D_ACT, 'Payroll');
  var ixDInvScope     = _reqHeaderIndex_(pIdx, CFG.D_INV_SCOPE, 'Payroll');
  var ixDPrimary      = _reqHeaderIndex_(pIdx, CFG.D_PRIMARY_AMT, 'Payroll');
  var ixDSecondary    = _reqHeaderIndex_(pIdx, CFG.D_SECONDARY_AMT, 'Payroll');
  var ixDInvAmt       = _reqHeaderIndex_(pIdx, CFG.D_INV_AMT, 'Payroll');

  var ixPName = (pIdx[CFG.P_NAME] != null) ? pIdx[CFG.P_NAME] : null;

  var pData = shP.getRange(
    CFG.PAYROLL_HEADER_ROW + 1, 1,
    pLastRow - CFG.PAYROLL_HEADER_ROW, pLastCol
  ).getValues();
  if (!pData.length) return;

  /* =======================
   * ABSENCES
   * ======================= */
  var absCols = 6;
  
  var aHeaders = shA
  .getRange(CFG.ABS_HEADER_ROW, 1, 1, absCols)
  .getValues()[0]
  .map(function(h) {
    return String(h || '')
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  });
  var aIdx = _buildIndex_(aHeaders);

  var ixAForce = _reqHeaderIndex_(aIdx, CFG.A_FORCE, 'Absences');
  var ixAMonth = _reqHeaderIndex_(aIdx, CFG.A_MONTH, 'Absences');
  var ixAType  = _reqHeaderIndex_(aIdx, CFG.A_TYPE, 'Absences');
  var ixAQty   = _reqHeaderIndex_(aIdx, CFG.A_QTY, 'Absences');

  var aLastRow = _lastDataRowByColumnA_(shA, CFG.ABS_HEADER_ROW);
  var absRows = Math.max(0, aLastRow - CFG.ABS_HEADER_ROW);
  var absAE = absRows
    ? shA.getRange(CFG.ABS_HEADER_ROW + 1, 1, absRows, absCols).getValues()
    : [];

  // source structure: calendar month
  var sums = new Map(); // fid|monthKey -> {vl, sl}
  for (var i = 0; i < absAE.length; i++) {
    var rowA = absAE[i];

    var fidA = _normId_(rowA[ixAForce]);
    if (!fidA) continue;

    var dA = _coerceDate_(rowA[ixAMonth]);
    if (!dA) continue;

    var m0A = _monthFirstLocal_(dA);
    if (!m0A) continue;

    var typeA = String(rowA[ixAType] || '').trim().toUpperCase();
    if (!typeA) continue;

    var baseType = typeA.replace(/^H/i, ''); // HVL/HSL -> VL/SL
    var qtyA = _toNum_(rowA[ixAQty]);
    if (!qtyA) continue;

    var keyA = fidA + '|' + _monthKeyUTC_(m0A);
    if (!sums.has(keyA)) sums.set(keyA, { vl: 0, sl: 0 });

    var acc = sums.get(keyA);
    if (baseType === 'VL') acc.vl += qtyA;
    else if (baseType === 'SL') acc.sl += qtyA;
  }

  /* =======================
   * MASTER
   * ======================= */
  var mLastRow = shM.getLastRow();
  var mLastCol = shM.getLastColumn();
  if (mLastRow <= CFG.MASTER_HEADER_ROW) throw new Error('No data in "' + CFG.MASTER_SHEET + '"');

  var mHeaders = shM.getRange(CFG.MASTER_HEADER_ROW, 1, 1, mLastCol).getValues()[0];
  var mIdx = _buildIndex_(mHeaders);

  var ixMForce     = _reqHeaderIndex_(mIdx, CFG.M_FORCE, 'Master');
  var ixMName      = _reqHeaderIndex_(mIdx, CFG.M_NAME, 'Master');
  var ixMAgr       = _reqHeaderIndex_(mIdx, CFG.M_AGR, 'Master');
  var ixMStart     = _reqHeaderIndex_(mIdx, CFG.M_START, 'Master');
  var ixMEnd       = _reqHeaderIndex_(mIdx, CFG.M_END, 'Master');
  var ixMVLLim     = _reqHeaderIndex_(mIdx, CFG.M_VL_LIMIT, 'Master');
  var ixMSLLim     = _reqHeaderIndex_(mIdx, CFG.M_SL_LIMIT, 'Master');
  var ixMProjectWL = _reqHeaderIndex_(mIdx, CFG.M_PROJECT_WL, 'Master');
  var ixMPrimary   = _reqHeaderIndex_(mIdx, CFG.M_PRIMARY_SERVICE, 'Master');
  var ixMSecondary = _reqHeaderIndex_(mIdx, CFG.M_SECONDARY_SERVICE, 'Master');

  var mData = shM.getRange(
    CFG.MASTER_HEADER_ROW + 1, 1,
    mLastRow - CFG.MASTER_HEADER_ROW, mLastCol
  ).getValues();

  var contractsByForce = new Map();
  var paramsByKey = new Map();

  for (var j = 0; j < mData.length; j++) {
    var rowM = mData[j];

    var fidM = _normId_(rowM[ixMForce]);
    if (!fidM) continue;

    var nameM = String(rowM[ixMName] || '').trim();
    var agrM = String(rowM[ixMAgr] || '').trim();
    if (!agrM) continue;

    var startM = _coerceDate_(rowM[ixMStart]);
    var endM   = _coerceDate_(rowM[ixMEnd]);
    if (!startM || !endM) continue;

    var s0 = _toMidnight_(startM);
    var e0 = _toMidnight_(endM);
    if (e0.getTime() < s0.getTime()) continue;

    if (!contractsByForce.has(fidM)) contractsByForce.set(fidM, []);
    contractsByForce.get(fidM).push({
      name: nameM,
      agr: agrM,
      start: s0,
      end: e0
    });

    var prmKey = fidM + '||' + agrM;
    if (paramsByKey.has(prmKey)) {
      var prev = paramsByKey.get(prmKey);
      _stopValidation_(
        'Validation error: Duplicate contract key in Master.\n\n' +
        'Person: ' + (nameM || prev.name || '(unknown)') + '\n' +
        'Force ID: ' + fidM + '\n' +
        'Agreement No: ' + agrM + '\n\n' +
        'Action:\nFix Master or Agreements dates to avoid overlap\n(or use different Force ID per concurrent contract).\n\n' +
        'Payroll update stopped to prevent incorrect calculations.'
      );
    }

    paramsByKey.set(prmKey, {
      name: nameM,
      start: s0,
      end: e0,
      vlLimit: _toNum_(rowM[ixMVLLim]),
      slLimit: _toNum_(rowM[ixMSLLim]),
      projectWorkload: _toNum_(rowM[ixMProjectWL]),
      primaryService: _toNum_(rowM[ixMPrimary]),
      secondaryService: _toNum_(rowM[ixMSecondary])
    });
  }

  _validateNoOverlapContracts_(contractsByForce);
  _validateNoPayrollDups_(pData, ixPForce, ixPAgr, ixPMonth, ixPName);

  /* =======================
   * PAYROLL ROWS BY CONTRACT
   * ======================= */
  var rowsByContract = new Map();

  for (var r = 0; r < pData.length; r++) {
    var rowP = pData[r];

    var fidP = _normId_(rowP[ixPForce]);
    if (!fidP) continue;

    var agrP = String(rowP[ixPAgr] || '').trim();
    if (!agrP) continue;

    var contractKey = fidP + '||' + agrP;
    var prm = paramsByKey.get(contractKey);
    if (!prm) continue;

    var serviceStartP = _coerceDate_(rowP[ixPServiceStart]);
    var serviceEndP   = _coerceDate_(rowP[ixPServiceEnd]);
    if (!serviceStartP || !serviceEndP) continue;

    serviceStartP = _toMidnight_(serviceStartP);
    serviceEndP   = _toMidnight_(serviceEndP);

    if (serviceEndP.getTime() < prm.start.getTime()) continue;
    if (serviceStartP.getTime() > prm.end.getTime()) continue;

    var dP = _coerceDate_(rowP[ixPMonth]);
    if (!dP) dP = new Date(serviceEndP.getFullYear(), serviceEndP.getMonth(), 1);

    var m0P = _monthFirstLocal_(dP);
    if (!m0P) continue;

    var monthKeyP = _monthKeyUTC_(m0P);

    if (!rowsByContract.has(contractKey)) rowsByContract.set(contractKey, []);
    rowsByContract.get(contractKey).push({
      r: r,
      monthKey: monthKeyP,
      monthDate: m0P,
      serviceStart: serviceStartP,
      serviceEnd: serviceEndP
    });
  }

  if (!rowsByContract.size) return;

  /* =======================
   * COMPUTE + WRITE
   * ======================= */
  var nowStr = Utilities.formatDate(new Date(), CFG.TS_TZ, CFG.TS_FMT);
  var out = pData.map(function(row) { return row.slice(); });
  var touched = 0;

  for (var it = rowsByContract.entries(), step = it.next(); !step.done; step = it.next()) {
    var contractKey2 = step.value[0];
    var arr = step.value[1];
    var prm2 = paramsByKey.get(contractKey2);
    if (!prm2) continue;

    arr.sort(function(a, b) {
      var ta = a.serviceStart.getTime();
      var tb = b.serviceStart.getTime();
      if (ta !== tb) return ta - tb;

      var ea = a.serviceEnd.getTime();
      var eb = b.serviceEnd.getTime();
      if (ea !== eb) return ea - eb;

      return a.monthKey < b.monthKey ? -1 : (a.monthKey > b.monthKey ? 1 : 0);
    });

    var vlLimit = _toNum_(prm2.vlLimit);
    var slLimit = _toNum_(prm2.slLimit);
    var remSL = slLimit;

    var cumVL = 0;
    var cumSL = 0;

    /* =======================
     * PASS 1: facts
     * ======================= */
    var calcRows = [];

    for (var k = 0; k < arr.length; k++) {
      var rowIndex = arr[k].r;

      var rowOld = pData[rowIndex];
      var fidKey = _normId_(rowOld[ixPForce]);
      var maxW = _toNum_(rowOld[ixPMax]);

      var periodStart = arr[k].serviceStart;
      var periodEnd   = arr[k].serviceEnd;

      var agg = _sumAbsencesForPeriod_(sums, fidKey, periodStart, periodEnd);
      var vlPeriod = _toNum_(agg.vl);
      var slPeriod = _toNum_(agg.sl);

      cumVL += vlPeriod;
      cumSL += slPeriod;

      var excessVL = Math.max(0, cumVL - vlLimit);
      var excessSL = Math.max(0, cumSL - slLimit);

      var planned = maxW - vlPeriod;
      var unpl = slPeriod;
      var actual = planned - unpl;

      var unpaidSL = Math.max(0, slPeriod - remSL);
      var paidSL = slPeriod - unpaidSL;
      remSL = Math.max(0, remSL - paidSL);

      var invScope = Math.max(0, maxW - vlPeriod - unpaidSL);

      calcRows.push({
        rowIndex: rowIndex,
        monthKey: arr[k].monthKey,
        serviceStart: periodStart,
        serviceEnd: periodEnd,

        maxW: maxW,
        vlPeriod: vlPeriod,
        slPeriod: slPeriod,
        unpaidSL: unpaidSL,

        wPlan: _roundToStep_(planned, CFG.ROUND_DAYS_STEP),
        wUnpl: _roundToStep_(unpl, CFG.ROUND_DAYS_STEP),
        wAct:  _roundToStep_(actual, CFG.ROUND_DAYS_STEP),
        wIS:   _roundToStep_(invScope, CFG.ROUND_DAYS_STEP),

        wCumVL: _roundToStep_(cumVL, CFG.ROUND_DAYS_STEP),
        wExcessVL: _roundToStep_(excessVL, CFG.ROUND_DAYS_STEP),
        wCumSL: _roundToStep_(cumSL, CFG.ROUND_DAYS_STEP),
        wExcessSL: _roundToStep_(excessSL, CFG.ROUND_DAYS_STEP)
      });
    }

    /* =======================
     * PASS 1.1. - SYNTHETIC VL TAIL
     * ======================= */
    var usedVL = 0;
    for (var sv = 0; sv < calcRows.length; sv++) {
      usedVL += _toNum_(calcRows[sv].vlPeriod);
    }

    var remainingVL = Math.max(0, _toNum_(prm2.vlLimit) - usedVL);

    if (remainingVL > 0) {
      var syntheticCapacityByCycle = calcRows.map(function(row) {
        var rowOld = pData[row.rowIndex];
        var isClosed = _isChecked_(rowOld[ixPFin]);

        if (isClosed) return 0;

        var maxW = _toNum_(row.maxW);
        var realVL = _toNum_(row.vlPeriod);

        var exactAllowed = maxW * CFG.TAIL_LIMIT;
        var remainingExactCapacity = Math.max(0, exactAllowed - realVL);

        return remainingExactCapacity / CFG.TAIL_LIMIT;
      });

      var maxSyntheticVL = syntheticCapacityByCycle.reduce(function(sum, wdEquivalent) {
        return sum + Math.floor(_toNum_(wdEquivalent) * CFG.TAIL_LIMIT * 2) / 2;
      }, 0);

      var vlToDistribute = Math.min(remainingVL, maxSyntheticVL);
      vlToDistribute = Math.floor(vlToDistribute * 2) / 2;

      if (vlToDistribute > 0) {
        var syntheticVL = distributeVacationDaysByServiceCycles(
          syntheticCapacityByCycle,
          vlToDistribute,
          CFG.TAIL_LIMIT
        );

        for (var sd = 0; sd < calcRows.length; sd++) {
          var rowOldSynthetic = pData[calcRows[sd].rowIndex];
          if (_isChecked_(rowOldSynthetic[ixPFin])) continue;

          var syntheticDays = _toNum_(syntheticVL.distribution[sd]);
          if (!syntheticDays) continue;

          calcRows[sd].vlPeriod = _roundToStep_(
            _toNum_(calcRows[sd].vlPeriod) + syntheticDays,
            CFG.ROUND_DAYS_STEP
          );

          var newPlanned = _toNum_(calcRows[sd].maxW) - _toNum_(calcRows[sd].vlPeriod);
          var newActual = newPlanned - _toNum_(calcRows[sd].slPeriod);
          var newInvScope = Math.max(
            0,
            _toNum_(calcRows[sd].maxW) -
            _toNum_(calcRows[sd].vlPeriod) -
            _toNum_(calcRows[sd].unpaidSL)
          );

          calcRows[sd].wPlan = _roundToStep_(newPlanned, CFG.ROUND_DAYS_STEP);
          calcRows[sd].wAct  = _roundToStep_(newActual, CFG.ROUND_DAYS_STEP);
          calcRows[sd].wIS   = _roundToStep_(newInvScope, CFG.ROUND_DAYS_STEP);
        }
      }
    }

    /* =======================
    * PASS 1.2. - VL RISK EVALUATION
    * ======================= */
    var riskCycles = calcRows.map(function(row) {
      var rowOld = pData[row.rowIndex];

      return {
        maxW: _toNum_(row.maxW),
        vlUsed: _toNum_(row.vlPeriod),
        isClosed: _isChecked_(rowOld[ixPFin])
      };
    });

    var vlRisk = evaluateVacationRisk({
      vlLimit: _toNum_(prm2.vlLimit),
      limitPerCycle: CFG.TAIL_LIMIT,
      cycles: riskCycles
    });

    /* =======================
    * PASS 1.3. - VL RISK EMAIL ALERT
    * ======================= */
    if (CFG.VL_ALERT_EMAIL_STATUSES.indexOf(vlRisk.status) !== -1) {
      _sendVacationRiskEmailIfNeeded_({
        cfg: CFG,
        risk: vlRisk,
        contractKey: contractKey2,
        personName: prm2.name || '',
        agreementNo: contractKey2.split('||')[1] || '',
        forceId: contractKey2.split('||')[0] || ''
      });
    }

    /* =======================
     * PASS 2: amounts
     * ======================= */
    var totalPW = _toNum_(prm2.projectWorkload);
    var totalPrimary = _toNum_(prm2.primaryService);
    var totalSecondary = _toNum_(prm2.secondaryService);

    _applyServiceAmountsByFormula_(calcRows, totalPrimary, totalPW, 'wPrimary');
    _applyServiceAmountsByFormula_(calcRows, totalSecondary, totalPW, 'wSecondary');

    for (var z = 0; z < calcRows.length; z++) {
      calcRows[z].wPrimary = _roundUsdCents_(calcRows[z].wPrimary);
      calcRows[z].wSecondary = _roundUsdCents_(calcRows[z].wSecondary);
      calcRows[z].wIA = _roundUsdCents_(
        _toNum_(calcRows[z].wPrimary) + _toNum_(calcRows[z].wSecondary)
      );
    }

    /* =======================
     * PASS 3: write
     * ======================= */
    for (var q = 0; q < calcRows.length; q++) {
      var rowCalc2 = calcRows[q];
      var rowIndex2 = rowCalc2.rowIndex;
      var rowOld2 = pData[rowIndex2];
      var isClosed = _isChecked_(rowOld2[ixPFin]);

      var rowChanged = false;

      rowChanged = _setCell_(out, rowIndex2, ixPCumVL, rowCalc2.wCumVL) || rowChanged;
      rowChanged = _setCell_(out, rowIndex2, ixPExcessVL, rowCalc2.wExcessVL) || rowChanged;
      rowChanged = _setCell_(out, rowIndex2, ixPCumSL, rowCalc2.wCumSL) || rowChanged;
      rowChanged = _setCell_(out, rowIndex2, ixPExcessSL, rowCalc2.wExcessSL) || rowChanged;

      var isLastOpenRowForRisk = _isLastOpenCalcRow_(calcRows, q, pData, ixPFin);

      if (isLastOpenRowForRisk) {
        rowChanged = _setCell_(out, rowIndex2, ixPVLRiskStatus, vlRisk.status) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLRemaining, vlRisk.remainingVL) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLRequiredMonthly, vlRisk.requiredMonthlyVL) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLHardCapacity, vlRisk.hardCapacity) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLRiskNote, _buildVacationRiskNote_(vlRisk)) || rowChanged;
      } else {
        rowChanged = _setCell_(out, rowIndex2, ixPVLRiskStatus, '') || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLRemaining, '') || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLRequiredMonthly, '') || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLHardCapacity, '') || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPVLRiskNote, '') || rowChanged;
      }

      if (!isClosed) {
        rowChanged = _setCell_(out, rowIndex2, ixPPlan, rowCalc2.wPlan) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPUnpl, rowCalc2.wUnpl) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPAct, rowCalc2.wAct) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPInvScope, rowCalc2.wIS) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPPrimary, rowCalc2.wPrimary) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPSecondary, rowCalc2.wSecondary) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixPInvAmt, rowCalc2.wIA) || rowChanged;

        if (_hasMeaningfulDeltaRow_(rowOld2, [ixDPlan, ixDUnpl, ixDAct, ixDInvScope, ixDPrimary, ixDSecondary, ixDInvAmt])) {
          rowChanged = _setCell_(out, rowIndex2, ixDPlan, '') || rowChanged;
          rowChanged = _setCell_(out, rowIndex2, ixDUnpl, '') || rowChanged;
          rowChanged = _setCell_(out, rowIndex2, ixDAct, '') || rowChanged;
          rowChanged = _setCell_(out, rowIndex2, ixDInvScope, '') || rowChanged;
          rowChanged = _setCell_(out, rowIndex2, ixDPrimary, '') || rowChanged;
          rowChanged = _setCell_(out, rowIndex2, ixDSecondary, '') || rowChanged;
          rowChanged = _setCell_(out, rowIndex2, ixDInvAmt, '') || rowChanged;
        }

      } else {
        var oldPlan = _toNum_(rowOld2[ixPPlan]);
        var oldUnpl = _toNum_(rowOld2[ixPUnpl]);
        var oldAct  = _toNum_(rowOld2[ixPAct]);
        var oldIS   = _toNum_(rowOld2[ixPInvScope]);
        var oldPrimary = _toNum_(rowOld2[ixPPrimary]);
        var oldSecondary = _toNum_(rowOld2[ixPSecondary]);

        var dPlan = _fixNegZero_(_roundToStep_(rowCalc2.wPlan - oldPlan, CFG.ROUND_DAYS_STEP));
        var dUnpl = _fixNegZero_(_roundToStep_(rowCalc2.wUnpl - oldUnpl, CFG.ROUND_DAYS_STEP));
        var dAct  = _fixNegZero_(_roundToStep_(rowCalc2.wAct - oldAct, CFG.ROUND_DAYS_STEP));
        var dIS   = _fixNegZero_(_roundToStep_(rowCalc2.wIS - oldIS, CFG.ROUND_DAYS_STEP));

        var dPrimary = _fixNegZero_(_roundUsdCents_(rowCalc2.wPrimary - oldPrimary));
        var dSecondary = _fixNegZero_(_roundUsdCents_(rowCalc2.wSecondary - oldSecondary));
        var dIA = _fixNegZero_(_roundUsdCents_(dPrimary + dSecondary));

        rowChanged = _setCell_(out, rowIndex2, ixDPlan, dPlan) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixDUnpl, dUnpl) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixDAct, dAct) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixDInvScope, dIS) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixDPrimary, dPrimary) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixDSecondary, dSecondary) || rowChanged;
        rowChanged = _setCell_(out, rowIndex2, ixDInvAmt, dIA) || rowChanged;
      }

      if (rowChanged) {
        out[rowIndex2][ixPLast] = nowStr;
        touched++;
      }
    }
  }

  if (!touched) return;

  shP.getRange(CFG.PAYROLL_HEADER_ROW + 1, 1, out.length, pLastCol).setValues(out);

  try {
    payroll_updateIndividualTimesheetsFromPayroll();
  } catch (e) {
    Logger.log('payroll_updateIndividualTimesheetsFromPayroll FAILED: ' + (e && e.stack ? e.stack : e));
    throw e;
  }
}

/* ===========================
 * VALIDATIONS
 * =========================== */

function _validateNoOverlapContracts_(contractsByForce) {
  for (var it = contractsByForce.entries(), s = it.next(); !s.done; s = it.next()) {
    var fid = s.value[0];
    var arr = s.value[1] || [];
    if (arr.length <= 1) continue;

    arr.sort(function(a, b) {
      var ta = a.start.getTime();
      var tb = b.start.getTime();
      return ta < tb ? -1 : (ta > tb ? 1 : 0);
    });

    for (var i = 1; i < arr.length; i++) {
      var prev = arr[i - 1];
      var cur = arr[i];

      if (cur.start.getTime() <= prev.end.getTime()) {
        var personName = cur.name || prev.name || '(unknown)';
        var msg =
          'Validation error: Overlapping contracts detected in Master.\n\n' +
          'Person: ' + personName + '\n' +
          'Force ID: ' + fid + '\n\n' +
          'Contract A:\n' +
          '  Agreement No: ' + prev.agr + '\n' +
          '  Period: ' + _fmtYmd_(prev.start) + ' → ' + _fmtYmd_(prev.end) + '\n\n' +
          'Contract B:\n' +
          '  Agreement No: ' + cur.agr + '\n' +
          '  Period: ' + _fmtYmd_(cur.start) + ' → ' + _fmtYmd_(cur.end) + '\n\n' +
          'Action:\n' +
          'Fix Master or Agreements dates to avoid overlap\n' +
          '(or use different Force ID per concurrent contract).\n\n' +
          'Payroll update stopped to prevent incorrect calculations.';
        _stopValidation_(msg);
      }
    }
  }
}

function _validateNoPayrollDups_(pData, ixForce, ixAgr, ixMonth, ixName) {
  var seen = new Map();
  for (var r = 0; r < pData.length; r++) {
    var row = pData[r];
    var fid = _normId_(row[ixForce]);
    if (!fid) continue;

    var agr = String(row[ixAgr] || '').trim();
    if (!agr) continue;

    var d = _coerceDate_(row[ixMonth]);
    if (!d) continue;

    var m0 = _monthFirstLocal_(d);
    if (!m0) continue;

    var key = fid + '||' + agr + '||' + _monthKeyUTC_(m0);
    if (seen.has(key)) {
      var name = '(unknown)';
      if (ixName != null) name = String(row[ixName] || '').trim() || '(unknown)';

      var msg =
        'Validation error: Duplicate Payroll rows detected.\n\n' +
        'Person: ' + name + '\n' +
        'Force ID: ' + fid + '\n' +
        'Agreement No: ' + agr + '\n' +
        'Month: ' + _monthKeyUTC_(m0) + '\n\n' +
        'Action:\n' +
        'Remove duplicates so that (Force ID, Agreement No, Month) is unique.\n\n' +
        'Payroll update stopped to prevent incorrect calculations.';
      _stopValidation_(msg);
    }
    seen.set(key, r);
  }
}

function _stopValidation_(message) {
  throw new Error(message);
}

/* ===========================
 * HELPERS
 * =========================== */

function _ensureHeaders_(sh, headerRow, lastCol, headersToEnsure) {
  var hdr = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var set = {};
  for (var i = 0; i < hdr.length; i++) {
    var h = String(hdr[i] || '').trim();
    if (h) set[h] = true;
  }

  var col = lastCol;
  for (var j = 0; j < headersToEnsure.length; j++) {
    var need = headersToEnsure[j];
    if (!set[need]) {
      col += 1;
      sh.getRange(headerRow, col, 1, 1).setValue(need);
      set[need] = true;
    }
  }
  return col;
}

function _lastDataRowByColumnA_(sh, headerRow) {
  var maxRows = sh.getMaxRows();
  if (maxRows <= headerRow) return headerRow;
  var lastA = sh.getRange(maxRows, 1).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  return Math.max(lastA, headerRow);
}

function _isChecked_(v) {
  if (v === true) return true;
  if (v === false || v === '' || v == null) return false;
  var s = String(v).trim().toLowerCase();
  return (s === 'true' || s === 'yes' || s === '1');
}

function _setCell_(out, r, c, v) {
  var oldV = out[r][c];
  if (_sameCell_(oldV, v)) return false;
  out[r][c] = v;
  return true;
}

function _isMeaningfulDelta_(v) {
  if (v === '' || v == null) return false;
  if (typeof v === 'number') return !Object.is(v, 0) && !Object.is(v, -0);
  var s = String(v).trim();
  return s !== '' && s !== '0' && s !== '-0';
}

function _hasMeaningfulDeltaRow_(row, idxs) {
  for (var i = 0; i < idxs.length; i++) {
    if (_isMeaningfulDelta_(row[idxs[i]])) return true;
  }
  return false;
}

function _fmtYmd_(d) {
  return Utilities.formatDate(d, 'Etc/UTC', 'yyyy-MM-dd');
}

function _sumAbsencesForPeriod_(sums, fid, start, end) {
  var keys = _listMonthKeysBetweenInclusive_(start, end);
  var vl = 0;
  var sl = 0;

  for (var i = 0; i < keys.length; i++) {
    var key = fid + '|' + keys[i];
    var acc = sums.has(key) ? sums.get(key) : null;
    if (!acc) continue;
    vl += _toNum_(acc.vl);
    sl += _toNum_(acc.sl);
  }

  return { vl: vl, sl: sl };
}

function _listMonthKeysBetweenInclusive_(start, end) {
  var s = _monthFirstLocal_(start);
  var e = _monthFirstLocal_(end);
  var out = [];

  var cur = new Date(s.getFullYear(), s.getMonth(), 1);
  while (cur.getTime() <= e.getTime()) {
    out.push(_monthKeyUTC_(cur));
    cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
  }

  return out;
}

function _applyServiceAmountsByFormula_(calcRows, serviceTotal, projectWorkload, targetField) {
  serviceTotal = _toNum_(serviceTotal);
  projectWorkload = _toNum_(projectWorkload);

  if (!calcRows || !calcRows.length) return;

  for (var i = 0; i < calcRows.length; i++) {
    calcRows[i][targetField] = 0;
  }

  if (serviceTotal <= 0 || projectWorkload <= 0) return;

  var sumInvScope = 0;
  var lastRowIdx = -1;

  for (var i = 0; i < calcRows.length; i++) {
    var inv = _toNum_(calcRows[i].wIS);
    if (inv !== 0) lastRowIdx = i;
    sumInvScope += inv;
  }

  if (lastRowIdx === -1) return;

  var adjustedTotal = Math.round(serviceTotal * sumInvScope / projectWorkload);

  var runningBaseSum = 0;

  for (var i = 0; i <= lastRowIdx; i++) {
    var maxW = _toNum_(calcRows[i].maxW);
    var invScope = _toNum_(calcRows[i].wIS);

    var baseRow = 0;

    if (maxW > 0) {
      var roundedMonthBase = Math.round(maxW * serviceTotal / projectWorkload);
      var reduction = Math.floor(
        roundedMonthBase * (maxW - invScope) / maxW
      );
      baseRow = roundedMonthBase - reduction;
    }

    if (i < lastRowIdx) {
      calcRows[i][targetField] = baseRow;
      runningBaseSum += _toNum_(baseRow);
    } else {
      calcRows[i][targetField] = adjustedTotal - runningBaseSum;
    }
  }

  for (var i = lastRowIdx + 1; i < calcRows.length; i++) {
    calcRows[i][targetField] = 0;
  }
}

function _isLastOpenCalcRow_(calcRows, currentIndex, pData, ixPFin) {
  for (var i = calcRows.length - 1; i >= 0; i--) {
    var rowOld = pData[calcRows[i].rowIndex];
    if (!_isChecked_(rowOld[ixPFin])) {
      return i === currentIndex;
    }
  }

  return false;
}

function _buildVacationRiskNote_(risk) {
  if (!risk) return '';

  if (risk.status === 'OK') {
    return '';
  }

  return [
    'VL risk detected',
    'Status: ' + risk.status,
    'Remaining VL: ' + risk.remainingVL,
    'Required VL / month: ' + risk.requiredMonthlyVL,
    'Hard capacity: ' + risk.hardCapacity,
    'Used VL: ' + risk.usedVL,
    'VL limit: ' + risk.vlLimit
  ].join('\n');
}

function _sendVacationRiskEmailIfNeeded_(o) {
  var cfg = o.cfg;
  var risk = o.risk;

  if (!risk || !risk.status) return;
  if (!cfg.VL_ALERT_EMAIL_TO) return;

  var cache = CacheService.getScriptCache();

  var todayKey = Utilities.formatDate(
    new Date(),
    cfg.TS_TZ || 'Europe/Warsaw',
    'yyyy-MM-dd'
  );

  var cacheKey = [
    'VL_RISK_EMAIL',
    todayKey,
    o.forceId,
    o.agreementNo,
    risk.status
  ].join('|');

  // Prevent duplicate emails for the same person/agreement/status during one day
  if (cache.get(cacheKey)) return;

  var subject =
    (cfg.VL_ALERT_SUBJECT_PREFIX || '[VL Risk]') +
    ' ' +
    risk.status +
    ' | ' +
    (o.personName || o.forceId) +
    ' | ' +
    o.agreementNo;

  var body = [
    'VL risk detected.',
    '',
    'Person: ' + (o.personName || '(unknown)'),
    'Force ID: ' + o.forceId,
    'Agreement No: ' + o.agreementNo,
    '',
    'Status: ' + risk.status,
    'VL limit: ' + risk.vlLimit,
    'Used VL: ' + risk.usedVL,
    'Remaining VL: ' + risk.remainingVL,
    'Required VL / month: ' + risk.requiredMonthlyVL,
    'Hard capacity: ' + risk.hardCapacity,
    '',
    'Action:',
    risk.status === 'CRITICAL'
      ? 'Immediate review is required. Remaining VL cannot be fully distributed within open payroll cycles.'
      : 'Please start planning VL usage to avoid a critical shortage near the end of the contract.',
    '',
    'This alert was generated automatically from Payroll recalculation.'
  ].join('\n');

  MailApp.sendEmail({
    to: cfg.VL_ALERT_EMAIL_TO,
    cc: cfg.VL_ALERT_EMAIL_CC || '',
    subject: subject,
    body: body
  });

  cache.put(cacheKey, '1', 21600); // 6 hours
}

function _getScriptProperty_(key) {
  return String(PropertiesService.getScriptProperties().getProperty(key) || '').trim();
}
