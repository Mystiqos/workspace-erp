/**
 * Evaluate vacation (VL) utilization risk for a contract.
 *
 * @param {Object} params
 * @param {number} params.vlLimit - Total VL allowed for contract
 * @param {Array<Object>} params.cycles - Payroll cycles (chronological)
 * @param {number} params.limitPerCycle - e.g. 0.7
 *
 * Each cycle object:
 * {
 *   maxW: number,        // max workload (working days)
 *   vlUsed: number,      // already allocated VL in this cycle
 *   isClosed: boolean    // Finance Check = true
 * }
 *
 * @returns {Object}
 */
function evaluateVacationRisk(params) {
  if (!params || typeof params !== 'object') {
    throw new Error('Invalid params');
  }

  const vlLimit = _num(params.vlLimit);
  const cycles = params.cycles || [];
  const limit = _num(params.limitPerCycle);

  if (!Array.isArray(cycles) || !cycles.length) {
    throw new Error('cycles must be non-empty array');
  }
  if (vlLimit < 0) {
    throw new Error('vlLimit must be >= 0');
  }
  if (!(limit > 0 && limit <= 1)) {
    throw new Error('limitPerCycle must be in (0,1]');
  }

  /* =========================
   * USED + REMAINING VL
   * ========================= */
  let usedVL = 0;
  for (let c of cycles) {
    usedVL += _num(c.vlUsed);
  }

  const remainingVL = Math.max(0, vlLimit - usedVL);

  /* =========================
   * OPEN CYCLES ONLY
   * ========================= */
  const openCycles = cycles.filter(c => !c.isClosed);

  const openMonths = openCycles.length;

  if (openMonths === 0) {
    return {
      status: remainingVL > 0 ? 'CRITICAL' : 'OK',
      reason: 'No open cycles',
      remainingVL,
      requiredMonthlyVL: 0,
      hardCapacity: 0
    };
  }

  /* =========================
   * REQUIRED RATE
   * ========================= */
  const requiredMonthlyVL = remainingVL / openMonths;

  /* =========================
   * HARD CAPACITY
   * ========================= */
  let hardCapacity = 0;

  for (let c of openCycles) {
    const maxW = _num(c.maxW);
    const vlUsed = _num(c.vlUsed);

    const exactCapacity = maxW * limit;
    const remainingCapacity = Math.max(0, exactCapacity - vlUsed);

    hardCapacity += remainingCapacity;
  }

  /* =========================
   * STATUS
   * ========================= */
  let status = 'OK';

  if (remainingVL > hardCapacity + 0.001) {
    status = 'CRITICAL';
  } else if (requiredMonthlyVL > 7) {
    status = 'WARNING';
  } else if (requiredMonthlyVL > 3) {
    status = 'EARLY_RISK';
  }

  /* =========================
   * DETAILS
   * ========================= */
  return {
    status,

    remainingVL,
    usedVL,
    vlLimit,

    openMonths,
    requiredMonthlyVL: _round2(requiredMonthlyVL),

    hardCapacity: _round2(hardCapacity),

    utilizationRatio: hardCapacity > 0
      ? _round2(remainingVL / hardCapacity)
      : 0
  };
}

/* =========================
 * HELPERS
 * ========================= */

function _num(v) {
  if (v == null || v === '') return 0;
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function _round2(v) {
  return Math.round(_num(v) * 100) / 100;
}