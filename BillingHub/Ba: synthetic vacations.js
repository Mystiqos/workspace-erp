/**
 * Distribute only unplanned remaining vacation days by service cycles.
 *
 * Strategy:
 * - does NOT redistribute already planned/approved VL
 * - receives remaining available capacity as workingDays-equivalent array
 * - allocates from the latest cycle backwards
 * - supports 1 and 0.5 day steps
 * - never exceeds limit per cycle
 *
 * @param {number[]} workingDaysByCycle - Working-days-equivalent capacity per cycle.
 *   In Payroll tail logic this should already represent remaining capacity only.
 *   Closed cycles should be passed as 0.
 * @param {number} vacationDays - Remaining unplanned VL to distribute.
 * @param {number} limit - Max vacation usage per cycle, e.g. 0.7.
 * @returns {Object}
 */
function distributeVacationDaysByServiceCycles(workingDaysByCycle, vacationDays, limit) {
  limit = limit == null ? 0.7 : Number(limit);

  if (!Array.isArray(workingDaysByCycle) || !workingDaysByCycle.length) {
    throw new Error('workingDaysByCycle must be a non-empty array.');
  }

  if (!(limit > 0 && limit <= 1)) {
    throw new Error('limit must be in range (0, 1].');
  }

  vacationDays = Number(vacationDays);
  if (!isFinite(vacationDays) || vacationDays < 0) {
    throw new Error('vacationDays must be a number >= 0.');
  }

  if (Math.round(vacationDays * 2) !== vacationDays * 2) {
    throw new Error('vacationDays must use 0.5 day step.');
  }

  var n = workingDaysByCycle.length;
  var distribution = new Array(n).fill(0);
  var exactCapacities = [];
  var halfStepCapacities = [];

  for (var i = 0; i < n; i++) {
    var wd = Number(workingDaysByCycle[i]);

    if (!isFinite(wd) || wd < 0) {
      throw new Error('All workingDaysByCycle values must be numbers >= 0.');
    }

    var exact = wd * limit;
    var halfStep = Math.floor(exact * 2) / 2;

    exactCapacities.push(exact);
    halfStepCapacities.push(halfStep);
  }

  var totalCapacity = halfStepCapacities.reduce(function(sum, v) {
    return sum + v;
  }, 0);

  var daysToAllocate = Math.min(vacationDays, totalCapacity);
  daysToAllocate = Math.floor(daysToAllocate * 2) / 2;

  var remaining = daysToAllocate;
  var chosenCycleIndexes = [];

  for (var idx = n - 1; idx >= 0; idx--) {
    if (remaining <= 0) break;

    var capacity = halfStepCapacities[idx];
    if (capacity <= 0) continue;

    var allocated = Math.min(capacity, remaining);
    allocated = Math.floor(allocated * 2) / 2;

    if (allocated <= 0) continue;

    distribution[idx] = allocated;
    chosenCycleIndexes.unshift(idx);
    remaining = Math.round((remaining - allocated) * 2) / 2;
  }

  var totalDistributed = distribution.reduce(function(sum, v) {
    return sum + v;
  }, 0);

  return {
    distribution: distribution,
    chosenCycleIndexes: chosenCycleIndexes,
    totalDistributed: totalDistributed,
    totalRequested: vacationDays,
    undistributed: Math.round((vacationDays - totalDistributed) * 2) / 2,
    capacities: {
      exact: exactCapacities,
      halfStep: halfStepCapacities
    },
    strategy: 'TAIL_ONLY_GREEDY_FROM_END'
  };
}

/**
 * Test helper for Apps Script Logger.
 */
function test_distributeVacationDaysByServiceCycles() {
  const result = distributeVacationDaysByServiceCycles(
    [21, 19, 20, 18, 22],
    26.5,
    0.7
  );

  Logger.log(JSON.stringify(result, null, 2));
}

/* ===========================
 * Helpers
 * =========================== */

function _validateVacationDistributionInput_(workingDaysByCycle, vacationDays, limit) {
  if (!Array.isArray(workingDaysByCycle) || workingDaysByCycle.length === 0) {
    throw new Error('workingDaysByCycle must be a non-empty array.');
  }

  for (let i = 0; i < workingDaysByCycle.length; i++) {
    const v = Number(workingDaysByCycle[i]);

    if (!isFinite(v) || v < 0) {
      throw new Error(`Invalid workingDaysByCycle[${i}]: must be a number >= 0.`);
    }
  }

  const vacation = Number(vacationDays);

  if (!isFinite(vacation) || vacation < 0) {
    throw new Error('vacationDays must be a number >= 0.');
  }

  if (!_isHalfStep_(vacation)) {
    throw new Error('vacationDays must use 0.5 step only.');
  }

  if (!isFinite(limit) || limit <= 0 || limit > 1) {
    throw new Error('limit must be in range (0, 1].');
  }
}

function _chooseLatestCyclesForFullDays_(fullDayCapacities, fullDaysToDistribute, halfDayRemainder, exactCapacities) {
  var chosen = [];
  var capacitySum = 0;

  for (var i = fullDayCapacities.length - 1; i >= 0; i--) {
    var hasFullCapacity = fullDayCapacities[i] > 0;
    var hasHalfCapacity = exactCapacities[i] >= 0.5;

    if (!hasFullCapacity && !hasHalfCapacity) continue;

    chosen.unshift(i);
    capacitySum += fullDayCapacities[i];

    var fullOk = capacitySum >= fullDaysToDistribute;

    if (fullOk) {
      if (halfDayRemainder !== 0.5) break;

      var canPlaceHalf = false;
      for (var j = 0; j < chosen.length; j++) {
        var idx = chosen[j];
        if (exactCapacities[idx] - fullDayCapacities[idx] >= 0.5) {
          canPlaceHalf = true;
          break;
        }
      }

      if (canPlaceHalf) break;
    }
  }

  if (!chosen.length && (fullDaysToDistribute > 0 || halfDayRemainder > 0)) {
    throw new Error('Vacation distribution impossible: no suitable cycles found.');
  }

  return chosen;
}

function _chooseHalfDayTargetCycle_(chosenCycleIndexes, exactCapacities, fullDayCapacities, fullDaysToDistribute) {
  for (var i = chosenCycleIndexes.length - 1; i >= 0; i--) {
    var cycleIndex = chosenCycleIndexes[i];

    var fullCap = fullDayCapacities[cycleIndex];
    var exactCap = exactCapacities[cycleIndex];

    if (exactCap - fullCap >= 0.5) {
      return cycleIndex;
    }
  }

  return -1;
}

function _isHalfStep_(value) {
  return Math.abs(value * 2 - Math.round(value * 2)) < 1e-9;
}

function _roundToHalf_(value) {
  return Math.round(Number(value) * 2) / 2;
}