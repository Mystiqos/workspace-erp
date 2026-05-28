function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Payroll Tools')
    .addItem('Reconcile Payroll vs Timesheets', 'payroll_reconcileAmountsWithTimesheets_v3')
    .addItem('Clear All Payroll Amount Notes', 'payroll_clearAllNotes')
    .addSeparator()
    .addItem('Create Payroll for Newly Signed', 'payroll_recalcNewRows')
    .addItem('Update Absences in Payroll and Timesheets', 'payroll_updateWorkloadsFromAbsences')
    .addToUi();
}


/* ===========================
 * GLOBAL HELPERS
 * =========================== */

function _indexByHeader_(hdrRow) {
  var m = {};
  for (var i = 0; i < hdrRow.length; i++) {
    m[String(hdrRow[i] || '').trim()] = i;
  }
  return m;
}

function _buildIndex_(headersRow) {
  var idx = {};
  for (var i = 0; i < headersRow.length; i++) {
    var h = String(headersRow[i] || '').trim();
    if (h) idx[h] = i;
  }
  return idx;
}

function _reqHeaderIndex_(idxMap, name, sheetLabel) {
  var ix = idxMap[name];
  if (ix == null) throw new Error(sheetLabel + ' header not found: "' + name + '"');
  return ix;
}

function _makeKey_(fid, agNo, monthKey) {
  return String(fid) + '|' + String(agNo) + '|' + String(monthKey);
}

function _makeKey3_(forceId, agreementNo, monthDate) {
  var fid = _normId_(forceId);
  var ag  = _normAgreement_(agreementNo);
  var d   = _coerceDate_(monthDate);
  if (!fid || !ag || !d) return '';
  var mk = _monthKeyUTC_(_monthFirstLocal_(d));
  return fid + '|' + ag + '|' + mk;
}

/**
 * Canonical Force ID normalization:
 * - trims
 * - normalizes unicode dashes to "-"
 * - collapses whitespace
 * - if purely digits: strips leading zeros
 * - else: uppercases, trims around hyphens, collapses multiple hyphens
 */
function _normId_(v) {
  if (v === null || v === undefined) return '';

  var s = String(v).trim();
  if (!s) return '';

  s = s
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, '-')
    .replace(/\s+/g, ' ');

  if (/^\d+$/.test(s)) {
    s = s.replace(/^0+/, '');
    return s === '' ? '0' : s;
  }

  s = s.toUpperCase()
    .replace(/\s*-\s*/g, '-')
    .replace(/-+/g, '-')
    .replace(/\s+/g, ' ');

  return s;
}

function _normAgreement_(v) {
  if (v === null || v === undefined) return '';

  var s = String(v).trim();
  if (!s) return '';

  s = s
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, '-')
    .replace(/\s+/g, ' ');

  return s;
}

function _coerceDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  if (typeof v === 'number') {
    // safer conversion for Google Sheets serial dates
    const base = new Date(Date.UTC(1899, 11, 30)); // Sheets epoch
    const d = new Date(base.getTime() + Math.round(v) * 86400000);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }

  const s = String(v || '').trim();
  if (!s) return null;

  const d = new Date(s);
  if (isNaN(d.getTime())) return null;

  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function _toMidnight_(d) {
  var x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function _monthFirstLocal_(d) {
  var x = _coerceDate_(d);
  if (!x) return null;
  var m = new Date(x.getFullYear(), x.getMonth(), 1);
  m.setHours(0, 0, 0, 0);
  return m;
}

function _endOfMonth_(d) {
  var x = new Date(d.getFullYear(), d.getMonth() + 1, 0);
  x.setHours(0, 0, 0, 0);
  return x;
}

function _monthKeyUTC_(d) {
  var x = _coerceDate_(d);
  if (!x) return '';
  var y = x.getFullYear();
  var m = x.getMonth() + 1;
  return y + '-' + (m < 10 ? '0' + m : m) + '-01';
}

function _toNum_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;

  var s = String(v).trim();
  if (!s) return 0;

  s = s.replace(/[^\d.,-]/g, '');

  var lastComma = s.lastIndexOf(',');
  var lastDot = s.lastIndexOf('.');

  if (lastComma !== -1 && lastDot !== -1) {
    if (lastDot > lastComma) {
      s = s.replace(/,/g, '');
    } else {
      s = s.replace(/\./g, '').replace(/,/g, '.');
    }
  } else if (lastComma !== -1) {
    s = s.replace(/,/g, '.');
  }

  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

function _roundToStep_(v, step) {
  var n = Number(v);
  if (!isFinite(n)) return 0;
  var s = Number(step) || 0;
  if (s <= 0) return n;
  return Math.round(n / s) * s;
}

function _roundUsdCents_(v) {
  var n = Number(v);
  if (!isFinite(n)) return 0;
  return Math.round(n * 100) / 100;
}

function _fixNegZero_(n) {
  return Object.is(n, -0) ? 0 : n;
}

function _sameCell_(a, b) {
  if (typeof a === 'number' && typeof b === 'number') return a === b;

  var A = (a === null || a === undefined) ? '' : a;
  var B = (b === null || b === undefined) ? '' : b;

  if (A instanceof Date && B instanceof Date) return A.getTime() === B.getTime();
  return String(A) === String(B);
}

function _same2D_(a, b) {
  if (!a || !b) return false;
  if (a.length !== b.length) return false;

  for (var r = 0; r < a.length; r++) {
    if (!a[r] || !b[r]) return false;
    if (a[r].length !== b[r].length) return false;

    for (var c = 0; c < a[r].length; c++) {
      var A = (a[r][c] == null) ? '' : a[r][c];
      var B = (b[r][c] == null) ? '' : b[r][c];

      if (A instanceof Date && B instanceof Date) {
        if (A.getTime() !== B.getTime()) return false;
      } else if (String(A) !== String(B)) {
        return false;
      }
    }
  }
  return true;
}

function _findTableHeaders_(sh, requiredHeaders, maxRows, maxCols) {
  var rows = Math.min(maxRows || 60, sh.getMaxRows());
  var cols = Math.min(maxCols || 26, sh.getMaxColumns());

  var grid = sh.getRange(1, 1, rows, cols).getValues();
  var want = new Set(requiredHeaders.map(function(x) { return String(x).trim(); }));

  for (var r = 0; r < grid.length; r++) {
    var row = grid[r].map(function(v) { return String(v || '').trim(); });
    var colByName = {};

    for (var c = 0; c < row.length; c++) {
      var h = row[c];
      if (want.has(h)) colByName[h] = c;
    }

    var ok = requiredHeaders.every(function(h) { return colByName[h] != null; });
    if (ok) return { headerRow: r + 1, colByName: colByName };
  }
  return null;
}

function _lastDataRowByColumn_(sh, col, headerRow) {
  var lastRow = sh.getLastRow();
  if (lastRow <= headerRow) return headerRow;

  var vals = sh.getRange(headerRow + 1, col, lastRow - headerRow, 1).getValues();

  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || '').trim() !== '') {
      return headerRow + 1 + i;
    }
  }

  return headerRow;
}

function _extractSpreadsheetId_(v) {
  if (v == null) return '';

  var s = String(v).trim();
  if (!s) return '';

  if (/^[a-zA-Z0-9-_]{20,}$/.test(s)) return s;

  var m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (m && m[1]) return m[1];

  return '';
}