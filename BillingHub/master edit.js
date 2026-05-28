/**
 * Master checkbox audit: logs who/when changed Finance Check or Synergy Check.
 *
 * Requirements:
 * - Sheet name: "Master"
 * - Header row: 2 (adjust if needed)
 * - Columns must exist (or will be auto-created at the end):
 *     Finance Check
 *     Finance Check Updated at
 *     Finance Check Updated by
 *     Synergy Check
 *     Synergy Check Updated at
 *     Synergy Check Updated by
 */
function master_auditCheckboxes_onEdit(e) {
  const CFG = {
    SHEET_NAME: 'Payroll',
    HEADER_ROW: 2,

    // checkbox headers
    FIN_CHECK: 'Finance Check',
    SYN_CHECK: 'Synergy Check',

    // audit headers
    FIN_AT: 'Finance Check Updated at',
    FIN_BY: 'Finance Check Updated by',
    SYN_AT: 'Synergy Check Updated at',
    SYN_BY: 'Synergy Check Updated by',
  };

  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row <= CFG.HEADER_ROW) return; // ignore header edits

  // Only single-cell edits
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  const lastCol = sh.getLastColumn();

  // Build header index
  let headers = sh.getRange(CFG.HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  let idx = Object.fromEntries(headers.map((h, i) => [h, i + 1])); // 1-based columns

  // Ensure audit columns exist (append if missing)
  const ensure = [CFG.FIN_AT, CFG.FIN_BY, CFG.SYN_AT, CFG.SYN_BY];
  let appended = false;
  for (const h of ensure) {
    if (!idx[h]) {
      headers.push(h);
      appended = true;
    }
  }
  if (appended) {
    sh.getRange(CFG.HEADER_ROW, 1, 1, headers.length).setValues([headers]);
    // refresh indices
    idx = Object.fromEntries(headers.map((h, i) => [h, i + 1]));
  }

  const colFin = idx[CFG.FIN_CHECK];
  const colSyn = idx[CFG.SYN_CHECK];

  // If edited col is neither checkbox, ignore
  if (col !== colFin && col !== colSyn) return;

  // Only act on actual value change
  const oldV = e.oldValue; // can be undefined
  const newV = e.value;    // can be undefined
  if (_sameCheckboxValue_(oldV, newV)) return;

  const now = new Date();
  const who = _getActorEmail_();

  if (col === colFin) {
    sh.getRange(row, idx[CFG.FIN_AT]).setValue(now);
    sh.getRange(row, idx[CFG.FIN_BY]).setValue(who);
  } else if (col === colSyn) {
    sh.getRange(row, idx[CFG.SYN_AT]).setValue(now);
    sh.getRange(row, idx[CFG.SYN_BY]).setValue(who);
  }
}

/** Helpers **/

function _getActorEmail_() {
  try {
    // Works best with installable triggers in Workspace
    return Session.getActiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}

function _sameCheckboxValue_(oldValue, newValue) {
  // e.value / e.oldValue may be "TRUE"/"FALSE", undefined, or actual booleans in some contexts
  const a = _toBoolLoose_(oldValue);
  const b = _toBoolLoose_(newValue);
  return a === b;
}

function _toBoolLoose_(v) {
  if (v === true) return true;
  if (v === false) return false;
  if (v == null) return null; // unknown / empty
  const s = String(v).trim().toLowerCase();
  if (s === 'true' || s === 'yes' || s === '1') return true;
  if (s === 'false' || s === 'no' || s === '0') return false;
  return null;
}
