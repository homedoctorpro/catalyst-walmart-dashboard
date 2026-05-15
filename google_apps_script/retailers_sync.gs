/**
 * Catalyst Pet — Retailers Sync (Google Apps Script)
 *
 * Backs the Retailers tab on the Walmart Sales Dashboard.
 * Paste this entire file into the Apps Script editor of a Google Sheet,
 * deploy as a Web App, and paste the URL into the dashboard.
 *
 * See SETUP.md for step-by-step instructions.
 *
 * Schema (columns 1..14):
 *   A retailer_id     stable ID from dashboard           (do not edit)
 *   B retailer_name   human-readable                      (gets overwritten on push)
 *   C channel         channel name                        (gets overwritten on push)
 *   D us_stores       store count                         (gets overwritten on push)
 *   E default_usw     channel default U/S/W               (gets overwritten on push)
 *   F usw_override    your override (blank = use default) EDIT ME
 *   G effective_usw   formula: usw_override if set, else default_usw
 *   H annual_units    formula: stores × effective_usw × 52
 *   I wholesale_opp   formula: annual_units × $10
 *   J retail_opp      formula: annual_units × $20
 *   K rep_firm        Unassigned / PSE / Brian Schlager   EDIT ME
 *   L status          in / pitched / target / declined    EDIT ME
 *   M next_steps      free text                            EDIT ME
 *   N updated_at      auto-stamped on every write
 */

const SHEET_NAME = 'Retailers';
const HEADERS = [
  'retailer_id', 'retailer_name', 'channel',
  'us_stores', 'default_usw', 'usw_override',
  'effective_usw', 'annual_units', 'wholesale_opp', 'retail_opp',
  'rep_firm', 'status', 'next_steps', 'updated_at',
];
const COL = {
  id: 1, name: 2, channel: 3,
  usStores: 4, defaultUsw: 5, uswOverride: 6,
  effectiveUsw: 7, annualUnits: 8, wholesaleOpp: 9, retailOpp: 10,
  repFirm: 11, status: 12, nextSteps: 13, updatedAt: 14,
};
const N_COLS = HEADERS.length;
const WHOLESALE_PRICE = 10;
const RETAIL_PRICE = 20;
const WEEKS_PER_YEAR = 52;

// ── Sheet management ────────────────────────────────────────────────────────

function formatHeader_(sh) {
  sh.getRange(1, 1, 1, N_COLS)
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function setColumnWidths_(sh) {
  const widths = [110, 220, 110, 80, 90, 100, 110, 110, 130, 130, 130, 100, 320, 160];
  for (let i = 0; i < widths.length; i++) sh.setColumnWidth(i + 1, widths[i]);
}

function setNumberFormats_(sh, lastDataRow) {
  if (lastDataRow < 2) return;
  const nRows = lastDataRow - 1;
  sh.getRange(2, COL.usStores,    nRows, 1).setNumberFormat('#,##0');
  sh.getRange(2, COL.defaultUsw,  nRows, 1).setNumberFormat('0.00');
  sh.getRange(2, COL.uswOverride, nRows, 1).setNumberFormat('0.00');
  sh.getRange(2, COL.effectiveUsw,nRows, 1).setNumberFormat('0.00');
  sh.getRange(2, COL.annualUnits, nRows, 1).setNumberFormat('#,##0');
  sh.getRange(2, COL.wholesaleOpp,nRows, 1).setNumberFormat('"$"#,##0');
  sh.getRange(2, COL.retailOpp,   nRows, 1).setNumberFormat('"$"#,##0');
  sh.getRange(2, COL.updatedAt,   nRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

function ensureSchema_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    formatHeader_(sh);
    setColumnWidths_(sh);
    return sh;
  }
  // Compare current row-1 headers; if they differ, wipe and reset.
  const lastCol = Math.max(sh.getLastColumn(), N_COLS);
  const current = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  let ok = current.length >= N_COLS;
  if (ok) {
    for (let i = 0; i < N_COLS; i++) {
      if (current[i] !== HEADERS[i]) { ok = false; break; }
    }
  }
  if (!ok) {
    sh.clear();
    sh.appendRow(HEADERS);
    formatHeader_(sh);
    setColumnWidths_(sh);
  }
  return sh;
}

// ── Read ────────────────────────────────────────────────────────────────────

function readAll_() {
  const sh = ensureSchema_();
  const last = sh.getLastRow();
  if (last < 2) return {};
  const rng = sh.getRange(2, 1, last - 1, N_COLS).getValues();
  const out = {};
  for (let i = 0; i < rng.length; i++) {
    const r = rng[i];
    const id = String(r[COL.id - 1] || '').trim();
    if (!id) continue;
    // skip the TOTAL row
    if (id === 'TOTAL') continue;
    const o = {};
    const rep = String(r[COL.repFirm - 1] || '').trim();
    const status = String(r[COL.status - 1] || '').trim();
    const nextSteps = String(r[COL.nextSteps - 1] || '').trim();
    const uswRaw = r[COL.uswOverride - 1];
    if (rep) o.repFirm = rep;
    if (status) o.status = status;
    if (nextSteps) o.nextSteps = nextSteps;
    if (uswRaw !== '' && uswRaw != null && !isNaN(Number(uswRaw))) {
      o.usw = Number(uswRaw);
    }
    if (Object.keys(o).length) out[id] = o;
  }
  return out;
}

// ── Write ───────────────────────────────────────────────────────────────────

function findRow_(sh, retailerId) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, COL.id, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === retailerId) return i + 2;
  }
  return -1;
}

function rowFormulas_(row) {
  // Formulas relative to the row index.
  return [
    '=IF(F' + row + '="",E' + row + ',F' + row + ')',
    '=D' + row + '*G' + row + '*' + WEEKS_PER_YEAR,
    '=H' + row + '*' + WHOLESALE_PRICE,
    '=H' + row + '*' + RETAIL_PRICE,
  ];
}

function writeStaticAndFormulas_(sh, row, item) {
  sh.getRange(row, COL.id).setValue(item.retailerId);
  if (item.retailerName != null) sh.getRange(row, COL.name).setValue(item.retailerName);
  if (item.channel != null)      sh.getRange(row, COL.channel).setValue(item.channel);
  if (item.usStores != null)     sh.getRange(row, COL.usStores).setValue(item.usStores);
  if (item.defaultUsw != null)   sh.getRange(row, COL.defaultUsw).setValue(item.defaultUsw);
  const usw = (item.fields && item.fields.usw != null && item.fields.usw !== '') ? item.fields.usw : '';
  sh.getRange(row, COL.uswOverride).setValue(usw);
  // Formulas are idempotent — write them every time
  const f = rowFormulas_(row);
  sh.getRange(row, COL.effectiveUsw, 1, 4).setFormulas([f]);
}

function writeEditable_(sh, row, fields) {
  if (!fields) fields = {};
  sh.getRange(row, COL.repFirm).setValue(fields.repFirm || '');
  sh.getRange(row, COL.status).setValue(fields.status || '');
  sh.getRange(row, COL.nextSteps).setValue(fields.nextSteps || '');
  sh.getRange(row, COL.updatedAt).setValue(new Date());
}

function upsert_(item) {
  if (!item || !item.retailerId) throw new Error('missing retailerId');
  const sh = ensureSchema_();
  let row = findRow_(sh, item.retailerId);
  if (row < 0) row = Math.max(sh.getLastRow() + 1, 2);
  writeStaticAndFormulas_(sh, row, item);
  writeEditable_(sh, row, item.fields);
  // Re-apply number formats around this row for new inserts
  setNumberFormats_(sh, sh.getLastRow());
}

function writeTotalsRow_(sh, lastDataRow) {
  const totalRow = lastDataRow + 1;
  // Label in column A
  sh.getRange(totalRow, COL.id).setValue('TOTAL');
  sh.getRange(totalRow, COL.usStores).setFormula(
    '=SUM(D2:D' + lastDataRow + ')');
  sh.getRange(totalRow, COL.annualUnits).setFormula(
    '=SUM(H2:H' + lastDataRow + ')');
  sh.getRange(totalRow, COL.wholesaleOpp).setFormula(
    '=SUM(I2:I' + lastDataRow + ')');
  sh.getRange(totalRow, COL.retailOpp).setFormula(
    '=SUM(J2:J' + lastDataRow + ')');
  sh.getRange(totalRow, 1, 1, N_COLS)
    .setFontWeight('bold')
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff');
  sh.getRange(totalRow, COL.usStores).setNumberFormat('#,##0');
  sh.getRange(totalRow, COL.annualUnits).setNumberFormat('#,##0');
  sh.getRange(totalRow, COL.wholesaleOpp).setNumberFormat('"$"#,##0');
  sh.getRange(totalRow, COL.retailOpp).setNumberFormat('"$"#,##0');
}

function bulkReplace_(items) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  // Full wipe + reseed
  sh.clear();
  sh.appendRow(HEADERS);
  formatHeader_(sh);
  setColumnWidths_(sh);

  if (!items || !items.length) return 0;

  // Write static columns (A–F) in one batch
  const staticRows = items.map(function (it) {
    return [
      it.retailerId || '',
      it.retailerName || '',
      it.channel || '',
      it.usStores != null ? it.usStores : '',
      it.defaultUsw != null ? it.defaultUsw : '',
      (it.fields && it.fields.usw != null && it.fields.usw !== '') ? it.fields.usw : '',
    ];
  });
  sh.getRange(2, 1, staticRows.length, 6).setValues(staticRows);

  // Write formula columns (G–J) in one batch
  const formulaRows = items.map(function (_, i) {
    return rowFormulas_(i + 2);
  });
  sh.getRange(2, COL.effectiveUsw, formulaRows.length, 4).setFormulas(formulaRows);

  // Write editable columns (K–N) in one batch
  const now = new Date();
  const editRows = items.map(function (it) {
    const f = it.fields || {};
    return [
      f.repFirm || '',
      f.status || '',
      f.nextSteps || '',
      now,
    ];
  });
  sh.getRange(2, COL.repFirm, editRows.length, 4).setValues(editRows);

  const lastDataRow = items.length + 1;
  setNumberFormats_(sh, lastDataRow);
  writeTotalsRow_(sh, lastDataRow);
  return items.length;
}

// ── HTTP entry points ───────────────────────────────────────────────────────

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = ((e && e.parameter && e.parameter.action) || 'get').toLowerCase();
    if (action === 'get')    return jsonOut_({ ok: true, data: readAll_() });
    if (action === 'health') return jsonOut_({ ok: true, status: 'ok', schema: HEADERS });
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '').toLowerCase();
    if (action === 'upsert') {
      upsert_(body.item || body);  // accept both flat and nested
      return jsonOut_({ ok: true });
    }
    if (action === 'bulkupsert') {
      // Same as upsert but for many items (no wipe)
      (body.items || []).forEach(function (it) { upsert_(it); });
      return jsonOut_({ ok: true, count: (body.items || []).length });
    }
    if (action === 'bulkreplace') {
      const count = bulkReplace_(body.items || []);
      return jsonOut_({ ok: true, count: count, data: readAll_() });
    }
    if (action === 'delete') {
      const sh = ensureSchema_();
      const row = findRow_(sh, body.retailerId);
      if (row > 0) sh.deleteRow(row);
      return jsonOut_({ ok: true });
    }
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}
