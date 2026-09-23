/**
 * Catalyst Pet — Retailers Sync (Google Apps Script)
 *
 * Backs the Retailers tab on the Walmart Sales Dashboard.
 * Paste this entire file into the Apps Script editor of a Google Sheet,
 * deploy as a Web App, and paste the URL into the dashboard.
 *
 * See SETUP.md for step-by-step instructions.
 *
 * Schema (columns 1..24):
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
 *   K rep_firm        Unassigned / Brian Schlager / Internal / Jeff Day / PSE / StoR   EDIT ME
 *   L status          in / pitched / target / non-target / declined    EDIT ME
 *   M next_steps      free text                            EDIT ME
 *   N updated_at      auto-stamped on every write
 *   O next_review     date (yyyy-mm-dd)                    EDIT ME
 *   P priority        1-10 (1 = highest), blank = none     EDIT ME
 *   Q sf_account_id   linked Salesforce Account           (written by Salesforce sync)
 *   R sf_account_name Salesforce Account name             (written by Salesforce sync)
 *   S contacts_json   up to 5 Account contacts as JSON    (written by Salesforce sync)
 *   T sf_synced_at    last successful Salesforce sync     (written by Salesforce sync)
 *   U sf_opportunity_id     Catalyst litter Opportunity   (written by Salesforce sync)
 *   V sf_opportunity_stage  its current stage             (written by Salesforce sync)
 *   W needs_distributor  1 = buys through a distributor       EDIT ME
 *   X distributor_name   which distributor                    EDIT ME
 *
 * Sheets created with fewer columns are migrated in place by appending the
 * missing headers; rows are kept.
 *
 * Salesforce two-way sync (optional): see the "Salesforce" section at the
 * bottom of this file and SETUP.md.
 */

const SHEET_NAME = 'Retailers';
const HEADERS = [
  'retailer_id', 'retailer_name', 'channel',
  'us_stores', 'default_usw', 'usw_override',
  'effective_usw', 'annual_units', 'wholesale_opp', 'retail_opp',
  'rep_firm', 'status', 'next_steps', 'updated_at',
  'next_review', 'priority',
  'sf_account_id', 'sf_account_name', 'contacts_json', 'sf_synced_at',
  'sf_opportunity_id', 'sf_opportunity_stage',
  'needs_distributor', 'distributor_name',
];
const COL = {
  id: 1, name: 2, channel: 3,
  usStores: 4, defaultUsw: 5, uswOverride: 6,
  effectiveUsw: 7, annualUnits: 8, wholesaleOpp: 9, retailOpp: 10,
  repFirm: 11, status: 12, nextSteps: 13, updatedAt: 14,
  nextReview: 15, priority: 16,
  sfAccountId: 17, sfAccountName: 18, contactsJson: 19, sfSyncedAt: 20,
  sfOppId: 21, sfOppStage: 22,
  needsDistributor: 23, distributorName: 24,
};
const N_SF_COLS = 6;  // Q..V, owned by the Salesforce sync
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
  const widths = [110, 220, 110, 80, 90, 100, 110, 110, 130, 130, 130, 100, 320, 160, 110, 70,
                  150, 200, 300, 140, 190, 130, 120, 180];
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
  sh.getRange(2, COL.nextReview,  nRows, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, COL.sfSyncedAt,  nRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
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
  // Migrate older layouts (14 = through updated_at, 15 = next_review, 16 = priority,
  // 20 = the Account sync block, 22 = the Opportunity columns): append the
  // missing trailing headers instead of wiping data
  for (const n of [14, 15, 16, 20, 22]) {
    let isLegacy = true;
    for (let i = 0; i < N_COLS; i++) {
      const want = i < n ? HEADERS[i] : '';
      const have = current[i] == null ? '' : current[i];
      if (have !== want) { isLegacy = false; break; }
    }
    if (isLegacy) {
      sh.getRange(1, n + 1, 1, N_COLS - n).setValues([HEADERS.slice(n)]);
      formatHeader_(sh);
      setColumnWidths_(sh);
      return sh;
    }
  }
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
    const reviewRaw = r[COL.nextReview - 1];
    const nextReview = (reviewRaw instanceof Date)
      ? Utilities.formatDate(reviewRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(reviewRaw || '').trim();
    if (nextReview) o.nextReview = nextReview;
    const prioRaw = r[COL.priority - 1];
    const prio = Number(prioRaw);
    if (prioRaw !== '' && prioRaw != null && prio >= 1 && prio <= 10) o.priority = Math.round(prio);
    if (String(r[COL.needsDistributor - 1] || '').trim()) o.needsDistributor = '1';
    const distName = String(r[COL.distributorName - 1] || '').trim();
    if (distName) o.distributorName = distName;
    if (rep) o.repFirm = rep;
    if (status) o.status = status;
    if (nextSteps) o.nextSteps = nextSteps;
    if (uswRaw !== '' && uswRaw != null && !isNaN(Number(uswRaw))) {
      o.usw = Number(uswRaw);
    }
    const sfId = String(r[COL.sfAccountId - 1] || '').trim();
    if (sfId) {
      o.sfAccountId = sfId;
      o.sfAccountName = String(r[COL.sfAccountName - 1] || '').trim();
      try { o.contacts = JSON.parse(r[COL.contactsJson - 1] || '[]'); } catch (e) { o.contacts = []; }
      const oppId = String(r[COL.sfOppId - 1] || '').trim();
      if (oppId) {
        o.sfOppId = oppId;
        o.sfOppStage = String(r[COL.sfOppStage - 1] || '').trim();
      }
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
  sh.getRange(row, COL.nextReview).setValue(fields.nextReview || '');
  sh.getRange(row, COL.priority).setValue(fields.priority || '');
  sh.getRange(row, COL.needsDistributor).setValue(fields.needsDistributor ? '1' : '');
  sh.getRange(row, COL.distributorName).setValue(fields.distributorName || '');
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
  // Linked retailers push straight to Salesforce. A failure here is not fatal:
  // the next syncSalesforce() run sees the Sheet differs from the snapshot and retries.
  if (sfConfigured_() && sh.getRange(row, COL.sfAccountId).getValue()) {
    try { sfPushRow_(sh, row); } catch (err) { console.warn('SF push failed: ' + err); }
  }
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
  // Keep the Salesforce link + contacts across the wipe; the dashboard doesn't send them
  const keptSf = {};
  if (sh.getLastRow() >= 2 && sh.getLastColumn() >= COL.sfSyncedAt) {
    const ids = sh.getRange(2, COL.id, sh.getLastRow() - 1, 1).getValues();
    const sfv = sh.getRange(2, COL.sfAccountId, sh.getLastRow() - 1, N_SF_COLS).getValues();
    ids.forEach(function (r, i) { if (r[0] && sfv[i][0]) keptSf[String(r[0]).trim()] = sfv[i]; });
  }
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
  sh.getRange(2, COL.nextReview, items.length, 2).setValues(items.map(function (it) {
    const f = it.fields || {};
    return [f.nextReview || '', f.priority || ''];
  }));
  sh.getRange(2, COL.needsDistributor, items.length, 2).setValues(items.map(function (it) {
    const f = it.fields || {};
    return [f.needsDistributor ? '1' : '', f.distributorName || ''];
  }));
  sh.getRange(2, COL.sfAccountId, items.length, N_SF_COLS).setValues(items.map(function (it) {
    return keptSf[it.retailerId] || ['', '', '', ''];
  }));

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
  // One writer at a time: dashboard saves and the 15-minute Salesforce sync
  // both rewrite rows, and interleaving them loses edits.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
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
    if (action === 'sfsync') {
      const summary = syncSalesforce_();
      return jsonOut_({ ok: true, summary: summary, data: readAll_() });
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
  } finally {
    lock.releaseLock();
  }
}

// ── Salesforce ──────────────────────────────────────────────────────────────
//
// Two-way sync between this Sheet and Salesforce Accounts. A retailer is
// linked when its Account's Dashboard_Retailer_ID__c holds the retailer_id.
//
// Only the dashboard-owned custom fields below are ever written. Name, door
// counts and the GUMU__ ERP fields are read-only here so the sync can't fight
// the ERP connector.
//
// Credentials live in Project Settings → Script Properties:
//   SF_CLIENT_ID, SF_CLIENT_SECRET   from the Salesforce External Client App
//   SF_DOMAIN (optional)             defaults to SF_DEFAULT_DOMAIN
//
// Conflicts use a per-field snapshot of the last synced value (hidden sheet
// SF_Sync): whichever side moved away from the snapshot wins; if both moved,
// the later edit wins (Sheet updated_at vs Account LastModifiedDate).

const SF_DEFAULT_DOMAIN = 'https://energex1.my.salesforce.com';
const SF_API = 'v62.0';
const SF_LINK_FIELD = 'Dashboard_Retailer_ID__c';
const SF_FIELDS = {           // Sheet field → Account field
  status:     'Retailer_Status__c',
  repFirm:    'Retailer_Rep_Firm__c',
  nextSteps:  'Retailer_Next_Steps__c',
  nextReview: 'Retailer_Next_Review__c',
  priority:   'Retailer_Priority__c',
};
// Dashboard status codes ↔ Salesforce picklist labels
const SF_STATUS = {
  'in': 'Currently In', 'pitched': 'Pitched', 'target': 'Target',
  'non-target': 'Non-Target', 'declined': 'Declined',
};
// Only these statuses put a retailer's fields into Salesforce. Everything
// else stays in the dashboard and the Sheet: the Account keeps its link and
// whatever it already had, and we never blank a field we stopped syncing.
const SF_PUSH_STATUSES = { 'target': 1, 'pitched': 1 };
// ── Opportunities ───────────────────────────────────────────────────────────
// One Catalyst cat-litter Opportunity per retailer, created the first time a
// retailer reaches Target or Pitched. Tagged with the org's own picklists:
// Product_Category__c = Pet Litter, Product_Label__c = Catalyst. Never created
// twice, never created when the Account already carries a litter Opportunity
// (open or closed), and never closed automatically — a retailer dropping off
// the target list leaves its Opportunity for a human to close.
const SF_OPP_CATEGORY = 'Pet Litter';
const SF_OPP_LABEL = 'Catalyst';
const SF_OPP_TYPE = 'New Business';
const SF_OPP_FORECAST = 'Omitted';        // keeps sizing estimates out of forecast totals
const SF_OPP_STAGE = { 'target': 'Qualification', 'pitched': 'Proposal' };
// Stages we set ourselves; anything past these is a human's call and is left alone.
const SF_OPP_OURS = { 'Qualification': 1, 'Proposal': 1 };
const SF_OPP_DEFAULT_DAYS = 90;           // close date when the row has no next review
// Dashboard channel → the org's Sales_Channel__c. Unmapped channels stay blank.
const SF_OPP_CHANNEL = {
  'Mass': 'Big Box', 'Club': 'Big Box',
  'Pet Specialty': 'Pet Specialty', 'Grocery': 'Grocery',
};

const SF_MAX_CONTACTS = 5;

function sfRowEligible_(statusCode) {
  return !!SF_PUSH_STATUSES[String(statusCode || '').trim()];
}
const SF_SNAPSHOT_SHEET = 'SF_Sync';
const SF_LOG_SHEET = 'SF_Log';       // overwrite / conflict warnings, newest first
const SF_LOG_KEEP = 500;
const SF_LINK_SHEET = 'SF_Link';
// Aggregate rows that will never map to one Account
const SF_UNLINKABLE = ['other-grocery', 'indie-via-distributors'];

function sfProps_() { return PropertiesService.getScriptProperties(); }

function sfConfigured_() {
  const p = sfProps_();
  return !!(p.getProperty('SF_CLIENT_ID') && p.getProperty('SF_CLIENT_SECRET'));
}

function sfDomain_() {
  return (sfProps_().getProperty('SF_DOMAIN') || SF_DEFAULT_DOMAIN).replace(/\/+$/, '');
}

function sfToken_(forceNew) {
  const cache = CacheService.getScriptCache();
  if (!forceNew) {
    const hit = cache.get('sf_token');
    if (hit) return hit;
  }
  const p = sfProps_();
  const res = UrlFetchApp.fetch(sfDomain_() + '/services/oauth2/token', {
    method: 'post',
    payload: {
      grant_type: 'client_credentials',
      client_id: p.getProperty('SF_CLIENT_ID'),
      client_secret: p.getProperty('SF_CLIENT_SECRET'),
    },
    muteHttpExceptions: true,
  });
  const j = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !j.access_token) {
    throw new Error('Salesforce login failed: ' + (j.error_description || j.error || res.getResponseCode()));
  }
  cache.put('sf_token', j.access_token, 1800);  // sessions last longer; 30 min is safe
  return j.access_token;
}

function sfFetch_(method, path, body) {
  const url = /^https?:/.test(path) ? path : sfDomain_() + path;
  for (let attempt = 0; attempt < 2; attempt++) {
    const opts = {
      method: method,
      headers: { Authorization: 'Bearer ' + sfToken_(attempt > 0) },
      contentType: 'application/json',
      muteHttpExceptions: true,
    };
    if (body != null) opts.payload = JSON.stringify(body);
    const res = UrlFetchApp.fetch(url, opts);
    const code = res.getResponseCode();
    if (code === 401 && attempt === 0) continue;  // expired token, retry once
    const text = res.getContentText();
    if (code >= 300) throw new Error('Salesforce ' + method + ' ' + path + ' → ' + code + ': ' + text.slice(0, 300));
    return text ? JSON.parse(text) : null;
  }
}

function sfQuery_(soql) {
  let j = sfFetch_('get', '/services/data/' + SF_API + '/query?q=' + encodeURIComponent(soql));
  let rows = j.records || [];
  while (!j.done && j.nextRecordsUrl) {
    j = sfFetch_('get', j.nextRecordsUrl);
    rows = rows.concat(j.records || []);
  }
  return rows;
}

function sfSoqlStr_(s) { return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'"); }

// Normalize to comparable strings so '3', 3 and 3.0 (or a Date cell and
// '2026-10-01') don't register as edits.
function sfNorm_(field, v) {
  if (v == null) return '';
  if (field === 'nextReview') {
    if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return String(v).trim().slice(0, 10);
  }
  if (field === 'priority') {
    const n = Number(v);
    return (v !== '' && n >= 1 && n <= 10) ? String(Math.round(n)) : '';
  }
  return String(v).trim();
}

function sfFromAccount_(field, v) {
  if (field === 'status') {
    const s = String(v || '').trim();
    for (const code in SF_STATUS) if (SF_STATUS[code] === s) return code;
    return s.toLowerCase();
  }
  return sfNorm_(field, v);
}

function sfToAccount_(field, v) {
  if (v === '') return null;
  if (field === 'status') return SF_STATUS[v] || v;
  if (field === 'priority') return Number(v);
  return v;
}

function sfSheetVals_(rowVals) {
  const out = {};
  for (const f in SF_FIELDS) out[f] = sfNorm_(f, rowVals[COL[f] - 1]);
  return out;
}

// Every write that lands on top of a value someone else had in Salesforce
// gets a row here, newest first. Set the Script Property SF_ALERT_EMAIL to
// also get an email when one happens.
function sfLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SF_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SF_LOG_SHEET);
    sh.appendRow(['when', 'retailer_id', 'account', 'field', 'kept', 'overwrote', 'winner', 'source']);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 8, 130);
  }
  return sh;
}

function sfLog_(entries) {
  if (!entries || !entries.length) return;
  const sh = sfLogSheet_();
  const now = new Date();
  const rows = entries.map(function (e) {
    return [now, e.id, e.account || '', e.field,
            e.kept === '' ? '(blank)' : e.kept,
            e.overwrote === '' ? '(blank)' : e.overwrote,
            e.winner, e.source];
  });
  sh.insertRowsAfter(1, rows.length);
  sh.getRange(2, 1, rows.length, 8).setValues(rows);
  sh.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  const last = sh.getLastRow();
  if (last > SF_LOG_KEEP + 1) sh.deleteRows(SF_LOG_KEEP + 2, last - SF_LOG_KEEP - 1);

  const to = sfProps_().getProperty('SF_ALERT_EMAIL');
  if (!to) return;
  const body = entries.map(function (e) {
    return e.id + ' · ' + e.field + ': kept "' + e.kept + '", overwrote "' + e.overwrote +
           '" (' + e.winner + ' won, via ' + e.source + ')';
  }).join('\n');
  try {
    MailApp.sendEmail(to, 'Catalyst Retailers sync: ' + entries.length + ' overwrite' +
      (entries.length === 1 ? '' : 's'),
      body + '\n\nFull history: ' + SpreadsheetApp.getActive().getUrl() + ' → ' + SF_LOG_SHEET + ' tab');
  } catch (err) { console.warn('alert email failed: ' + err); }
}

function sfSnapshotSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SF_SNAPSHOT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SF_SNAPSHOT_SHEET);
    sh.appendRow(['retailer_id', 'snapshot_json']);
    sh.hideSheet();
  }
  return sh;
}

function sfReadSnapshots_() {
  const sh = sfSnapshotSheet_();
  const out = {};
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (!r[0]) return;
    try { out[String(r[0])] = JSON.parse(r[1] || '{}'); } catch (e) { /* treat as unsynced */ }
  });
  return out;
}

function sfWriteSnapshots_(snaps) {
  const sh = sfSnapshotSheet_();
  if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const rows = Object.keys(snaps).map(function (id) { return [id, JSON.stringify(snaps[id])]; });
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
}

// Push one row's dashboard fields to its Account right after a dashboard save.
function sfPushRow_(sh, row, source) {
  const vals = sh.getRange(row, 1, 1, N_COLS).getValues()[0];
  const accountId = String(vals[COL.sfAccountId - 1] || '').trim();
  if (!accountId) return;
  const id = String(vals[COL.id - 1]).trim();
  const cur = sfSheetVals_(vals);
  const snaps = sfReadSnapshots_();
  const snap = snaps[id];

  // Not Target/Pitched → nothing goes to Salesforce. If the row used to be
  // synced, say so once in the log so the stale Account values are visible.
  if (!sfRowEligible_(cur.status)) {
    if (snap) {
      delete snaps[id];
      sfWriteSnapshots_(snaps);
      sfLog_([{ id: id, account: String(vals[COL.sfAccountName - 1] || ''), field: '(left synced set)',
                kept: 'status ' + (cur.status || 'blank') + ' — dashboard only from here',
                overwrote: 'Salesforce keeps its last synced values',
                winner: 'none', source: source || 'dashboard edit' }]);
    }
    return 0;
  }

  // Read the Account first so we can warn about anything we are about to
  // replace that someone else changed in Salesforce since the last sync.
  const fieldNames = Object.keys(SF_FIELDS).map(function (f) { return SF_FIELDS[f]; });
  const found = sfQuery_('SELECT Id, Name, ' + fieldNames.join(', ') +
                         " FROM Account WHERE Id = '" + sfSoqlStr_(accountId) + "'");
  const a = found[0];
  const warnings = [];
  if (a) {
    for (const f in SF_FIELDS) {
      const fv = sfFromAccount_(f, a[SF_FIELDS[f]]);
      if (fv === '' || fv === cur[f]) continue;          // nothing lost
      if (snap && fv === snap[f]) continue;              // unchanged since last sync
      warnings.push({ id: id, account: a.Name || '', field: f, kept: cur[f],
                      overwrote: fv, winner: 'dashboard', source: source || 'dashboard edit' });
    }
  }

  const body = {};
  for (const f in SF_FIELDS) body[SF_FIELDS[f]] = sfToAccount_(f, cur[f]);
  sfFetch_('patch', '/services/data/' + SF_API + '/sobjects/Account/' + accountId, body);
  snaps[id] = cur;
  sfWriteSnapshots_(snaps);
  sh.getRange(row, COL.sfSyncedAt).setValue(new Date());
  sfLog_(warnings);
  return warnings.length;
}

// Time-driven entry point (see installSalesforceTrigger).
function syncSalesforce() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return syncSalesforce_(); } finally { lock.releaseLock(); }
}

function syncSalesforce_() {
  if (!sfConfigured_()) return { skipped: 'Salesforce credentials not set in Script Properties' };
  const sh = ensureSchema_();
  const last = sh.getLastRow();
  if (last < 2) return { skipped: 'Sheet is empty' };

  const selectFields = ['Id', 'Name', 'LastModifiedDate', SF_LINK_FIELD]
    .concat(Object.keys(SF_FIELDS).map(function (f) { return SF_FIELDS[f]; }));
  const accounts = sfQuery_(
    'SELECT ' + selectFields.join(', ') +
    ', (SELECT Name, Title, Email, Phone, MobilePhone FROM Contacts' +
    ' ORDER BY LastModifiedDate DESC LIMIT ' + SF_MAX_CONTACTS + ')' +
    ' FROM Account WHERE ' + SF_LINK_FIELD + ' != null');
  const byRetailer = {};
  accounts.forEach(function (a) { byRetailer[String(a[SF_LINK_FIELD]).trim()] = a; });

  const data = sh.getRange(2, 1, last - 1, N_COLS).getValues();
  const snaps = sfReadSnapshots_();
  const now = new Date();
  const patches = [];
  const warnings = [];
  const summary = { linked: 0, pulled: 0, pushed: 0, conflicts: 0, unlinked: 0, warned: 0, skipped: 0 };

  data.forEach(function (r) {
    const id = String(r[COL.id - 1] || '').trim();
    if (!id || id === 'TOTAL') return;
    const a = byRetailer[id];
    if (!a) {
      if (r[COL.sfAccountId - 1]) {   // link was removed in Salesforce
        for (let k = 0; k < N_SF_COLS; k++) r[COL.sfAccountId - 1 + k] = '';
        delete snaps[id];
        summary.unlinked++;
      }
      return;
    }
    summary.linked++;
    const sheet = sfSheetVals_(r);
    const snap = snaps[id];

    // Not Target/Pitched: no push, no pull. Account name and contacts still
    // refresh below so the dashboard hover card stays current.
    if (!sfRowEligible_(sheet.status)) {
      if (snap) {
        delete snaps[id];
        warnings.push({ id: id, account: a.Name || '', field: '(left synced set)',
                        kept: 'status ' + (sheet.status || 'blank') + ' — dashboard only from here',
                        overwrote: 'Salesforce keeps its last synced values',
                        winner: 'none', source: 'scheduled sync' });
      }
      summary.skipped++;
      r[COL.sfAccountId - 1] = a.Id;
      r[COL.sfAccountName - 1] = a.Name || '';
      r[COL.contactsJson - 1] = JSON.stringify(
        ((a.Contacts && a.Contacts.records) || []).map(function (c) {
          return { name: c.Name || '', title: c.Title || '', email: c.Email || '',
                   phone: c.Phone || c.MobilePhone || '' };
        }));
      r[COL.sfSyncedAt - 1] = now;
      return;
    }
    const sheetEditedAt = r[COL.updatedAt - 1] instanceof Date ? r[COL.updatedAt - 1].getTime() : 0;
    const sfEditedAt = new Date(a.LastModifiedDate).getTime();
    const merged = {};
    const push = {};
    let pulledAny = false;
    for (const f in SF_FIELDS) {
      const sv = sheet[f];
      const fv = sfFromAccount_(f, a[SF_FIELDS[f]]);
      let winner, why = '';
      if (sv === fv) winner = 'same';
      else if (!snap) {                                    // first link: keep whichever side has data, Sheet first
        winner = sv !== '' ? 'sheet' : 'sf';
        if (winner === 'sheet' && fv !== '') why = 'first link';
      }
      else if (fv === snap[f]) winner = 'sheet';
      else if (sv === snap[f]) winner = 'sf';
      else {
        summary.conflicts++;
        winner = sheetEditedAt >= sfEditedAt ? 'sheet' : 'sf';
        why = 'both sides changed, later edit won';
      }
      if (why) {
        warnings.push({ id: id, account: a.Name || '', field: f,
                        kept: winner === 'sheet' ? sv : fv,
                        overwrote: winner === 'sheet' ? fv : sv,
                        winner: winner === 'sheet' ? 'dashboard/sheet' : 'salesforce',
                        source: 'scheduled sync (' + why + ')' });
      }
      if (winner === 'sf') {
        merged[f] = fv;
        r[COL[f] - 1] = f === 'priority' && fv !== '' ? Number(fv) : fv;
        pulledAny = true;
      } else {
        merged[f] = sv;
        if (winner === 'sheet') push[SF_FIELDS[f]] = sfToAccount_(f, sv);
      }
    }
    if (pulledAny) { r[COL.updatedAt - 1] = now; summary.pulled++; }
    if (Object.keys(push).length) {
      push.attributes = { type: 'Account' };
      push.id = a.Id;
      patches.push({ id: id, record: push });
    }
    snaps[id] = merged;

    const contacts = ((a.Contacts && a.Contacts.records) || []).map(function (c) {
      return { name: c.Name || '', title: c.Title || '', email: c.Email || '',
               phone: c.Phone || c.MobilePhone || '' };
    });
    r[COL.sfAccountId - 1] = a.Id;
    r[COL.sfAccountName - 1] = a.Name || '';
    r[COL.contactsJson - 1] = JSON.stringify(contacts);
    r[COL.sfSyncedAt - 1] = now;
  });

  // Composite PATCH, 200 records per call. Failed records drop their snapshot
  // so the next run retries them instead of pulling the stale SF value back.
  for (let i = 0; i < patches.length; i += 200) {
    const chunk = patches.slice(i, i + 200);
    const res = sfFetch_('patch', '/services/data/' + SF_API + '/composite/sobjects',
      { allOrNone: false, records: chunk.map(function (p) { return p.record; }) });
    (res || []).forEach(function (rr, k) {
      if (rr.success) summary.pushed++;
      else {
        delete snaps[chunk[k].id];
        console.warn('SF push failed for ' + chunk[k].id + ': ' + JSON.stringify(rr.errors));
      }
    });
  }

  // Write back only the editable + Salesforce columns; G–J are formulas.
  sh.getRange(2, COL.repFirm, data.length, 4).setValues(data.map(function (r) {
    return r.slice(COL.repFirm - 1, COL.updatedAt);
  }));
  sh.getRange(2, COL.nextReview, data.length, 2 + N_SF_COLS).setValues(data.map(function (r) {
    return r.slice(COL.nextReview - 1, COL.sfOppStage);
  }));
  sfWriteSnapshots_(snaps);

  // Opportunities: create for Target/Pitched, keep our own stages in step.
  try {
    warnings.push.apply(warnings, sfSyncOpportunities_(data, byRetailer, summary) || []);
    sh.getRange(2, COL.sfOppId, data.length, 2).setValues(data.map(function (r) {
      return [r[COL.sfOppId - 1] || '', r[COL.sfOppStage - 1] || ''];
    }));
  } catch (err) {
    console.warn('opportunity sync failed: ' + err);
    warnings.push({ id: '(all)', account: '', field: '(opportunity sync failed)',
                    kept: String(err).slice(0, 200), overwrote: '', winner: 'none',
                    source: 'scheduled sync' });
  }

  summary.warned = warnings.length;
  sfLog_(warnings);
  return summary;
}

// ── Opportunities ──────────────────────────────────────────────────────────

function sfOppDate_(v) {
  const d = (v instanceof Date) ? v : (String(v || '').trim() ? new Date(String(v).trim()) : null);
  if (d && !isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fallback = new Date();
  fallback.setDate(fallback.getDate() + SF_OPP_DEFAULT_DAYS);
  return Utilities.formatDate(fallback, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// Runs at the end of every sync. `rows` is the Retailers grid, already merged.
function sfSyncOpportunities_(rows, byRetailer, summary) {
  const existing = {};
  sfQuery_('SELECT Id, Name, StageName, IsClosed, CloseDate, Account.' + SF_LINK_FIELD +
    ' FROM Opportunity WHERE Account.' + SF_LINK_FIELD + " != null" +
    " AND Product_Category__c INCLUDES ('" + sfSoqlStr_(SF_OPP_CATEGORY) + "')" +
    ' ORDER BY CreatedDate DESC').forEach(function (o) {
      const rid = String(o.Account[SF_LINK_FIELD] || '').trim();
      if (rid && !existing[rid]) existing[rid] = o;   // newest wins
    });

  const creates = [], stageMoves = [], notes = [];

  rows.forEach(function (r) {
    const id = String(r[COL.id - 1] || '').trim();
    const a = byRetailer[id];
    if (!id || id === 'TOTAL' || !a) return;
    const status = sfNorm_('status', r[COL.status - 1]);
    const opp = existing[id];

    if (opp) {                                  // remember it, keep the stage honest
      r[COL.sfOppId - 1] = opp.Id;
      r[COL.sfOppStage - 1] = opp.StageName + (opp.IsClosed ? '' : '');
      if (!opp.IsClosed && sfRowEligible_(status)) {
        const want = SF_OPP_STAGE[status];
        if (want && opp.StageName !== want && SF_OPP_OURS[opp.StageName]) {
          stageMoves.push({ attributes: { type: 'Opportunity' }, id: opp.Id, StageName: want });
          r[COL.sfOppStage - 1] = want;
        }
      }
      return;
    }

    if (!sfRowEligible_(status)) return;        // only Target/Pitched create one

    const amount = Number(r[COL.wholesaleOpp - 1]);
    const rec = {
      attributes: { type: 'Opportunity' },
      Name: 'Catalyst Cat Litter - ' + String(r[COL.name - 1] || id).trim() + ' ' +
            sfOppDate_(r[COL.nextReview - 1]).slice(0, 4),
      AccountId: a.Id,
      StageName: SF_OPP_STAGE[status],
      CloseDate: sfOppDate_(r[COL.nextReview - 1]),
      Type: SF_OPP_TYPE,
      ForecastCategoryName: SF_OPP_FORECAST,
      Product_Category__c: SF_OPP_CATEGORY,
      Product_Label__c: SF_OPP_LABEL,
    };
    if (isFinite(amount) && amount > 0) rec.Amount = amount;
    const channel = SF_OPP_CHANNEL[String(r[COL.channel - 1] || '').trim()];
    if (channel) rec.Sales_Channel__c = channel;
    const steps = sfNorm_('nextSteps', r[COL.nextSteps - 1]);
    if (steps) rec.Problem_Statement__c = steps;
    creates.push({ id: id, row: r, record: rec });
  });

  for (let i = 0; i < stageMoves.length; i += 200) {
    sfFetch_('patch', '/services/data/' + SF_API + '/composite/sobjects',
      { allOrNone: false, records: stageMoves.slice(i, i + 200) });
  }

  for (let i = 0; i < creates.length; i += 200) {
    const chunk = creates.slice(i, i + 200);
    const res = sfFetch_('post', '/services/data/' + SF_API + '/composite/sobjects',
      { allOrNone: false, records: chunk.map(function (c) { return c.record; }) });
    (res || []).forEach(function (rr, k) {
      const c = chunk[k];
      if (rr.success) {
        c.row[COL.sfOppId - 1] = rr.id;
        c.row[COL.sfOppStage - 1] = c.record.StageName;
        notes.push({ id: c.id, account: c.record.Name, field: '(opportunity created)',
                     kept: c.record.StageName + ' · closes ' + c.record.CloseDate +
                           (c.record.Amount ? ' · $' + c.record.Amount.toLocaleString() : ' · no amount'),
                     overwrote: '', winner: 'dashboard', source: 'opportunity sync' });
      } else {
        notes.push({ id: c.id, account: c.record.Name, field: '(opportunity FAILED)',
                     kept: JSON.stringify(rr.errors).slice(0, 200), overwrote: '',
                     winner: 'none', source: 'opportunity sync' });
      }
    });
  }

  summary.oppsCreated = creates.length;
  summary.oppsRestaged = stageMoves.length;
  return notes;
}

// Run by hand to see what the next sync would create, without writing.
function previewOpportunities() {
  if (!sfConfigured_()) throw new Error('Salesforce credentials not set in Script Properties');
  const sh = ensureSchema_();
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, N_COLS).getValues();
  const linked = {};
  sfQuery_('SELECT Id, Name, ' + SF_LINK_FIELD + ' FROM Account WHERE ' + SF_LINK_FIELD + ' != null')
    .forEach(function (a) { linked[String(a[SF_LINK_FIELD]).trim()] = a; });
  const has = {};
  sfQuery_('SELECT Id, Name, StageName, Account.' + SF_LINK_FIELD + ' FROM Opportunity WHERE Account.' +
    SF_LINK_FIELD + " != null AND Product_Category__c INCLUDES ('" + sfSoqlStr_(SF_OPP_CATEGORY) + "')")
    .forEach(function (o) { has[String(o.Account[SF_LINK_FIELD]).trim()] = o; });

  const lines = [];
  rows.forEach(function (r) {
    const id = String(r[COL.id - 1] || '').trim();
    if (!id || id === 'TOTAL') return;
    const status = sfNorm_('status', r[COL.status - 1]);
    if (!sfRowEligible_(status)) return;
    if (!linked[id]) { lines.push(id + ': ' + status + ' but no linked Account — skipped'); return; }
    if (has[id]) { lines.push(id + ': already has ' + has[id].Name + ' (' + has[id].StageName + ')'); return; }
    const amt = Number(r[COL.wholesaleOpp - 1]);
    lines.push('CREATE  ' + id + ' · ' + SF_OPP_STAGE[status] + ' · closes ' +
      sfOppDate_(r[COL.nextReview - 1]) + ' · ' + (amt > 0 ? '$' + amt.toLocaleString() : 'no amount') +
      ' · ' + linked[id].Name);
  });
  console.log(lines.length ? lines.join('\n') : 'No Target/Pitched retailers yet — nothing to create');
  return lines;
}

// ── Linking retailers to Accounts (run from the Apps Script editor) ─────────

// Step 1: lists candidate Accounts for every unlinked retailer on an SF_Link
// sheet. Tick "confirm" on the right Account for each retailer.
function suggestSalesforceLinks() {
  const sh = ensureSchema_();
  const last = sh.getLastRow();
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, N_COLS).getValues();
  const out = [['retailer_id', 'retailer_name', 'confirm', 'account_id', 'account_name',
                'billing_city', 'billing_state', 'account_type', 'contacts']];
  rows.forEach(function (r) {
    const id = String(r[COL.id - 1] || '').trim();
    if (!id || id === 'TOTAL' || r[COL.sfAccountId - 1] || SF_UNLINKABLE.indexOf(id) >= 0) return;
    const name = String(r[COL.name - 1] || '');
    const full = name.replace(/\(.*?\)/g, '').split('/')[0].trim();
    const word = full.split(/\s+/).filter(function (w) { return w.length >= 4; })
      .sort(function (x, y) { return y.length - x.length; })[0];
    let hits = [];
    [full, word].forEach(function (term) {
      if (hits.length || !term) return;
      hits = sfQuery_("SELECT Id, Name, BillingCity, BillingState, Account_Type__c," +
        " (SELECT Id FROM Contacts) FROM Account WHERE Name LIKE '%" + sfSoqlStr_(term) + "%'" +
        ' ORDER BY Name LIMIT 8');
    });
    if (!hits.length) out.push([id, name, false, '', '(no match: paste an Account Id here)', '', '', '', '']);
    hits.forEach(function (a) {
      out.push([id, name, hits.length === 1, a.Id, a.Name, a.BillingCity || '', a.BillingState || '',
                a.Account_Type__c || '', (a.Contacts && a.Contacts.totalSize) || 0]);
    });
  });
  const ss = SpreadsheetApp.getActive();
  let ls = ss.getSheetByName(SF_LINK_SHEET);
  if (ls) ls.clear(); else ls = ss.insertSheet(SF_LINK_SHEET);
  ls.getRange(1, 1, out.length, out[0].length).setValues(out);
  if (out.length > 1) ls.getRange(2, 3, out.length - 1, 1).insertCheckboxes();
  ls.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
  ls.setFrozenRows(1);
  return out.length - 1;
}

// Step 2: writes the retailer_id onto each confirmed Account, then syncs.
function applySalesforceLinks() {
  const ls = SpreadsheetApp.getActive().getSheetByName(SF_LINK_SHEET);
  if (!ls || ls.getLastRow() < 2) throw new Error('Run suggestSalesforceLinks first');
  const rows = ls.getRange(2, 1, ls.getLastRow() - 1, 4).getValues();
  const seen = {};
  const records = [];
  rows.forEach(function (r) {
    const id = String(r[0] || '').trim();
    const acct = String(r[3] || '').trim();
    if (r[2] !== true || !id || !acct) return;
    if (seen[id]) throw new Error(id + ' is confirmed on more than one Account');
    seen[id] = true;
    const rec = { attributes: { type: 'Account' }, id: acct };
    rec[SF_LINK_FIELD] = id;
    records.push(rec);
  });
  const errors = [];
  for (let i = 0; i < records.length; i += 200) {
    const res = sfFetch_('patch', '/services/data/' + SF_API + '/composite/sobjects',
      { allOrNone: false, records: records.slice(i, i + 200) });
    (res || []).forEach(function (rr, k) {
      if (!rr.success) errors.push(records[i + k][SF_LINK_FIELD] + ': ' + JSON.stringify(rr.errors));
    });
  }
  const summary = syncSalesforce();
  console.log('Linked ' + (records.length - errors.length) + ' of ' + records.length +
              (errors.length ? '. Errors: ' + errors.join('; ') : '') + ' · sync: ' + JSON.stringify(summary));
  return { linked: records.length - errors.length, errors: errors, sync: summary };
}

// Run once from the editor: syncs every 15 minutes.
function installSalesforceTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncSalesforce') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncSalesforce').timeBased().everyMinutes(15).create();
}

// Run once from the editor: pushes an edit typed straight into the Sheet to
// Salesforce within seconds, instead of waiting for the 15-minute sync.
function installSheetEditTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sfOnSheetEdit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sfOnSheetEdit').forSpreadsheet(ss).onEdit().create();
}

// Both triggers in one go.
function installAllTriggers() {
  installSalesforceTrigger();
  installSheetEditTrigger();
  console.log('Installed: 15-minute sync + instant push on Sheet edits');
}

// Installable onEdit handler. Fires only for edits a person makes; writes the
// script itself makes (a pull, a dashboard push) never re-trigger it.
function sfOnSheetEdit(e) {
  if (!e || !e.range || !sfConfigured_()) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_NAME) return;

  const editable = {};
  editable[COL.repFirm] = 1; editable[COL.status] = 1; editable[COL.nextSteps] = 1;
  editable[COL.nextReview] = 1; editable[COL.priority] = 1; editable[COL.uswOverride] = 1;
  editable[COL.needsDistributor] = 1; editable[COL.distributorName] = 1;

  const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
  let touched = false;
  for (let c = c1; c <= c2 && !touched; c++) if (editable[c]) touched = true;
  if (!touched) return;

  const r1 = Math.max(e.range.getRow(), 2), r2 = e.range.getLastRow();
  for (let row = r1; row <= r2; row++) {
    const id = String(sh.getRange(row, COL.id).getValue() || '').trim();
    if (!id || id === 'TOTAL') continue;
    sh.getRange(row, COL.updatedAt).setValue(new Date());
    try {
      sfPushRow_(sh, row, 'sheet edit');
    } catch (err) {
      console.warn('sheet-edit push failed for ' + id + ': ' + err);
      sfLog_([{ id: id, account: '', field: '(push failed)', kept: String(err).slice(0, 120),
                overwrote: '', winner: 'none', source: 'sheet edit' }]);
    }
  }
}

// Optional: blank the five dashboard fields on linked Accounts whose retailer
// is no longer Target/Pitched, so Salesforce doesn't keep a stale status.
// Run by hand — the sync never does this on its own. Logs every clear.
function clearUnsyncedFields() {
  if (!sfConfigured_()) throw new Error('Salesforce credentials not set in Script Properties');
  const sh = ensureSchema_();
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const fieldNames = Object.keys(SF_FIELDS).map(function (f) { return SF_FIELDS[f]; });
  const accounts = sfQuery_('SELECT Id, Name, ' + SF_LINK_FIELD + ', ' + fieldNames.join(', ') +
    ' FROM Account WHERE ' + SF_LINK_FIELD + ' != null');
  const byRetailer = {};
  accounts.forEach(function (a) { byRetailer[String(a[SF_LINK_FIELD]).trim()] = a; });

  const records = [], entries = [];
  sh.getRange(2, 1, last - 1, N_COLS).getValues().forEach(function (r) {
    const id = String(r[COL.id - 1] || '').trim();
    const a = byRetailer[id];
    if (!id || id === 'TOTAL' || !a) return;
    if (sfRowEligible_(sfNorm_('status', r[COL.status - 1]))) return;

    const rec = { attributes: { type: 'Account' }, id: a.Id };
    let any = false;
    for (const f in SF_FIELDS) {
      const fv = sfFromAccount_(f, a[SF_FIELDS[f]]);
      if (fv === '') continue;
      rec[SF_FIELDS[f]] = null;
      any = true;
      entries.push({ id: id, account: a.Name || '', field: f, kept: '(blank)',
                     overwrote: fv, winner: 'dashboard', source: 'clearUnsyncedFields' });
    }
    if (any) records.push(rec);
  });

  for (let i = 0; i < records.length; i += 200) {
    sfFetch_('patch', '/services/data/' + SF_API + '/composite/sobjects',
      { allOrNone: false, records: records.slice(i, i + 200) });
  }
  sfLog_(entries);
  console.log('Cleared ' + entries.length + ' field value(s) on ' + records.length + ' Account(s)');
  return entries.length;
}

// Read-only survey of how this org models Opportunities, so Opportunity
// creation can be designed against what is actually there. Writes nothing.
function describeOpportunitySetup() {
  if (!sfConfigured_()) throw new Error('Salesforce credentials not set in Script Properties');
  const out = [];

  const d = sfFetch_('get', '/services/data/' + SF_API + '/sobjects/Opportunity/describe');

  out.push('— Record types —');
  (d.recordTypeInfos || []).forEach(function (rt) {
    if (rt.available) out.push('  ' + rt.name + (rt.defaultRecordTypeMapping ? '  (default)' : ''));
  });

  out.push('— Required fields on create —');
  (d.fields || []).forEach(function (f) {
    if (f.createable && !f.nillable && !f.defaultedOnCreate) out.push('  ' + f.name + ' (' + f.type + ')');
  });

  ['StageName', 'Type', 'ForecastCategoryName', 'LeadSource'].forEach(function (name) {
    const f = (d.fields || []).filter(function (x) { return x.name === name; })[0];
    if (!f) return;
    out.push('— ' + name + ' values —');
    (f.picklistValues || []).forEach(function (v) {
      if (v.active) out.push('  ' + v.label + (v.defaultValue ? '  (default)' : ''));
    });
  });

  out.push('— Custom fields (non-GUMU) —');
  (d.fields || []).forEach(function (f) {
    if (!f.custom || /^GUMU__/.test(f.name)) return;
    out.push('  ' + f.name + ' (' + f.type + (f.picklistValues && f.picklistValues.length
      ? ': ' + f.picklistValues.filter(function (v) { return v.active; })
          .map(function (v) { return v.label; }).slice(0, 12).join(' | ')
      : '') + ')');
  });
  const gumu = (d.fields || []).filter(function (f) { return /^GUMU__/.test(f.name); });
  out.push('  [' + gumu.length + ' GUMU__ fields — never written]');

  out.push('— Products matching litter / catalyst —');
  const prods = sfQuery_("SELECT Id, Name, ProductCode, Family, IsActive FROM Product2 " +
    "WHERE Name LIKE '%litter%' OR Name LIKE '%Catalyst%' OR ProductCode LIKE '%CATALYST%' " +
    "OR Family LIKE '%itter%' ORDER BY Name LIMIT 50");
  if (!prods.length) out.push('  (none)');
  prods.forEach(function (pr) {
    out.push('  ' + pr.Name + ' · code ' + (pr.ProductCode || '—') + ' · family ' +
             (pr.Family || '—') + (pr.IsActive ? '' : ' · INACTIVE'));
  });

  out.push('— Product families in use —');
  const fams = sfQuery_('SELECT Family, COUNT(Id) n FROM Product2 WHERE Family != null GROUP BY Family ORDER BY COUNT(Id) DESC LIMIT 25');
  fams.forEach(function (f) { out.push('  ' + f.Family + ' · ' + f.n); });

  out.push('— Price books —');
  sfQuery_('SELECT Id, Name, IsActive, IsStandard FROM Pricebook2 ORDER BY Name LIMIT 25')
    .forEach(function (b) {
      out.push('  ' + b.Name + (b.IsStandard ? ' (standard)' : '') + (b.IsActive ? '' : ' INACTIVE'));
    });

  out.push('— Open Opportunities on the 36 linked Accounts —');
  const opps = sfQuery_('SELECT Id, Name, StageName, Amount, CloseDate, Account.Name, ' +
    'Account.' + SF_LINK_FIELD + ' FROM Opportunity WHERE IsClosed = false AND Account.' +
    SF_LINK_FIELD + ' != null ORDER BY Account.Name LIMIT 100');
  if (!opps.length) out.push('  (none)');
  opps.forEach(function (o) {
    out.push('  ' + o.Account.Name + ' · ' + o.Name + ' · ' + o.StageName + ' · ' +
             (o.Amount == null ? 'no amount' : o.Amount) + ' · closes ' + o.CloseDate);
  });

  console.log(out.join('\n'));
  return out.length;
}

// Dry run: logs what the next sync would write, without sending anything.
function previewSalesforceSync() {
  if (!sfConfigured_()) throw new Error('Salesforce credentials not set in Script Properties');
  const sh = ensureSchema_();
  const last = sh.getLastRow();
  if (last < 2) return 'Sheet is empty';

  const selectFields = ['Id', 'Name', 'LastModifiedDate', SF_LINK_FIELD]
    .concat(Object.keys(SF_FIELDS).map(function (f) { return SF_FIELDS[f]; }));
  const accounts = sfQuery_('SELECT ' + selectFields.join(', ') +
    ' FROM Account WHERE ' + SF_LINK_FIELD + ' != null');
  const byRetailer = {};
  accounts.forEach(function (a) { byRetailer[String(a[SF_LINK_FIELD]).trim()] = a; });

  const snaps = sfReadSnapshots_();
  const lines = [];
  sh.getRange(2, 1, last - 1, N_COLS).getValues().forEach(function (r) {
    const id = String(r[COL.id - 1] || '').trim();
    const a = byRetailer[id];
    if (!id || id === 'TOTAL' || !a) return;
    const sheet = sfSheetVals_(r);
    const snap = snaps[id];
    for (const f in SF_FIELDS) {
      const sv = sheet[f], fv = sfFromAccount_(f, a[SF_FIELDS[f]]);
      if (sv === fv) continue;
      let winner;
      if (!snap) winner = sv !== '' ? 'sheet' : 'sf';
      else if (fv === snap[f]) winner = 'sheet';
      else if (sv === snap[f]) winner = 'sf';
      else winner = 'CONFLICT';
      lines.push(id + ' · ' + f + ': sheet "' + sv + '" vs salesforce "' + fv + '" → ' + winner);
    }
  });
  console.log(lines.length ? lines.join('\n') : 'Nothing to write — both sides match');
  return lines;
}

// Run from the editor to check credentials.
function testSalesforceConnection() {
  const rows = sfQuery_('SELECT Id FROM Account WHERE ' + SF_LINK_FIELD + ' != null LIMIT 200');
  console.log('Connected to ' + sfDomain_() + ' · ' + rows.length + ' linked Accounts');
  return rows.length;
}
