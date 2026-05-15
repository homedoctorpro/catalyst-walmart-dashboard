/**
 * Catalyst Pet — Retailers Sync (Google Apps Script)
 *
 * Backs the Retailers tab on the Walmart Sales Dashboard.
 * Paste this entire file into the Apps Script editor of a Google Sheet,
 * deploy as a Web App, and paste the URL into the dashboard.
 *
 * See SETUP.md for step-by-step instructions.
 */

const SHEET_NAME = 'Retailers';
const HEADERS = [
  'retailer_id',     // stable ID from the dashboard (do not edit)
  'retailer_name',   // human-readable name (mirrors the dashboard)
  'channel',         // channel name (mirrors the dashboard)
  'rep_firm',        // EDIT ME — PSE / Brian Schlager / Unassigned (or blank)
  'status',          // EDIT ME — in / pitched / target / declined (or blank)
  'next_steps',      // EDIT ME — free text
  'usw_override',    // EDIT ME — leave blank to use channel default
  'updated_at',      // auto-stamped on every write
];

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, HEADERS.length, 150);
  }
  return sh;
}

function readAll_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return {};
  const rng = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const out = {};
  for (let i = 0; i < rng.length; i++) {
    const r = rng[i];
    const id = String(r[0] || '').trim();
    if (!id) continue;
    const o = {};
    const rep = String(r[3] || '').trim();
    const status = String(r[4] || '').trim();
    const nextSteps = String(r[5] || '').trim();
    const uswRaw = r[6];
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

function findRow_(sh, retailerId) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === retailerId) return i + 2;
  }
  return -1;
}

function upsert_(retailerId, retailerName, channel, fields) {
  if (!retailerId) throw new Error('missing retailerId');
  const sh = getSheet_();
  let row = findRow_(sh, retailerId);
  if (row < 0) {
    row = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(row, 1).setValue(retailerId);
    sh.getRange(row, 2).setValue(retailerName || '');
    sh.getRange(row, 3).setValue(channel || '');
  } else {
    if (retailerName) sh.getRange(row, 2).setValue(retailerName);
    if (channel) sh.getRange(row, 3).setValue(channel);
  }
  // 1=id, 2=name, 3=channel, 4=rep, 5=status, 6=next, 7=usw, 8=updated
  const colMap = { repFirm: 4, status: 5, nextSteps: 6, usw: 7 };
  Object.keys(fields).forEach(function (k) {
    const col = colMap[k];
    if (!col) return;
    const v = fields[k];
    sh.getRange(row, col).setValue(v == null || v === '' ? '' : v);
  });
  sh.getRange(row, 8).setValue(new Date());
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = ((e && e.parameter && e.parameter.action) || 'get').toLowerCase();
    if (action === 'get')    return jsonOut_({ ok: true, data: readAll_() });
    if (action === 'health') return jsonOut_({ ok: true, status: 'ok' });
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
      upsert_(body.retailerId, body.retailerName, body.channel, body.fields || {});
      return jsonOut_({ ok: true });
    }
    if (action === 'bulkupsert') {
      (body.items || []).forEach(function (it) {
        upsert_(it.retailerId, it.retailerName, it.channel, it.fields || {});
      });
      return jsonOut_({ ok: true, count: (body.items || []).length });
    }
    if (action === 'delete') {
      const sh = getSheet_();
      const row = findRow_(sh, body.retailerId);
      if (row > 0) sh.deleteRow(row);
      return jsonOut_({ ok: true });
    }
    return jsonOut_({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.message || err) });
  }
}
