/**
 * One-off helper: tick the confirm box on the SF_Link tab.
 *
 * Paste this into the Sheet's Apps Script project as a second file
 * (+ → Script), then run `applyLinkPicks` from the function dropdown.
 * It only writes column C of SF_Link. Nothing in Salesforce changes until
 * you run `applySalesforceLinks` afterwards.
 *
 * Picks made 2026-09-21: corporate Account over store-level, preferring the
 * record that carries contacts. Retailers left out stay unticked (no usable
 * match in the org) — add a line here or tick the row by hand if one turns up.
 */

var LINK_PICKS = {
  // auto-ticked by suggestSalesforceLinks and kept
  'meijer':            '0014V000041wnPqQAI',
  'atwoods':           '0014V00003CPOUWQA5',
  'murdochs':          '0014V00003CPPJaQAP',
  'publix':            '001UK0000056a7aYAA',
  'sprouts':           '001UK000005mpYQYAY',
  'wegmans':           '001UK000005ZmuvYAC',
  'family-dollar':     '001UK000007GWcdYAG',
  'pet-food-experts':  '001UK00000AahK2YAJ',
  'phillips-pet':      '001UK000003EmXxYAK',
  'unfi':              '0014V00003CPPlnQAH',

  // picked from multiple candidates
  'walmart':           '0014V00003CPPnkQAH',
  'target':            '0014V00003yU2cfQAC',
  'petsmart':          '0014V00003CPSjzQAH',
  'petco':             '0014V00003CPSjmQAH',
  'hollywood-feed':    '0014V00003CPRg7QAH',
  'costco':            '0014V00003CPOhZQAX',
  'sams-club':         '0014V00003CNoDPQA1',
  'bjs':               '0014V00003qxyaNQAQ',
  'tractor-supply':    '0014V00003CPPkCQAX',
  'rural-king':        '0016g00000LS94zAAD',
  'bomgaars':          '0016g00000QPETBAA5',
  'fleet-farm':        '0014V00003CPPHlQAP',
  'farm-fleet':        '0014V00003CPQYCQA5',
  'cal-ranch':         '0014V00003CPQhcQAH',
  'big-r':             '0014V00003yTANVQA4',   // Account is flagged CREDIT HOLD
  'kroger':            '0014V00003CPPhbQAH',
  'albertsons':        '0014V00003CPORWQA5',
  'hyvee':             '0014V00003CPRjDQAX',
  'giant-eagle':       '0014V00003CPOuiQAH',
  'winco':             '001UK00000Ab1sKYAR',
  'home-depot':        '0014V00003CPP1gQAH',
  'lowes':             '0014V00003CPPCgQAP',
  'menards':           '0014V00003CPSGmQAP',
  'cvs':               '0014V00003CPSvTQAX',
  'dollar-general':    '001UK00000R26rTYAR',
  'bradley-caldwell':  '0014V00003CPOZGQA5',
};

function applyLinkPicks() {
  var sh = SpreadsheetApp.getActive().getSheetByName('SF_Link');
  if (!sh) throw new Error('No SF_Link tab — run suggestSalesforceLinks first');
  var last = sh.getLastRow();
  if (last < 2) throw new Error('SF_Link is empty — run suggestSalesforceLinks first');

  var rows = sh.getRange(2, 1, last - 1, 4).getValues();
  var flags = [];
  var ticked = {}, cleared = 0;

  rows.forEach(function (r) {
    var id = String(r[0] || '').trim();
    var acct = String(r[3] || '').trim();
    var want = LINK_PICKS[id] && acct && LINK_PICKS[id] === acct;
    if (want) ticked[id] = acct;
    if (!want && r[2] === true) cleared++;
    flags.push([!!want]);
  });

  sh.getRange(2, 3, flags.length, 1).setValues(flags);

  var missing = Object.keys(LINK_PICKS).filter(function (id) { return !ticked[id]; });
  console.log('Ticked ' + Object.keys(ticked).length + ' of ' + Object.keys(LINK_PICKS).length +
              ' picks · cleared ' + cleared + ' other tick(s)');
  if (missing.length) console.log('NOT FOUND in SF_Link (check the Account Id): ' + missing.join(', '));
  return Object.keys(ticked).length;
}
