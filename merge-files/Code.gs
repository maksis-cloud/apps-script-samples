/**
 * Demo 3 — merge several spreadsheets into one master sheet + summary
 * Typical ask: "we have 6 separate files, we want one place to see everything".
 * Reads the same-named tab from each source file, maps columns by header name
 * (so column order may differ between files), adds a Source column,
 * drops exact duplicates and rebuilds a small summary by month and source.
 *
 * Setup: list file IDs in the "Sources" tab (column A = file ID, B = label),
 * run setup() once. Rebuild runs every morning at 06:00 and from the menu.
 */

const SOURCE_TAB = 'Bookings';
const MASTER_TAB = 'Master';
const SUMMARY_TAB = 'Summary';
const COLUMNS = ['Date', 'Guest', 'Room', 'Nights', 'Amount', 'Status'];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Master')
    .addItem('Rebuild now', 'rebuildMaster')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  [MASTER_TAB, SUMMARY_TAB, 'Sources'].forEach(n => { if (!ss.getSheetByName(n)) ss.insertSheet(n); });
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'rebuildMaster')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('rebuildMaster').timeBased().atHour(6).everyDays(1).create();
}

function rebuildMaster() {
  const ss = SpreadsheetApp.getActive();
  const sources = ss.getSheetByName('Sources').getDataRange().getValues()
    .filter(r => String(r[0]).trim());
  const rows = [];
  const problems = [];
  const seen = {};

  sources.forEach(([fileId, label]) => {
    try {
      const tab = SpreadsheetApp.openById(String(fileId).trim()).getSheetByName(SOURCE_TAB);
      if (!tab) throw new Error('no tab ' + SOURCE_TAB);
      const values = tab.getDataRange().getValues();
      const header = values[0].map(h => String(h).trim());
      const idx = COLUMNS.map(c => header.indexOf(c));
      const missing = COLUMNS.filter((c, i) => idx[i] < 0);
      if (missing.length) throw new Error('missing columns: ' + missing.join(', '));
      values.slice(1).forEach(v => {
        if (!v[idx[0]]) return; // skip empty lines
        const out = idx.map(i => v[i]);
        const key = out.join('|');
        if (seen[key]) return;
        seen[key] = true;
        rows.push(out.concat([label || fileId]));
      });
    } catch (err) {
      problems.push((label || fileId) + ': ' + err.message);
    }
  });

  rows.sort((a, b) => new Date(a[0]) - new Date(b[0]));
  const master = ss.getSheetByName(MASTER_TAB);
  master.clearContents();
  const header = COLUMNS.concat(['Source']);
  master.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length) master.getRange(2, 1, rows.length, header.length).setValues(rows);
  master.setFrozenRows(1);

  writeSummary_(ss.getSheetByName(SUMMARY_TAB), rows, problems);
}

/** Month x source totals: bookings, nights, amount. Errors listed underneath. */
function writeSummary_(sheet, rows, problems) {
  const tz = Session.getScriptTimeZone();
  const agg = {};
  rows.forEach(r => {
    const month = Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM');
    const k = month + '|' + r[6];
    if (!agg[k]) agg[k] = [month, r[6], 0, 0, 0];
    agg[k][2] += 1;
    agg[k][3] += Number(r[3]) || 0;
    agg[k][4] += Number(r[4]) || 0;
  });
  const table = Object.keys(agg).sort().map(k => agg[k]);
  sheet.clearContents();
  const head = [['Month', 'Source', 'Bookings', 'Nights', 'Amount']];
  sheet.getRange(1, 1, 1, 5).setValues(head).setFontWeight('bold');
  if (table.length) sheet.getRange(2, 1, table.length, 5).setValues(table);
  const at = table.length + 3;
  sheet.getRange(at, 1).setValue('Last rebuild: ' + new Date());
  if (problems.length) {
    sheet.getRange(at + 1, 1, problems.length, 1).setValues(problems.map(p => ['Problem: ' + p]));
  }
}
