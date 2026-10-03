/**
 * Demo 2 — API -> Google Sheets, every hour
 * Pulls Stripe charges (paged, 100 per call) into a sheet.
 * Existing rows are updated by ID, new rows are added, nothing is duplicated.
 * The API key lives in Script Properties, never in the code or the sheet.
 *
 * Setup:
 * 1. Project Settings -> Script Properties -> STRIPE_KEY = restricted read-only key
 * 2. Run setup() once.
 * The same pattern works for Shopify, HubSpot or any JSON API with paging.
 */

const SHEET_NAME = 'Stripe charges';
const HEADERS = ['id', 'created', 'amount', 'currency', 'status', 'customer_email', 'description', 'synced_at'];
const LOOKBACK_DAYS = 30;

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncCharges')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncCharges').timeBased().everyHours(1).create();
}

function syncCharges() {
  const key = PropertiesService.getScriptProperties().getProperty('STRIPE_KEY');
  if (!key) throw new Error('Set STRIPE_KEY in Script Properties');

  const since = Math.floor(Date.now() / 1000) - LOOKBACK_DAYS * 86400;
  const charges = fetchAll_(key, since);

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const existing = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues()
    : [];
  const rowById = {};
  existing.forEach((row, i) => { rowById[row[0]] = i; });

  const now = new Date();
  charges.forEach(c => {
    const row = [
      c.id,
      new Date(c.created * 1000),
      c.amount / 100,
      String(c.currency).toUpperCase(),
      c.status,
      (c.billing_details && c.billing_details.email) || c.receipt_email || '',
      c.description || '',
      now
    ];
    if (c.id in rowById) existing[rowById[c.id]] = row;
    else { rowById[c.id] = existing.length; existing.push(row); }
  });

  if (existing.length) {
    existing.sort((a, b) => b[1] - a[1]); // newest first
    sheet.getRange(2, 1, existing.length, HEADERS.length).setValues(existing);
  }
  console.log('Synced ' + charges.length + ' charges, sheet has ' + existing.length + ' rows');
}

/** Follows Stripe paging until has_more is false. Retries on 429/5xx. */
function fetchAll_(key, since) {
  const out = [];
  let startingAfter = null;
  for (let page = 0; page < 50; page++) {
    let url = 'https://api.stripe.com/v1/charges?limit=100&created[gte]=' + since;
    if (startingAfter) url += '&starting_after=' + encodeURIComponent(startingAfter);
    const body = fetchWithRetry_(url, { headers: { Authorization: 'Bearer ' + key } });
    out.push.apply(out, body.data);
    if (!body.has_more || !body.data.length) break;
    startingAfter = body.data[body.data.length - 1].id;
  }
  return out;
}

function fetchWithRetry_(url, options) {
  const opts = Object.assign({ muteHttpExceptions: true }, options);
  for (let attempt = 1; attempt <= 4; attempt++) {
    const res = UrlFetchApp.fetch(url, opts);
    const code = res.getResponseCode();
    if (code === 200) return JSON.parse(res.getContentText());
    if (code !== 429 && code < 500) {
      throw new Error('API ' + code + ': ' + res.getContentText().slice(0, 300));
    }
    Utilities.sleep(1000 * Math.pow(2, attempt)); // 2s, 4s, 8s, 16s
  }
  throw new Error('API kept failing after 4 attempts: ' + url);
}
