/**
 * Page Watcher — Telegram / email alert when a public web page changes.
 * Runs inside the client's own Google account (Google Sheets + Apps Script).
 *
 * Sheets:
 *   Settings : key | value  (TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID, ALERT_EMAIL, CHECK_EVERY_MINUTES)
 *   Watch    : Active | Label | URL | Rule | Value | Pattern | Last status | Last checked
 *   Log      : Time | Label | URL | Result | Detail
 *
 * Rules:
 *   contains      alert when Value text appears on the page
 *   not_contains  alert when Value text disappears (e.g. "Sold out")
 *   number_below  alert when the number on the page drops below Value
 *   number_above  alert when the number on the page rises above Value
 *   changed       alert when the visible page text changes at all
 * Pattern (optional): number rules -> regex whose first group is the number, e.g. Price:\s*\$([\d.,]+)
 *                    changed rule  -> regex for the part of the page to watch, e.g. Results: \d+
 *
 * An alert fires once when the condition becomes true, not on every check.
 */

var SHEET_SETTINGS = 'Settings';
var SHEET_WATCH = 'Watch';
var SHEET_LOG = 'Log';
var WATCH_HEADERS = ['Active', 'Label', 'URL', 'Rule', 'Value', 'Pattern', 'Last status', 'Last checked'];
var LOG_HEADERS = ['Time', 'Label', 'URL', 'Result', 'Detail'];
var RULES = ['contains', 'not_contains', 'number_below', 'number_above', 'changed'];
var ALLOWED_MINUTES = [1, 5, 10, 15, 30];
var ERRORS_BEFORE_ALERT = 3;
var USER_AGENT = 'Mozilla/5.0 (compatible; PageWatcher/1.0; +Google Apps Script)';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Page watcher')
    .addItem('1. Set up / restart', 'setup')
    .addItem('2. Send test alert', 'sendTestAlert')
    .addItem('3. Check now', 'checkAll')
    .addItem('Stop watching', 'stopWatching')
    .addToUi();
}

/* ---------- setup ---------- */

function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_SETTINGS, ['Key', 'Value'], [
    ['TELEGRAM_BOT_TOKEN', ''],
    ['TELEGRAM_CHAT_ID', ''],
    ['ALERT_EMAIL', ''],
    ['CHECK_EVERY_MINUTES', 15]
  ]);
  ensureSheet_(ss, SHEET_WATCH, WATCH_HEADERS, [
    [true, 'Example: book under 60', 'https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html',
      'number_below', 60, '', '', '']
  ]);
  ensureSheet_(ss, SHEET_LOG, LOG_HEADERS, []);

  var minutes = Number(getSettings_().CHECK_EVERY_MINUTES) || 15;
  if (ALLOWED_MINUTES.indexOf(minutes) < 0) minutes = 15;
  removeTriggers_();
  ScriptApp.newTrigger('checkAll').timeBased().everyMinutes(minutes).create();
  log_('setup', '', 'OK', 'Checking every ' + minutes + ' min');
  return minutes;
}

function stopWatching() {
  removeTriggers_();
  log_('stop', '', 'OK', 'Automatic checks switched off');
}

function removeTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkAll') ScriptApp.deleteTrigger(t);
  });
}

function ensureSheet_(ss, name, headers, exampleRows) {
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.appendRow(headers);
  exampleRows.forEach(function (r) { sh.appendRow(r); });
  sh.setFrozenRows(1);
  return sh;
}

/* ---------- main check ---------- */

function checkAll() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return [];
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET_WATCH);
    if (!sh) throw new Error('Run "Set up" first: sheet "' + SHEET_WATCH + '" is missing.');
    var rows = sh.getDataRange().getValues();
    var settings = getSettings_();
    var props = PropertiesService.getScriptProperties();
    var now = new Date();
    var results = [];

    for (var i = 1; i < rows.length; i++) {
      var w = rowToWatch_(rows[i]);
      if (!w.active || !w.url) continue;
      var res = checkOne_(w, props);
      sh.getRange(i + 1, 7, 1, 2).setValues([[res.status, now]]);
      if (res.alert) {
        sendAlert_(settings, w, res.message);
        log_(w.label, w.url, 'ALERT', res.message);
      } else if (res.status.indexOf('ERROR') === 0) {
        log_(w.label, w.url, 'ERROR', res.status);
      }
      results.push(res);
    }
    return results;
  } finally {
    lock.releaseLock();
  }
}

function rowToWatch_(r) {
  return {
    active: r[0] === true || String(r[0]).toUpperCase() === 'TRUE' || String(r[0]).toUpperCase() === 'YES',
    label: String(r[1] || r[2] || '').trim(),
    url: String(r[2] || '').trim(),
    rule: String(r[3] || 'changed').trim().toLowerCase(),
    value: r[4],
    pattern: String(r[5] || '').trim()
  };
}

function checkOne_(w, props) {
  var key = 'w_' + hash_(w.url + '|' + w.rule + '|' + w.value + '|' + w.pattern);
  var state = JSON.parse(props.getProperty(key) || '{}');

  if (RULES.indexOf(w.rule) < 0) {
    return { alert: false, status: 'ERROR: unknown rule "' + w.rule + '" (use ' + RULES.join(', ') + ')' };
  }

  var page;
  try {
    page = fetchText_(w.url);
  } catch (e) {
    return handleError_(w, state, key, props, String(e.message || e));
  }
  if (page.code >= 400) return handleError_(w, state, key, props, 'HTTP ' + page.code);

  var wasDown = (state.errors || 0) >= ERRORS_BEFORE_ALERT;
  state.errors = 0;

  var met = false, detail = '';
  if (w.rule === 'contains' || w.rule === 'not_contains') {
    var needle = String(w.value || '').trim();
    if (!needle) return { alert: false, status: 'ERROR: Value is empty' };
    var found = page.text.toLowerCase().indexOf(needle.toLowerCase()) >= 0;
    met = (w.rule === 'contains') ? found : !found;
    detail = (w.rule === 'contains')
      ? (found ? '"' + needle + '" is now on the page' : '"' + needle + '" not on the page')
      : (found ? '"' + needle + '" still on the page' : '"' + needle + '" is gone from the page');
  } else if (w.rule === 'number_below' || w.rule === 'number_above') {
    var num = extractNumber_(page.text, w.pattern);
    var limit = parseNumber_(String(w.value));
    if (num === null) return { alert: false, status: 'ERROR: number not found on page (set Pattern)' };
    if (limit === null) return { alert: false, status: 'ERROR: Value is not a number' };
    met = (w.rule === 'number_below') ? num < limit : num > limit;
    detail = 'Now ' + num + ' (' + (w.rule === 'number_below' ? 'alert below ' : 'alert above ') + limit + ')';
    state.lastNumber = num;
  } else if (w.rule === 'changed') {
    // With a Pattern, only the matched part counts, so clocks/ads elsewhere on the page do not trigger alerts.
    var part = page.text;
    if (w.pattern) {
      var pm = page.text.match(new RegExp(w.pattern, 'i'));
      part = pm ? pm[0] : '';
    }
    var h = hash_(part);
    if (!state.hash) {
      state.hash = h;
      props.setProperty(key, JSON.stringify(state));
      return { alert: false, status: 'Baseline saved' };
    }
    met = (h !== state.hash);
    detail = met ? 'Page text changed' : 'No change';
    state.hash = h;
    if (met) {
      // "changed" is an event, not a state: alert every time the text differs from the last check.
      props.setProperty(key, JSON.stringify(state));
      return { alert: true, status: 'CHANGED', message: detail };
    }
  }

  var alert = met && !state.met;
  state.met = met;
  props.setProperty(key, JSON.stringify(state));

  var status = (met ? 'MET: ' : 'Waiting: ') + detail;
  if (wasDown && !alert) status += ' (page reachable again)';
  return { alert: alert, status: status, message: detail };
}

function handleError_(w, state, key, props, reason) {
  state.errors = (state.errors || 0) + 1;
  props.setProperty(key, JSON.stringify(state));
  if (state.errors === ERRORS_BEFORE_ALERT) {
    return {
      alert: true,
      status: 'ERROR: ' + reason,
      message: 'Cannot read the page ' + ERRORS_BEFORE_ALERT + ' times in a row (' + reason + '). Watching continues.'
    };
  }
  return { alert: false, status: 'ERROR: ' + reason + ' (' + state.errors + 'x)' };
}

/* ---------- page reading ---------- */

function fetchText_(url) {
  var resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: { 'User-Agent': USER_AGENT }
  });
  return { code: resp.getResponseCode(), text: htmlToText_(resp.getContentText()) };
}

function htmlToText_(html) {
  return String(html)
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<(script|style|noscript|svg|head)[\s\S]*?<\/\1>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;|&apos;/g, "'")
    .replace(/&pound;/g, '£')
    .replace(/&euro;/g, '€')
    .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(Number(d)); })
    .replace(/\s+/g, ' ')
    .trim();
}

function extractNumber_(text, pattern) {
  var re = pattern
    ? new RegExp(pattern, 'i')
    : /(?:[$£€¥]|USD|EUR|GBP)\s?(\d[\d\s,.]*\d|\d)/i;
  var m = text.match(re);
  if (!m) return null;
  return parseNumber_(m[1] !== undefined ? m[1] : m[0]);
}

function parseNumber_(s) {
  var t = String(s).replace(/[^\d.,-]/g, '');
  if (!t) return null;
  var lastDot = t.lastIndexOf('.'), lastComma = t.lastIndexOf(',');
  if (lastDot >= 0 && lastComma >= 0) {
    // the later separator is the decimal one: 1,299.00 or 1.299,00
    t = lastDot > lastComma ? t.replace(/,/g, '') : t.replace(/\./g, '').replace(',', '.');
  } else if (lastComma >= 0) {
    // 12,50 -> decimal; 1,299 -> thousands
    t = /,\d{1,2}$/.test(t) ? t.replace(',', '.') : t.replace(/,/g, '');
  }
  var n = parseFloat(t);
  return isNaN(n) ? null : n;
}

/* ---------- alerts ---------- */

function sendTestAlert() {
  var sent = sendAlert_(getSettings_(), { label: 'Test', url: '' }, 'Page watcher is connected. Real alerts will look like this.');
  log_('test', '', sent ? 'OK' : 'NOT SENT', sent ? 'Test alert sent' : 'Fill TELEGRAM_* or ALERT_EMAIL in Settings');
  return sent;
}

function sendAlert_(settings, w, message) {
  var time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var text = '🔔 ' + w.label + '\n' + message + (w.url ? '\n' + w.url : '') + '\nChecked ' + time;
  var sent = 0;

  var token = String(settings.TELEGRAM_BOT_TOKEN || '').trim();
  var chats = String(settings.TELEGRAM_CHAT_ID || '').split(',').map(function (c) { return c.trim(); }).filter(String);
  if (token && chats.length) {
    chats.forEach(function (chat) {
      var r = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post',
        payload: { chat_id: chat, text: text, disable_web_page_preview: 'true' },
        muteHttpExceptions: true
      });
      if (r.getResponseCode() === 200) sent++;
      else log_(w.label, w.url, 'TELEGRAM ERROR', r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
    });
  }

  var email = String(settings.ALERT_EMAIL || '').trim();
  if (email) {
    MailApp.sendEmail(email, 'Page watcher: ' + w.label, text);
    sent++;
  }
  return sent > 0;
}

/* ---------- helpers ---------- */

function getSettings_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
  var out = {};
  if (!sh) return out;
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0]) out[String(r[0]).trim()] = r[1];
  });
  return out;
}

function log_(label, url, result, detail) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  if (sh) sh.appendRow([new Date(), label, url, result, detail]);
}

function hash_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, String(s), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
