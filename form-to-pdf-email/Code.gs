/**
 * Demo 1 — Google Form -> Sheet -> PDF -> email
 * Every new form answer fills a Google Docs template, saves it as PDF
 * in a Drive folder, emails it to the person and writes the status back
 * to the sheet. Safe to re-run: rows that already have a PDF are skipped.
 *
 * Setup:
 * 1. Response sheet "Form Responses 1" with columns:
 *    Timestamp | Name | Email | Service | Date | Amount
 * 2. A Google Doc template with placeholders {{Name}} {{Service}} {{Date}} {{Amount}} {{Number}}
 * 3. Fill CONFIG, run setup() once, approve permissions.
 */

const CONFIG = {
  SHEET_NAME: 'Form Responses 1',
  TEMPLATE_DOC_ID: 'PUT_TEMPLATE_DOC_ID_HERE',
  OUTPUT_FOLDER_ID: 'PUT_DRIVE_FOLDER_ID_HERE',
  EMAIL_SUBJECT: 'Your confirmation',
  STATUS_HEADER: 'PDF status',
  LINK_HEADER: 'PDF link'
};

/** Run once: adds status columns and the form-submit trigger. */
function setup() {
  const sheet = getSheet_();
  ensureColumn_(sheet, CONFIG.STATUS_HEADER);
  ensureColumn_(sheet, CONFIG.LINK_HEADER);
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'onFormSubmitHandler');
  if (!exists) {
    ScriptApp.newTrigger('onFormSubmitHandler')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
  }
}

/** Trigger entry point. Processes every row without a PDF, not only the new one. */
function onFormSubmitHandler() {
  processPending();
}

function processPending() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // another run is busy
  try {
    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);
    const statusCol = col(CONFIG.STATUS_HEADER);
    const linkCol = col(CONFIG.LINK_HEADER);
    if (statusCol < 0 || linkCol < 0) throw new Error('Run setup() first');

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (row[statusCol] === 'Sent') continue;
      const record = {};
      headers.forEach((h, i) => { record[h] = row[i]; });
      record.Number = 'C-' + String(r).padStart(4, '0');
      try {
        if (!isValidEmail_(record.Email)) throw new Error('Bad email: ' + record.Email);
        const pdf = buildPdf_(record);
        MailApp.sendEmail({
          to: record.Email,
          subject: CONFIG.EMAIL_SUBJECT + ' ' + record.Number,
          body: 'Hi ' + record.Name + ',\n\nYour confirmation is attached.\n',
          attachments: [pdf.getAs(MimeType.PDF)]
        });
        sheet.getRange(r + 1, statusCol + 1).setValue('Sent');
        sheet.getRange(r + 1, linkCol + 1).setValue(pdf.getUrl());
      } catch (err) {
        sheet.getRange(r + 1, statusCol + 1).setValue('Error: ' + err.message);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function buildPdf_(record) {
  const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const copy = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID)
    .makeCopy(record.Number + ' ' + record.Name, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();
  Object.keys(record).forEach(key => {
    body.replaceText('\\{\\{' + key + '\\}\\}', formatValue_(record[key]));
  });
  doc.saveAndClose();
  const pdf = folder.createFile(copy.getAs(MimeType.PDF)).setName(record.Number + '.pdf');
  copy.setTrashed(true); // keep only the PDF
  return pdf;
}

function formatValue_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v === null || v === undefined ? '' : String(v);
}

function isValidEmail_(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || ''));
}

function getSheet_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('Sheet not found: ' + CONFIG.SHEET_NAME);
  return sheet;
}

function ensureColumn_(sheet, header) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf(header) < 0) {
    sheet.getRange(1, headers.length + 1).setValue(header);
  }
}
