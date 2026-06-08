/**
 * DocFill — Google Sheets history backend (Google Apps Script Web App)
 * ====================================================================
 *
 * SETUP (one time):
 *   1. Create / open a Google Sheet (this will hold the history).
 *   2. Extensions ▸ Apps Script  →  delete any code, paste THIS file.
 *   3. Change TOKEN below to a long random string (keep it secret).
 *   4. Deploy ▸ New deployment ▸ type "Web app":
 *        - Execute as:      Me
 *        - Who has access:  Anyone
 *      Click Deploy, authorize, and COPY the "/exec" Web app URL.
 *   5. In DocFill's .streamlit/secrets.toml add:
 *        GSHEET_WEBAPP_URL = "https://script.google.com/macros/s/XXXX/exec"
 *        GSHEET_TOKEN      = "the-same-TOKEN-you-set-below"
 *
 * After that, DocFill automatically uses Google Sheets for history
 * (cloud + local). No service account or OAuth keys needed.
 */

const TOKEN = 'aalsfhoaiqowfihqwiohoiqw';
const SHEET_NAME = 'History';
const HEADERS = [
  'id', 'ts', 'batch_id', 'source_file', 'english_name', 'arabic_name', 'iban',
  'account_number', 'swift', 'currency', 'bank_name_ar', 'bank_name_en',
  'nationality', 'emirates_id', 'date_of_birth', 'id_expiry', 'emirate', 'iban_valid'
];

function _sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); }
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); }
  return sh;
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.token !== TOKEN) return _json({ ok: false, error: 'bad token' });

    const sh = _sheet();
    const action = req.action;

    if (action === 'add') {
      const recs = req.records || [];
      recs.forEach(function (r) {
        const id = Utilities.getUuid();
        const row = HEADERS.map(function (h) {
          return h === 'id' ? id : (r[h] != null ? r[h] : '');
        });
        sh.appendRow(row);
      });
      return _json({ ok: true, added: recs.length });
    }

    if (action === 'list') {
      const values = sh.getDataRange().getValues();
      const out = [];
      for (var i = 1; i < values.length; i++) {
        var obj = {};
        for (var j = 0; j < HEADERS.length; j++) { obj[HEADERS[j]] = values[i][j]; }
        out.push(obj);
      }
      return _json({ ok: true, records: out });
    }

    if (action === 'delete') {
      const ids = req.ids || [];
      const values = sh.getDataRange().getValues();
      var deleted = 0;
      for (var i = values.length - 1; i >= 1; i--) {
        if (ids.indexOf(values[i][0]) !== -1) { sh.deleteRow(i + 1); deleted++; }
      }
      return _json({ ok: true, deleted: deleted });
    }

    if (action === 'clear') {
      const last = sh.getLastRow();
      const removed = Math.max(0, last - 1);
      if (last > 1) { sh.deleteRows(2, last - 1); }
      return _json({ ok: true, removed: removed });
    }

    return _json({ ok: false, error: 'unknown action: ' + action });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return _json({ ok: true, service: 'DocFill history', sheet: SHEET_NAME });
}
