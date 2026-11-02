/**
 * Halloween costume vote — Google Apps Script backend
 * Bind this script to the Google Sheet that will store the votes
 * (Extensions > Apps Script), then deploy as a Web app.
 */

const SHEET_NAME = 'Votes';
const HEADERS = ['Timestamp', 'Email', 'Name', 'Best Costume Male', 'Best Costume Female'];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // keep emails as plain text
  }
  return sh;
}

function normalize_(email) {
  return String(email || '').trim().toLowerCase();
}

function emailExists_(sh, email) {
  const last = sh.getLastRow();
  if (last < 2) return false;
  const emails = sh.getRange(2, 2, last - 1, 1).getValues().flat();
  return emails.some(function (e) { return normalize_(e) === email; });
}

// Stops values like =HYPERLINK(...) from being run as formulas in the sheet
function clean_(value) {
  const s = String(value || '').trim().slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Used by the page to check an email OR submit a vote (GET avoids CORS preflight)
function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === 'check') {
    const email = normalize_(p.email);
    return json_({ exists: email ? emailExists_(getSheet_(), email) : false });
  }

  if (p.action === 'vote') {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);
      const email  = normalize_(p.email);
      const name   = String(p.name   || '').trim();
      const male   = String(p.male   || '').trim();
      const female = String(p.female || '').trim();

      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false, error: 'invalid_email' });
      if (!name || !male || !female) return json_({ ok: false, error: 'missing_fields' });

      const sh = getSheet_();
      if (emailExists_(sh, email)) return json_({ ok: false, error: 'duplicate' });

      sh.appendRow([new Date(), email, clean_(name), clean_(male), clean_(female)]);
      return json_({ ok: true, saved: true });
    } catch (err) {
      return json_({ ok: false, error: 'server_error' });
    } finally {
      try { lock.releaseLock(); } catch (x) {}
    }
  }

  // Unknown action — likely a stale deployment
  return json_({ ok: false, error: 'unknown_action' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // prevents two simultaneous submissions slipping past the duplicate check
    const data = JSON.parse(e.postData.contents);

    const email = normalize_(data.email);
    const name = String(data.name || '').trim();
    const male = String(data.male || '').trim();
    const female = String(data.female || '').trim();

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false, error: 'invalid_email' });
    if (!name || !male || !female) return json_({ ok: false, error: 'missing_fields' });

    const sh = getSheet_();
    if (emailExists_(sh, email)) return json_({ ok: false, error: 'duplicate' });

    sh.appendRow([new Date(), email, clean_(name), clean_(male), clean_(female)]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'server_error' });
  } finally {
    try { lock.releaseLock(); } catch (x) { }
  }
}