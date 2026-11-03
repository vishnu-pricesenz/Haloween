/**
 * Halloween Event Feedback — Google Apps Script backend
 * Extensions > Apps Script > paste > Deploy as Web App (Anyone)
 *
 * Sheet columns (auto-created if sheet 'Feedback' doesn't exist):
 *   A: Timestamp
 *   B: Email          <-- duplicate check
 *   C: Name
 *   D: Program Rating  (Good / Kind Satisfied / Average / Not Satisfied)
 *   E: Feedback        (short description)
 *   F: Next Event      (chosen/typed suggestion)
 */

const SHEET_NAME = 'Feedback';
const HEADERS    = ['Timestamp', 'Email', 'Name', 'Program Rating', 'Feedback', 'Next Event'];

// ── helpers ─────────────────────────────────────────────────────

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');   // emails as plain text
  }
  return sh;
}

function normalize_(v) { return String(v || '').trim().toLowerCase(); }

function emailExists_(sh, email) {
  const last = sh.getLastRow();
  if (last < 2) return false;
  const emails = sh.getRange(2, 2, last - 1, 1).getValues().flat();
  return emails.some(function (e) { return normalize_(e) === email; });
}

function clean_(v) {
  const s = String(v || '').trim().slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── doGet ────────────────────────────────────────────────────────
// All requests come in as GET to avoid CORS preflight issues.

function doGet(e) {
  const p = (e && e.parameter) || {};

  // ── Check duplicate email ────────────────────────────
  if (p.action === 'check') {
    const email = normalize_(p.email);
    return json_({ exists: email ? emailExists_(getSheet_(), email) : false });
  }

  // ── Submit feedback ──────────────────────────────────
  if (p.action === 'feedback') {
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(20000);

      const email    = normalize_(p.email);
      const name     = String(p.name     || '').trim();
      const rating   = String(p.rating   || '').trim();
      const feedback = String(p.feedback || '').trim();
      const next     = String(p.next     || '').trim();

      // Validate
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))
        return json_({ ok: false, error: 'invalid_email' });
      if (!name || !rating || !feedback || !next)
        return json_({ ok: false, error: 'missing_fields' });

      // Duplicate check
      const sh = getSheet_();
      if (emailExists_(sh, email))
        return json_({ ok: false, error: 'duplicate' });

      // Write row
      sh.appendRow([
        new Date(),
        email,
        clean_(name),
        clean_(rating),
        clean_(feedback),
        clean_(next)
      ]);

      return json_({ ok: true, saved: true });

    } catch (err) {
      return json_({ ok: false, error: 'server_error' });
    } finally {
      try { lock.releaseLock(); } catch (x) {}
    }
  }

  // Unknown action
  return json_({ ok: false, error: 'unknown_action' });
}

// ── doPost (kept as fallback, not used by current page) ──────────
function doPost(e) {
  return json_({ ok: false, error: 'use_get' });
}
