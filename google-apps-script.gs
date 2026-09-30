// Saumos · synchronisation du suivi (Google Sheets)
// À coller dans Extensions > Apps Script d'une Google Sheet vide.
const SHEET = 'Suivi';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(['type', 'date', 'task', 'by', 'at', 'text']); }
  return s;
}
function day_(v) {
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const out = { entries: [], comments: [], config: null };
  rows.forEach(r => {
    if (r[0] === 'config') { try { out.config = JSON.parse(r[5]); } catch (x) {} return; }
    const o = { date: day_(r[1]), task: r[2], by: r[3], at: String(r[4]), text: r[5] };
    (r[0] === 'comment' ? out.comments : out.entries).push(o);
  });
  return json_(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const b = JSON.parse(e.postData.contents);
    const s = sheet_();
    if (b.type === 'entry') {
      const done = s.getDataRange().getValues().some(r => r[0] === 'entry' && day_(r[1]) === b.date && r[2] === b.task);
      if (done) return json_({ ok: false, reason: 'already' });
    }
    s.appendRow([b.type, "'" + b.date, b.task || '', b.by || '', b.at || new Date().toISOString(), b.text || '']);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
