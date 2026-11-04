// Google Apps Script – מסד נתונים מרכזי למסך בית הכנסת (Google Drive)
// חובה להחליף את הסיסמה לפני פריסה. אין להשתמש בסיסמה שמופיעה בדוגמה.
const PASSWORD = 'CHANGE-THIS-TO-A-LONG-UNIQUE-PASSWORD';
const FILE = 'synagogue-config.json';
const SHEET_NAME = 'מסך בית הכנסת – נתונים';

function file_() {
  const it = DriveApp.getFilesByName(FILE);
  return it.hasNext() ? it.next() : DriveApp.createFile(FILE, '{}', 'application/json');
}
function out_(o) {
  return ContentService.createTextOutput(typeof o === 'string' ? o : JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  // קריאה ציבורית למסכים: מחזירה את התצורה האחרונה שנשמרה.
  return out_(file_().getBlob().getDataAsString() || '{}');
}
function writeTabs_(tabs) {
  const it = DriveApp.getFilesByName(SHEET_NAME);
  const ss = it.hasNext() ? SpreadsheetApp.open(it.next()) : SpreadsheetApp.create(SHEET_NAME);
  const names = Object.keys(tabs);
  names.forEach(n => {
    if (!Array.isArray(tabs[n]) || !tabs[n].length) return;
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    sh.clear();
    const rows = tabs[n];
    const w = Math.max(1, ...rows.map(r => Array.isArray(r) ? r.length : 1));
    const fixed = rows.map(r => { const x = Array.isArray(r) ? r.slice() : [String(r)]; while (x.length < w) x.push(''); return x; });
    sh.getRange(1, 1, fixed.length, w).setNumberFormat('@').setValues(fixed);
    sh.getRange(1, 1, 1, w).setFontWeight('bold');
    sh.setRightToLeft(true);
  });
  ss.getSheets().filter(s => names.indexOf(s.getName()) < 0).forEach(s => {
    if (ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  return ss.getUrl();
}
function doPost(e) {
  let result = { ok: false };
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('בקשה ריקה');
    const body = JSON.parse(e.postData.contents);
    if (!body || typeof body !== 'object' || Array.isArray(body)) {
      return out_({ ok: false, err: 'מבנה הבקשה אינו תקין' });
    }
    if (body.pw !== PASSWORD) return out_({ ok: false, err: 'סיסמה שגויה' });
    if (body.tabs && typeof body.tabs === 'object' && !Array.isArray(body.tabs)) {
      return out_({ ok: true, url: writeTabs_(body.tabs) });
    }
    // תאימות לאחור: מקבלים גם data וגם config/payload, וכן אובייקט הגדרות ישן עם pw בשורש.
    let config = body.data || body.config || body.payload;
    if (!config) {
      config = Object.assign({}, body);
      delete config.pw;
      delete config.t;
    }
    if (!config || typeof config !== 'object' || Array.isArray(config) || !Object.keys(config).length) {
      return out_({ ok: false, err: 'נתונים לא תקינים: לא נמצא אובייקט הגדרות לשמירה' });
    }
    lock.waitLock(15000);
    file_().setContent(JSON.stringify({ t: Date.now(), data: config }));
    result = { ok: true, savedAt: Date.now() };
  } catch (err) {
    result = { ok: false, err: String(err && err.message ? err.message : err) };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch (_) {}
  }
  return out_(result);
}
