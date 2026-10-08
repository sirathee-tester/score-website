/**
 * ระบบกรอกคะแนน Best Practice 2569 — ส่วนเก็บข้อมูลใน Google Sheets
 *
 * ชีต "กรรมการ"  : คอลัมน์ A = ชื่อกรรมการ, B = PIN  (แก้ไขรายชื่อ/PIN ได้ที่ชีตนี้)
 * ชีต "คะแนน"    : ระบบเขียนให้อัตโนมัติ 1 แถว ต่อ กรรมการ 1 ท่าน ต่อ ครู 1 คน
 *
 * ติดตั้งครั้งแรก: เลือกฟังก์ชัน setup แล้วกด Run หนึ่งครั้ง
 */
const SHEET_JUDGES = 'กรรมการ';
const SHEET_SCORES = 'คะแนน';
const ITEMS = ['1.1','1.2','2.1','2.2','2.3','2.4','2.5','3.1','3.2','3.3','4.1','4.2','4.3','5.1','5.2'];
const HEAD = ['กรรมการ','ช่วงชั้น','ลำดับ','ชื่อ-นามสกุล','โรงเรียน'];
const FIRST_ITEM_COL = HEAD.length;               // 0-based index of 1.1
const TOTAL_COL = FIRST_ITEM_COL + ITEMS.length;  // 0-based index of รวม

function setup() {
  const ss = SpreadsheetApp.getActive();
  let j = ss.getSheetByName(SHEET_JUDGES);
  if (!j) {
    j = ss.insertSheet(SHEET_JUDGES);
    j.getRange(1, 1, 1, 2).setValues([['ชื่อกรรมการ', 'PIN']]).setFontWeight('bold');
    j.getRange(2, 1, 3, 2).setValues([
      ['กรรมการคนที่ 1', '1111'],
      ['กรรมการคนที่ 2', '2222'],
      ['กรรมการคนที่ 3', '3333'],
    ]);
    j.getRange('B:B').setNumberFormat('@');
    j.setColumnWidth(1, 260);
  }
  let s = ss.getSheetByName(SHEET_SCORES);
  if (!s) {
    s = ss.insertSheet(SHEET_SCORES);
    s.getRange(1, 1, 1, TOTAL_COL + 2).setValues([HEAD.concat(ITEMS, ['รวม', 'บันทึกล่าสุด'])]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.setColumnWidth(1, 200); s.setColumnWidth(4, 220); s.setColumnWidth(5, 220);
  }
}

function doGet(e) {
  return out(handle({ action: (e && e.parameter && e.parameter.action) || 'judges' }));
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad_request' }); }
  return out(handle(req));
}

function handle(req) {
  try {
    if (req.action === 'judges') return { ok: true, judges: readJudges().map(j => j.name) };
    const judge = auth(req.judge, req.pin);
    if (!judge) return { ok: false, error: 'auth' };
    if (req.action === 'login') return { ok: true };
    if (req.action === 'scores') return { ok: true, judges: readJudges().map(j => j.name), scores: readScores() };
    if (req.action === 'save') { saveRows(judge, req.rows || []); return { ok: true }; }
    return { ok: false, error: 'unknown_action' };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function readJudges() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_JUDGES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .map(r => ({ name: String(r[0]).trim(), pin: String(r[1]).trim() }))
    .filter(j => j.name);
}

function auth(name, pin) {
  const j = readJudges().find(x => x.name === String(name || '').trim());
  return j && j.pin === String(pin || '').trim() ? j.name : null;
}

function readScores() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SCORES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, TOTAL_COL).getValues().map(r => {
    const s = {};
    ITEMS.forEach((id, i) => {
      const v = r[FIRST_ITEM_COL + i];
      if (v !== '' && v !== null && !isNaN(Number(v))) s[id] = Number(v);
    });
    return { judge: String(r[0]), level: Number(r[1]), no: Number(r[2]), s: s };
  });
}

function saveRows(judge, rows) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SCORES);
    const n = sh.getLastRow() - 1;
    const keys = n > 0 ? sh.getRange(2, 1, n, 3).getValues().map(r => r[0] + '|' + r[1] + '|' + r[2]) : [];
    const now = new Date();
    rows.forEach(row => {
      const level = Number(row.level), no = Number(row.no), sc = row.s || {};
      const vals = ITEMS.map(id => {
        const v = Number(sc[id]);
        return sc[id] === undefined || sc[id] === null || sc[id] === '' || isNaN(v) ? '' : v;
      });
      const nums = vals.filter(v => v !== '');
      const total = nums.length ? nums.reduce((a, b) => a + b, 0) : '';
      const line = [judge, level, no, String(row.teacher || ''), String(row.school || '')].concat(vals, [total, now]);
      const idx = keys.indexOf(judge + '|' + level + '|' + no);
      if (idx >= 0) {
        sh.getRange(idx + 2, 1, 1, line.length).setValues([line]);
      } else {
        sh.appendRow(line);
        keys.push(judge + '|' + level + '|' + no);
      }
    });
  } finally {
    lock.releaseLock();
  }
}
