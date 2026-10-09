/**
 * ระบบแจ้งประกาศดับไฟ - การไฟฟ้าสาขาวัดสิงห์
 * ไฟล์ในโปรเจกต์: Code.gs, Index.html, Admin.html, Shared.html
 * Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
 *   - หน้าประชาชน:    <URL>/exec
 *   - หน้าเจ้าหน้าที่: <URL>/exec?page=admin
 * รหัสผ่านเจ้าหน้าที่ดึงจากชีต "รหัส" (คอลัมน์ A = รหัส, B = ตำแหน่ง)
 */

const SHEET_NAME = 'Outages';
const HEADERS = ['id', 'area', 'reason', 'date', 'startTime', 'endTime', 'remark', 'geojson', 'status', 'createdAt', 'updatedBy', 'slots'];

// ชีตเก็บรหัสผ่าน
const AUTH_SHEET_ID = '1SKqwYU8GBEBiv9K_ZDZ4UOT8-pDqID24VlcU-ihLnkU';
const AUTH_SHEET_NAME = 'รหัส';

// ข้อความบนภาพประกาศ (แก้ได้)
const CONFIG = {
  branch: 'การไฟฟ้าส่วนภูมิภาควัดสิงห์',
  callCenter: 'PEA Call Center 1129',
  line: 'บริการไฟฟ้าวัดสิงห์',
  phone: '056-461-476',
  web: 'www.pea.co.th',
  defaultRemark: 'กรณีมีผู้ป่วยที่ต้องใช้ไฟฟ้าในการเดินเครื่องมือแพทย์ ในบริเวณพื้นที่ที่งดจ่ายกระแสไฟฟ้า โปรดแจ้งให้การไฟฟ้าส่วนภูมิภาควัดสิงห์ทราบล่วงหน้า เมื่อดำเนินการแล้วเสร็จจะจ่ายกระแสไฟฟ้ากลับคืนทันที จึงขออภัยในความไม่สะดวกมา ณ ที่นี้'
};

// (ไม่บังคับ) ใส่ ID ไฟล์โลโก้ใน Google Drive (png/jpg) เพื่อให้โลโก้ขึ้นบนภาพประกาศ
// เปิดลิงก์ไฟล์ใน Drive แล้วเอาข้อความระหว่าง /d/ กับ /view มาใส่
const LOGO_FILE_ID = '';

function doGet(e) {
  const isAdmin = e && e.parameter && e.parameter.page === 'admin';
  const t = HtmlService.createTemplateFromFile(isAdmin ? 'Admin' : 'Index');
  const base = ScriptApp.getService().getUrl();
  t.adminUrl = base + '?page=admin';
  t.publicUrl = base;
  return t.evaluate()
    .setTitle(isAdmin ? 'จัดการประกาศดับไฟ' : 'ประกาศดับไฟ - กฟภ.สาขาวัดสิงห์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

function getConfig() {
  const c = Object.assign({}, CONFIG, { logo: '' });
  if (LOGO_FILE_ID) {
    try {
      const b = DriveApp.getFileById(LOGO_FILE_ID).getBlob();
      c.logo = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
    } catch (err) { /* ไม่มีโลโก้ก็ใช้ตราแทน */ }
  }
  return c;
}

/* ---------------- ตรวจรหัสผ่านจากชีต ---------------- */

/** คืนค่า "ตำแหน่ง" ของรหัสที่ตรงกัน ถ้าไม่ตรงจะ throw error */
function checkAuth_(password) {
  const pw = String(password == null ? '' : password).trim();
  if (!pw) throw new Error('กรุณากรอกรหัสผ่าน');
  const sh = SpreadsheetApp.openById(AUTH_SHEET_ID).getSheetByName(AUTH_SHEET_NAME);
  if (!sh) throw new Error('ไม่พบชีต "' + AUTH_SHEET_NAME + '" ในไฟล์รหัสผ่าน');
  const last = sh.getLastRow();
  if (last < 1) throw new Error('ยังไม่มีรหัสผ่านในชีต');
  const values = sh.getRange(1, 1, last, 2).getDisplayValues();
  const headerWords = ['รหัส', 'รหัสผ่าน', 'password'];
  for (let i = 0; i < values.length; i++) {
    const code = String(values[i][0]).trim();
    if (!code || headerWords.indexOf(code.toLowerCase()) >= 0) continue;
    if (code === pw) return String(values[i][1]).trim() || 'เจ้าหน้าที่';
  }
  throw new Error('รหัสผ่านไม่ถูกต้อง');
}

/** รันครั้งแรกเพื่อให้ระบบขอสิทธิ์ และทดสอบการอ่านชีตรหัส */
function testAuth() {
  const sh = SpreadsheetApp.openById(AUTH_SHEET_ID).getSheetByName(AUTH_SHEET_NAME);
  Logger.log(sh ? 'อ่านชีต "รหัส" ได้ จำนวนแถว: ' + sh.getLastRow() : 'ไม่พบชีต "รหัส"');
}

/* ---------------- ฐานข้อมูลประกาศ ---------------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  const hdr = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (hdr.join('|') !== HEADERS.join('|')) {          // เติมหัวตาราง/คอลัมน์ใหม่ให้ครบ
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), HEADERS.length).setNumberFormat('@'); // เก็บเป็นข้อความ
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues().map(r => {
    const o = {};
    HEADERS.forEach((h, i) => o[h] = String(r[i] == null ? '' : r[i]));
    let slots = [];
    try { slots = JSON.parse(o.slots || '[]'); } catch (e) {}
    if (!Array.isArray(slots) || !slots.length) {      // ข้อมูลเก่าที่มีช่วงเดียว
      slots = o.date ? [{ date: o.date, start: o.startTime, end: o.endTime }] : [];
    }
    o.slots = slots;
    return o;
  });
}

/** หน้าประชาชน: ไม่รวมรายการที่ยกเลิก และไม่ส่งข้อมูลภายใน */
function getOutages() {
  return readAll_().filter(o => o.status !== 'ยกเลิก').map(o => { delete o.updatedBy; return o; });
}

/** หน้าเจ้าหน้าที่ (ใช้ล็อกอินด้วย) คืน {position, rows} */
function adminList(password) {
  const position = checkAuth_(password);
  return { position: position, rows: readAll_() };
}

function adminSave(password, d) {
  const position = checkAuth_(password);
  const slots = (Array.isArray(d.slots) ? d.slots : [])
    .filter(s => s && /^\d{4}-\d{2}-\d{2}$/.test(String(s.date)))
    .map(s => ({ date: String(s.date), start: String(s.start || ''), end: String(s.end || '') }))
    .sort((a, b) => (a.date + a.start).localeCompare(b.date + b.start));
  if (!d.area || !slots.length) throw new Error('กรุณากรอกบริเวณและวันที่อย่างน้อย 1 ช่วง');
  const rec = Object.assign({}, d, {
    date: slots[0].date, startTime: slots[0].start, endTime: slots[0].end,
    slots: JSON.stringify(slots), updatedBy: position
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const row = HEADERS.map(h => rec[h] == null ? '' : String(rec[h]));
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
    const idx = d.id ? ids.indexOf(d.id) : -1;
    const ci = HEADERS.indexOf('createdAt');
    if (idx >= 0) {
      row[ci] = sh.getRange(idx + 2, ci + 1).getValue();
      sh.getRange(idx + 2, 1, 1, HEADERS.length).setNumberFormat('@').setValues([row]);
    } else {
      row[0] = Utilities.getUuid().slice(0, 8);
      row[ci] = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
      if (!row[HEADERS.indexOf('status')]) row[HEADERS.indexOf('status')] = 'ประกาศ';
      sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length).setNumberFormat('@').setValues([row]);
    }
    return true;
  } finally {
    lock.releaseLock();
  }
}

function adminDelete(password, id) {
  checkAuth_(password);
  const sh = getSheet_();
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
  const idx = ids.indexOf(id);
  if (idx < 0) throw new Error('ไม่พบรายการ');
  sh.deleteRow(idx + 2);
  return true;
}