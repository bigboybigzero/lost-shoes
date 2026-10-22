/**
 * @OnlyCurrentDoc
 * ระบบตามหารองเท้าหาย: Backend (Google Apps Script Web App)
 * วางโค้ดนี้ใน Google Sheet > ส่วนขยาย > Apps Script แล้ว Deploy เป็น Web App
 * ชีต "รายการแจ้ง" และ "สมาชิก" จะถูกสร้างให้อัตโนมัติในการใช้งานครั้งแรก
 * ตั้งใครเป็น admin: เปิดชีต "สมาชิก" แล้วเปลี่ยนช่องสิทธิ์ของคนนั้นจาก user เป็น admin
 */

const REPORT_SHEET = 'รายการแจ้ง';
const USER_SHEET = 'สมาชิก';
const REPORT_HEADERS = ['เวลาที่แจ้ง', 'ชื่อ-นามสกุล', 'ระดับชั้น', 'เวลาที่หาย', 'สถานที่', 'เบอร์โทร'];
const USER_HEADERS = ['ชื่อ-นามสกุล', 'Username', 'Password (เข้ารหัสแล้ว)', 'Salt', 'สิทธิ์ (user/admin)'];
const SESSION_SECONDS = 6 * 60 * 60; // อยู่ในระบบได้ 6 ชั่วโมง

// GET: ส่งรายการแจ้งรองเท้าหายทั้งหมด (ใหม่สุดขึ้นก่อน)
function doGet() {
  const sheet = getSheet_(REPORT_SHEET, REPORT_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return json_([]);

  const tz = Session.getScriptTimeZone();
  const rows = sheet.getRange(2, 1, lastRow - 1, REPORT_HEADERS.length).getValues();
  const data = rows
    .map((r, i) => ({
      rowNumber: i + 2,
      timestamp: r[0],
      studentName: String(r[1]),
      studentClass: String(r[2]),
      loseTime: r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'HH:mm') : String(r[3]),
      location: String(r[4]),
      tel: String(r[5])
    }))
    .filter(item => item.studentName !== '');
  return json_(data.reverse());
}

// POST: สมัครสมาชิก / เข้าสู่ระบบ / แจ้ง / แก้ไข / ลบ
function doPost(e) {
  const p = (e && e.parameter) || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    switch (p.action) {
      case 'register': return json_(register_(p));
      case 'login': return json_(login_(p));
      case 'create': return json_(createReport_(p));
      case 'update': return json_(updateReport_(p));
      case 'delete': return json_(deleteReport_(p));
      default: return json_(fail_('ไม่รู้จักคำสั่งนี้'));
    }
  } catch (err) {
    return json_(fail_('เกิดข้อผิดพลาดที่เซิร์ฟเวอร์: ' + err.message));
  } finally {
    lock.releaseLock();
  }
}

function register_(p) {
  const fullName = clean_(p.fullName);
  const username = clean_(p.username).toLowerCase();
  const password = String(p.password || '');
  if (!fullName || !username || !password) return fail_('กรุณากรอกข้อมูลให้ครบ');
  if (!/^\S{3,30}$/.test(username)) return fail_('Username ต้องยาว 3-30 ตัวอักษร และห้ามมีช่องว่าง');
  if (password.length < 4) return fail_('Password ต้องยาวอย่างน้อย 4 ตัวอักษร');
  if (findUser_(username)) return fail_('Username นี้มีคนใช้แล้ว');

  // สมัครใหม่เป็น user เสมอ ไม่เชื่อค่า role ที่ส่งมาจากหน้าเว็บ
  const salt = Utilities.getUuid();
  getSheet_(USER_SHEET, USER_HEADERS)
    .appendRow([text_(fullName), text_(username), text_(hash_(password, salt)), text_(salt), 'user']);
  return ok_('สมัครสมาชิกสำเร็จ');
}

function login_(p) {
  const user = findUser_(clean_(p.username).toLowerCase());
  if (!user || user.hash !== hash_(String(p.password || ''), user.salt)) {
    return fail_('Username หรือ Password ไม่ถูกต้อง');
  }
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, user.username, SESSION_SECONDS);
  return {
    result: 'success',
    user: { fullName: user.fullName, username: user.username, role: user.role, token: token }
  };
}

function createReport_(p) {
  const values = reportValues_(p);
  if (!values) return fail_('กรุณากรอกข้อมูลให้ครบ');
  getSheet_(REPORT_SHEET, REPORT_HEADERS).appendRow([new Date()].concat(values));
  return ok_('ส่งข้อมูลแจ้งรองเท้าหายเรียบร้อยแล้ว');
}

function updateReport_(p) {
  const denied = requireAdmin_(p.token);
  if (denied) return fail_(denied);
  const sheet = getSheet_(REPORT_SHEET, REPORT_HEADERS);
  const row = Number(p.rowNumber);
  if (!isReportRow_(sheet, row)) return fail_('ไม่พบรายการนี้ อาจถูกลบไปแล้ว');
  const values = reportValues_(p);
  if (!values) return fail_('กรุณากรอกข้อมูลให้ครบ');
  sheet.getRange(row, 2, 1, values.length).setValues([values]);
  return ok_('แก้ไขข้อมูลเรียบร้อยแล้ว');
}

function deleteReport_(p) {
  const denied = requireAdmin_(p.token);
  if (denied) return fail_(denied);
  const sheet = getSheet_(REPORT_SHEET, REPORT_HEADERS);
  const row = Number(p.rowNumber);
  if (!isReportRow_(sheet, row)) return fail_('ไม่พบรายการนี้ อาจถูกลบไปแล้ว');
  sheet.deleteRow(row);
  return ok_('ลบข้อมูลเรียบร้อยแล้ว');
}

// ตรวจสิทธิ์ที่เซิร์ฟเวอร์ คืนข้อความเหตุผลถ้าไม่ผ่าน หรือ '' ถ้าเป็น admin
function requireAdmin_(token) {
  const username = token ? CacheService.getScriptCache().get('session_' + token) : null;
  if (!username) return 'กรุณาเข้าสู่ระบบใหม่อีกครั้ง';
  const user = findUser_(username);
  if (!user || user.role !== 'admin') return 'เฉพาะ admin เท่านั้นที่แก้ไขหรือลบข้อมูลได้';
  return '';
}

function findUser_(username) {
  if (!username) return null;
  const sheet = getSheet_(USER_SHEET, USER_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const rows = sheet.getRange(2, 1, lastRow - 1, USER_HEADERS.length).getValues();
  for (const r of rows) {
    if (String(r[1]).trim().toLowerCase() === username) {
      return {
        fullName: String(r[0]),
        username: username,
        hash: String(r[2]),
        salt: String(r[3]),
        role: String(r[4]).trim().toLowerCase() || 'user'
      };
    }
  }
  return null;
}

function reportValues_(p) {
  const v = [p.student_name, p.student_class, p.lose_time, p.student_shoe, p.student_tel].map(clean_);
  if (v.some(x => !x)) return null;
  return v.map(text_);
}

function isReportRow_(sheet, row) {
  return Number.isInteger(row) && row >= 2 && row <= sheet.getLastRow();
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function hash_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function clean_(v) {
  return String(v == null ? '' : v).trim().slice(0, 200);
}

// ใส่ ' นำหน้าให้ชีตเก็บเป็นข้อความ: เบอร์โทรไม่หาย 0 นำหน้า เวลาไม่ถูกแปลง และสูตรที่พิมพ์มาไม่ถูกรัน
function text_(v) {
  return "'" + v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ok_(msg) {
  return { result: 'success', msg: msg };
}

function fail_(msg) {
  return { result: 'error', msg: msg };
}
