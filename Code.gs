/**
 * ระบบสรุปยอดเงินโอนรายวัน (Slip Reader)
 * Backend : Google Apps Script + Google Sheets
 * อ่านสลิป : Gemini API (อ่านรูป → JSON)
 *
 * ชีต: Users | Accounts | Slips | Logs  (สร้างอัตโนมัติด้วย setup())
 *
 * ข้อมูลหลักของสลิป 3 ส่วน = (1) วันที่+เวลา (2) จำนวนเงิน (3) บัญชีร้าน
 * ถ้าทั้ง 3 ส่วนตรงกับรายการที่มีอยู่แล้ว → สถานะ "ซ้ำ? รอยืนยัน" และเด้ง popup ให้ผู้ใช้ยืนยันตัดออก
 */

// ======================== CONFIG ========================
const CFG = {
  SHOP_NAME: 'บ้านยาเขวา',
  APP_NAME: 'ระบบสรุปยอดเงินโอนรายวัน',
  TZ: 'Asia/Bangkok',
  SESSION_SECONDS: 6 * 60 * 60,      // อายุ session (สูงสุดของ CacheService = 6 ชม.)
  MAX_LOGIN_FAIL: 5,
  LOCK_MINUTES: 15,
  HASH_ROUNDS: 1000,
  IMAGE_FOLDER: 'SlipReader_Images',
  MAX_AMOUNT: 5000000,
  DASHBOARD_MAX_DAYS: 800,
};

const SH = { USERS: 'Users', ACCOUNTS: 'Accounts', SLIPS: 'Slips', LOGS: 'Logs' };

const HEADERS = {
  Users: ['Username', 'DisplayName', 'Role', 'Active', 'PasswordHash', 'Salt', 'CreatedAt', 'LastLogin'],
  Accounts: ['Code', 'DisplayName', 'Bank', 'MatchKeys', 'Color', 'Active', 'SortOrder'],
  Slips: ['ID', 'BatchID', 'UploadedAt', 'UploadedBy', 'TxnDate', 'TxnTime', 'AccountCode', 'Amount',
          'ReceiverName', 'ReceiverAccount', 'ReceiverBank', 'SenderName', 'SenderBank', 'RefNo',
          'Status', 'Note', 'FileName', 'ImageFileId', 'ImageHash', 'AIConfidence', 'RawJSON',
          'EditedBy', 'EditedAt', 'ReceiverDetails', 'DupOf', 'BaseStatus'],
  Logs: ['Timestamp', 'User', 'Action', 'Detail'],
};

/*
 * สถานะสลิป
 *  OK          ยืนยันแล้ว — นับในยอด
 *  REVIEW      รอตรวจสอบ — อ่านไม่ครบ/จับคู่บัญชีไม่ได้ (ไม่นับจนกว่าจะยืนยัน)
 *  DUP_PENDING ซ้ำ? รอยืนยัน — 3 ส่วนหลักตรงกับรายการเดิม (ไม่นับ) รอผู้ใช้กดตัดออก/เก็บไว้
 *  DUPLICATE   ตัดออกแล้ว (ซ้ำ)
 *  NOT_SLIP    ไม่ใช่สลิป
 *  CANCELLED   admin ยกเลิก
 *  DRAFT       อ่านแล้ว รอผู้ใช้ตรวจ/กด "บันทึกข้อมูล" (ไม่นับ)
 *  DISCARDED   ผู้ใช้ลบออกก่อนบันทึก (ไม่เกี่ยวข้อง/เลือกรูปผิด) — เก็บแถวไว้ตรวจย้อนหลัง
 */
const ACTIVE_STATUSES = ['OK', 'REVIEW'];

// ค่าเริ่มต้นบัญชีร้าน (ตามที่ปรากฏบนสลิปจริง) — แก้/เพิ่มได้ในหน้า ผู้ดูแลระบบ
const DEFAULT_ACCOUNTS = [
  ['A', 'BBL-Banyakhwao', 'กรุงเทพ', 'Banyakhwao\nหมายเลขร้านค้า1 006001030989217', '#1d4ed8', true, 1],
  ['B', 'SCB-Gift:Banya', 'ไทยพาณิชย์', 'พร้อมเพย์\nนาง สายสุดา สายรัตน์\nXXX XXX 3724', '#7c3aed', true, 2],
  ['C', 'TTB-Tam:Banya', 'ทหารไทยธนชาต', 'พร้อมเพย์\nน.ส. อรณี ปัตถาติ\nXXX XXX 3223', '#ea580c', true, 3],
];

// ======================== WEB APP ========================
function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.shopName = CFG.SHOP_NAME;
  t.appName = CFG.APP_NAME;
  return t.evaluate()
    .setTitle(CFG.APP_NAME + ' — ' + CFG.SHOP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Slip Reader')
    .addItem('ตั้งค่าเริ่มต้น / อัปเดตโครงสร้างชีต', 'setup')
    .addItem('รีเซ็ตรหัสผ่าน admin', 'resetAdminPassword')
    .addToUi();
}

// ======================== SETUP ========================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  migrateAccountsSheet_(ss);
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const h = HEADERS[name];
    if (sh.getMaxColumns() < h.length) sh.insertColumnsAfter(sh.getMaxColumns(), h.length - sh.getMaxColumns());
    // เขียน/เติมหัวตาราง (คอลัมน์ใหม่จะต่อท้าย ข้อมูลเดิมไม่เลื่อน)
    sh.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight('bold').setBackground('#12302f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });

  const slips = ss.getSheetByName(SH.SLIPS);
  ['ID', 'BatchID', 'TxnDate', 'TxnTime', 'AccountCode', 'ReceiverAccount', 'RefNo', 'ImageHash', 'DupOf']
    .forEach(h => slips.getRange(2, col_('Slips', h), slips.getMaxRows() - 1, 1).setNumberFormat('@'));
  slips.getRange(2, col_('Slips', 'Amount'), slips.getMaxRows() - 1, 1).setNumberFormat('#,##0.00');

  const acc = ss.getSheetByName(SH.ACCOUNTS);
  acc.getRange(2, col_('Accounts', 'MatchKeys'), acc.getMaxRows() - 1, 1).setNumberFormat('@').setWrap(true);
  if (acc.getLastRow() < 2) {
    acc.getRange(2, 1, DEFAULT_ACCOUNTS.length, HEADERS.Accounts.length).setValues(DEFAULT_ACCOUNTS);
  }

  const users = ss.getSheetByName(SH.USERS);
  let msg = 'ตั้งค่าชีตเรียบร้อย';
  if (users.getLastRow() < 2) {
    const pw = randomPassword_();
    createUser_('admin', 'ผู้ดูแลระบบ', 'admin', pw);
    msg += '\n\nสร้างผู้ใช้ admin แล้ว\nUsername: admin\nรหัสผ่านชั่วคราว: ' + pw +
           '\n\nกรุณาจดไว้ และเปลี่ยนรหัสผ่านหลังเข้าสู่ระบบครั้งแรก';
  }
  getImageFolder_();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* รันจาก editor */ }
}

/** ย้ายข้อมูลจากชีต Accounts รุ่นแรก (Identifiers + NameKeywords) → MatchKeys */
function migrateAccountsSheet_(ss) {
  const sh = ss.getSheetByName(SH.ACCOUNTS);
  if (!sh || sh.getLastRow() < 1) return;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (head[3] !== 'Identifiers') return;
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
  const out = rows.map(r => [r[0], r[1], r[2], [r[3], r[4]].join(',').split(/[,\n;]/).map(s => String(s).trim()).filter(Boolean).join('\n'),
                             r[5], r[6], r[7]]);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.Accounts.length).setValues([HEADERS.Accounts]);
  if (out.length) sh.getRange(2, 1, out.length, HEADERS.Accounts.length).setValues(out);
}

function resetAdminPassword() {
  const pw = randomPassword_();
  const u = findUser_('admin');
  if (!u) { createUser_('admin', 'ผู้ดูแลระบบ', 'admin', pw); }
  else {
    const salt = Utilities.getUuid();
    const sh = sheet_(SH.USERS);
    sh.getRange(u.row, col_('Users', 'PasswordHash')).setValue(hashPassword_(pw, salt));
    sh.getRange(u.row, col_('Users', 'Salt')).setValue(salt);
    sh.getRange(u.row, col_('Users', 'Active')).setValue(true);
    sh.getRange(u.row, col_('Users', 'Role')).setValue('admin');
    bumpSessionVersion_('admin');
  }
  const msg = 'รหัสผ่านใหม่ของ admin: ' + pw;
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}

// ======================== HELPERS ========================
function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต "' + name + '" — กรุณารัน setup() ก่อน');
  return sh;
}
function col_(sheetName, header) { return HEADERS[sheetName].indexOf(header) + 1; }
function now_() { return new Date(); }
function today_() { return Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd'); }
function fmtDateTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CFG.TZ, 'yyyy-MM-dd HH:mm');
  return String(v || '');
}
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function pad2_(n) { return ('0' + n).slice(-2); }
function round2_(n) { return Math.round((Number(n) || 0) * 100) / 100; }
function randomPassword_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function run_(fn) {
  try { return { ok: true, data: fn() }; }
  catch (e) {
    console.error(e && e.stack || e);
    return { ok: false, error: String(e && e.message || e), code: (e && e.code) || '' };
  }
}
function err_(msg, code) { const e = new Error(msg); e.code = code; return e; }

function log_(user, action, detail) {
  try { sheet_(SH.LOGS).appendRow([now_(), user || '', action, String(detail || '').slice(0, 500)]); }
  catch (e) { console.error('log failed', e); }
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function normalizeDate_(s) {
  if (s instanceof Date) return Utilities.formatDate(s, CFG.TZ, 'yyyy-MM-dd');
  const m = String(s || '').match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!m) return '';
  let y = +m[1]; const mo = +m[2], d = +m[3];
  if (y > 2400) y -= 543;                         // พ.ศ. → ค.ศ.
  if (y < 2015 || y > 2100 || mo < 1 || mo > 12 || d < 1 || d > 31) return '';
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCMonth() !== mo - 1) return '';
  return y + '-' + pad2_(mo) + '-' + pad2_(d);
}
function normalizeTime_(s) {
  if (s instanceof Date) return Utilities.formatDate(s, CFG.TZ, 'HH:mm');
  const m = String(s || '').match(/(\d{1,2})[:.](\d{2})/);
  if (!m) return '';
  const h = +m[1], mi = +m[2];
  if (h > 23 || mi > 59) return '';
  return pad2_(h) + ':' + pad2_(mi);
}
function addDays_(iso, n) {
  const p = iso.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1) + '-' + pad2_(d.getUTCDate());
}
function daysBetween_(a, b) {
  const pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}

// ======================== AUTH ========================
function hashPassword_(pw, salt) {
  let h = salt + '|' + pw;
  for (let i = 0; i < CFG.HASH_ROUNDS; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  }
  return h;
}

function readUsers_() {
  const sh = sheet_(SH.USERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.Users.length).getValues().map((r, i) => ({
    row: i + 2,
    username: String(r[0]).trim().toLowerCase(),
    name: String(r[1] || r[0]),
    role: String(r[2]) === 'admin' ? 'admin' : 'user',
    active: bool_(r[3]),
    hash: String(r[4]),
    salt: String(r[5]),
    createdAt: fmtDateTime_(r[6]),
    lastLogin: fmtDateTime_(r[7]),
  })).filter(u => u.username);
}
function findUser_(username) {
  username = String(username || '').trim().toLowerCase();
  return readUsers_().find(u => u.username === username) || null;
}
function createUser_(username, name, role, pw) {
  const salt = Utilities.getUuid();
  sheet_(SH.USERS).appendRow([username, name, role, true, hashPassword_(pw, salt), salt, now_(), '']);
}

function sessionVersion_(username) {
  return PropertiesService.getScriptProperties().getProperty('SV_' + username) || '1';
}
function bumpSessionVersion_(username) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('SV_' + username, String(Number(p.getProperty('SV_' + username) || 1) + 1));
}

function newSession_(user) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '').slice(0, 12);
  CacheService.getScriptCache().put('S_' + token, JSON.stringify({
    u: user.username, n: user.name, r: user.role, v: sessionVersion_(user.username),
  }), CFG.SESSION_SECONDS);
  return token;
}

function requireSession_(token, adminOnly) {
  const authErr = () => err_('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่', 'AUTH');
  if (!token) throw authErr();
  const cache = CacheService.getScriptCache();
  const raw = cache.get('S_' + token);
  if (!raw) throw authErr();
  const s = JSON.parse(raw);
  if (s.v !== sessionVersion_(s.u)) { cache.remove('S_' + token); throw authErr(); }
  cache.put('S_' + token, raw, CFG.SESSION_SECONDS);
  if (adminOnly && s.r !== 'admin') throw err_('สิทธิ์ไม่เพียงพอ (เฉพาะผู้ดูแลระบบ)', 'FORBIDDEN');
  return s;
}

function apiLogin(username, password) {
  return run_(() => {
    username = String(username || '').trim().toLowerCase();
    if (!username || !password) throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');
    const cache = CacheService.getScriptCache();
    const failKey = 'F_' + username;
    const fails = Number(cache.get(failKey) || 0);
    if (fails >= CFG.MAX_LOGIN_FAIL) {
      throw new Error('ใส่รหัสผิดเกิน ' + CFG.MAX_LOGIN_FAIL + ' ครั้ง กรุณารอ ' + CFG.LOCK_MINUTES + ' นาที');
    }
    const user = findUser_(username);
    if (!user || !user.active || hashPassword_(String(password), user.salt) !== user.hash) {
      cache.put(failKey, String(fails + 1), CFG.LOCK_MINUTES * 60);
      log_(username, 'LOGIN_FAIL', '');
      throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    }
    cache.remove(failKey);
    sheet_(SH.USERS).getRange(user.row, col_('Users', 'LastLogin')).setValue(now_());
    const token = newSession_(user);
    log_(user.username, 'LOGIN', '');
    return {
      token,
      user: { username: user.username, name: user.name, role: user.role },
      accounts: publicAccounts_(user.role === 'admin'),
      today: today_(),
    };
  });
}

function apiLogout(token) {
  return run_(() => { CacheService.getScriptCache().remove('S_' + token); return true; });
}

function apiChangePassword(token, oldPw, newPw) {
  return run_(() => {
    const s = requireSession_(token);
    if (String(newPw || '').length < 6) throw new Error('รหัสผ่านใหม่ต้องยาวอย่างน้อย 6 ตัวอักษร');
    const u = findUser_(s.u);
    if (!u || hashPassword_(String(oldPw || ''), u.salt) !== u.hash) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    const salt = Utilities.getUuid();
    const sh = sheet_(SH.USERS);
    sh.getRange(u.row, col_('Users', 'PasswordHash')).setValue(hashPassword_(String(newPw), salt));
    sh.getRange(u.row, col_('Users', 'Salt')).setValue(salt);
    log_(s.u, 'CHANGE_PASSWORD', '');
    return true;
  });
}

// ======================== USERS (admin) ========================
function apiListUsers(token) {
  return run_(() => {
    requireSession_(token, true);
    return readUsers_().map(u => ({
      username: u.username, name: u.name, role: u.role, active: u.active,
      createdAt: u.createdAt, lastLogin: u.lastLogin,
    }));
  });
}

function apiSaveUser(token, data) {
  return run_(() => {
    const s = requireSession_(token, true);
    const username = String(data.username || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('Username ใช้ a-z, 0-9, . _ - ยาว 3–30 ตัว');
    const name = String(data.name || '').trim() || username;
    const role = data.role === 'admin' ? 'admin' : 'user';
    const active = data.active !== false;
    const pw = String(data.password || '');
    if (pw && pw.length < 6) throw new Error('รหัสผ่านต้องยาวอย่างน้อย 6 ตัวอักษร');

    return withLock_(() => {
      const users = readUsers_();
      const ex = users.find(u => u.username === username);
      const sh = sheet_(SH.USERS);
      if (!ex) {
        if (!pw) throw new Error('ผู้ใช้ใหม่ต้องกำหนดรหัสผ่าน');
        createUser_(username, name, role, pw);
        if (!active) { const nu = findUser_(username); sh.getRange(nu.row, col_('Users', 'Active')).setValue(false); }
        log_(s.u, 'USER_CREATE', username + ' (' + role + ')');
        return true;
      }
      if (username === s.u && (role !== 'admin' || !active)) {
        throw new Error('ไม่สามารถลดสิทธิ์หรือปิดใช้งานบัญชีของตัวเองได้');
      }
      const otherAdmins = users.filter(u => u.username !== username && u.role === 'admin' && u.active).length;
      if (ex.role === 'admin' && (role !== 'admin' || !active) && otherAdmins === 0) {
        throw new Error('ต้องมีผู้ดูแลระบบที่ใช้งานอยู่อย่างน้อย 1 คน');
      }
      sh.getRange(ex.row, col_('Users', 'DisplayName')).setValue(name);
      sh.getRange(ex.row, col_('Users', 'Role')).setValue(role);
      sh.getRange(ex.row, col_('Users', 'Active')).setValue(active);
      if (pw) {
        const salt = Utilities.getUuid();
        sh.getRange(ex.row, col_('Users', 'PasswordHash')).setValue(hashPassword_(pw, salt));
        sh.getRange(ex.row, col_('Users', 'Salt')).setValue(salt);
      }
      if (pw || role !== ex.role || active !== ex.active) bumpSessionVersion_(username);
      log_(s.u, 'USER_UPDATE', username + ' role=' + role + ' active=' + active + (pw ? ' (reset pw)' : ''));
      return true;
    });
  });
}

// ======================== ACCOUNTS ========================
function listAccounts_(includeInactive) {
  const sh = sheet_(SH.ACCOUNTS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.Accounts.length).getValues()
    .map((r, i) => ({
      row: i + 2,
      code: String(r[0]).trim().toUpperCase(),
      name: String(r[1] || ('บัญชี ' + r[0])),
      bank: String(r[2] || ''),
      keys: String(r[3] || '').split(/\n|,|;/).map(s => s.trim()).filter(Boolean),
      color: /^#[0-9a-f]{6}$/i.test(String(r[4])) ? String(r[4]) : '#64748b',
      active: bool_(r[5]),
      sort: Number(r[6]) || 99,
    }))
    .filter(a => a.code && (includeInactive || a.active))
    .sort((a, b) => a.sort - b.sort || a.code.localeCompare(b.code));
}
/** ข้อมูลบัญชีสำหรับส่งให้หน้าเว็บ — ผู้ใช้ทั่วไปไม่เห็นคำค้น (MatchKeys) */
function publicAccounts_(isAdmin) {
  return listAccounts_(true).map(a => {
    const o = { code: a.code, name: a.name, bank: a.bank, color: a.color, active: a.active, sort: a.sort };
    if (isAdmin) o.keys = a.keys;
    return o;
  });
}

function apiGetAccounts(token) {
  return run_(() => {
    const s = requireSession_(token);
    return publicAccounts_(s.r === 'admin');
  });
}

function apiSaveAccount(token, a) {
  return run_(() => {
    const s = requireSession_(token, true);
    const code = String(a.code || '').trim().toUpperCase();
    if (!/^[A-Z0-9_-]{1,10}$/.test(code)) throw new Error('รหัสบัญชีใช้ A-Z, 0-9 ยาวไม่เกิน 10 ตัว เช่น A, B, E');
    const keys = String(a.keys || '').split(/\n/).map(x => x.trim()).filter(Boolean);
    if (!keys.length) throw new Error('กรุณาใส่ข้อมูลที่ปรากฏบนสลิปอย่างน้อย 1 บรรทัด');
    const row = [
      code,
      String(a.name || '').trim() || ('บัญชี ' + code),
      String(a.bank || '').trim(),
      keys.join('\n'),
      /^#[0-9a-f]{6}$/i.test(String(a.color)) ? a.color : '#64748b',
      a.active !== false,
      Number(a.sort) || 99,
    ];
    const saved = withLock_(() => {
      const sh = sheet_(SH.ACCOUNTS);
      const ex = listAccounts_(true).find(x => x.code === code);
      const r = ex ? ex.row : sh.getLastRow() + 1;
      sh.getRange(r, col_('Accounts', 'MatchKeys')).setNumberFormat('@');
      sh.getRange(r, 1, 1, row.length).setValues([row]);
      log_(s.u, ex ? 'ACCOUNT_UPDATE' : 'ACCOUNT_CREATE', code + ' keys=' + keys.join(' | '));
      return true;
    });
    // จับคู่รายการ "รอตรวจสอบ" ที่ยังไม่มีบัญชีใหม่อีกรอบ
    const rematch = rematchPending_(s.u);
    return { accounts: publicAccounts_(true), rematch: rematch, saved: saved };
  });
}

// ======================== SETTINGS (admin) ========================
/*
 * วิธีอ่านสลิป (OCR_ENGINE)
 *  drive  = Google Drive OCR (ค่าเริ่มต้น) — ฟรี ไม่ต้องมี API key รูปไม่ออกนอกบัญชี Google ของร้าน
 *  hybrid = Drive OCR ก่อน ถ้าอ่านไม่ครบค่อยใช้ Gemini เฉพาะใบนั้น (ต้องมี Gemini key)
 *  gemini = Gemini อ่านทุกใบ (ต้องมี Gemini key)
 *
 * โหมดฟรีเท่านั้น (บังคับเสมอ):
 *  - Gemini ใช้ Key แบบไม่ผูกบัตร → เกินโควตา Google จะปฏิเสธ (HTTP 429) ไม่เรียกเก็บเงิน
 *  - เจอ 429 รายวัน → หยุดใช้ Gemini ทันทีจนถึงเวลารีเซ็ต (เที่ยงคืนเวลาสหรัฐฯ) แล้วอ่านด้วย Drive OCR แทน
 *  - เพดานรายวันของระบบเอง (GEMINI_DAILY_CAP) เป็นตัวล็อกชั้นที่สอง
 *  - เลือกรุ่น Flash-Lite อัตโนมัติ (โควตาฟรีมากที่สุด) และไม่ใช้รุ่นที่ต้องจ่ายเงิน (Pro ฯลฯ)
 */
const ENGINES = ['drive', 'hybrid', 'gemini'];
const GEMINI_DEFAULT_CAP = 100;   // เพดานการเรียก Gemini ต่อวันของระบบ (ปรับได้ในหน้าตั้งค่า)
function getEngine_() {
  const e = PropertiesService.getScriptProperties().getProperty('OCR_ENGINE');
  return ENGINES.indexOf(e) >= 0 ? e : 'drive';
}

// ---------- โควตา Gemini (นับตามวันของ Google = เวลาแปซิฟิก) ----------
function quotaDay_() { return Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd'); }
/** เวลาที่โควตารายวันของ Google รีเซ็ตครั้งถัดไป (เที่ยงคืนเวลาแปซิฟิก) */
function nextQuotaReset_() {
  const now = new Date();
  const tomorrowLA = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), 'America/Los_Angeles', 'yyyy-MM-dd');
  const off = Utilities.formatDate(now, 'America/Los_Angeles', 'Z');            // เช่น -0700
  const t = new Date(tomorrowLA + 'T00:00:00' + off.slice(0, 3) + ':' + off.slice(3));
  return new Date(t.getTime() + 5 * 60 * 1000);                                 // เผื่อ 5 นาที
}
function geminiCap_() {
  const raw = PropertiesService.getScriptProperties().getProperty('GEMINI_DAILY_CAP');
  const v = Number(raw);
  return (raw === null || raw === '' || !isFinite(v) || v < 0) ? GEMINI_DEFAULT_CAP : v;
}
function geminiUsedToday_() {
  return Number(PropertiesService.getScriptProperties().getProperty('GEM_USE_' + quotaDay_()) || 0);
}
/** นับการใช้ +1 (ลบตัวนับของวันก่อนๆ ทิ้ง) */
function countGeminiUse_() {
  const p = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const k = 'GEM_USE_' + quotaDay_();
    p.setProperty(k, String(Number(p.getProperty(k) || 0) + 1));
    p.getKeys().forEach(x => { if (x.indexOf('GEM_USE_') === 0 && x !== k) p.deleteProperty(x); });
  } finally { lock.releaseLock(); }
}
/** บันทึกว่าโควตาเต็มถึงเวลาใด (รายวัน = ถึงรอบรีเซ็ต, รายนาที = 1 นาที) */
function markGeminiQuotaOut_(daily, msg) {
  const until = daily ? nextQuotaReset_() : new Date(Date.now() + 65 * 1000);
  PropertiesService.getScriptProperties().setProperty('GEMINI_QUOTA_OUT', JSON.stringify({ until: until.getTime(), daily: daily }));
  log_('system', 'GEMINI_QUOTA_OUT', (daily ? 'daily' : 'per-minute') + ' until ' + fmtDateTime_(until) + ' ' + String(msg || '').slice(0, 200));
}
function geminiQuotaOut_() {
  try {
    const q = JSON.parse(PropertiesService.getScriptProperties().getProperty('GEMINI_QUOTA_OUT') || 'null');
    return q && q.until > Date.now() ? q : null;
  } catch (e) { return null; }
}
/** ตอนนี้ใช้ Gemini ได้หรือไม่ → { ok, reason } */
function geminiAvailable_() {
  if (!PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY')) return { ok: false, reason: 'ยังไม่มี Gemini API Key' };
  const q = geminiQuotaOut_();
  if (q) return { ok: false, reason: 'โควตาฟรีของ Gemini เต็ม ' + (q.daily ? 'ถึง ' + fmtDateTime_(new Date(q.until)) : 'ชั่วคราว 1 นาที') };
  if (geminiUsedToday_() >= geminiCap_()) return { ok: false, reason: 'ครบเพดานของระบบวันนี้ (' + geminiCap_() + ' ครั้ง)' };
  return { ok: true, reason: '' };
}

function apiGetSettings(token) {
  return run_(() => {
    requireSession_(token, true);
    const p = PropertiesService.getScriptProperties();
    const key = p.getProperty('GEMINI_API_KEY') || '';
    const q = geminiQuotaOut_();
    return {
      engine: getEngine_(),
      hasKey: !!key,
      keyHint: key ? '••••' + key.slice(-4) : '',
      model: p.getProperty('GEMINI_MODEL') || '',
      autoModel: p.getProperty('GEMINI_MODEL_FREE') || '',
      cap: geminiCap_(),
      usedToday: geminiUsedToday_(),
      quotaOutUntil: q ? fmtDateTime_(new Date(q.until)) : '',
      quotaDaily: q ? !!q.daily : false,
      nextReset: fmtDateTime_(nextQuotaReset_()),
    };
  });
}

function apiSaveSettings(token, data) {
  return run_(() => {
    const s = requireSession_(token, true);
    const p = PropertiesService.getScriptProperties();
    const engine = ENGINES.indexOf(data.engine) >= 0 ? data.engine : getEngine_();
    const key = String(data.apiKey || '').trim();
    if (key) { p.setProperty('GEMINI_API_KEY', key); p.deleteProperty('GEMINI_QUOTA_OUT'); }
    if (data.clearKey) p.deleteProperty('GEMINI_API_KEY');
    if (engine !== 'drive' && !p.getProperty('GEMINI_API_KEY')) {
      throw new Error('โหมดนี้ต้องใส่ Gemini API Key ก่อน');
    }
    p.setProperty('OCR_ENGINE', engine);
    const model = String(data.model || '').trim();
    if (model && !/^gemini-[a-z0-9.\-]+$/i.test(model)) throw new Error('ชื่อโมเดลไม่ถูกต้อง');
    if (model && /pro|ultra|image|veo|imagen/i.test(model)) throw new Error('รุ่นนี้ไม่มีแบบฟรี — ระบบอนุญาตเฉพาะรุ่น Flash / Flash-Lite');
    if (model) p.setProperty('GEMINI_MODEL', model); else p.deleteProperty('GEMINI_MODEL');
    p.deleteProperty('GEMINI_MODEL_FREE');
    if (data.cap !== undefined && data.cap !== '') {
      const cap = Math.floor(Number(data.cap));
      if (!(cap >= 0 && cap <= 1000)) throw new Error('เพดานต่อวันต้องอยู่ระหว่าง 0–1000');
      p.setProperty('GEMINI_DAILY_CAP', String(cap));
    }
    log_(s.u, 'SETTINGS_UPDATE', 'engine=' + engine + (key ? '; api key changed' : '') +
         (data.clearKey ? '; api key removed' : '') + '; model=' + (model || 'auto') + '; cap=' + geminiCap_());
    return true;
  });
}

// ======================== อ่านสลิป: ตัวเลือกวิธีอ่าน ========================
/**
 * อ่านสลิป 1 รูปตามวิธีที่ตั้งไว้ → { ext (ข้อมูลที่ normalize แล้ว), raw (เก็บลงชีต), engine }
 * Gemini ใช้ไม่ได้ (โควตาเต็ม/ครบเพดาน/ผิดพลาด) → ใช้ Drive OCR แทนเสมอ (ฟรี) ไม่หยุดทำงาน
 */
function readSlip_(base64, bytes, mime) {
  const engine = getEngine_();
  const tryGemini = () => {
    const av = geminiAvailable_();
    if (!av.ok) return { skipped: av.reason };
    try {
      const g = callGemini_(base64, mime);
      return { g: g, ext: normalizeExtraction_(g.data) };
    } catch (e) { return { skipped: String(e.message || e).slice(0, 200) }; }
  };

  if (engine === 'gemini') {
    const r = tryGemini();
    if (r.g) return { ext: r.ext, raw: { engine: 'gemini:' + r.g.model, data: r.g.data }, engine: 'gemini' };
    const text = driveOcr_(bytes, mime);
    return { ext: normalizeExtraction_(parseSlipText_(text)),
             raw: { engine: 'drive', text: text.slice(0, 3000), aiSkipped: r.skipped }, engine: 'drive' };
  }

  const text = driveOcr_(bytes, mime);
  const ext = normalizeExtraction_(parseSlipText_(text));
  const out = { ext: ext, raw: { engine: 'drive', text: text.slice(0, 3000) }, engine: 'drive' };
  if (engine === 'hybrid' && needsAi_(ext)) {
    const r = tryGemini();
    if (r.g && (!needsAi_(r.ext) || score_(r.ext) > score_(ext))) {
      return { ext: r.ext, raw: { engine: 'gemini:' + r.g.model, data: r.g.data, driveText: text.slice(0, 2000) }, engine: 'gemini' };
    }
    if (r.skipped) out.raw.aiSkipped = r.skipped;     // ใช้ผล Drive OCR ต่อ (ใบนี้จะเป็น "รอตรวจสอบ" ถ้าอ่านไม่ครบ)
  }
  return out;
}
function needsAi_(x) {
  if (!x.isSlip) return true;
  return x.amount === null || !x.date || !x.time || x.confidence === 'low' ||
         !matchAccount_(x, listAccounts_(false)).code;
}
function score_(x) {
  return (x.isSlip ? 1 : 0) + (x.amount !== null ? 2 : 0) + (x.date ? 1 : 0) + (x.time ? 1 : 0) +
         (matchAccount_(x, listAccounts_(false)).code ? 2 : 0);
}

// ======================== Google Drive OCR (ฟรี) ========================
function getOcrTempFolderId_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('OCR_TMP_FOLDER_ID');
  if (id) { try { DriveApp.getFolderById(id); return id; } catch (e) {} }
  const f = getImageFolder_().createFolder('_ocr_temp');
  p.setProperty('OCR_TMP_FOLDER_ID', f.getId());
  return f.getId();
}

/** ส่งรูปให้ Google Drive แปลงเป็น Google Doc พร้อม OCR ภาษาไทย → อ่านข้อความ → ลบไฟล์ชั่วคราวทิ้ง */
function driveOcr_(bytes, mime) {
  if (bytes.length > 2 * 1024 * 1024) throw new Error('รูปใหญ่เกิน 2 MB สำหรับ Drive OCR');
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const boundary = 'slipreader' + Date.now();
  const meta = { name: 'ocr_' + Date.now(), mimeType: 'application/vnd.google-apps.document',
                 parents: [getOcrTempFolderId_()] };
  const head = Utilities.newBlob('--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(meta) + '\r\n--' + boundary + '\r\nContent-Type: ' + mime + '\r\n\r\n').getBytes();
  const tail = Utilities.newBlob('\r\n--' + boundary + '--').getBytes();
  const payload = head.concat(Array.prototype.slice.call(bytes), tail);

  let fileId = '', lastErr = '';
  for (let attempt = 0; attempt < 3 && !fileId; attempt++) {
    const res = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&ocrLanguage=th&fields=id', {
        method: 'post', contentType: 'multipart/related; boundary=' + boundary,
        payload: payload, headers: auth, muteHttpExceptions: true,
      });
    const code = res.getResponseCode();
    if (code === 200) { fileId = JSON.parse(res.getContentText()).id; break; }
    lastErr = 'HTTP ' + code + ' ' + res.getContentText().slice(0, 150);
    if (code === 403 || code === 429 || code >= 500) { Utilities.sleep(2000 * (attempt + 1)); continue; }
    break;
  }
  if (!fileId) throw new Error('Drive OCR ไม่สำเร็จ (' + lastErr + ')');

  try {
    const r = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/export?mimeType=text/plain',
                                { headers: auth, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('อ่านข้อความจาก Drive OCR ไม่ได้ (HTTP ' + r.getResponseCode() + ')');
    return r.getContentText('UTF-8').replace(/^﻿/, '').replace(/\r/g, '');
  } finally {
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId,
                      { method: 'delete', headers: auth, muteHttpExceptions: true });
  }
}

/**
 * ทดสอบ Drive OCR จาก editor: อัปโหลดรูปสลิป 1 รูปไว้ใน Google Drive แล้วใส่ชื่อไฟล์ด้านล่าง → Run
 * ดูผลใน Execution log (ข้อความที่อ่านได้ + ข้อมูลที่ระบบแยกออกมา + บัญชีที่จับคู่ได้)
 */
function testDriveOcr() {
  const FILE_NAME = 'slip-test.jpg';            // ← แก้เป็นชื่อไฟล์รูปสลิปใน Drive
  const it = DriveApp.getFilesByName(FILE_NAME);
  if (!it.hasNext()) throw new Error('ไม่พบไฟล์ ' + FILE_NAME + ' ใน Google Drive');
  const blob = it.next().getBlob();
  const text = driveOcr_(blob.getBytes(), blob.getContentType());
  const ext = normalizeExtraction_(parseSlipText_(text));
  Logger.log('--- ข้อความที่อ่านได้ ---\n' + text);
  Logger.log('--- ข้อมูลที่แยกได้ ---\n' + JSON.stringify(ext, null, 1));
  Logger.log('--- บัญชีร้าน ---\n' + JSON.stringify(matchAccount_(ext, listAccounts_(false))));
}

// ======================== แยกข้อมูลจากข้อความสลิป ========================
const TH_MONTHS = {
  'มค': 1, 'มกราคม': 1, 'กพ': 2, 'กุมภาพันธ์': 2, 'มีค': 3, 'มีนาคม': 3, 'เมย': 4, 'เมษายน': 4,
  'พค': 5, 'พฤษภาคม': 5, 'มิย': 6, 'มิถุนายน': 6, 'กค': 7, 'กรกฎาคม': 7, 'สค': 8, 'สิงหาคม': 8,
  'กย': 9, 'กันยายน': 9, 'ตค': 10, 'ตุลาคม': 10, 'พย': 11, 'พฤศจิกายน': 11, 'ธค': 12, 'ธันวาคม': 12,
};
const EN_MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

function yearToAD_(y) {
  y = Number(y);
  if (y >= 2400) return y - 543;            // 2569
  if (y >= 1900) return y;                   // 2026
  if (y >= 50) return 2500 + y - 543;        // 69 → 2569 → 2026
  return 2000 + y;                           // 26 → 2026
}
function isoDate_(y, m, d) {
  return normalizeDate_(yearToAD_(y) + '-' + pad2_(m) + '-' + pad2_(d));
}

/** หาวันที่จากข้อความ → { date, index } */
function findDate_(lines) {
  for (let i = 0; i < lines.length; i++) {
    const L = lines[i];
    let m;
    // 27 ต.ค. 69 / 27 ต.ค. 2569 / 27 ตุลาคม 2569
    const reTh = /(\d{1,2})\s*([ก-๙][ก-๙.\s]{0,12}?)\s*(\d{4}|\d{2})(?!\d)/g;
    while ((m = reTh.exec(L))) {
      const key = m[2].replace(/[.\s]/g, '');
      if (TH_MONTHS[key]) { const d = isoDate_(m[3], TH_MONTHS[key], m[1]); if (d) return { date: d, index: i }; }
    }
    // 27 Oct 2026 / 27 Oct 26
    const reEn = /(\d{1,2})\s*([A-Za-z]{3})[a-z]*\.?,?\s*(\d{4}|\d{2})(?!\d)/g;
    while ((m = reEn.exec(L))) {
      const mo = EN_MONTHS[m[2].toLowerCase()];
      if (mo) { const d = isoDate_(m[3], mo, m[1]); if (d) return { date: d, index: i }; }
    }
    // 27/10/2569 , 27-10-26
    const reNum = /(?:^|[^\d])(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4}|\d{2})(?!\d)/g;
    while ((m = reNum.exec(L))) {
      if (+m[2] >= 1 && +m[2] <= 12) { const d = isoDate_(m[3], +m[2], m[1]); if (d) return { date: d, index: i }; }
    }
  }
  return { date: '', index: -1 };
}

/** หาเวลา: ให้ความสำคัญกับบรรทัดเดียวกับวันที่ → บรรทัดถัดไป → ทั้งข้อความ */
function findTime_(lines, dateIdx) {
  const reColon = /(?:^|[^\d])([01]?\d|2[0-3]):([0-5]\d)(?::[0-5]\d)?(?!\d)/;
  const reDotN = /(?:^|[^\d])([01]?\d|2[0-3])\.([0-5]\d)\s*(?:น\.?|นาฬิกา)/;
  const order = [];
  if (dateIdx >= 0) order.push(dateIdx, dateIdx + 1, dateIdx - 1);
  for (let i = 0; i < lines.length; i++) if (order.indexOf(i) < 0) order.push(i);
  for (const i of order) {
    const L = lines[i]; if (!L) continue;
    const m = L.match(reColon) || L.match(reDotN);
    if (m) return pad2_(+m[1]) + ':' + m[2];
  }
  return '';
}

const MONEY_RE = /(?:^|[^\d.,])(\d{1,3}(?:,\d{3})+|\d+)\.(\d{2})(?![\d])/g;
const AMOUNT_LABEL = /(จำนวนเงิน|จํานวนเงิน|จำนวน|จํานวน|ยอดเงิน|ยอดโอน|ยอดชำระ|ยอดชําระ|ยอดรวม|amount|total)/i;
const NOT_AMOUNT = /(ค่าธรรมเนียม|ค่าธรรม|fee|คงเหลือ|balance|available|วงเงิน)/i;
function moneyIn_(L) {
  const out = []; let m; MONEY_RE.lastIndex = 0;
  while ((m = MONEY_RE.exec(L))) {
    const v = Number(m[1].replace(/,/g, '') + '.' + m[2]);
    if (v > 0) out.push(v);
  }
  return out;
}

/** หาจำนวนเงิน → { amount, confidence } */
function findAmount_(lines) {
  const labeled = [];
  for (let i = 0; i < lines.length; i++) {
    if (!AMOUNT_LABEL.test(lines[i]) || NOT_AMOUNT.test(lines[i])) continue;
    for (let j = i; j <= Math.min(i + 2, lines.length - 1); j++) {
      if (j > i && (NOT_AMOUNT.test(lines[j]) || AMOUNT_LABEL.test(lines[j]))) break;
      const v = moneyIn_(lines[j]);
      if (v.length) { labeled.push(v[0]); break; }
    }
  }
  const uniq = a => a.filter((v, i) => a.indexOf(v) === i);
  const lab = uniq(labeled);
  if (lab.length === 1) return { amount: lab[0], confidence: 'high' };
  if (lab.length > 1) return { amount: lab[0], confidence: 'low' };

  const baht = [];
  lines.forEach((L, i) => {
    if (NOT_AMOUNT.test(L) || (i > 0 && NOT_AMOUNT.test(lines[i - 1]) && !/\d/.test(lines[i - 1]))) return;
    if (/(บาท|thb|฿)/i.test(L) || (lines[i + 1] && /^\s*(บาท|thb|฿)\s*$/i.test(lines[i + 1]))) baht.push.apply(baht, moneyIn_(L));
  });
  const b = uniq(baht);
  if (b.length === 1) return { amount: b[0], confidence: 'medium' };
  if (b.length > 1) return { amount: Math.max.apply(null, b), confidence: 'low' };
  return { amount: null, confidence: 'low' };
}

const REF_LABEL = /(เลขที่รายการ|รหัสอ้างอิง|หมายเลขอ้างอิง|เลขที่อ้างอิง|เลขอ้างอิง|รหัสรายการ|หมายเลขรายการ|transaction\s*(?:id|no\.?|ref\.?)|ref(?:erence)?\s*(?:no\.?|id|number)?)/i;
function findRef_(lines) {
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(REF_LABEL);
    if (!m) continue;
    const rest = lines[i].slice(m.index + m[0].length).replace(/[:：\s]/g, '');
    let r = (rest.match(/[A-Za-z0-9]{6,}/) || [])[0];
    if (!r && lines[i + 1]) r = (lines[i + 1].replace(/[\s:]/g, '').match(/^[A-Za-z0-9]{6,}/) || [])[0];
    if (r) return r;
  }
  return '';
}

// หมายเหตุ: \b ใช้กับอักษรไทยไม่ได้ จึงใช้ (?=ช่องว่าง|:|ท้ายบรรทัด) แทน
const TO_LABEL = /^(?:ไปยัง|ไปที่|ผู้รับเงิน|ผู้รับ|บัญชีปลายทาง|ปลายทาง|(?:ไป|ถึง|ร้านค้า|รับเงิน|to|merchant|payee|receiver)(?=[\s:：]|$))/i;
const FROM_LABEL = /^(?:จาก|ผู้โอน|บัญชีต้นทาง|ต้นทาง|(?:from|sender|payer)(?=[\s:：]|$))/i;
const STOP_LABEL = /(จำนวน|จํานวน|amount|ค่าธรรมเนียม|fee|เลขที่รายการ|รหัสอ้างอิง|ref|บันทึกช่วยจำ|memo|หมายเหตุ|สแกน|scan|qr)/i;

/** แยกส่วนผู้รับ/ผู้โอนจากข้อความ (ใช้ป้ายกำกับ ไปยัง/จาก ถ้ามี) */
function splitParties_(lines) {
  const block = start => {
    const out = [];
    const first = lines[start].replace(TO_LABEL, '').replace(FROM_LABEL, '').replace(/^[\s:：]+/, '');
    if (first) out.push(first);
    for (let j = start + 1; j < Math.min(start + 5, lines.length); j++) {
      if (TO_LABEL.test(lines[j]) || FROM_LABEL.test(lines[j]) || STOP_LABEL.test(lines[j])) break;
      if (lines[j]) out.push(lines[j]);
    }
    return out;
  };
  const toIdx = lines.findIndex(L => TO_LABEL.test(L));
  const fromIdx = lines.findIndex(L => FROM_LABEL.test(L));
  const sender = fromIdx >= 0 ? block(fromIdx) : [];
  let receiver;
  if (toIdx >= 0) receiver = block(toIdx);
  else if (fromIdx >= 0) {        // มีแต่ "จาก" → ใช้ข้อความทั้งหมดยกเว้นส่วนผู้โอน
    const skip = {}; for (let j = fromIdx; j < fromIdx + 1 + sender.length; j++) skip[j] = 1;
    receiver = lines.filter((L, j) => !skip[j]);
  } else receiver = lines.slice();  // ไม่มีป้ายกำกับ → ใช้ทั้งข้อความ (ตรงหลายบัญชีจะให้คนตรวจ)
  return { receiver: receiver, sender: sender };
}

/**
 * แปลงข้อความ OCR ของสลิปเป็นข้อมูลรูปแบบเดียวกับผลจาก AI
 */
function parseSlipText_(text) {
  const lines = String(text || '').split('\n').map(s => s.replace(/\s+/g, ' ').trim()).filter(Boolean);
  const all = lines.join('\n');
  if (all.replace(/\s/g, '').length < 15) {
    return { is_slip: true, amount: null, transaction_date: null, transaction_time: null,
             receiver_details: '', reference_no: null, confidence: 'low', remark: 'อ่านข้อความจากรูปไม่ได้' };
  }
  const failed = /(ไม่สำเร็จ|ไม่สําเร็จ|ล้มเหลว|unsuccessful|failed|declined)/i.test(all);
  const dt = findDate_(lines);
  const time = findTime_(lines, dt.index);
  const amt = findAmount_(lines);
  const ref = findRef_(lines);
  const parties = splitParties_(lines);
  const slipWords = /(สำเร็จ|สําเร็จ|successful|success|completed|โอนเงิน|ชำระเงิน|ชําระเงิน|เลขที่รายการ|รหัสอ้างอิง|พร้อมเพย์|promptpay|transfer|payment|e-?slip)/i.test(all);
  const isSlip = !failed && (slipWords || (amt.amount !== null && !!dt.date));

  let conf = amt.confidence;
  if (conf === 'high' && (!dt.date || !time)) conf = 'medium';
  return {
    is_slip: isSlip,
    amount: amt.amount,
    transaction_date: dt.date || null,
    transaction_time: time || null,
    sender_name: parties.sender[0] || null,
    sender_bank: null,
    receiver_name: null,
    receiver_account: null,
    receiver_bank: null,
    receiver_details: parties.receiver.join(' | ').slice(0, 1000),
    full_text: all,
    reference_no: ref || null,
    confidence: conf,
    remark: failed ? 'สลิปแสดงว่ารายการไม่สำเร็จ' : (isSlip ? null : 'ไม่พบลักษณะของสลิปโอนเงิน'),
  };
}

// ======================== GEMINI (ตัวเลือกเสริม) ========================
const SLIP_PROMPT = [
  'You read Thai bank transfer / QR payment e-slips (K PLUS, SCB EASY, Krungthai NEXT, Bualuang mBanking, krungsri, ttb touch, GSB MyMo, BAAC, PromptPay, TrueMoney, etc.).',
  'Extract ONLY what is printed on the slip. Never guess. Use null when a field is missing, cut off or unreadable.',
  'Rules:',
  '- is_slip: true only if the image is a completed transfer/payment receipt (สำเร็จ). false for any other image, failed/pending transfers, a bare QR code, or a chat screenshot without a slip.',
  '- amount: the transferred amount (จำนวนเงิน / จำนวน / Amount) as a plain number, e.g. 1250.50 or 0.10. Never use the fee (ค่าธรรมเนียม), balance (ยอดคงเหลือ) or any other number.',
  '- transaction_date: Gregorian YYYY-MM-DD. Thai slips use the Buddhist Era: 2569 means 2026; a two-digit year like "69" means BE 2569 = 2026; "68" = 2025.',
  '  Thai months: ม.ค.=01 ก.พ.=02 มี.ค.=03 เม.ย.=04 พ.ค.=05 มิ.ย.=06 ก.ค.=07 ส.ค.=08 ก.ย.=09 ต.ค.=10 พ.ย.=11 ธ.ค.=12.',
  '- transaction_time: 24-hour HH:MM exactly as printed.',
  '- receiver_* is the destination / payee (ไปยัง / ไปที่ / ถึง / To / ผู้รับ / ร้านค้า / Merchant). sender_* is the source (จาก / From / ผู้โอน).',
  '- receiver_account: the receiver account / PromptPay / merchant number exactly as printed, including masking (x, X, *, •), spaces and dashes.',
  '- receiver_details: copy VERBATIM every piece of text printed about the receiver/payee, joined with " | ": names, shop/merchant names, account or PromptPay numbers (with masks), labels such as "พร้อมเพย์" or "PromptPay", merchant/biller IDs with their labels (e.g. "หมายเลขร้านค้า1 006001030989217"), bank name. Do not include sender information.',
  '- reference_no: the transaction reference (เลขที่รายการ / รหัสอ้างอิง / หมายเลขอ้างอิง / Ref No. / Transaction ID) exactly as printed, without spaces.',
  '- confidence: "low" if amount, date, time or receiver is blurry, cropped, covered or uncertain; "medium" if minor doubt; otherwise "high".',
  'Return JSON only.',
].join('\n');

const SLIP_SCHEMA = {
  type: 'OBJECT',
  properties: {
    is_slip: { type: 'BOOLEAN' },
    amount: { type: 'NUMBER', nullable: true },
    transaction_date: { type: 'STRING', nullable: true },
    transaction_time: { type: 'STRING', nullable: true },
    sender_name: { type: 'STRING', nullable: true },
    sender_bank: { type: 'STRING', nullable: true },
    receiver_name: { type: 'STRING', nullable: true },
    receiver_account: { type: 'STRING', nullable: true },
    receiver_bank: { type: 'STRING', nullable: true },
    receiver_details: { type: 'STRING', nullable: true },
    reference_no: { type: 'STRING', nullable: true },
    confidence: { type: 'STRING', enum: ['high', 'medium', 'low'] },
    remark: { type: 'STRING', nullable: true },
  },
  required: ['is_slip', 'amount', 'transaction_date', 'transaction_time', 'receiver_name',
             'receiver_account', 'receiver_details', 'reference_no', 'confidence'],
};

/**
 * เลือกโมเดล Gemini แบบฟรีอัตโนมัติ: Flash-Lite รุ่นล่าสุดก่อน (โควตาฟรีมากที่สุด) แล้วจึง Flash
 * ไม่เลือกรุ่น Pro / preview / experimental / image ซึ่งไม่มีแบบฟรีหรือเปลี่ยนแปลงบ่อย
 */
function pickGeminiModel_(key, force) {
  const p = PropertiesService.getScriptProperties();
  const manual = p.getProperty('GEMINI_MODEL');
  if (manual && !force) return manual;
  const cached = p.getProperty('GEMINI_MODEL_FREE');
  if (cached && !force) return cached;
  const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200',
                                { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('ตรวจสอบโมเดล Gemini ไม่ได้ (HTTP ' + res.getResponseCode() + ') — API Key ถูกต้องหรือไม่?');
  const ver = n => Number((n.match(/gemini-(\d+(?:\.\d+)?)/) || [0, '0'])[1]);
  const models = (JSON.parse(res.getContentText()).models || [])
    .filter(m => (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0)
    .map(m => m.name.replace(/^models\//, ''))
    .filter(n => /^gemini-[\d.]+-flash(-lite)?$/.test(n));
  if (!models.length) throw new Error('บัญชีนี้ไม่มีโมเดล Gemini Flash ที่ใช้ได้');
  models.sort((a, b) => (/-lite$/.test(b) - /-lite$/.test(a)) || (ver(b) - ver(a)));   // lite ก่อน, รุ่นใหม่ก่อน
  p.setProperty('GEMINI_MODEL_FREE', models[0]);
  return models[0];
}

/** 429 รายวันหรือรายนาที? (ข้อความของ Google ระบุ metric เช่น ...PerDay / per_day / FreeTier...) */
function isDailyQuota_(msg) { return /per.?day|perday|daily|requests_per_day|RPD/i.test(String(msg)); }

function callGemini_(base64, mime) {
  const p = PropertiesService.getScriptProperties();
  const key = p.getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('ยังไม่ได้ตั้งค่า Gemini API Key');
  const av = geminiAvailable_();
  if (!av.ok) throw err_(av.reason, 'QUOTA');
  let model = pickGeminiModel_(key, false);
  const body = {
    contents: [{ role: 'user', parts: [
      { inline_data: { mime_type: mime, data: base64 } },
      { text: SLIP_PROMPT },
    ] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json', responseSchema: SLIP_SCHEMA },
  };
  let lastErr = '', repicked = false;
  for (let attempt = 0; attempt < 3; attempt++) {
    countGeminiUse_();
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
    const res = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-goog-api-key': key }, payload: JSON.stringify(body),
    });
    const code = res.getResponseCode();
    if (code === 200) {
      const j = JSON.parse(res.getContentText());
      const cand = (j.candidates || [])[0] || {};
      const text = ((cand.content || {}).parts || [])
        .filter(pt => !pt.thought && pt.text).map(pt => pt.text).join('')
        .replace(/^```(?:json)?\s*|\s*```$/g, '').trim();
      if (text) {
        try { return { data: JSON.parse(text), model: model }; }
        catch (e) { lastErr = 'อ่านผลลัพธ์ AI ไม่ได้'; }
      } else {
        lastErr = 'AI ไม่ส่งผลลัพธ์ (' + (cand.finishReason || 'unknown') + ')';
      }
      break;                                     // ไม่ลองซ้ำ เพื่อประหยัดโควตา
    }
    let msg = res.getContentText();
    try { msg = JSON.parse(msg).error.message; } catch (e) {}
    if (code === 429) {                          // โควตาฟรีเต็ม → หยุดใช้ทันที ไม่รอ ไม่ลองซ้ำ
      markGeminiQuotaOut_(isDailyQuota_(msg) || attempt > 0, msg);
      throw err_('โควตาฟรีของ Gemini เต็ม — ใช้ Drive OCR แทน', 'QUOTA');
    }
    if (code >= 500) { lastErr = 'HTTP ' + code; Utilities.sleep(1500); continue; }
    if (!repicked && !p.getProperty('GEMINI_MODEL') && (code === 404 || /not found|not supported|no longer|not available/i.test(msg))) {
      repicked = true; model = pickGeminiModel_(key, true); continue;
    }
    throw new Error('Gemini API error ' + code + ': ' + String(msg).slice(0, 200));
  }
  throw new Error('AI อ่านสลิปไม่สำเร็จ (' + lastErr + ')');
}

function normalizeExtraction_(x) {
  x = x || {};
  const str = v => (v === null || v === undefined) ? '' : String(v).trim();
  let amt = x.amount;
  if (typeof amt === 'string') amt = Number(amt.replace(/[^\d.]/g, ''));
  amt = Number(amt);
  return {
    isSlip: x.is_slip === true,
    amount: (isFinite(amt) && amt > 0) ? round2_(amt) : null,
    date: normalizeDate_(x.transaction_date),
    time: normalizeTime_(x.transaction_time),
    senderName: str(x.sender_name),
    senderBank: str(x.sender_bank),
    receiverName: str(x.receiver_name),
    receiverAccount: str(x.receiver_account),
    receiverBank: str(x.receiver_bank),
    receiverDetails: str(x.receiver_details).slice(0, 1000),
    fullText: str(x.full_text).slice(0, 4000),
    refNo: str(x.reference_no).replace(/\s+/g, ''),
    confidence: ['high', 'medium', 'low'].indexOf(x.confidence) >= 0 ? x.confidence : 'low',
    remark: str(x.remark),
  };
}

// ======================== จับคู่บัญชีร้าน ========================
/*
 * กติกา: สลิปจะนับเป็นบัญชีใด ต้องพบ "ข้อมูลที่ปรากฏบนสลิป" (MatchKeys) ของบัญชีนั้น "ครบทุกบรรทัด"
 *   เช่น BBL-Banyakhwao = Banyakhwao + หมายเลขร้านค้า1 006001030989217 (ต้องพบทั้ง 2 อย่าง)
 *        SCB-Gift:Banya = พร้อมเพย์ + นาง สายสุดา สายรัตน์ + XXX XXX 3724 (ต้องพบทั้ง 3 อย่าง)
 * วิธีเทียบแต่ละบรรทัด:
 *   - บรรทัดที่มีตัวเลข ≥3 หลัก → เทียบชุดตัวเลข (ไม่สนช่องว่าง ขีด หรือ X ที่ปิดบัง)
 *   - บรรทัดที่เป็นข้อความ → ไม่สนคำนำหน้า/ช่องว่าง/ตัวพิมพ์เล็กใหญ่, "PromptPay" = "พร้อมเพย์"
 *     ชื่อ-นามสกุล ยอมรับสลิปที่ตัดนามสกุลเหลืออักษรแรก (เช่น นาง สายสุดา ส)
 * ลำดับการค้น: ส่วนผู้รับก่อน → ถ้าไม่ครบจึงค้นทั้งสลิป
 * ครบเพียงบัญชีเดียว = ใช้บัญชีนั้น • ครบหลายบัญชี/ไม่ครบเลย = ไม่ผ่านเงื่อนไข (ให้ผู้ใช้ตรวจ)
 */
const TITLE_RE = /(นางสาว|น\.ส\.|นาง|นาย|ด\.ช\.|ด\.ญ\.|mrs\.?|mr\.?|ms\.?|miss)/g;
function normText_(s) {
  return String(s || '').toLowerCase().replace(TITLE_RE, ' ')
    .replace(/promptpay|prompt pay/g, 'พร้อมเพย์')
    .replace(/[\s.\-_,:;()\/\\|'"]+/g, '');
}
function parseOne_(raw) {
  const runs = (String(raw).match(/\d+/g) || []).filter(r => r.length >= 3);
  if (runs.length) return { type: 'num', runs: runs, id: 'n:' + runs.join('-') };
  const words = String(raw).toLowerCase().replace(TITLE_RE, ' ')
    .replace(/promptpay|prompt pay/g, 'พร้อมเพย์').split(/[\s.]+/).filter(Boolean);
  const full = normText_(raw);
  if (full.length < 2) return null;
  const pats = [full];
  if (words.length >= 2) pats.push(normText_(words[0]) + normText_(words[1]).charAt(0));
  return { type: 'text', pats: pats, id: 't:' + full };
}
/** 1 บรรทัด = 1 รายละเอียดที่ต้องพบ; ใช้ " | " คั่นรูปแบบที่ใช้แทนกันได้ เช่น "น.ส. อรณี ปัตถาติ | ORANEE P" */
function parseKey_(raw) {
  const alts = String(raw).split('|').map(s => s.trim()).filter(Boolean).map(parseOne_).filter(Boolean);
  if (!alts.length) return null;
  return { raw: String(raw).trim(), alts: alts, id: alts.map(x => x.id).sort().join('||') };
}
function evalAccounts_(blob, parsed, shared) {
  const txt = normText_(blob);
  const digits = String(blob).replace(/[\s\-.]/g, '');     // เก็บ x ไว้ ไม่ให้ตัวเลขคนละช่วงต่อกัน
  const hitOne = k => k.type === 'num'
    ? k.runs.every(r => digits.indexOf(r) >= 0)
    : k.pats.some(pt => txt.indexOf(pt) >= 0);
  return parsed.map(p => {
    const hits = [], missing = []; let uniqueHits = 0;
    p.keys.forEach(k => {
      if (k.alts.some(hitOne)) { hits.push(k.raw); if (!shared[k.id]) uniqueHits++; }
      else missing.push(k.raw);
    });
    return { code: p.a.code, name: p.a.name, hits: hits, missing: missing, total: p.keys.length, uniqueHits: uniqueHits };
  });
}

function matchAccount_(ext, accounts) {
  const parsed = accounts.map(a => ({ a: a, keys: a.keys.map(parseKey_).filter(Boolean) })).filter(p => p.keys.length);
  if (!parsed.length) return { code: '', reason: 'ยังไม่ได้ตั้งค่าข้อมูลบัญชีร้าน' };
  // รายละเอียดที่มีในหลายบัญชี (เช่น "พร้อมเพย์") ไม่ใช้เป็นเบาะแสว่าใกล้เคียงบัญชีไหน
  const cnt = {}, shared = {};
  parsed.forEach(p => { const seen = {}; p.keys.forEach(k => { if (!seen[k.id]) { seen[k.id] = 1; cnt[k.id] = (cnt[k.id] || 0) + 1; } }); });
  Object.keys(cnt).forEach(id => { if (cnt[id] > 1) shared[id] = 1; });

  const recv = [ext.receiverName, ext.receiverAccount, ext.receiverDetails, ext.receiverBank].join(' | ');
  const blobs = [recv];
  if (ext.fullText && ext.fullText !== recv) blobs.push(ext.fullText);
  let best = null;
  for (let i = 0; i < blobs.length; i++) {
    const r = evalAccounts_(blobs[i], parsed, shared);
    const full = r.filter(x => x.missing.length === 0);
    if (full.length === 1) return { code: full[0].code, reason: 'ข้อมูลครบตามบัญชี ' + full[0].code + ' ' + full[0].name };
    if (full.length > 1) {
      return { code: '', reason: 'ข้อมูลครบมากกว่า 1 บัญชี (' + full.map(x => x.code + ' ' + x.name).join(', ') +
                                 ') — อาจเป็นการโอนระหว่างบัญชีร้าน' };
    }
    r.forEach(x => { if (x.uniqueHits && (!best || x.uniqueHits > best.uniqueHits)) best = x; });
  }
  if (best) {
    return { code: '', partial: best.code,
             reason: 'ข้อมูลบัญชีไม่ครบ — ใกล้เคียง ' + best.code + ' ' + best.name + ' แต่ไม่พบ "' + best.missing.join('", "') + '"' };
  }
  return { code: '', reason: 'ไม่พบข้อมูลบัญชีร้านในสลิป' };
}

// ======================== SLIPS: บันทึก/อ่าน ========================
function slipFromRow_(r) {
  const o = {};
  HEADERS.Slips.forEach((h, i) => o[h] = r[i]);
  const amt = (o.Amount === '' || o.Amount === null || o.Amount === undefined) ? null : Number(o.Amount);
  const s = v => (v === undefined || v === null) ? '' : String(v);
  return {
    id: s(o.ID), batchId: s(o.BatchID),
    uploadedAt: fmtDateTime_(o.UploadedAt), uploadedBy: s(o.UploadedBy),
    date: normalizeDate_(o.TxnDate), time: normalizeTime_(o.TxnTime),
    account: s(o.AccountCode).toUpperCase(),
    amount: isFinite(amt) ? amt : null,
    receiverName: s(o.ReceiverName), receiverAccount: s(o.ReceiverAccount),
    receiverBank: s(o.ReceiverBank), receiverDetails: s(o.ReceiverDetails),
    senderName: s(o.SenderName), senderBank: s(o.SenderBank), refNo: s(o.RefNo),
    status: s(o.Status), note: s(o.Note), fileName: s(o.FileName),
    hasImage: !!o.ImageFileId, confidence: s(o.AIConfidence),
    editedBy: s(o.EditedBy), editedAt: fmtDateTime_(o.EditedAt),
    dupOf: s(o.DupOf), baseStatus: s(o.BaseStatus),
  };
}

function readAllSlips_() {
  const sh = sheet_(SH.SLIPS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.Slips.length).getValues()
    .filter(r => r[0] !== '').map(slipFromRow_);
}

function findSlipRows_(sh, header, value) {
  if (!value) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, col_('Slips', header), last - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).findAll()
    .map(rg => {
      const row = rg.getRow();
      return { row, slip: slipFromRow_(sh.getRange(row, 1, 1, HEADERS.Slips.length).getValues()[0]) };
    });
}

/** อ่านเฉพาะแถวของวันที่ที่ระบุ (อ่านช่วงแถวเดียว ไม่ต้องอ่านทั้งชีต) */
function slipsOnDate_(sh, date) {
  const last = sh.getLastRow();
  if (last < 2 || !date) return [];
  const vals = sh.getRange(2, col_('Slips', 'TxnDate'), last - 1, 1).getDisplayValues();
  let lo = -1, hi = -1;
  vals.forEach((v, i) => { if (String(v[0]).slice(0, 10) === date) { if (lo < 0) lo = i; hi = i; } });
  if (lo < 0) return [];
  return sh.getRange(2 + lo, 1, hi - lo + 1, HEADERS.Slips.length).getValues()
    .map((r, i) => ({ row: 2 + lo + i, slip: slipFromRow_(r) }))
    .filter(x => x.slip.date === date);
}

/** กุญแจตรวจซ้ำ = วันที่ + เวลา + จำนวนเงิน + บัญชีร้าน */
function dupKey_(date, time, amount, account) {
  return [date, time, (amount === null || amount === '') ? '' : Number(amount).toFixed(2), account || ''].join('|');
}
function findDuplicate_(sh, date, time, amount, account, excludeId, statuses) {
  if (!date || !time || amount === null) return null;
  const key = dupKey_(date, time, amount, account);
  const hit = slipsOnDate_(sh, date).find(x => x.slip.id !== excludeId &&
    statuses.indexOf(x.slip.status) >= 0 &&
    dupKey_(x.slip.date, x.slip.time, x.slip.amount, x.slip.account) === key);
  return hit ? hit.slip : null;
}

function appendSlip_(sh, obj) {
  const values = HEADERS.Slips.map(h => (obj[h] === undefined || obj[h] === null) ? '' : obj[h]);
  const fmts = HEADERS.Slips.map(h =>
    h === 'Amount' ? '#,##0.00' : (h === 'UploadedAt' || h === 'EditedAt') ? 'yyyy-mm-dd hh:mm:ss' : '@');
  const r = sh.getLastRow() + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  sh.getRange(r, 1, 1, values.length).setNumberFormats([fmts]).setValues([values]);
  return r;
}

function getImageFolder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('IMAGE_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(CFG.IMAGE_FOLDER);
  p.setProperty('IMAGE_FOLDER_ID', f.getId());
  return f;
}
function getMonthFolder_(month) {
  const p = PropertiesService.getScriptProperties();
  const key = 'IMG_' + month;
  const id = p.getProperty(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const root = getImageFolder_();
  const it = root.getFoldersByName(month);
  const f = it.hasNext() ? it.next() : root.createFolder(month);
  p.setProperty(key, f.getId());
  return f;
}

function newSlipId_() {
  return 'S' + Utilities.formatDate(new Date(), CFG.TZ, 'yyMMddHHmmss') +
         Utilities.getUuid().replace(/-/g, '').slice(0, 4).toUpperCase();
}

/** ตรวจความครบถ้วนของ 3 ส่วนหลัก → OK หรือ REVIEW พร้อมเหตุผล */
function assessSlip_(ext, match) {
  const issues = [];
  if (ext.amount === null) issues.push('อ่านจำนวนเงินไม่ได้');
  else if (ext.amount > CFG.MAX_AMOUNT) issues.push('จำนวนเงินสูงผิดปกติ');
  if (!ext.date) issues.push('อ่านวันที่ไม่ได้');
  else if (ext.date > addDays_(today_(), 1)) issues.push('วันที่อยู่ในอนาคต');
  if (!ext.time) issues.push('อ่านเวลาไม่ได้');
  if (!match.code) issues.push(match.reason);
  if (ext.confidence === 'low') issues.push(ext.remark || 'ระบบไม่มั่นใจในการอ่าน (ตรวจรูปอีกครั้ง)');
  return { status: issues.length ? 'REVIEW' : 'OK', note: issues.length ? issues.join(' • ') : match.reason };
}

/**
 * ขั้นที่ 1 — อ่านสลิป 1 รูป แล้วเก็บเป็น "ฉบับร่าง" (DRAFT) ยังไม่นับยอด
 * ผู้ใช้ตรวจ/แก้/ลบ ในหน้าเว็บก่อน แล้วกด "บันทึกข้อมูล" → apiCommitDrafts (ขั้นที่ 2)
 * payload: { batchId, fileName, mimeType, data }
 */
function apiProcessSlip(token, payload) {
  return run_(() => {
    const s = requireSession_(token);
    const mime = /^image\/(jpeg|png|webp)$/.test(payload.mimeType) ? payload.mimeType : 'image/jpeg';
    const bytes = Utilities.base64Decode(payload.data);
    if (bytes.length > 8 * 1024 * 1024) throw new Error('ไฟล์ใหญ่เกิน 8 MB');
    const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, bytes)
      .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
    const sh = sheet_(SH.SLIPS);
    const rec = {
      ID: newSlipId_(), BatchID: String(payload.batchId || ''), UploadedAt: now_(), UploadedBy: s.u,
      FileName: String(payload.fileName || '').slice(0, 200), ImageHash: hash, Status: 'DRAFT',
    };

    // รูปเดียวกันเคยอ่านแล้ว → ใช้ข้อมูลเดิม ไม่ต้องอ่านซ้ำ (การตรวจซ้ำจะเกิดตอนกดบันทึก)
    const prev = findSlipRows_(sh, 'ImageHash', hash).map(x => x.slip)
      .filter(x => x.status !== 'DISCARDED').sort((a, b) => (a.status === 'OK' ? -1 : 0) - (b.status === 'OK' ? -1 : 0))[0];
    let ext, match, assess;
    if (prev) {
      ext = { isSlip: prev.status !== 'NOT_SLIP', amount: prev.amount, date: prev.date, time: prev.time,
              confidence: 'high', remark: '' };
      match = prev.account ? { code: prev.account, reason: 'ใช้ข้อมูลจากรูปเดียวกัน (' + prev.id + ')' }
                           : { code: '', reason: 'รูปเดียวกับรายการ ' + prev.id + ' ซึ่งยังไม่มีบัญชีร้าน' };
      Object.assign(rec, { ReceiverName: prev.receiverName, ReceiverAccount: prev.receiverAccount,
        ReceiverDetails: prev.receiverDetails, SenderName: prev.senderName, RefNo: prev.refNo,
        RawJSON: JSON.stringify({ engine: 'reuse', of: prev.id }) });
    } else {
      const rd = readSlip_(payload.data, bytes, mime);
      ext = rd.ext;
      match = ext.isSlip ? matchAccount_(ext, listAccounts_(false)) : { code: '', reason: '' };
      Object.assign(rec, { ReceiverName: ext.receiverName, ReceiverAccount: ext.receiverAccount,
        ReceiverBank: ext.receiverBank, ReceiverDetails: ext.receiverDetails, SenderName: ext.senderName,
        SenderBank: ext.senderBank, RefNo: ext.refNo, AIConfidence: ext.confidence,
        RawJSON: JSON.stringify(rd.raw).slice(0, 4000) });
    }
    assess = ext.isSlip ? assessSlip_(ext, match)
      : { status: 'NOT_SLIP', note: 'ไม่ใช่สลิปโอนเงินที่สำเร็จ' + (ext.remark ? ' (' + ext.remark + ')' : '') };
    // วันที่ในสลิปไม่ตรงกับวันที่อัปโหลด → ให้ผู้ใช้ยืนยันก่อนนับยอด (อาจเป็นสลิปเก่า/ส่งซ้ำ/อ่านวันที่ผิด)
    const upDate = today_();
    if (ext.isSlip && ext.date && ext.date !== upDate) {
      const msg = 'วันที่ในสลิปไม่ตรงกับวันที่อัปโหลด (สลิป ' + ext.date + ', อัปโหลด ' + upDate + ')';
      assess = { status: 'REVIEW', note: assess.status === 'OK' ? msg : assess.note + ' • ' + msg };
    }
    Object.assign(rec, { TxnDate: ext.date, TxnTime: ext.time, AccountCode: match.code, Amount: ext.amount,
                         BaseStatus: assess.status, Note: assess.note });

    try {
      const file = getMonthFolder_((rec.TxnDate || today_()).slice(0, 7))
        .createFile(Utilities.newBlob(bytes, mime, (rec.TxnDate || 'nodate') + '_' + rec.ID + '.jpg'));
      rec.ImageFileId = file.getId();
    } catch (e) { rec.Note += ' • บันทึกรูปไม่สำเร็จ'; }
    withLock_(() => appendSlip_(sh, rec));
    const out = resultFromRecord_(rec, null);
    out.partial = match.partial || '';
    return out;
  });
}

/** ฉบับร่างที่ยังไม่ได้กดบันทึก (ของผู้ใช้คนนี้ / admin เห็นทั้งหมด) */
function apiGetDrafts(token) {
  return run_(() => {
    const s = requireSession_(token);
    return readAllSlips_().filter(x => x.status === 'DRAFT' && (s.r === 'admin' || x.uploadedBy === s.u))
      .sort((a, b) => a.uploadedAt.localeCompare(b.uploadedAt)).slice(0, 300);
  });
}

/**
 * ขั้นที่ 2 — บันทึกข้อมูล: ใช้การตัดสินใจของผู้ใช้กับฉบับร่าง แล้วตรวจซ้ำ
 * decisions: [{ id, action: 'keep' | 'fix' | 'delete', date, time, amount, account, refNo, reason }]
 *   keep   = บันทึกตามที่อ่านได้ (ใช้กับใบที่ผ่านเงื่อนไข)
 *   fix    = ผู้ใช้แก้ วันที่/เวลา/จำนวนเงิน/บัญชี ให้ถูกต้องแล้ว → ผ่านเงื่อนไข
 *   delete = ลบออก (ไม่เกี่ยวข้อง/เลือกรูปผิด) → DISCARDED + ย้ายรูปไปถังขยะ
 * ผลลัพธ์: OK / REVIEW / NOT_SLIP / DUP_PENDING (ซ้ำ → เด้ง popup ตรวจซ้ำต่อ)
 */
function apiCommitDrafts(token, decisions) {
  return run_(() => {
    const s = requireSession_(token);
    const codes = listAccounts_(true).map(a => a.code);
    decisions = (decisions || []).map(d => {          // ตรวจข้อมูลทั้งหมดก่อนเขียนอะไรลงชีต
      const o = { id: String(d.id), action: ['keep', 'fix', 'delete'].indexOf(d.action) >= 0 ? d.action : 'keep',
                  reason: String(d.reason || '').slice(0, 120) };
      if (o.action === 'fix') {
        o.date = normalizeDate_(d.date); o.time = normalizeTime_(d.time);
        o.amount = round2_(d.amount); o.account = String(d.account || '').toUpperCase();
        o.refNo = d.refNo === undefined ? undefined : String(d.refNo || '').replace(/\s+/g, '');
        if (!o.date || !o.time || !(o.amount > 0) || codes.indexOf(o.account) < 0) {
          throw new Error('ข้อมูลที่แก้ไขไม่ครบ (วันที่ เวลา จำนวนเงิน บัญชีร้าน) — รายการ ' + o.id);
        }
      }
      return o;
    });

    return withLock_(() => {
      const sh = sheet_(SH.SLIPS);
      const last = sh.getLastRow();
      const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, HEADERS.Slips.length).getValues();
      const C = h => HEADERS.Slips.indexOf(h);
      const byId = {}, keyMap = {}, hashMap = {}, refMap = {};
      const index = (x, hash) => {
        if (x.date && x.time && x.amount !== null) { const k = dupKey_(x.date, x.time, x.amount, x.account); if (!keyMap[k]) keyMap[k] = x.id; }
        if (hash && !hashMap[hash]) hashMap[hash] = x.id;
        if (x.refNo && x.amount !== null) { const k = x.refNo + '|' + x.amount; if (!refMap[k]) refMap[k] = x.id; }
      };
      rows.forEach((r, i) => {
        const x = slipFromRow_(r);
        byId[x.id] = i;
        if (ACTIVE_STATUSES.indexOf(x.status) >= 0) index(x, String(r[C('ImageHash')]));
      });

      // ตรวจสิทธิ์ทุกรายการก่อนเขียน (ไม่ให้บันทึกได้ครึ่งๆ กลางๆ)
      decisions.forEach(d => {
        const i = byId[d.id];
        if (i === undefined) return;
        const x = slipFromRow_(rows[i]);
        if (x.status === 'DRAFT' && s.r !== 'admin' && x.uploadedBy !== s.u) throw new Error('ไม่มีสิทธิ์บันทึกรายการ ' + d.id);
      });
      const results = [], sum = { ok: 0, review: 0, dup: 0, deleted: 0, notSlip: 0 };
      decisions.forEach(d => {
        const i = byId[d.id];
        if (i === undefined) return;
        const r = rows[i];
        const x0 = slipFromRow_(r);
        if (x0.status !== 'DRAFT') { results.push(x0); return; }     // บันทึกไปแล้ว
        r[C('EditedBy')] = s.u; r[C('EditedAt')] = now_();

        if (d.action === 'delete') {
          r[C('Status')] = 'DISCARDED';
          r[C('Note')] = ['ลบก่อนบันทึก: ' + (d.reason || 'ไม่เกี่ยวข้อง/เลือกรูปผิด'), x0.note].filter(Boolean).join(' • ').slice(0, 300);
          try { if (r[C('ImageFileId')]) DriveApp.getFileById(r[C('ImageFileId')]).setTrashed(true); } catch (e) {}
          sum.deleted++;
        } else {
          let status = String(r[C('BaseStatus')] || 'REVIEW');
          if (d.action === 'fix') {
            r[C('TxnDate')] = d.date; r[C('TxnTime')] = d.time; r[C('Amount')] = d.amount; r[C('AccountCode')] = d.account;
            if (d.refNo !== undefined) r[C('RefNo')] = d.refNo;
            r[C('Note')] = ('แก้ไขก่อนบันทึกโดย ' + s.u + (x0.note ? ' • เดิม: ' + x0.note : '')).slice(0, 300);
            status = 'OK';
          }
          r[C('BaseStatus')] = status;
          if (status === 'OK' || status === 'REVIEW') {
            const x = slipFromRow_(r), hash = String(r[C('ImageHash')]);
            const orig = (hash && hashMap[hash]) ||
                         (x.date && x.time && x.amount !== null && keyMap[dupKey_(x.date, x.time, x.amount, x.account)]) ||
                         (x.refNo && x.amount !== null && refMap[x.refNo + '|' + x.amount]) || '';
            if (orig) {
              r[C('Status')] = 'DUP_PENDING'; r[C('DupOf')] = orig;
              r[C('Note')] = (hash && hashMap[hash] === orig ? 'รูปเดียวกับรายการเดิม' : 'วันที่-เวลา จำนวนเงิน และบัญชี ตรงกับรายการเดิม').slice(0, 300);
              sum.dup++;
            } else {
              r[C('Status')] = status; index(x, hash);
              if (status === 'OK') sum.ok++; else sum.review++;
            }
          } else { r[C('Status')] = status; sum.notSlip++; }
        }
        sh.getRange(2 + i, 1, 1, HEADERS.Slips.length).setValues([r]);
        results.push(slipFromRow_(r));
      });
      log_(s.u, 'COMMIT', 'บันทึก ' + sum.ok + ', รอตรวจ ' + sum.review + ', ซ้ำ ' + sum.dup + ', ลบ ' + sum.deleted + ', ไม่ใช่สลิป ' + sum.notSlip);
      return { results: results, summary: sum };
    });
  });
}

function resultFromRecord_(rec, orig) {
  const r = slipFromRow_(HEADERS.Slips.map(h => rec[h] === undefined ? '' : rec[h]));
  r.dupOrig = orig || null;
  return r;
}

function findSlipById_(sh, id) {
  const hit = findSlipRows_(sh, 'ID', id)[0];
  if (!hit) throw new Error('ไม่พบรายการ ' + id);
  return hit;
}

// ======================== ข้อมูลซ้ำ: popup ยืนยัน ========================
/** รายการ "ซ้ำ? รอยืนยัน" พร้อมรายการต้นฉบับ (ids = เฉพาะรายการที่ระบุ, ว่าง = ทั้งหมด) */
function apiGetDupPending(token, ids) {
  return run_(() => {
    requireSession_(token);
    const all = readAllSlips_();
    const byId = {}; all.forEach(x => byId[x.id] = x);
    const want = (ids && ids.length) ? ids : null;
    return all.filter(x => x.status === 'DUP_PENDING' && (!want || want.indexOf(x.id) >= 0))
      .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
      .slice(0, 300)
      .map(x => ({ slip: x, orig: byId[x.dupOf] || null }));
  });
}

/**
 * ยืนยันผลการตรวจซ้ำ
 * decisions: [{ id, remove: true|false }]  remove=true → ตัดออก (DUPLICATE), false → เก็บไว้ (กลับเป็นสถานะเดิม)
 */
function apiResolveDuplicates(token, decisions) {
  return run_(() => {
    const s = requireSession_(token);
    return withLock_(() => {
      const sh = sheet_(SH.SLIPS);
      let removed = 0, kept = 0;
      (decisions || []).forEach(d => {
        const hit = findSlipRows_(sh, 'ID', d.id)[0];
        if (!hit || hit.slip.status !== 'DUP_PENDING') return;
        const set = (h, v) => sh.getRange(hit.row, col_('Slips', h)).setValue(v);
        if (d.remove) {
          set('Status', 'DUPLICATE');
          set('Note', ('ตัดออก: ซ้ำกับ ' + hit.slip.dupOf + ' • ' + hit.slip.note).slice(0, 300));
          removed++;
        } else {
          const back = hit.slip.baseStatus === 'OK' ? 'OK' : 'REVIEW';
          set('Status', back);
          set('Note', ('ยืนยันว่าไม่ซ้ำกับ ' + hit.slip.dupOf + ' • ' + hit.slip.note).slice(0, 300));
          kept++;
        }
        set('EditedBy', s.u); set('EditedAt', now_());
      });
      log_(s.u, 'DUP_RESOLVE', 'ตัดออก ' + removed + ', เก็บไว้ ' + kept);
      return { removed, kept };
    });
  });
}

// ======================== แก้ไข / ยืนยัน / ยกเลิก ========================
/** แก้ไข/ยืนยันข้อมูลสลิป (ผู้ใช้ทุกระดับ) → สถานะเป็น OK */
function apiUpdateSlip(token, id, ch) {
  return run_(() => {
    const s = requireSession_(token);
    const date = normalizeDate_(ch.date);
    if (!date) throw new Error('วันที่ไม่ถูกต้อง');
    const time = normalizeTime_(ch.time);
    if (!time) throw new Error('กรุณาใส่เวลาตามสลิป');
    const amount = round2_(ch.amount);
    if (!(amount > 0)) throw new Error('จำนวนเงินต้องมากกว่า 0');
    const account = String(ch.account || '').toUpperCase();
    if (!listAccounts_(true).some(a => a.code === account)) throw new Error('กรุณาเลือกบัญชีร้าน');
    const refNo = String(ch.refNo || '').replace(/\s+/g, '');
    const note = String(ch.note || '').slice(0, 300);

    return withLock_(() => {
      const sh = sheet_(SH.SLIPS);
      const hit = findSlipById_(sh, id);
      if (hit.slip.status === 'CANCELLED' && s.r !== 'admin') throw new Error('รายการนี้ถูกยกเลิกแล้ว');
      if (hit.slip.status === 'DRAFT' || hit.slip.status === 'DISCARDED') throw new Error('รายการนี้ยังไม่ได้บันทึก — ตรวจและกด "บันทึกข้อมูล" ในหน้าอัปโหลดก่อน');
      const dup = findDuplicate_(sh, date, time, amount, account, id, ['OK']);
      if (dup && !ch.allowDup) {
        throw err_('ข้อมูลนี้ซ้ำกับรายการที่ยืนยันแล้ว (' + dup.id + ' อัปโหลดโดย ' + dup.uploadedBy + ' ' + dup.uploadedAt +
                   (dup.refNo ? ', เลขอ้างอิง ' + dup.refNo : '') + ')', 'DUP');
      }
      const set = (h, v) => sh.getRange(hit.row, col_('Slips', h)).setValue(v);
      set('TxnDate', date); set('TxnTime', time); set('AccountCode', account);
      set('Amount', amount); set('RefNo', refNo);
      set('Note', note || (dup ? 'ยืนยันว่าไม่ซ้ำกับ ' + dup.id : 'ยืนยันโดยผู้ใช้'));
      set('Status', 'OK'); set('BaseStatus', 'OK'); set('EditedBy', s.u); set('EditedAt', now_());
      log_(s.u, 'SLIP_CONFIRM', id + ' ' + date + ' ' + time + ' ' + account + ' ' + amount +
           (hit.slip.amount !== amount ? ' (เดิม ' + hit.slip.amount + ')' : '') +
           (hit.slip.account !== account ? ' (บัญชีเดิม ' + (hit.slip.account || '-') + ')' : ''));
      return slipFromRow_(sh.getRange(hit.row, 1, 1, HEADERS.Slips.length).getValues()[0]);
    });
  });
}

/** ยกเลิกรายการ (admin) — ไม่ลบแถว เก็บไว้ตรวจย้อนหลัง */
function apiCancelSlip(token, id, reason) {
  return run_(() => {
    const s = requireSession_(token, true);
    return withLock_(() => {
      const sh = sheet_(SH.SLIPS);
      const hit = findSlipById_(sh, id);
      sh.getRange(hit.row, col_('Slips', 'Status')).setValue('CANCELLED');
      sh.getRange(hit.row, col_('Slips', 'Note')).setValue(('ยกเลิก: ' + String(reason || '-') + ' • ' + hit.slip.note).slice(0, 300));
      sh.getRange(hit.row, col_('Slips', 'EditedBy')).setValue(s.u);
      sh.getRange(hit.row, col_('Slips', 'EditedAt')).setValue(now_());
      log_(s.u, 'SLIP_CANCEL', id + ' ' + (reason || ''));
      return slipFromRow_(sh.getRange(hit.row, 1, 1, HEADERS.Slips.length).getValues()[0]);
    });
  });
}

function apiGetSlipImage(token, id) {
  return run_(() => {
    requireSession_(token);
    const sh = sheet_(SH.SLIPS);
    const hit = findSlipById_(sh, id);
    let fileId = sh.getRange(hit.row, col_('Slips', 'ImageFileId')).getValue();
    if (!fileId && hit.slip.dupOf) {                    // รูปซ้ำ → ใช้รูปของรายการต้นฉบับ
      const o = findSlipRows_(sh, 'ID', hit.slip.dupOf)[0];
      if (o) fileId = sh.getRange(o.row, col_('Slips', 'ImageFileId')).getValue();
    }
    if (!fileId) throw new Error('ไม่มีรูปของรายการนี้');
    const blob = DriveApp.getFileById(fileId).getBlob();
    // ข้อความที่ระบบอ่านได้ (ช่วย admin ดูว่าควรเพิ่มคำอะไรในบัญชีร้าน)
    let readText = '', engine = '';
    try {
      const raw = JSON.parse(sh.getRange(hit.row, col_('Slips', 'RawJSON')).getValue() || '{}');
      engine = String(raw.engine || '');
      readText = raw.text || raw.driveText || (raw.data ? JSON.stringify(raw.data, null, 1) : '');
    } catch (e) {}
    return { mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()),
             readText: String(readText).slice(0, 3000), engine: engine };
  });
}

/** จับคู่บัญชีใหม่ให้รายการ "รอตรวจสอบ" ที่ยังไม่มีบัญชี (เรียกอัตโนมัติหลัง admin แก้บัญชี) */
function rematchPending_(user) {
  return withLock_(() => {
    const sh = sheet_(SH.SLIPS);
    const last = sh.getLastRow();
    if (last < 2) return { checked: 0, matched: 0, confirmed: 0 };
    const accounts = listAccounts_(false);
    const rows = sh.getRange(2, 1, last - 1, HEADERS.Slips.length).getValues();
    let checked = 0, matched = 0, confirmed = 0;
    rows.forEach((r, i) => {
      const x = slipFromRow_(r);
      if (x.status !== 'REVIEW' || x.account) return;
      checked++;
      const ext = {
        receiverName: x.receiverName, receiverAccount: x.receiverAccount, receiverBank: x.receiverBank,
        receiverDetails: x.receiverDetails, amount: x.amount, date: x.date, time: x.time,
        confidence: x.confidence || 'high',
      };
      const m = matchAccount_(ext, accounts);
      if (!m.code) return;
      matched++;
      const row = i + 2;
      const a = assessSlip_(ext, m);
      let status = a.status, note = a.note;
      let dupOf = '';
      if (status === 'OK') {
        const d = findDuplicate_(sh, x.date, x.time, x.amount, m.code, x.id, ACTIVE_STATUSES);
        if (d) { status = 'DUP_PENDING'; dupOf = d.id; note = 'วันที่-เวลา จำนวนเงิน และบัญชี ตรงกับรายการเดิม'; }
        else confirmed++;
      }
      const set = (h, v) => sh.getRange(row, col_('Slips', h)).setValue(v);
      set('AccountCode', m.code); set('Status', status); set('Note', note);
      set('BaseStatus', a.status); if (dupOf) set('DupOf', dupOf);
    });
    if (matched) log_(user, 'REMATCH', 'checked ' + checked + ', matched ' + matched + ', confirmed ' + confirmed);
    return { checked, matched, confirmed };
  });
}

// ======================== สรุปรายวัน (ทุกสิทธิ์) ========================
function apiGetDaily(token, date) {
  return run_(() => {
    const s = requireSession_(token);
    date = normalizeDate_(date) || today_();
    const all = readAllSlips_();
    const accounts = listAccounts_(true);
    const day = all.filter(x => x.date === date && x.status !== 'DRAFT' && x.status !== 'DISCARDED')
      .sort((a, b) => (a.time || '99').localeCompare(b.time || '99'));

    const groups = {};
    accounts.forEach(a => groups[a.code] = {
      code: a.code, name: a.name, bank: a.bank, color: a.color, active: a.active,
      count: 0, total: 0, items: [],
    });
    day.filter(x => x.status === 'OK').forEach(x => {
      if (!groups[x.account]) groups[x.account] = {
        code: x.account || '?', name: '(ไม่อยู่ในรายการบัญชี)', bank: '', color: '#64748b', active: true,
        count: 0, total: 0, items: [],
      };
      const g = groups[x.account];
      g.count++; g.total += x.amount || 0;
      g.items.push({ id: x.id, time: x.time, amount: x.amount, sender: x.senderName, refNo: x.refNo });
    });
    const list = Object.keys(groups).map(k => groups[k])
      .filter(g => g.active || g.count > 0)
      .map(g => Object.assign(g, { total: round2_(g.total) }));

    const cnt = st => day.filter(x => x.status === st).length;
    return {
      date,
      accounts: list,
      grandCount: list.reduce((a, g) => a + g.count, 0),
      grandTotal: round2_(list.reduce((a, g) => a + g.total, 0)),
      records: day,
      counts: { review: cnt('REVIEW'), dupPending: cnt('DUP_PENDING'), duplicate: cnt('DUPLICATE'),
                notSlip: cnt('NOT_SLIP'), cancelled: cnt('CANCELLED') },
      pendingAll: all.filter(x => x.status === 'REVIEW').length,
      dupPendingAll: all.filter(x => x.status === 'DUP_PENDING').length,
      draftsMine: all.filter(x => x.status === 'DRAFT' && (s.r === 'admin' || x.uploadedBy === s.u)).length,
    };
  });
}

/** รายการรอตรวจสอบทั้งหมด (รวมที่อ่านวันที่ไม่ได้) */
function apiGetPending(token) {
  return run_(() => {
    requireSession_(token);
    return readAllSlips_().filter(x => x.status === 'REVIEW')
      .sort((a, b) => b.uploadedAt.localeCompare(a.uploadedAt)).slice(0, 300);
  });
}

// ======================== DASHBOARD (admin เท่านั้น) ========================
function checkRange_(from, to) {
  from = normalizeDate_(from); to = normalizeDate_(to);
  if (!from || !to) throw new Error('กรุณาเลือกช่วงวันที่');
  if (from > to) { const t = from; from = to; to = t; }
  const n = daysBetween_(from, to) + 1;
  if (n > CFG.DASHBOARD_MAX_DAYS) throw new Error('เลือกช่วงได้ไม่เกิน ' + CFG.DASHBOARD_MAX_DAYS + ' วัน');
  return { from, to, n };
}

function apiGetDashboard(token, from, to) {
  return run_(() => {
    requireSession_(token, true);
    const rg = checkRange_(from, to); from = rg.from; to = rg.to; const n = rg.n;

    const days = [];
    for (let i = 0; i < n; i++) days.push(addDays_(from, i));
    const idx = {}; days.forEach((d, i) => idx[d] = i);

    const all = readAllSlips_();
    const accounts = listAccounts_(true);
    const okIn = all.filter(x => x.status === 'OK' && x.date >= from && x.date <= to);

    const acc = {};
    accounts.forEach(a => acc[a.code] = {
      code: a.code, name: a.name, bank: a.bank, color: a.color, active: a.active,
      total: 0, count: 0, daily: days.map(() => 0), dailyCount: days.map(() => 0),
    });
    okIn.forEach(x => {
      if (!acc[x.account]) acc[x.account] = {
        code: x.account || '?', name: '(ไม่อยู่ในรายการบัญชี)', bank: '', color: '#64748b', active: true,
        total: 0, count: 0, daily: days.map(() => 0), dailyCount: days.map(() => 0),
      };
      const a = acc[x.account];
      a.total += x.amount || 0; a.count++;
      a.daily[idx[x.date]] += x.amount || 0; a.dailyCount[idx[x.date]]++;
    });
    const list = Object.keys(acc).map(k => acc[k]).filter(a => a.active || a.count > 0).map(a => {
      a.total = round2_(a.total);
      a.daily = a.daily.map(round2_);
      a.avgPerTxn = a.count ? round2_(a.total / a.count) : 0;
      return a;
    });

    const dayTotals = days.map((d, i) => round2_(list.reduce((s, a) => s + a.daily[i], 0)));
    const dayCounts = days.map((d, i) => list.reduce((s, a) => s + a.dailyCount[i], 0));
    const grandTotal = round2_(dayTotals.reduce((a, b) => a + b, 0));
    const grandCount = dayCounts.reduce((a, b) => a + b, 0);
    let maxI = -1; dayTotals.forEach((t, i) => { if (t > 0 && (maxI < 0 || t > dayTotals[maxI])) maxI = i; });

    const pTo = addDays_(from, -1), pFrom = addDays_(from, -n);
    const prev = all.filter(x => x.status === 'OK' && x.date >= pFrom && x.date <= pTo);

    const hours = Array.from({ length: 24 }, () => 0);
    okIn.forEach(x => { if (x.time) hours[Number(x.time.slice(0, 2))] += x.amount || 0; });

    return {
      from, to, days, accounts: list, dayTotals, dayCounts, grandTotal, grandCount,
      activeDays: dayTotals.filter(t => t > 0).length,
      maxDay: maxI >= 0 ? { date: days[maxI], total: dayTotals[maxI] } : null,
      prev: { from: pFrom, to: pTo, total: round2_(prev.reduce((s, x) => s + (x.amount || 0), 0)), count: prev.length },
      pendingInRange: all.filter(x => (x.status === 'REVIEW' || x.status === 'DUP_PENDING') && x.date >= from && x.date <= to).length,
      pendingNoDate: all.filter(x => x.status === 'REVIEW' && !x.date).length,
      hours: hours.map(round2_),
    };
  });
}

/** รายการที่ยืนยันแล้วในช่วงวันที่ สำหรับส่งออก Excel (admin) */
function apiExportRows(token, from, to) {
  return run_(() => {
    const s = requireSession_(token, true);
    const rg = checkRange_(from, to);
    const rows = readAllSlips_()
      .filter(x => x.status === 'OK' && x.date >= rg.from && x.date <= rg.to)
      .sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time))
      .map(x => ({ id: x.id, date: x.date, time: x.time, account: x.account, amount: x.amount,
                   sender: x.senderName, senderBank: x.senderBank, refNo: x.refNo,
                   uploadedBy: x.uploadedBy, uploadedAt: x.uploadedAt, editedBy: x.editedBy }));
    log_(s.u, 'EXPORT', rg.from + ' to ' + rg.to + ' (' + rows.length + ' rows)');
    return rows;
  });
}
