/**
 * ระบบใบลานักเรียน (Google Apps Script, container-bound กับ Google Sheet)
 * ข้อมูล: Google Sheet | รูปหลักฐาน: Google Drive | แจ้งเตือน: LINE Official Account | พิมพ์: PDF
 */

var COLS = ['id','createdAt','loginKey','name','grade','no','count','type','start','end','days','reason','guardian','photoId','status','note','decidedAt'];
var TYPES = { sick: 'ลาป่วย', personal: 'ลากิจ', other: 'ลาอื่น ๆ' };
var STATUS_TH = { pending: 'รออนุมัติ', approved: 'อนุมัติ', rejected: 'ไม่อนุมัติ' };
var FILTER_TH = { pending: 'รออนุมัติ', today: 'ลาวันนี้', all: 'ทั้งหมด', summary: 'สรุปรายคน' };
var MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
var TZ = 'Asia/Bangkok';
var MAX_PHOTO_B64 = 4500000; // ~3.3 MB หลังถอดรหัส
var MAX_PENDING = 5;         // ใบลาที่รออนุมัติได้พร้อมกันต่อหนึ่งชื่อ/เลขที่
var FONT = 'Sarabun';

/* ---------- เว็บแอป ---------- */

// LINE webhook: ใครส่งข้อความหาบอท จะถูกเพิ่มชื่อในชีต Notify (active = FALSE) ให้ครูเปลี่ยนเป็น TRUE เอง
function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var sh = sheet_('Notify');
    var have = {};
    var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) have[String(v[i][1]).trim()] = true;
    (body.events || []).forEach(function (ev) {
      var uid = ev && ev.source && ev.source.userId;
      if (!uid || have[uid] || sh.getLastRow() > 60) return;
      var name = '';
      try { name = JSON.parse(lineFetch_('get', '/v2/bot/profile/' + encodeURIComponent(uid)).getContentText()).displayName || ''; } catch (err) {}
      sh.appendRow([name, uid, 'FALSE']);
      have[uid] = true;
    });
  } catch (err) { log_('webhook-error', err.message); }
  return ContentService.createTextOutput('OK');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ใบลานักเรียน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- เมนูใน Google Sheet ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบใบลา')
    .addItem('1. ตั้งค่าเริ่มต้น (สร้างชีตและโฟลเดอร์ / ขอสิทธิ์)', 'setup')
    .addItem('2. ตั้งรหัสครู', 'setTeacherPin')
    .addItem('3. ตั้ง LINE Channel Access Token', 'setLineToken')
    .addItem('4. ทดสอบแจ้งเตือน LINE', 'testNotify')
    .addItem('ดูที่อยู่เว็บแอป', 'showUrl')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  var moved = '';
  var old = ss.getSheetByName('Leaves');
  if (old && old.getLastRow() > 0) {
    var head = old.getRange(1, 1, 1, Math.max(1, old.getLastColumn())).getValues()[0].slice(0, COLS.length).join(',');
    if (head !== COLS.join(',')) {
      moved = 'Leaves_เก่า_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
      old.setName(moved);
    }
  }
  var oldSt = ss.getSheetByName('Students');
  var movedSt = '';
  if (oldSt && String(oldSt.getRange(1, 1).getValue()).trim() === 'studentId') {
    movedSt = 'Students_เก่า_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
    oldSt.setName(movedSt);
  }
  ensure_(ss, 'Students', ['name', 'grade', 'no']);
  ensure_(ss, 'Leaves', COLS);
  ensure_(ss, 'Notify', ['name', 'userId', 'active']);
  ensure_(ss, 'Log', ['time', 'event', 'detail']);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('DRIVE_FOLDER_ID')) {
    props.setProperty('DRIVE_FOLDER_ID', DriveApp.createFolder('ใบลา-รูปภาพ').getId());
  }
  ui_('ตั้งค่าแล้ว', 'สร้างชีต Students, Leaves, Notify, Log และโฟลเดอร์รูปใน Drive เรียบร้อย'
    + (moved ? '\nชีต Leaves เดิมมีโครงสร้างคนละแบบกับเวอร์ชันนี้ จึงย้ายไปเก็บไว้ในชีต "' + moved + '" (ข้อมูลไม่หาย)' : '')
    + (movedSt ? '\nชีต Students เดิม (แบบมี PIN) ย้ายไปเก็บเป็นชีต "' + movedSt + '" ให้คัดลอกชื่อ ห้อง เลขที่ มาวางในชีต Students ใหม่' : '')
    + '\nขั้นต่อไป: ใส่รายชื่อนักเรียนในชีต Students (คอลัมน์ name, grade, no) ตั้งรหัสครู แล้ว Deploy เวอร์ชันใหม่');
}

function setTeacherPin() {
  var r = promptText_('ตั้งรหัสครู', 'พิมพ์รหัสผ่านสำหรับเข้าหน้าครู (อย่างน้อย 6 ตัว)');
  if (r === null) return;
  if (r.length < 6) return ui_('รหัสสั้นเกินไป', 'ใช้อย่างน้อย 6 ตัวอักษร');
  PropertiesService.getScriptProperties().setProperty('TEACHER_PIN', r);
  ui_('บันทึกแล้ว', 'ตั้งรหัสครูเรียบร้อย');
}

function setLineToken() {
  var r = promptText_('LINE Channel Access Token', 'วาง Channel access token (long-lived) ของ Messaging API channel');
  if (r === null) return;
  PropertiesService.getScriptProperties().setProperty('LINE_TOKEN', r);
  ui_('บันทึกแล้ว', 'เก็บ token ไว้ใน Script Properties (ไม่อยู่ในชีต)');
}

function showUrl() {
  var url = ScriptApp.getService().getUrl();
  ui_('ที่อยู่เว็บแอป', url || 'ยังไม่ได้ Deploy ให้ไปที่ Deploy > New deployment > Web app ก่อน');
}

function ensure_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // เก็บเป็นข้อความทั้งหมด กัน Sheets แปลงวันที่และตัดเลข 0 หน้ารหัส
  if (sh.getMaxRows() > 1) sh.getRange(2, 1, sh.getMaxRows() - 1, headers.length).setNumberFormat('@');
  return sh;
}

function ui_(title, text) { SpreadsheetApp.getUi().alert(title, text, SpreadsheetApp.getUi().ButtonSet.OK); }
function promptText_(title, text) {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText().trim() : null;
}

/* ---------- ตัวช่วย ---------- */

function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('ยังไม่ได้ตั้งค่าระบบ (เมนู ระบบใบลา > ตั้งค่าเริ่มต้น)');
  return sh;
}
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function str_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max); }
function normKey_(v) { return str_(v, 60).toLowerCase().replace(/\s+/g, ' '); }
function log_(event, detail) {
  try { sheet_('Log').appendRow([new Date().toISOString(), event, String(detail).slice(0, 500)]); } catch (e) {}
}
function pd_(s) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : NaN;
}
function weekdays_(a, b) {
  var x = pd_(a), y = pd_(b);
  if (isNaN(x) || isNaN(y) || y < x) return 0;
  var n = 0, span = Math.min(Math.round((y - x) / 864e5), 400);
  for (var i = 0; i <= span; i++) {
    var d = new Date(x + i * 864e5).getUTCDay();
    if (d !== 0 && d !== 6) n++;
  }
  return n;
}
function fmtThai_(s) {
  var t = new Date(pd_(s));
  return t.getUTCDate() + ' ' + MONTHS[t.getUTCMonth()] + ' ' + (t.getUTCFullYear() + 543);
}
function range_(a, b) { return a === b ? fmtThai_(a) : fmtThai_(a) + ' – ' + fmtThai_(b); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function isoThai_(iso) {
  try { return fmtThai_(Utilities.formatDate(new Date(iso), TZ, 'yyyy-MM-dd')); } catch (e) { return ''; }
}

// เพิ่มแถวเป็นข้อความล้วน กัน Sheets แปลงวันที่/ตัวเลขเอง (ใช้ได้แม้เกินจำนวนแถวเริ่มต้น)
function appendText_(sh, values) {
  var r = sh.getLastRow() + 1;
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
  var rng = sh.getRange(r, 1, 1, values.length);
  rng.setNumberFormat('@');
  rng.setValues([values]);
}

function toObj_(row) {
  var o = {};
  COLS.forEach(function (c, i) { o[c] = row[i] === undefined ? '' : String(row[i]); });
  o.days = Number(o.days) || 0;
  o.count = Number(o.count) || 0;
  o.hasPhoto = !!o.photoId;
  delete o.photoId;
  if (!o.status) o.status = 'pending';
  return o;
}
function photoIdOf_(vals) { return String(vals[COLS.indexOf('photoId')] || ''); }
function allRows_() {
  var v = sheet_('Leaves').getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) if (v[i][0]) out.push({ row: i + 1, vals: v[i] });
  return out;
}
function byNewest_(a, b) { return b.createdAt.localeCompare(a.createdAt); }

/* ---------- ครู: เข้าสู่ระบบ ---------- */

function teacherLogin(pin) {
  var real = prop_('TEACHER_PIN');
  if (!real) throw new Error('ยังไม่ได้ตั้งรหัสครู (เมนู ระบบใบลา > ตั้งรหัสครู)');
  var cache = CacheService.getScriptCache();
  if (Number(cache.get('t_fail') || 0) >= 5) throw new Error('ลองผิดหลายครั้งเกินไป รอ 10 นาทีแล้วลองใหม่');
  if (str_(pin, 60) !== real) {
    cache.put('t_fail', String(Number(cache.get('t_fail') || 0) + 1), 600);
    throw new Error('รหัสครูไม่ถูกต้อง');
  }
  cache.remove('t_fail');
  var token = Utilities.getUuid();
  cache.put('tk_' + token, '1', 21600);
  return token;
}
function requireTeacher_(token) {
  if (!token || !CacheService.getScriptCache().get('tk_' + token)) throw new Error('หมดเวลาเข้าสู่ระบบ ให้เข้าด้วยรหัสครูอีกครั้ง');
}

/* ---------- นักเรียน ---------- */

// รายชื่อนักเรียนจากชีต Students (คอลัมน์ name, grade, no) ถ้าชีตว่าง ระบบเปิดให้กรอกเอง
function roster_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Students');
  if (!sh) return [];
  var v = sh.getDataRange().getValues(), out = [], seen = {};
  for (var i = 1; i < v.length; i++) {
    var s = { name: str_(v[i][0], 80), grade: str_(v[i][1], 20), no: str_(v[i][2], 5) };
    if (!s.name || !s.grade || !s.no) continue;
    s.sid = sidOf_(s);
    if (seen[s.sid]) continue;
    seen[s.sid] = true;
    out.push(s);
  }
  return out;
}
function sidOf_(s) { return normKey_(s.name) + '|' + str_(s.grade, 20).toLowerCase() + '|' + str_(s.no, 5); }
function rosterFind_(roster, sid) {
  for (var i = 0; i < roster.length; i++) if (roster[i].sid === sid) return roster[i];
  return null;
}

// ข้อมูลที่คืนให้นักเรียน: สถานะใบลาของตัวเองเท่านั้น (ไม่คืนเหตุผล รูป หรือชื่อคนอื่น)
function session_(sid, st) {
  var mine = allRows_().map(function (r) { return toObj_(r.vals); }).filter(function (o) { return o.loginKey === sid; });
  var leaves = mine.map(function (o) {
    return { id: o.id, createdAt: o.createdAt, count: o.count, type: o.type, start: o.start, end: o.end, days: o.days, status: o.status, note: o.note };
  }).sort(byNewest_);
  var used = 0;
  mine.forEach(function (o) { if (o.status === 'approved') used += o.days; });
  return {
    mode: st ? 'roster' : 'manual',
    sid: sid,
    profile: st ? { name: st.name, grade: st.grade, no: st.no } : null,
    leaves: leaves,
    nextCount: mine.filter(function (o) { return o.status !== 'rejected'; }).length + 1,
    usedDays: used
  };
}

// เข้าใช้งานด้วย "ชื่อ หรือ เลขที่" แล้วระบบหาชื่อ ห้อง เลขที่ จากรายชื่อให้เอง
// ถ้าเจอหลายคน คืน mode 'choose' ให้เลือกห้อง แล้วเรียกซ้ำพร้อม grade
function studentLogin(key, grade) {
  key = normKey_(key);
  grade = str_(grade, 20).toLowerCase();
  if (!key) throw new Error('พิมพ์ชื่อหรือเลขที่ก่อน');
  var roster = roster_();
  if (!roster.length) return session_(key, null); // ครูยังไม่ใส่รายชื่อ: กรอกเอง
  var numeric = /^\d{1,3}$/.test(key);
  var m = roster.filter(function (s) { return numeric ? s.no === key : normKey_(s.name) === key; });
  if (!m.length && !numeric && key.length >= 2) m = roster.filter(function (s) { return normKey_(s.name).indexOf(key) >= 0; });
  if (grade) m = m.filter(function (s) { return s.grade.toLowerCase() === grade; });
  if (!m.length) throw new Error('ไม่พบ' + (numeric ? 'เลขที่' : 'ชื่อ') + 'นี้ในรายชื่อนักเรียน ตรวจตัวสะกดอีกครั้ง หรือแจ้งครูให้เพิ่มรายชื่อ');
  if (m.length > 1) {
    return { mode: 'choose', numeric: numeric, choices: m.slice(0, 12).map(function (s) { return numeric ? { grade: s.grade } : { grade: s.grade, name: s.name }; }) };
  }
  return session_(m[0].sid, m[0]);
}

// โหลดใบลาของตัวเองซ้ำหลังส่งหรือถอนใบลา
function studentSession(sid) {
  sid = str_(sid, 120);
  var roster = roster_();
  if (!roster.length) return session_(sid, null);
  var st = rosterFind_(roster, sid);
  if (!st) throw new Error('ไม่พบรายชื่อนักเรียนนี้ ให้เข้าใช้งานใหม่');
  return session_(sid, st);
}

function submitLeave(p) {
  var sid = str_(p.sid, 120);
  if (!sid) throw new Error('เข้าใช้งานด้วยชื่อหรือเลขที่ก่อน');
  var roster = roster_(), name, grade, no;
  if (roster.length) {
    var st = rosterFind_(roster, sid);
    if (!st) throw new Error('ไม่พบรายชื่อนักเรียนนี้ ให้เข้าใช้งานใหม่');
    name = st.name; grade = st.grade; no = st.no; // ใช้ข้อมูลจากรายชื่อเสมอ ไม่เชื่อค่าจากหน้าเว็บ
  } else {
    name = str_(p.name, 80); grade = str_(p.grade, 20); no = str_(p.no, 5);
    if (!name) throw new Error('กรอกชื่อ-สกุล');
    if (!grade) throw new Error('กรอกห้อง');
    if (!/^\d{1,3}$/.test(no)) throw new Error('เลขที่ต้องเป็นตัวเลข');
  }
  if (!TYPES[p.type]) throw new Error('เลือกประเภทการลา');
  var start = str_(p.start, 10), end = str_(p.end, 10);
  if (isNaN(pd_(start)) || isNaN(pd_(end))) throw new Error('วันที่ไม่ถูกต้อง');
  if (pd_(end) < pd_(start)) throw new Error('วันที่สิ้นสุดต้องไม่ก่อนวันเริ่มลา');
  var days = weekdays_(start, end); // นับวันเรียนฝั่งเซิร์ฟเวอร์ ไม่รวมเสาร์–อาทิตย์
  if (!days) throw new Error('ช่วงวันที่นี้เป็นวันเสาร์–อาทิตย์ทั้งหมด ไม่ต้องยื่นใบลา');
  var reason = str_(p.reason, 500);
  if (!reason) throw new Error('เขียนสาเหตุการลาสั้น ๆ');
  var guardian = str_(p.guardian, 80);
  if (!p.photo) throw new Error('ต้องแนบรูปหลักฐานการลา (ถ่ายรูปหรือเลือกรูป)');
  if (String(p.photo).length > MAX_PHOTO_B64) throw new Error('รูปใหญ่เกินไป ลองถ่ายใหม่หรือเลือกรูปที่เล็กกว่านี้');
  var bytes = Utilities.base64Decode(p.photo);
  if (!(bytes[0] === -1 && bytes[1] === -40)) throw new Error('รองรับเฉพาะรูป JPEG');

  var mine = allRows_().map(function (r) { return toObj_(r.vals); }).filter(function (o) { return o.loginKey === sid; });
  if (mine.filter(function (o) { return o.status === 'pending'; }).length >= MAX_PENDING) {
    throw new Error('มีใบลารออนุมัติอยู่ ' + MAX_PENDING + ' ใบแล้ว รอครูพิจารณาก่อนหรือถอนใบเก่า');
  }
  var count = mine.filter(function (o) { return o.status !== 'rejected'; }).length + 1;

  var id = 'L' + Utilities.getUuid().slice(0, 8);
  var folder = DriveApp.getFolderById(prop_('DRIVE_FOLDER_ID'));
  var photoId = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', id + '_' + no + '.jpg')).getId();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    appendText_(sheet_('Leaves'), [id, new Date().toISOString(), sid, name, grade, no, String(count),
      p.type, start, end, String(days), reason, guardian, photoId, 'pending', '', '']);
  } finally { lock.releaseLock(); }

  var text = '📝 ใบลาใหม่ (ครั้งที่ ' + count + ')\n' + name + ' ห้อง ' + grade + ' เลขที่ ' + no + '\n'
    + TYPES[p.type] + ' ' + range_(start, end) + ' (' + days + ' วัน)\n'
    + 'สาเหตุ: ' + reason + '\n📎 แนบรูปหลักฐานแล้ว'
    + '\nอนุมัติที่: ' + (ScriptApp.getService().getUrl() || '');
  var notified = { sent: 0, failed: 0 };
  try { notified = notifyTeachers_(text); } catch (e) { log_('notify-error', e.message); }
  return { ok: true, id: id, notified: notified.sent > 0 };
}

function withdrawLeave(sid, id) {
  sid = str_(sid, 120);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var rows = allRows_();
    for (var i = 0; i < rows.length; i++) {
      var o = toObj_(rows[i].vals);
      if (o.id === id && o.loginKey === sid) {
        if (o.status !== 'pending') throw new Error('ใบลานี้ครูตัดสินแล้ว ถอนไม่ได้');
        var pid = photoIdOf_(rows[i].vals);
        sheet_('Leaves').deleteRow(rows[i].row);
        if (pid) { try { DriveApp.getFileById(pid).setTrashed(true); } catch (e) {} }
        return { ok: true };
      }
    }
  } finally { lock.releaseLock(); }
  throw new Error('ไม่พบใบลานี้');
}

/* ---------- ครู ---------- */

function getAllLeaves(token) {
  requireTeacher_(token);
  return allRows_().map(function (r) { return toObj_(r.vals); }).sort(byNewest_);
}

function getPhoto(token, id) {
  requireTeacher_(token);
  var rows = allRows_();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].vals[0]) === id) {
      var pid = photoIdOf_(rows[i].vals);
      if (!pid) throw new Error('ใบลานี้ไม่มีรูป');
      return 'data:image/jpeg;base64,' + Utilities.base64Encode(DriveApp.getFileById(pid).getBlob().getBytes());
    }
  }
  throw new Error('ไม่พบใบลานี้');
}

function decideLeave(token, id, status, note) {
  requireTeacher_(token);
  if (status !== 'approved' && status !== 'rejected') throw new Error('สถานะไม่ถูกต้อง');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var rows = allRows_();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i].vals[0]) === id) {
        var rng = sheet_('Leaves').getRange(rows[i].row, COLS.indexOf('status') + 1, 1, 3);
        rng.setNumberFormat('@');
        rng.setValues([[status, str_(note, 120), new Date().toISOString()]]);
        return { ok: true };
      }
    }
  } finally { lock.releaseLock(); }
  throw new Error('ไม่พบใบลานี้ (นักเรียนอาจถอนไปแล้ว)');
}

/* ---------- PDF ---------- */

function filterRows_(rows, filter, q) {
  var t = today_();
  q = String(q || '').trim().toLowerCase();
  return rows.filter(function (r) {
    if (q && ((r.name || '') + ' ' + (r.grade || '')).toLowerCase().indexOf(q) < 0) return false;
    if (filter === 'pending') return r.status === 'pending';
    if (filter === 'today') return r.status === 'approved' && r.start <= t && t <= r.end;
    return true;
  });
}

function summaryRows_(rows, q) {
  var m = {};
  filterRows_(rows, 'all', q).forEach(function (r) {
    if (r.status !== 'approved') return;
    var k = [r.name, r.grade, r.no].join('|');
    var x = m[k] || (m[k] = { name: r.name, grade: r.grade, no: r.no, sick: 0, personal: 0, other: 0 });
    x[r.type] = (x[r.type] || 0) + r.days;
  });
  return Object.keys(m).map(function (k) { return m[k]; })
    .sort(function (a, b) { return a.name.localeCompare(b.name, 'th'); });
}

function pdfFolder_() {
  var root = DriveApp.getFolderById(prop_('DRIVE_FOLDER_ID'));
  var it = root.getFoldersByName('PDF');
  return it.hasNext() ? it.next() : root.createFolder('PDF');
}

function Builder_(doc, landscape) {
  this.body = doc.getBody();
  this.first = true;
  this.body.clear();
  this.body.setPageWidth(landscape ? 842 : 595);
  this.body.setPageHeight(landscape ? 595 : 842);
  this.body.setMarginTop(40);
  this.body.setMarginBottom(40);
  this.body.setMarginLeft(landscape ? 36 : 56);
  this.body.setMarginRight(landscape ? 36 : 56);
}
Builder_.prototype.p = function (text, o) {
  o = o || {};
  var p;
  text = text === '' ? ' ' : String(text);
  if (this.first) { p = this.body.getParagraphs()[0]; p.setText(text); this.first = false; }
  else p = this.body.appendParagraph(text);
  var t = p.editAsText();
  t.setFontFamily(FONT);
  t.setFontSize(o.size || 12);
  t.setBold(!!o.bold);
  p.setAlignment(o.align || DocumentApp.HorizontalAlignment.LEFT);
  p.setSpacingAfter(o.after == null ? 6 : o.after);
  return p;
};
Builder_.prototype.table = function (data, widths) {
  var cells = data.map(function (row) { return row.map(function (c) { return c === '' || c == null ? ' ' : String(c); }); });
  var tb = this.body.appendTable(cells);
  var t = tb.editAsText();
  t.setFontFamily(FONT);
  t.setFontSize(10);
  for (var c = 0; c < widths.length; c++) tb.setColumnWidth(c, widths[c]);
  var head = tb.getRow(0);
  head.editAsText().setBold(true);
  for (var i = 0; i < head.getNumCells(); i++) head.getCell(i).setBackgroundColor('#dfe5f5');
  return tb;
};

function buildListDoc_(doc, rows, filter, q) {
  var b = new Builder_(doc, true);
  var C = DocumentApp.HorizontalAlignment.CENTER;
  b.p('รายงานใบลานักเรียน', { size: 20, bold: true, align: C, after: 2 });
  b.p('รายการ: ' + FILTER_TH[filter] + (q ? ' · ค้นหา "' + q + '"' : '') + ' · ' + rows.length + ' ใบ · พิมพ์เมื่อ ' + fmtThai_(today_()), { size: 11, align: C, after: 10 });
  var data = [['ที่', 'ชื่อ-สกุล', 'ห้อง / เลขที่', 'ประเภท', 'วันที่ลา', 'วัน', 'เหตุผล', 'สถานะ']];
  rows.forEach(function (r, i) {
    data.push([String(i + 1), r.name, r.grade + ' / ' + r.no, TYPES[r.type] || r.type, range_(r.start, r.end), String(r.days), r.reason, STATUS_TH[r.status] || r.status]);
  });
  b.table(data, [30, 140, 80, 70, 150, 35, 190, 70]);
  if (!rows.length) b.p('ไม่มีรายการในตัวกรองนี้', { size: 11, after: 0 });
}

function buildSummaryDoc_(doc, rows, q) {
  var b = new Builder_(doc, false);
  var C = DocumentApp.HorizontalAlignment.CENTER;
  var list = summaryRows_(rows, q);
  b.p('สรุปวันลารายบุคคล', { size: 20, bold: true, align: C, after: 2 });
  b.p('นับเฉพาะใบลาที่อนุมัติแล้ว' + (q ? ' · ค้นหา "' + q + '"' : '') + ' · พิมพ์เมื่อ ' + fmtThai_(today_()), { size: 11, align: C, after: 10 });
  var data = [['ที่', 'ชื่อ-สกุล', 'ห้อง / เลขที่', 'ป่วย', 'กิจ', 'อื่น ๆ', 'รวม (วัน)']];
  list.forEach(function (x, i) {
    data.push([String(i + 1), x.name, x.grade + ' / ' + x.no, String(x.sick), String(x.personal), String(x.other), String(x.sick + x.personal + x.other)]);
  });
  b.table(data, [30, 170, 90, 50, 50, 50, 60]);
  if (!list.length) b.p('ยังไม่มีใบลาที่อนุมัติ', { size: 11, after: 0 });
}

function buildSlipDoc_(doc, r, photoId) {
  var b = new Builder_(doc, false);
  var C = DocumentApp.HorizontalAlignment.CENTER, R = DocumentApp.HorizontalAlignment.RIGHT;
  b.p('ใบลา', { size: 22, bold: true, align: C, after: 10 });
  b.p('ลาครั้งที่ ' + (r.count || 1) + '        วันที่ ' + isoThai_(r.createdAt), { align: R, after: 10 });
  b.p('เรียน คุณครูที่ปรึกษาและครูผู้สอน', { after: 10 });
  b.p('ข้าพเจ้า ' + r.name + '  ห้อง ' + r.grade + '  เลขที่ ' + r.no, { after: 8 });
  b.p('มีความประสงค์ขออนุญาต ' + (TYPES[r.type] || r.type), { after: 8 });
  b.p('เนื่องจากสาเหตุ: ' + r.reason, { after: 8 });
  b.p('ในวันที่ ' + range_(r.start, r.end) + ' รวมจำนวนวันทั้งสิ้น ' + r.days + ' วัน', { after: 8 });
  if (r.guardian) b.p('ผู้ปกครองที่ติดต่อได้: ' + r.guardian, { size: 11, after: 14 });
  b.p('จึงเรียนมาเพื่อโปรดพิจารณาอนุญาต', { align: C, after: 4 });
  b.p('ด้วยความเคารพเป็นอย่างสูง', { align: C, after: 22 });
  b.p('(ลงชื่อ) ..............................................', { align: C, after: 2 });
  b.p('[' + r.name + ']', { align: C, after: 16 });
  b.body.appendHorizontalRule();
  b.p('ผลการพิจารณา: ' + (STATUS_TH[r.status] || r.status) + (r.decidedAt ? ' เมื่อ ' + isoThai_(r.decidedAt) : ''), { bold: true, after: 4 });
  if (r.note) b.p('หมายเหตุจากครู: ' + r.note, { size: 11, after: 4 });
  b.p('ลงชื่อ ..............................................  ครูที่ปรึกษา', { align: R, after: 14 });
  b.p('รูปหลักฐาน', { bold: true, after: 4 });
  if (photoId) {
    var img = b.body.appendImage(DriveApp.getFileById(photoId).getBlob());
    var w = img.getWidth(), h = img.getHeight(), k = Math.min(1, 300 / w, 330 / h);
    img.setWidth(Math.round(w * k));
    img.setHeight(Math.round(h * k));
  } else {
    b.p('ไม่มีรูปหลักฐาน', { size: 11, after: 0 });
  }
}

// opts: {kind:'list'|'summary'|'slip', filter, q, id}
function makePdf(token, opts) {
  requireTeacher_(token);
  opts = opts || {};
  var kind = opts.kind === 'slip' || opts.kind === 'summary' ? opts.kind : 'list';
  var filter = FILTER_TH[opts.filter] ? opts.filter : 'all';
  var q = str_(opts.q, 60);
  var raw = allRows_();
  var rows = raw.map(function (r) { return toObj_(r.vals); }).sort(byNewest_);
  var doc = null, name;
  try {
    doc = DocumentApp.create('tmp-leave-pdf');
    if (kind === 'slip') {
      var found = null, photoId = '';
      for (var i = 0; i < raw.length; i++) if (String(raw[i].vals[0]) === opts.id) { found = toObj_(raw[i].vals); photoId = photoIdOf_(raw[i].vals); }
      if (!found) throw new Error('ไม่พบใบลานี้ (นักเรียนอาจถอนไปแล้ว)');
      buildSlipDoc_(doc, found, photoId);
      name = 'ใบลา_' + found.name + '_' + found.start;
    } else if (kind === 'summary') {
      buildSummaryDoc_(doc, rows, q);
      name = 'สรุปวันลา_' + today_();
    } else {
      buildListDoc_(doc, filterRows_(rows, filter, q), filter, q);
      name = 'รายงานใบลา_' + FILTER_TH[filter] + '_' + today_();
    }
    name = name.replace(/[\\\/:*?"<>|]/g, '-');
    doc.saveAndClose();
    var blob = DriveApp.getFileById(doc.getId()).getAs('application/pdf').setName(name + '.pdf');
    var saved = pdfFolder_().createFile(blob);
    return { name: name + '.pdf', data: Utilities.base64Encode(blob.getBytes()), driveUrl: saved.getUrl() };
  } finally {
    if (doc) { try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch (e) {} }
  }
}

/* ---------- LINE ---------- */

function lineFetch_(method, path, payload) {
  var token = prop_('LINE_TOKEN');
  if (!token) throw new Error('ยังไม่ได้ตั้ง LINE Channel Access Token');
  var o = { method: method, muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + token } };
  if (payload) { o.contentType = 'application/json'; o.payload = JSON.stringify(payload); }
  var r = UrlFetchApp.fetch('https://api.line.me' + path, o);
  if (r.getResponseCode() >= 300) {
    var m = '';
    try { m = JSON.parse(r.getContentText()).message; } catch (e) {}
    throw new Error(m || ('HTTP ' + r.getResponseCode()));
  }
  return r;
}

function sendLine_(to, text) {
  return lineFetch_('post', '/v2/bot/message/push', { to: to, messages: [{ type: 'text', text: text.slice(0, 4800) }] });
}

function notifyTeachers_(text) {
  var sent = 0, failed = 0;
  var v = sheet_('Notify').getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    var to = String(v[i][1]).trim();
    if (!to || String(v[i][2]).trim().toUpperCase() !== 'TRUE') continue;
    try { sendLine_(to, text); sent++; }
    catch (e) { failed++; log_('line-fail', String(v[i][0]) + ': ' + e.message); }
  }
  return { sent: sent, failed: failed };
}

function testNotify() {
  try {
    var r = notifyTeachers_('✅ ทดสอบแจ้งเตือนจากระบบใบลานักเรียน');
    ui_('ผลทดสอบ', 'ส่งสำเร็จ ' + r.sent + ' คน, ไม่สำเร็จ ' + r.failed + ' คน' + (r.failed ? '\nดูสาเหตุในชีต Log' : '') + (!r.sent && !r.failed ? '\nยังไม่มีแถวที่ active เป็น TRUE ในชีต Notify' : ''));
  } catch (e) { ui_('ทดสอบไม่สำเร็จ', e.message); }
}
