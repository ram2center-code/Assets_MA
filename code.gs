// Google Apps Script — วางใน Extensions > Apps Script ของ Google Sheet
const TOKEN = 'AssetsRam2#10240'; // ต้องตรงกับช่อง Token ในหน้าเว็บ
const SHEET_NAME = 'Assets';
const HEADERS = ['ID', 'Tag', 'S/N', 'แผนก', 'ผู้บันทึก', 'วันเวลาที่บันทึก', 'เวลาส่งเข้าชีต'];

// ---- ตั้งค่า Master ----
const MASTER_FILE_ID = '';   // ว่าง = Master เป็นแท็บในไฟล์นี้ / ถ้าอยู่คนละไฟล์ ใส่ ID จาก URL ของไฟล์นั้น
const MASTER_SHEET = 'Master-Old'; // ชื่อแท็บ Master
const M_TAG = 'TAG';   // ชื่อหัวคอลัมน์แถวแรกของ Master (ไม่สนตัวพิมพ์เล็ก/ใหญ่)
const M_SN = 'S/N';
const M_DEPT = 'แผนก';  // แผนกเดิมใน Master (ไม่มีคอลัมน์นี้ก็ได้)

// ---- รายชื่อแผนก (แท็บ "แผนก" ในไฟล์นี้ ดึงจากคอลัมน์ที่กำหนด) ----
const DEPT_SHEET = 'แผนก';
const DEPT_COL = 2;           // 2 = คอลัมน์ B (A = ID, B = ชื่อแผนก)

// ---- รายชื่อผู้บันทึก (แท็บ "User" ในไฟล์นี้) ----
const USER_SHEET = 'User';
const USER_COL = 2;           // คอลัมน์ที่เก็บชื่อผู้บันทึก (2 = B ตามชีตของคุณ: A = ID, B = Name)

// หาแท็บโดยไม่สนตัวพิมพ์เล็ก/ใหญ่และช่องว่างหัวท้าย
function getTab_(ss, name) {
  const want = String(name).trim().toLowerCase();
  return ss.getSheets().find(sh => sh.getName().trim().toLowerCase() === want) || null;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = getTab_(ss, SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange('A:F').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ตัดช่องว่าง แปลงเป็นตัวพิมพ์ใหญ่ และกรองค่าว่างเทียม เช่น "-", "N/A", "ไม่มี"
function norm_(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase().replace(/\s+/g, '');
  return /^(-+|N\/?A|NONE|NULL|ไม่มี)$/i.test(s) ? '' : s;
}

function readMaster_() {
  const ss = MASTER_FILE_ID ? SpreadsheetApp.openById(MASTER_FILE_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const sh = getTab_(ss, MASTER_SHEET);
  if (!sh) throw new Error('ไม่พบแท็บ "' + MASTER_SHEET + '"');
  const vals = sh.getDataRange().getDisplayValues();
  if (!vals.length) throw new Error('แท็บ "' + MASTER_SHEET + '" ไม่มีข้อมูล');
  const head = vals[0].map(h => String(h).trim());
  const H = head.map(h => h.toUpperCase());
  const iTag = H.indexOf(M_TAG.toUpperCase()), iSn = H.indexOf(M_SN.toUpperCase()), iDept = H.indexOf(M_DEPT.toUpperCase());
  if (iTag < 0 && iSn < 0) throw new Error('ไม่พบหัวคอลัมน์ "' + M_TAG + '" หรือ "' + M_SN + '" ในแท็บ Master');
  const rows = vals.slice(1).filter(r => r.some(c => String(c).trim() !== ''));
  return { head, rows, iTag, iSn, iDept };
}

function readDepts_() {
  const sh = getTab_(SpreadsheetApp.getActiveSpreadsheet(), DEPT_SHEET);
  if (!sh) throw new Error('ไม่พบแท็บ "' + DEPT_SHEET + '"');
  const n = Math.max(sh.getLastRow(), 1);
  const col = sh.getLastColumn() >= DEPT_COL ? DEPT_COL : 1;
  const vals = sh.getRange(1, col, n, 1).getDisplayValues().flat().map(v => String(v).trim()).filter(Boolean);
  if (vals.length && /^(id|รหัส|รหัสแผนก|แผนก|ชื่อแผนก|department|dept|name)$/i.test(vals[0])) vals.shift();
  return Array.from(new Set(vals));
}

function readUsers_() {
  const sh = getTab_(SpreadsheetApp.getActiveSpreadsheet(), USER_SHEET);
  if (!sh) throw new Error('ไม่พบแท็บ "' + USER_SHEET + '"');
  const n = Math.max(sh.getLastRow(), 1);
  const vals = sh.getRange(1, USER_COL, n, 1).getDisplayValues().flat().map(v => String(v).trim()).filter(Boolean);
  if (vals.length && /^(user|users|name|ชื่อ|ชื่อผู้ใช้|ผู้บันทึก|ชื่อผู้บันทึก)$/i.test(vals[0])) vals.shift();
  return Array.from(new Set(vals));
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if ((p.token || '') !== TOKEN) return out_({ ok: false, error: 'Token ไม่ถูกต้อง' });
  if (p.action === 'lists') { // ส่งรายชื่อแผนกและผู้บันทึกในครั้งเดียว (แท็บไหนไม่มีก็ข้าม)
    const res = { ok: true, depts: [], users: [], errors: [] };
    try { res.depts = readDepts_(); } catch (err) { res.errors.push(String(err.message || err)); }
    try { res.users = readUsers_(); } catch (err) { res.errors.push(String(err.message || err)); }
    return out_(res);
  }
  if (p.action === 'depts') {
    try { return out_({ ok: true, depts: readDepts_() }); }
    catch (err) { return out_({ ok: false, error: String(err.message || err) }); }
  }
  if (p.action === 'lookup') {
    try {
      const m = readMaster_(), t = norm_(p.tag), s = norm_(p.sn);
      if (!t && !s) return out_({ ok: true, found: false });
      const hit = m.rows.find(r =>
        (t && m.iTag >= 0 && norm_(r[m.iTag]) === t) || (s && m.iSn >= 0 && norm_(r[m.iSn]) === s));
      if (!hit) return out_({ ok: true, found: false });
      const info = m.head.map((h, i) => [h, hit[i], i])
        .filter(x => x[2] !== m.iTag && String(x[1]).trim() !== '')
        .slice(0, 4).map(x => (x[2] === m.iDept ? 'แผนกเดิม' : x[0]) + ': ' + x[1]).join(' | ');
      const dept = m.iDept >= 0 ? String(hit[m.iDept] || '').trim() : '';
      return out_({ ok: true, found: true, info: info, dept: dept });
    } catch (err) {
      return out_({ ok: false, error: String(err.message || err) });
    }
  }
  return out_({ ok: true, rows: Math.max(getSheet_().getLastRow() - 1, 0) });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return out_({ ok: false, error: 'Token ไม่ถูกต้อง' });
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = getSheet_();
      const last = sh.getLastRow();
      const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().flat().map(String) : [];
      const seen = new Set(ids);
      const rows = (body.records || [])
        .filter(r => !seen.has(String(r.id)))
        .map(r => [r.id, r.tag, r.sn, r.dept, r.rec, String(r.dt).replace('T', ' '), new Date()]);
      if (rows.length) {
        sh.getRange(last + 1, 1, rows.length, 6).setNumberFormat('@');
        sh.getRange(last + 1, 1, rows.length, 1).offset(0, 6).setNumberFormat('yyyy-mm-dd hh:mm:ss');
        sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
      return out_({ ok: true, added: rows.length });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// ---- เมนูเทียบกับ Master ----
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📦 Asset').addItem('เทียบกับ Master', 'compareWithMaster').addToUi();
}

function compareWithMaster() {
  const ui = SpreadsheetApp.getUi();
  let m;
  try { m = readMaster_(); } catch (err) { ui.alert(err.message); return; }

  const a = getSheet_(), last = a.getLastRow();
  const scans = last > 1 ? a.getRange(2, 1, last - 1, 6).getDisplayValues() : [];
  const byTag = new Map(), bySn = new Map();
  scans.forEach((r, i) => { // ใช้การสแกนล่าสุดของแต่ละเลข
    const t = norm_(r[1]), s = norm_(r[2]);
    if (t) byTag.set(t, i);
    if (s) bySn.set(s, i);
  });

  const masterTags = new Set(), masterSns = new Set();
  const matchedScanIdx = new Set();
  const data = [], colors = [];
  let found = 0, missing = 0, deptDiff = 0;

  m.rows.forEach(r => {
    const mt = m.iTag >= 0 ? r[m.iTag] : '', ms = m.iSn >= 0 ? r[m.iSn] : '', md = m.iDept >= 0 ? r[m.iDept] : '';
    const t = norm_(mt), s = norm_(ms);
    if (t) masterTags.add(t);
    if (s) masterSns.add(s);
    const idx = t && byTag.has(t) ? byTag.get(t) : (s && bySn.has(s) ? bySn.get(s) : -1);
    if (idx < 0) {
      missing++;
      data.push(['✘ ยังไม่พบ', mt, ms, md, '', '', '']);
      colors.push('#f8d7da');
    } else {
      matchedScanIdx.add(idx);
      const sc = scans[idx];
      const diff = m.iDept >= 0 && norm_(md) !== norm_(sc[3]);
      if (diff) { deptDiff++; } else { found++; }
      data.push([diff ? '🔄 ย้ายแผนก' : '✔ พบแล้ว', mt, ms, md, sc[3], sc[4], sc[5]]);
      colors.push(diff ? '#fff3cd' : '#d4edda');
    }
  });

  let extra = 0;
  const seenExtra = new Set();
  // วนจากท้ายสุดเพื่อให้ได้ข้อมูลสแกนล่าสุดของรายการนอก Master และไม่แสดงซ้ำ
  for (let i = scans.length - 1; i >= 0; i--) {
    if (matchedScanIdx.has(i)) continue;
    const sc = scans[i];
    const t = norm_(sc[1]), s = norm_(sc[2]);
    if ((t && masterTags.has(t)) || (s && masterSns.has(s))) continue;
    const key = (t || '') + '|' + (s || '');
    if (key !== '|' && seenExtra.has(key)) continue;
    if (key !== '|') seenExtra.add(key);
    extra++;
    data.push(['⚠ ไม่อยู่ใน Master', sc[1], sc[2], '', sc[3], sc[4], sc[5]]);
    colors.push('#ffe5b4');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Compare') || ss.insertSheet('Compare');
  sh.clear();
  const head = ['สถานะ', 'Tag', 'S/N', 'แผนกเดิม (Master-old)', 'แผนกปัจจุบัน (สแกน)', 'ผู้บันทึก', 'วันเวลาที่สแกน'];
  sh.getRange(1, 1, 1, 7).setValues([head]).setFontWeight('bold').setBackground('#e0e7ff');
  sh.setFrozenRows(1);
  if (data.length) {
    sh.getRange(2, 1, data.length, 7).setNumberFormat('@').setValues(data)
      .setBackgrounds(colors.map(c => Array(7).fill(c)));
  }
  const total = m.rows.length;
  sh.getRange('I1:J6').setValues([
    ['สรุป', 'จำนวน'],
    ['ใน Master ทั้งหมด', total],
    ['✔ พบแล้ว', found],
    ['🔄 ย้ายแผนก', deptDiff],
    ['✘ ยังไม่พบ', missing],
    ['⚠ สแกนแต่ไม่อยู่ใน Master', extra]
  ]);
  sh.getRange('I1:J1').setFontWeight('bold').setBackground('#e0e7ff');
  sh.autoResizeColumns(1, 10);
  ss.setActiveSheet(sh);
  ui.alert('เทียบเสร็จ\n\nใน Master: ' + total + '\nพบแล้ว: ' + found + '\nแผนกไม่ตรง: ' + deptDiff +
    '\nยังไม่พบ: ' + missing + '\nสแกนแต่ไม่อยู่ใน Master: ' + extra);
}
