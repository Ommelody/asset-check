/***********************************************************************
 *  ระบบตรวจนับครุภัณฑ์ด้วย QR Code — Google Apps Script (Backend)
 *  เขียนผลการตรวจกลับเข้า Google Sheet โดยจับคู่ด้วยรหัสครุภัณฑ์ (คอลัม B)
 *  - คอลัม AJ = สติกเกอร์ (สถานะ)
 *  - คอลัม AK = ชั้น   |   AL = แผนก   |   AM = ห้อง
 *
 *  วิธีติดตั้งอยู่ที่ไฟล์  backend/INSTALL.md
 ***********************************************************************/

var SHEET_NAME = 'ALL';   // ชื่อชีตที่เก็บข้อมูล (เปลี่ยนได้ถ้าใช้ชีตอื่น)
var COL_CODE   = 2;       // B  = รหัสครุภัณฑ์
var COL_AJ     = 36;      // AJ = สติกเกอร์ (สถานะ)
var COL_AK     = 37;      // AK = ชั้น
var COL_AL     = 38;      // AL = แผนก
var COL_AM     = 39;      // AM = ห้อง
// คอลัมรายละเอียด (สำหรับดึงค่าจริงจาก XLOOKUP มาแสดงในป๊อบอัพ)
var COL_NAME   = 3;       // C  = ชื่อทรัพย์สิน
var COL_NOTE   = 27;      // AA = หมายเหตุ
var COL_ORIG   = 28;      // AB = ที่ตั้งเดิม
var COL_PLACE  = 33;      // AG = New! Location
var COL_WORK   = 41;      // AO = ประเภทงาน

/** ตัดอักขระที่ไม่ใช่ตัวเลข/ตัวอักษรออก แล้วทำเป็นตัวพิมพ์ใหญ่ (จับคู่แบบยืดหยุ่น) */
function sk(v) { return String(v == null ? '' : v).toUpperCase().replace(/[^0-9A-Z]/g, ''); }

/** สร้าง index: ทั้งแบบตรงเป๊ะ และแบบตัดอักขระ -> เลขแถวจริง */
function buildIndex(sh, lastRow) {
  var codes = sh.getRange(2, COL_CODE, lastRow - 1, 1).getDisplayValues();
  var exact = {}, stripped = {};
  for (var i = 0; i < codes.length; i++) {
    var raw = String(codes[i][0]).trim();
    if (!raw) continue;
    var row = i + 2;
    if (exact[raw] == null) exact[raw] = row;
    var k = sk(raw);
    if (k && stripped[k] == null) stripped[k] = row;
  }
  return { exact: exact, stripped: stripped };
}

/** หาเลขแถวจากรหัสที่ส่งมา (ลองตรงเป๊ะก่อน แล้วค่อยแบบตัดอักขระ) */
function findRow(idx, want) {
  want = String(want == null ? '' : want).trim();
  if (idx.exact[want] != null) return idx.exact[want];
  var k = sk(want);
  if (k && idx.stripped[k] != null) return idx.stripped[k];
  return -1;
}

/** รับข้อมูลอัปเดตจากแอป แล้วเขียนกลับเข้า Sheet */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body    = JSON.parse(e.postData.contents);
    var updates = body.updates || [];
    var sh      = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return json({ ok: false, error: 'ไม่พบชีตชื่อ ' + SHEET_NAME });

    var lastRow = sh.getLastRow();
    var idx     = buildIndex(sh, lastRow);

    var results = [];
    for (var u = 0; u < updates.length; u++) {
      var up   = updates[u];
      var code = String(up.code).trim();
      var row  = findRow(idx, code);
      if (row < 0) { results.push({ code: code, ok: false, error: 'not_found' }); continue; }

      if (up.status != null) sh.getRange(row, COL_AJ).setValue(up.status);
      if (up.floor  != null) sh.getRange(row, COL_AK).setValue(up.floor);
      if (up.dept   != null) sh.getRange(row, COL_AL).setValue(up.dept);
      if (up.room   != null) sh.getRange(row, COL_AM).setValue(up.room);
      results.push({ code: code, ok: true, row: row });
    }
    return json({ ok: true, written: results.length, results: results });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** ทดสอบการเชื่อมต่อ / ดึงสถานะปัจจุบันเพื่อ sync กลับเข้าแอป */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ping';
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) return json({ ok: false, error: 'ไม่พบชีตชื่อ ' + SHEET_NAME });
  var lastRow = sh.getLastRow();

  if (action === 'ping') {
    return json({ ok: true, sheet: SHEET_NAME, rows: lastRow - 1 });
  }

  if (action === 'item') {
    // ดึงรายละเอียดรายตัวด้วยรหัส คืนค่าจริง (รวมผลลัพธ์ XLOOKUP) แบบสด ๆ
    var want = String((e.parameter && e.parameter.code) || '').trim();
    if (!want) return json({ ok: false, error: 'no code' });
    var rowIdx = findRow(buildIndex(sh, lastRow), want);
    if (rowIdx < 0) return json({ ok: false, error: 'not_found', wanted: want });
    // getDisplayValues = ค่าที่แสดงจริงบนหน้าจอ (ผ่านสูตรแล้ว)
    var row = sh.getRange(rowIdx, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
    var g = function (c) { return String(row[c - 1] || ''); };
    return json({ ok: true, item: {
      code:     g(COL_CODE),  name:     g(COL_NAME),  note:    g(COL_NOTE),
      origLoc:  g(COL_ORIG),  placeLoc: g(COL_PLACE), workType: g(COL_WORK),
      status:   g(COL_AJ),    floor:    g(COL_AK),    dept:     g(COL_AL),  room: g(COL_AM)
    } });
  }

  if (action === 'pull') {
    // คืนเฉพาะคอลัมที่แอปต้องใช้ sync: รหัส + สถานะ + ที่ตั้ง
    var n = lastRow - 1;
    var code   = sh.getRange(2, COL_CODE, n, 1).getValues();
    var status = sh.getRange(2, COL_AJ,   n, 1).getValues();
    var ak     = sh.getRange(2, COL_AK,   n, 1).getValues();
    var al     = sh.getRange(2, COL_AL,   n, 1).getValues();
    var am     = sh.getRange(2, COL_AM,   n, 1).getValues();
    var out = [];
    for (var i = 0; i < n; i++) {
      var c = String(code[i][0]).trim();
      if (!c) continue;
      out.push({
        code: c,
        status: String(status[i][0] || ''),
        floor:  String(ak[i][0] || ''),
        dept:   String(al[i][0] || ''),
        room:   String(am[i][0] || '')
      });
    }
    return json({ ok: true, count: out.length, rows: out });
  }

  return json({ ok: false, error: 'unknown action' });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
