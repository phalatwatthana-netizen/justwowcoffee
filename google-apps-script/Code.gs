/**
 * JUSTWOW COFFEE — Google Sheets backend (Apps Script Web App)
 * ------------------------------------------------------------------
 * เก็บข้อมูลแต่ละชุด (stock / ledger / certificates / receipts / history)
 * เป็น "JSON" ในชีตชื่อ _DATA  →  ทุกฟิลด์ถูกบันทึกครบถาวร
 * รวมถึง certNo (เลขที่ใบรับรอง JW-xxxx), เลขที่สลิป PV-xxxx, slip, items ฯลฯ
 * และฟิลด์ใหม่ในอนาคตก็ไม่หาย (ไม่ต้องแก้สคริปต์อีก)
 *
 * ลายเซ็นยังอ่านจากชีตชื่อ "SIGN" (คอลัมน์หัวตาราง: name | signature) เหมือนเดิม
 *
 * วิธีติดตั้ง:
 *  1) เปิด Google Sheet ของร้าน → เมนู Extensions → Apps Script
 *  2) ⚠️ สำรองข้อมูลก่อน: File → Make a copy ของ Sheet ทั้งไฟล์
 *  3) ลบโค้ดเดิมทั้งหมดใน Code.gs แล้ววางโค้ดนี้แทน → Save
 *  4) Deploy → Manage deployments → แก้ deployment เดิม (เวอร์ชันใหม่)
 *     หรือ Deploy → New deployment → Web app
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     แล้วใช้ URL เดิม (ถ้าแก้ deployment เดิม URL จะไม่เปลี่ยน)
 *
 * หมายเหตุการย้ายข้อมูล:
 *  - สคริปต์นี้ใช้รูปแบบจัดเก็บใหม่ (JSON ในชีต _DATA) ข้อมูลเดิมที่อยู่คนละรูปแบบ
 *    จะยังไม่ถูกอ่านโดยอัตโนมัติ ระบบจะเริ่มสะสมข้อมูลใหม่ลง _DATA
 *  - ตัวแอปมี localStorage สำรองอยู่ เมื่อบันทึกครั้งถัดไปข้อมูลในเครื่องจะถูกเขียนลง _DATA
 */

var DATA_SHEET = '_DATA';   // ชีตเก็บ JSON (จะถูกสร้างอัตโนมัติ ซ่อนได้)
var SIGN_SHEET = 'SIGN';    // ชีตลายเซ็น (หัวตาราง: name | signature)

/* ---------- helpers ---------- */
function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function _dataSheet() {
  var ss = _ss();
  var sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.getRange('A1').setValue('key');
    sh.getRange('B1').setValue('json');
  }
  return sh;
}

function _readCollection(key) {
  var sh = _dataSheet();
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) {
      try { return JSON.parse(values[i][1] || '[]'); } catch (e) { return []; }
    }
  }
  return [];
}

function _writeCollection(key, arr) {
  var sh = _dataSheet();
  var json = JSON.stringify(arr || []);
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === key) { sh.getRange(i + 1, 2).setValue(json); return; }
  }
  sh.appendRow([key, json]);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _readSignatures() {
  var sh = _ss().getSheetByName(SIGN_SHEET);
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var header = values[0].map(function (h) { return String(h).toLowerCase().trim(); });
  var nameCol = header.indexOf('name');
  var sigCol = header.indexOf('signature');
  if (nameCol < 0) nameCol = 0;
  if (sigCol < 0) sigCol = 1;
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var name = values[i][nameCol];
    var sig = values[i][sigCol];
    if (name || sig) out.push({ name: String(name || ''), signature: String(sig || '') });
  }
  return out;
}

/* ---------- GET ---------- */
function doGet(e) {
  var action = ((e && e.parameter && e.parameter.action) || '').trim();
  switch (action) {
    case 'getStock':        return _json(_readCollection('stock'));
    case 'getLedger':       return _json(_readCollection('ledger'));
    case 'getCertificates': return _json(_readCollection('certificates')); // มี certNo ครบ
    case 'getReceipts':     return _json(_readCollection('receipts'));
    case 'get':             return _json(_readCollection('history'));
    case 'getSignatures':   return _json(_readSignatures());
    default:                return _json([]);
  }
}

/* ---------- POST ---------- */
function doPost(e) {
  var p = (e && e.parameter) || {};
  var action = (p.action || '').trim();
  var data = null;
  try { data = p.data ? JSON.parse(p.data) : null; } catch (err) { data = null; }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) {}
  try {
    switch (action) {
      case 'saveStock':        _writeCollection('stock', data); break;
      case 'saveLedger':       _writeCollection('ledger', data); break;
      case 'saveCertificates': _writeCollection('certificates', data); break; // เก็บ certNo + ทุกฟิลด์
      case 'saveReceipts':     _writeCollection('receipts', data); break;

      case 'save': { // ต่อท้ายประวัติ 1 รายการ
        var hs = _readCollection('history');
        if (data) hs.push(data);
        _writeCollection('history', hs);
        break;
      }
      case 'editHistory': { // แก้ไขประวัติตาม id
        var he = _readCollection('history');
        if (data && data.id != null) {
          var idx = -1;
          for (var i = 0; i < he.length; i++) { if (String(he[i].id) === String(data.id)) { idx = i; break; } }
          if (idx >= 0) he[idx] = data; else he.push(data);
        }
        _writeCollection('history', he);
        break;
      }
      case 'delete': { // ลบสต๊อกตาม id
        var st = _readCollection('stock').filter(function (x) { return String(x.id) !== String(p.id); });
        _writeCollection('stock', st);
        break;
      }
    }
  } finally {
    try { lock.releaseLock(); } catch (err) {}
  }
  return _json({ ok: true });
}
