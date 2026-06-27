// ==========================================
// 🛠️ ตั้งค่า ID ของ Google Sheets และ Folder
// ==========================================
const SPREADSHEET_ID = '1IaZR3AmvIjNse8nywXPLOU15KECnrmBiF6Z06_gTEco';

// เปลี่ยนโฟลเดอร์ ID หากต้องการสร้างโฟลเดอร์แยก เช่น "1A2B3C4D5E6F7G8H9I0J"
// ถ้าไม่ใส่ ระบบจะสร้างโฟลเดอร์ชื่อ "JWC_Receipts" ให้เองในหน้าแรกของ Drive
const FOLDER_ID = '';

// ตัวช่วยตอบกลับเป็น JSON
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ตัวช่วยอ่านชีตทั้งแผ่นแปลงเป็น array of objects
function sheetToObjects(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ==========================================
// 📡 ดึงข้อมูลจาก Sheets (GET Request)
// ==========================================
function doGet(e) {
  // ไม่มี action -> เสิร์ฟหน้าเว็บ index.html
  if (!e || !e.parameter || !e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('JUST WOW COFFEE - Cost Calculator')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const action = e.parameter.action;

  if (action === 'getStock')        return jsonOut(sheetToObjects(ss, 'Stock'));
  if (action === 'getLedger')       return jsonOut(sheetToObjects(ss, 'Ledger'));
  if (action === 'get')             return jsonOut(sheetToObjects(ss, 'History'));
  if (action === 'getReceipts')     return jsonOut(sheetToObjects(ss, 'Receipts'));
  if (action === 'getCertificates') return jsonOut(sheetToObjects(ss, 'Certificates'));
  if (action === 'getSettings')     return jsonOut(sheetToObjects(ss, 'Settings'));

  // 🌟 ดึงลายเซ็นจากชีต "sign" : คอลัมน์ A = รูป/ลิงก์, คอลัมน์ B = ชื่อ (แถวแรกเป็นหัวตาราง)
  if (action === 'getSignatures') {
    const sheet = ss.getSheetByName('sign');
    if (!sheet) return jsonOut([]);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return jsonOut([]);
    const result = data.slice(1)
      .map(row => ({ signature: row[0], name: row[1] }))
      .filter(item => item.signature || item.name);
    return jsonOut(result);
  }

  // action ที่ไม่รู้จัก -> ตอบ array ว่าง (ป้องกัน error ฝั่งหน้าเว็บ)
  return jsonOut([]);
}

// ==========================================
// 🖼️ แปลง Base64 เป็นไฟล์รูปและเก็บลง Drive (คืน URL)
// ==========================================
function saveImageToDrive(base64Data, filename) {
  try {
    var dataTypeMatch = base64Data.match(/^data:(image\/[a-zA-Z]+);base64,/);
    var dataType = dataTypeMatch ? dataTypeMatch[1] : 'image/jpeg';
    var base64String = base64Data.split(',')[1];
    var blob = Utilities.newBlob(Utilities.base64Decode(base64String), dataType, filename);

    var folder;
    if (FOLDER_ID) {
      folder = DriveApp.getFolderById(FOLDER_ID);
    } else {
      var folderName = "JWC_Receipts";
      var folders = DriveApp.getFoldersByName(folderName);
      if (folders.hasNext()) {
        folder = folders.next();
      } else {
        folder = DriveApp.createFolder(folderName);
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      }
    }

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    Logger.log("Image Upload Error: " + err.message);
    return null; // อัปโหลดพลาด -> คืน null (จะได้ไม่ยัด Base64 ลง Sheet)
  }
}

// ==========================================
// 🗑️ ลบไฟล์ใน Drive ที่ไม่ถูกอ้างอิงแล้ว
//    (เช่น เมื่อมีการลบ/แก้ไขรายการในแอป รูปเก่าจะถูกย้ายลงถังขยะ Drive)
// ==========================================
function _extractDriveId(url) {
  if (!url) return null;
  url = String(url);
  var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return null;
}

function _collectDriveIds(arr, fields) {
  var set = {};
  (arr || []).forEach(function (obj) {
    if (!obj) return;
    fields.forEach(function (f) {
      var v = obj[f];
      if (v == null || v === '') return;
      // บางฟิลด์ (เช่น slip) เก็บเป็น array หรือ JSON string
      if (typeof v === 'string' && v.charAt(0) === '[') { try { v = JSON.parse(v); } catch (e) {} }
      if (Array.isArray(v)) {
        v.forEach(function (u) { var id = _extractDriveId(u); if (id) set[id] = true; });
      } else {
        var id = _extractDriveId(v); if (id) set[id] = true;
      }
    });
  });
  return set;
}

// เทียบ "ของเดิมในชีต" กับ "ของใหม่ที่กำลังจะบันทึก" แล้วทิ้งไฟล์ที่หายไป
function _trashRemovedFiles(oldArr, newArr, fields) {
  try {
    var oldSet = _collectDriveIds(oldArr, fields);
    var newSet = _collectDriveIds(newArr, fields);
    Object.keys(oldSet).forEach(function (id) {
      if (!newSet[id]) {
        try { DriveApp.getFileById(id).setTrashed(true); }
        catch (e) { Logger.log('Trash fail ' + id + ': ' + e.message); }
      }
    });
  } catch (e) { Logger.log('trashRemoved error: ' + e.message); }
}

// เขียน array ของ object ลงชีต (เคลียร์แล้วเขียนใหม่ทั้งแผ่น)
// ✅ แก้ไข: รวมหัวคอลัมน์จาก "ทุก record" (ไม่ใช่แค่ตัวแรก)
//    เพื่อไม่ให้ฟิลด์ที่บางแถวไม่มี เช่น certNo / เลขที่เอกสาร ตกหล่นหายไป
function writeObjects(sheet, arr) {
  sheet.clear();
  if (arr.length === 0) return;

  // รวมคีย์จากทุก object ตามลำดับที่พบ
  const headers = [];
  const seen = {};
  arr.forEach(obj => {
    Object.keys(obj || {}).forEach(k => {
      if (!seen[k]) { seen[k] = true; headers.push(k); }
    });
  });

  sheet.appendRow(headers);
  const rows = arr.map(obj => headers.map(h => {
    let val = obj[h];
    if (typeof val === 'object' && val !== null) return JSON.stringify(val);
    return (val !== undefined && val !== null) ? val : '';
  }));
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ==========================================
// 💾 บันทึกข้อมูลลง Sheets (POST Request)
// ==========================================
function doPost(e) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const action = e.parameter.action;

  // ---- บันทึกคลังสินค้า (อัปโหลดรูปลง Drive) ----
  if (action === 'saveStock') {
    let sheet = ss.getSheetByName('Stock') || ss.insertSheet('Stock');
    let stockArr = [];
    try { stockArr = JSON.parse(e.parameter.data); }
    catch (err) { return ContentService.createTextOutput("JSON Parse Error"); }

    const _oldStock = sheetToObjects(ss, 'Stock'); // เก็บของเดิมไว้เทียบเพื่อลบไฟล์ Drive ที่หายไป

    for (let i = 0; i < stockArr.length; i++) {
      if (stockArr[i].image && String(stockArr[i].image).startsWith('data:image')) {
        let dateOnly = stockArr[i].date ? String(stockArr[i].date).split(' ')[0].replace(/\//g, '-') : 'NoDate';
        let cleanName = stockArr[i].name ? String(stockArr[i].name).replace(/[\\/:*?"<>|]/g, '_') : 'ไม่ระบุ';
        let fileName = 'Stock_' + dateOnly + '_' + cleanName + '.jpg';
        const url = saveImageToDrive(stockArr[i].image, fileName);
        if (url) stockArr[i].image = url;
      }
    }
    _trashRemovedFiles(_oldStock, stockArr, ['image']);
    writeObjects(sheet, stockArr);
    return ContentService.createTextOutput("Stock Sync Success");
  }

  // ---- บันทึกบัญชีการเงิน (อัปโหลดใบเสร็จ/สลิปลง Drive) ----
  if (action === 'saveLedger') {
    let sheet = ss.getSheetByName('Ledger') || ss.insertSheet('Ledger');
    let ledgerArr = [];
    try { ledgerArr = JSON.parse(e.parameter.data); }
    catch (err) { return ContentService.createTextOutput("JSON Parse Error"); }

    const _oldLedger = sheetToObjects(ss, 'Ledger');

    for (let i = 0; i < ledgerArr.length; i++) {
      if (ledgerArr[i].receipt && String(ledgerArr[i].receipt).startsWith('data:image')) {
        let dateOnly = ledgerArr[i].date ? String(ledgerArr[i].date).split(' ')[0].replace(/\//g, '-') : 'NoDate';
        let cleanDesc = ledgerArr[i].desc ? String(ledgerArr[i].desc).replace(/[\\/:*?"<>|]/g, '_') : 'ไม่มีรายละเอียด';
        let fileName = dateOnly + '_' + cleanDesc + '.jpg';
        const url = saveImageToDrive(ledgerArr[i].receipt, fileName);
        ledgerArr[i].receipt = url || ""; // อัปโหลดพลาด -> เคลียร์ทิ้ง
      }
    }
    _trashRemovedFiles(_oldLedger, ledgerArr, ['receipt']);
    writeObjects(sheet, ledgerArr);
    return ContentService.createTextOutput("Ledger Sync Success");
  }

  // ---- บันทึกประวัติคำนวณสูตร (เพิ่มทีละแถว) ----
  if (action === 'save') {
    let sheet = ss.getSheetByName('History') || ss.insertSheet('History');
    const d = JSON.parse(e.parameter.data);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['id', 'name', 'cost', 'price', 'margin', 'date', 'ingredients']);
    } else {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.indexOf('ingredients') === -1) sheet.getRange(1, headers.length + 1).setValue('ingredients');
    }
    sheet.appendRow([d.id, d.name, d.cost, d.price, d.margin, d.date, d.ingredients || '']);
    return ContentService.createTextOutput("Success");
  }

  // ---- อัปเดตประวัติคำนวณสูตรเดิม ----
  if (action === 'editHistory') {
    let sheet = ss.getSheetByName('History');
    if (!sheet) return ContentService.createTextOutput("Not Found");
    const d = JSON.parse(e.parameter.data);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == d.id) {
        sheet.getRange(i + 1, 2).setValue(d.name);
        sheet.getRange(i + 1, 3).setValue(d.cost);
        sheet.getRange(i + 1, 4).setValue(d.price);
        sheet.getRange(i + 1, 5).setValue(d.margin);
        sheet.getRange(i + 1, 6).setValue(d.date);
        sheet.getRange(i + 1, 7).setValue(d.ingredients || '');
        break;
      }
    }
    return ContentService.createTextOutput("History Edited Success");
  }

  // ---- ลบประวัติคำนวณสูตร ----
  if (action === 'delete') {
    let sheet = ss.getSheetByName('History');
    if (!sheet) return ContentService.createTextOutput("Not Found");
    const targetId = e.parameter.id;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == targetId) { sheet.deleteRow(i + 1); break; }
    }
    return ContentService.createTextOutput("Deleted");
  }

  // ---- บันทึกใบเสร็จรับเงิน ----
  if (action === 'saveReceipts') {
    let sheet = ss.getSheetByName('Receipts') || ss.insertSheet('Receipts');
    let arr = [];
    try { arr = JSON.parse(e.parameter.data); }
    catch (err) { return ContentService.createTextOutput("JSON Parse Error"); }

    const _oldReceipts = sheetToObjects(ss, 'Receipts');

    // อัปโหลดลายเซ็นที่เป็น Base64 ขึ้น Drive แล้วเก็บเป็น URL (กันเกินขีดจำกัด 50,000 ตัวอักษร/ช่อง)
    for (let i = 0; i < arr.length; i++) {
      if (arr[i].signature && String(arr[i].signature).startsWith('data:image')) {
        const url = saveImageToDrive(arr[i].signature, 'Signature_' + (arr[i].receiptNo || Date.now()) + '.png');
        if (url) arr[i].signature = url;
      }
    }
    _trashRemovedFiles(_oldReceipts, arr, ['signature']);
    writeObjects(sheet, arr);
    return ContentService.createTextOutput("Receipts Sync Success");
  }

  // ---- บันทึกใบรับรองแทนใบเสร็จรับเงิน ----
  if (action === 'saveCertificates') {
    let sheet = ss.getSheetByName('Certificates') || ss.insertSheet('Certificates');
    let arr = [];
    try { arr = JSON.parse(e.parameter.data); }
    catch (err) { return ContentService.createTextOutput("JSON Parse Error"); }

    const _oldCerts = sheetToObjects(ss, 'Certificates');

    for (let i = 0; i < arr.length; i++) {
      // ✅ อัปโหลดรูปเอกสาร (image) ขึ้น Drive — base64 ยาวเกิน 50,000 ตัวอักษร/ช่อง จะทำให้บันทึกพัง
      if (arr[i].image && String(arr[i].image).startsWith('data:image')) {
        const url = saveImageToDrive(arr[i].image, 'CertDoc_' + (arr[i].certNo || arr[i].id || Date.now()) + '.jpg');
        if (url) arr[i].image = url;
      }
      // อัปโหลดสลิป (array ของ Base64) ขึ้น Drive แล้วเก็บเป็น array ของ URL
      if (Array.isArray(arr[i].slip)) {
        arr[i].slip = arr[i].slip.map((img, idx) => {
          if (img && String(img).startsWith('data:image')) {
            const url = saveImageToDrive(img, 'CertSlip_' + (arr[i].certNo || arr[i].id || Date.now()) + '_' + idx + '.jpg');
            return url || img;
          }
          return img;
        });
      }
    }
    _trashRemovedFiles(_oldCerts, arr, ['image', 'slip']);
    writeObjects(sheet, arr);
    return ContentService.createTextOutput("Certs Sync Success");
  }

  // ---- บันทึกการตั้งค่า ----
  if (action === 'saveSettings') {
    let sheet = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
    let arr = [];
    try { arr = JSON.parse(e.parameter.data); }
    catch (err) { return ContentService.createTextOutput("JSON Parse Error"); }
    writeObjects(sheet, arr);
    return ContentService.createTextOutput("Settings Sync Success");
  }

  // action ที่ไม่รู้จัก
  return ContentService.createTextOutput("Unknown action: " + action);
}
