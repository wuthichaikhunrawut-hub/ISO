/**
 * ระบบจัดการไฟล์มาตรฐาน ISO 27001
 * DB.js - คลาสตัวช่วยในการจัดระเบียบและจัดการข้อมูลใน Google Sheets แบบฐานข้อมูล (CRUD Operations)
 */

var DB = (function() {
  
  /**
   * แปลงข้อมูลใน Sheet เป็น Array of Objects
   */
  function getTableData(sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return []; // ไม่มีข้อมูล (แถวแรกเป็น Header)
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    var result = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var obj = {};
      // เก็บค่า Row Number ไว้เพื่ออ้างอิงตอนอัปเดต/ลบ
      obj._rowNum = i + 2; 
      
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        if (key) {
          var val = row[j];
          // จัดรูปแบบ Date ให้ไม่พังเมื่อส่งไปหา Client
          if (val instanceof Date) {
            obj[key] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            obj[key] = val;
          }
        }
      }
      result.push(obj);
    }
    return result;
  }
  
  /**
   * ค้นหาแถวข้อมูลตามฟิลด์ที่กำหนด
   */
  function getRowByField(sheetName, fieldName, fieldValue) {
    var data = getTableData(sheetName);
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][fieldName]).trim() === String(fieldValue).trim()) {
        return data[i];
      }
    }
    return null;
  }
  
  /**
   * เพิ่มแถวข้อมูลใหม่ (Insert)
   * รองรับการรัน ID อัตโนมัติ (Auto-increment) ถ้าคอลัมน์แรกลงท้ายด้วย _id
   */
  function insertRow(sheetName, dataObject) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("ไม่พบชีต: " + sheetName);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // จัดตำแหน่งข้อมูลให้ตรงกับหัวตาราง
    var newRowValues = [];
    
    // ตรวจสอบระบบรัน ID อัตโนมัติในคอลัมน์แรก (เช่น evidence_id, trail_id)
    var pkName = headers[0];
    if (pkName && pkName.endsWith("_id") && !dataObject[pkName]) {
      var nextId = 1;
      if (lastRow >= 2) {
        var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r) { return Number(r[0]); });
        var maxId = Math.max.apply(null, ids.filter(function(v) { return !isNaN(v); }));
        if (maxId > 0) nextId = maxId + 1;
      }
      dataObject[pkName] = nextId;
    }
    
    for (var i = 0; i < headers.length; i++) {
      var val = dataObject[headers[i]];
      newRowValues.push(val === undefined ? "" : val);
    }
    
    sheet.appendRow(newRowValues);
    Logger.log("แทรกแถวใหม่ใน " + sheetName + " สำเร็จ");
    return dataObject;
  }
  
  /**
   * แก้ไขแถวข้อมูล (Update) ค้นหาจาก Key Field
   */
  function updateRow(sheetName, keyFieldName, keyFieldValue, updatedFieldsObject) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("ไม่พบชีต: " + sheetName);
    
    var rowData = getRowByField(sheetName, keyFieldName, keyFieldValue);
    if (!rowData) throw new Error("ไม่พบข้อมูลที่ต้องการอัปเดตในชีต " + sheetName + " ด้วยค่า " + keyFieldValue);
    
    var rowNum = rowData._rowNum;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // อัปเดตเฉพาะคอลัมน์ที่ส่งมา
    for (var i = 0; i < headers.length; i++) {
      var key = headers[i];
      if (key in updatedFieldsObject && key !== keyFieldName) {
        var colIndex = i + 1;
        var val = updatedFieldsObject[key];
        sheet.getRange(rowNum, colIndex).setValue(val === undefined ? "" : val);
      }
    }
    
    Logger.log("อัปเดตแถวที่ " + rowNum + " ใน " + sheetName + " สำเร็จ");
    return true;
  }
  
  /**
   * ลบแถวข้อมูล (Delete) ค้นหาจาก Key Field
   */
  function deleteRow(sheetName, keyFieldName, keyFieldValue) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error("ไม่พบชีต: " + sheetName);
    
    var rowData = getRowByField(sheetName, keyFieldName, keyFieldValue);
    if (!rowData) throw new Error("ไม่พบข้อมูลที่ต้องการลบในชีต " + sheetName + " ด้วยค่า " + keyFieldValue);
    
    var rowNum = rowData._rowNum;
    sheet.deleteRow(rowNum);
    Logger.log("ลบแถวที่ " + rowNum + " ใน " + sheetName + " สำเร็จ");
    return true;
  }
  
  return {
    getTableData: getTableData,
    getRowByField: getRowByField,
    insertRow: insertRow,
    updateRow: updateRow,
    deleteRow: deleteRow
  };
  
})();
