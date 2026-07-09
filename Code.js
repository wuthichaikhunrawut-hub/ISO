/**
 * ระบบจัดการไฟล์มาตรฐาน ISO 27001
 * Code.js - จุดเริ่มต้นหลักและการกำหนดเราท์และฟังก์ชันสำหรับ Web App API
 */

/**
 * ฟังก์ชันหลักที่ทริกเกอร์เมื่อมีผู้เรียกลิงก์ Web App
 */
function doGet(e) {
  // ตรวจสอบและอัปเกรดโครงสร้างคอลัมน์ให้อัตโนมัติ (Schema Migration)
  try {
    ensureUserCodeColumnExists();
    ensureMasterDocumentsSchemaUpgraded();
    ensureMasterDocumentsHeadersExist();
    fixValidationRules();
  } catch(err) {
    Logger.log("Migration Error: " + err.message);
  }

  var template = HtmlService.createTemplateFromFile("Index");
  
  // โหลดสิทธิ์เพื่อนำไปแสดงผลฝั่งไคลเอนต์
  var user = getCurrentUserRole();
  template.userEmail = user.email;
  template.userRole = user.role;
  template.userName = user.name;
  template.userAllowed = user.allowed;
  
  return template.evaluate()
    .setTitle("ระบบจัดการไฟล์มาตรฐาน ISO 27001")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * ฟังก์ชันรวมไฟล์ย่อย (CSS/JS) ลงใน HTML หลัก
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ------------------------------------------------------------
 * RPC / API ENDPOINTS (เรียกใช้งานผ่าน google.script.run)
 * ------------------------------------------------------------
 */

function apiGetCurrentUserRole() {
  return getCurrentUserRole();
}

function apiGetComplianceStats() {
  return getComplianceStats();
}

function apiGetISOControls() {
  return DB.getTableData("ISO_Controls");
}

function apiGetMasterDocuments() {
  return getMasterDocuments();
}

function apiGetEvidenceLogs() {
  // ดึงประวัติหลักฐานทั้งหมด
  return DB.getTableData("Evidence_Logs");
}

function apiGetAuditTrails() {
  // ดึงประวัติกิจกรรม (เฉพาะ Admin เท่านั้น)
  var user = getCurrentUserRole();
  if (user.role !== "Admin") {
    return [];
  }
  return DB.getTableData("Audit_Trails");
}

function apiGetUserAccessMatrix() {
  // ดึงรายการผู้ใช้เพื่อใช้แมปแผนกและสิทธิ์การแสดงผลในหน้าเว็บแอป
  return DB.getTableData("User_Access_Matrix");
}

function apiSaveMasterDocument(docData, selectedControls, ipAddress) {
  return saveMasterDocument(docData, selectedControls, ipAddress);
}

function apiDeleteMasterDocument(docId, ipAddress) {
  return deleteMasterDocument(docId, ipAddress);
}

function apiUploadMasterDocumentFile(fileName, mimeType, base64Data) {
  return uploadMasterDocumentFile(fileName, mimeType, base64Data);
}

function apiUploadEvidence(docId, controlCode, period, details, fileData, ipAddress) {
  return uploadEvidence(docId, controlCode, period, details, fileData, ipAddress);
}

function apiUpdateEvidenceLog(evidenceId, period, details, docId, fileData, controlCode, ipAddress, deleteOldFile) {
  return updateEvidenceLog(evidenceId, period, details, docId, fileData, controlCode, ipAddress, deleteOldFile);
}

function apiDeleteEvidenceLog(evidenceId, ipAddress) {
  return deleteEvidenceLog(evidenceId, ipAddress);
}

function apiUpdateControlStatus(controlCode, status, gapAnalysis, ipAddress) {
  return updateControlStatus(controlCode, status, gapAnalysis, ipAddress);
}

/**
 * ฟังก์ชันสำหรับเพิ่มข้อมูลผู้ใช้ใหม่ (เฉพาะแอดมิน)
 */
function apiSaveUser(userData, ipAddress) {
  var user = getCurrentUserRole();
  if (user.role !== "Admin") {
    throw new Error("คุณไม่มีสิทธิ์จัดระเบียบรายชื่อผู้ใช้");
  }
  
  var existing = DB.getRowByField("User_Access_Matrix", "user_email", userData.user_email);
  userData.last_review_date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  if (existing) {
    DB.updateRow("User_Access_Matrix", "user_email", userData.user_email, userData);
    logAuditTrail("Update", "User_Access_Matrix", userData.user_email, "Updated user: " + userData.name, ipAddress);
  } else {
    DB.insertRow("User_Access_Matrix", userData);
    logAuditTrail("Create", "User_Access_Matrix", userData.user_email, "Added new user: " + userData.name, ipAddress);
  }
  return { success: true };
}

/**
 * ลบผู้ใช้ออกจากตาราง (เฉพาะแอดมิน)
 */
function apiDeleteUser(email, ipAddress) {
  var user = getCurrentUserRole();
  if (user.role !== "Admin") {
    throw new Error("คุณไม่มีสิทธิ์ลบรายชื่อผู้ใช้");
  }
  
  var existing = DB.getRowByField("User_Access_Matrix", "user_email", email);
  if (!existing) throw new Error("ไม่พบรายชื่อผู้ใช้");
  
  DB.deleteRow("User_Access_Matrix", "user_email", email);
  logAuditTrail("Delete", "User_Access_Matrix", email, "Deleted user: " + existing.name, ipAddress);
  return { success: true };
}

/**
 * ฟังก์ชันสำหรับติดตั้งโครงสร้างตารางข้อมูลเริ่มต้น (Setup Database Sheets)
 */
function initDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // นิยามตาราง หัวตาราง ตัวอย่างข้อมูล และรูปแบบ Validation
  var tables = {
    "ISO_Controls": {
      headers: ["control_code", "domain", "requirement", "assessment_status", "gap_analysis"],
      sample: ["A.5.15", "Organizational", "Access control rules and rights shall be established and implemented.", "In Progress", "Need to define remote access process"],
      validations: {
        "D2:D1000": { type: "list", values: ["Ready", "In Progress", "Missing", "N/A"] }
      }
    },
    "Master_Documents": {
      headers: ["doc_type", "doc_id", "title", "owner", "edition", "revision", "effective_date", "policy", "drive_link", "review_frequency_months", "next_review_date", "evidence_frequency", "status"],
      sample: ["Procedure", "ISMS-PROC-02", "Access Control Procedure", "System IT Dept", "1", "0", "2026-07-09", "A.5.15", "https://drive.google.com/open?id=1xyz...", 12, "2027-07-09", "Monthly", "Approved"],
      validations: {
        "A2:A1000": { type: "list", values: ["Policy", "Procedure", "Form", "Report"] },
        "G2:G1000": { type: "date" },
        "K2:K1000": { type: "date" },
        "L2:L1000": { type: "list", values: ["Monthly", "Quarterly", "Annually", "Ad-hoc"] },
        "M2:M1000": { type: "list", values: ["Draft", "Approved", "Obsolete"] }
      }
    },
    "Document_Control_Mapping": {
      headers: ["doc_id", "control_code"],
      sample: ["ISMS-PROC-02", "A.5.15"],
      validations: {
        "A2:A1000": { type: "range", sourceSheet: "Master_Documents", sourceRange: "B2:B1000" },
        "B2:B1000": { type: "range", sourceSheet: "ISO_Controls", sourceRange: "A2:A1000" }
      }
    },
    "Evidence_Logs": {
      headers: ["evidence_id", "recorded_date", "doc_id", "evidence_period", "details", "drive_link", "recorded_by"],
      sample: [1, "2026-07-09 10:00:00", "ISMS-PROC-02", "2026-06", "June Access Log Review Evidence", "https://drive.google.com/open?id=2abc...", "staff@company.com"],
      validations: {
        "C2:C1000": { type: "range", sourceSheet: "Master_Documents", sourceRange: "B2:B1000" }
      }
    },
    "User_Access_Matrix": {
      headers: ["user_email", "user_code", "name", "department", "role", "account_status", "last_review_date"],
      sample: [Session.getActiveUser().getEmail() || "admin@company.com", "EMP-000", "System Administrator", "ISMS Office", "Admin", "Active", "2026-07-09"],
      validations: {
        "E2:E1000": { type: "list", values: ["Admin", "Owner", "Auditor"] },
        "F2:F1000": { type: "list", values: ["Active", "Inactive"] },
        "G2:G1000": { type: "date" }
      }
    },
    "Audit_Trails": {
      headers: ["trail_id", "timestamp", "action", "ref_table", "ref_id", "performed_by", "change_details", "ip_address"],
      sample: [1, "2026-07-09 10:05:00", "Create", "Evidence_Logs", "1", "staff@company.com", "Uploaded June access review log", "192.168.1.100"],
      validations: {
        "C2:C1000": { type: "list", values: ["Create", "Update", "Delete", "Import"] }
      }
    }
  };

  Logger.log("เริ่มการสร้างโครงสร้างตารางระบบ ISO 27001...");

  for (var sheetName in tables) {
    var sheet = ss.getSheetByName(sheetName);
    
    // ถ้ายังไม่มีชีตนี้ ให้สร้างใหม่
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      Logger.log("สร้างชีตใหม่: " + sheetName);
    } else {
      // ถ้ามีอยู่แล้ว เคลียร์ข้อมูลเดิมเพื่อเริ่มเซ็ตใหม่
      sheet.clear();
      sheet.clearValidation();
      Logger.log("พบชีตเดิม เคลียร์รูปแบบและข้อมูลเก่า: " + sheetName);
    }

    var tableInfo = tables[sheetName];

    // 1. เขียนหัวคอลัมน์ (Headers)
    var headerRange = sheet.getRange(1, 1, 1, tableInfo.headers.length);
    headerRange.setValues([tableInfo.headers]);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#E0E0E0"); // สีเทาอ่อนสำหรับหัวตาราง
    headerRange.setHorizontalAlignment("left");

    // 2. ใส่ข้อมูลตัวอย่าง 1 แถว (Sample Data Row)
    var sampleRange = sheet.getRange(2, 1, 1, tableInfo.sample.length);
    sampleRange.setValues([tableInfo.sample]);

    // 3. ตกแต่งตารางเบื้องต้น
    sheet.setFrozenRows(1); // ล็อกแถวแรก
    
    // ปรับความกว้างคอลัมน์อัตโนมัติ
    for (var col = 1; col <= tableInfo.headers.length; col++) {
      sheet.autoResizeColumn(col);
    }
  }

  // 4. สั่งทำ Data Validation (Dropdown & Dates) หลังจากทุกชีตถูกสร้างเสร็จแล้ว
  Logger.log("กำลังสร้างกฎการเลือกข้อมูล (Data Validation Rules)...");
  for (var sheetName in tables) {
    var sheet = ss.getSheetByName(sheetName);
    var validations = tables[sheetName].validations;
    
    if (validations) {
      for (var rangeStr in validations) {
        var ruleInfo = validations[rangeStr];
        var cellRange = sheet.getRange(rangeStr);
        var ruleBuilder = SpreadsheetApp.newDataValidation().setAllowInvalid(false);

        if (ruleInfo.type === "list") {
          // Dropdown จากลิสต์ข้อความที่กำหนด
          ruleBuilder.requireValueInList(ruleInfo.values, true);
          cellRange.setDataValidation(ruleBuilder.build());
          Logger.log("  - ตั้งค่า Dropdown สำหรับ " + sheetName + " (" + rangeStr + ")");
        } 
        else if (ruleInfo.type === "date") {
          // ตรวจสอบว่าเป็นวันที่
          ruleBuilder.requireDateOnOrAfter(new Date(1970, 0, 1));
          cellRange.setDataValidation(ruleBuilder.build());
          Logger.log("  - ตั้งค่า Date Validation สำหรับ " + sheetName + " (" + rangeStr + ")");
        } 
        else if (ruleInfo.type === "range") {
          // Dropdown ดึงค่าแบบ Dynamic จากคอลัมน์ในชีตอื่น (เช่น doc_id ดึงมาจาก Master_Documents!A2:A)
          var sourceSheet = ss.getSheetByName(ruleInfo.sourceSheet);
          var sourceRange = sourceSheet.getRange(ruleInfo.sourceRange);
          ruleBuilder.requireValueInRange(sourceRange, true);
          cellRange.setDataValidation(ruleBuilder.build());
          Logger.log("  - ตั้งค่า Dynamic Dropdown สำหรับ " + sheetName + " (" + rangeStr + ") -> อ้างอิง " + ruleInfo.sourceSheet);
        }
      }
    }
  }

  Logger.log("เสร็จสิ้นขั้นตอนการติดตั้งตาราง! คุณสามารถนำโค้ด Web App มาต่อยอดได้เลยครับ");
}

/**
 * ฟังก์ชันช่วยตรวจสอบและอัปเกรดโครงสร้างตารางโดยการเพิ่มคอลัมน์ user_code หากไม่มี (Schema Migration)
 */
function ensureUserCodeColumnExists() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("User_Access_Matrix");
  if (!sheet) return;
  
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var codeIdx = headers.indexOf("user_code");
  
  if (codeIdx === -1) {
    // แทรกคอลัมน์ใหม่ที่ตำแหน่งคอลัมน์ที่ 2 (ระหว่าง A: user_email และ B: name เดิม)
    sheet.insertColumnBefore(2);
    sheet.getRange(1, 2).setValue("user_code");
    
    // ตั้งค่าข้อมูลพนักงานเดิมด้วยรหัสแบบรันเรียง
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      for (var r = 2; r <= lastRow; r++) {
        var num = r - 2;
        var nextCode = "EMP-" + String(num).padStart(3, '0');
        sheet.getRange(r, 2).setValue(nextCode);
      }
    }
    SpreadsheetApp.flush();
    Logger.log("อัปเกรดโครงสร้างตารางเรียบร้อย: เพิ่มคอลัมน์ user_code");
  }
}

/**
 * ฟังก์ชันตรวจสอบและอัปเกรดตาราง Master_Documents และแปลงข้อมูลพนักงานตามฟีดที่ผู้ใช้ต้องการ
 */
function ensureMasterDocumentsSchemaUpgraded() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Documents");
  if (!sheet) return;
  
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // หาก doc_type ไม่ได้เป็นคอลัมน์แรก แสดงว่าต้องทำ Migration โครงสร้างใหม่
  if (headers[0] !== "doc_type") {
    // 1. อ่านข้อมูลทั้งหมดเดิมเก็บไว้ในหน่วยความจำ
    var lastRow = sheet.getLastRow();
    var oldData = [];
    if (lastRow >= 2) {
      var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      for (var r = 0; r < values.length; r++) {
        var rowObj = {};
        for (var c = 0; c < headers.length; c++) {
          rowObj[headers[c]] = values[r][c];
        }
        oldData.push(rowObj);
      }
    }
    
    // ดึงข้อมูล Mapping เดิมมาแมปเข้าด้วยกัน (ถ้ามีชีต Document_Control_Mapping)
    var mapSheet = ss.getSheetByName("Document_Control_Mapping");
    var docToControls = {};
    if (mapSheet) {
      var mapLastRow = mapSheet.getLastRow();
      if (mapLastRow >= 2) {
        var mapValues = mapSheet.getRange(2, 1, mapLastRow - 1, 2).getValues();
        mapValues.forEach(function(row) {
          var dId = row[0];
          var cCode = row[1];
          if (!docToControls[dId]) {
            docToControls[dId] = [];
          }
          docToControls[dId].push(cCode);
        });
      }
    }
    
    // 2. เคลียร์ข้อมูลและการตั้งค่าเดิมในชีต
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.setDataValidations([]);
    
    // 3. เขียนหัวตารางชุดใหม่
    var newHeaders = ["doc_type", "doc_id", "title", "owner", "edition", "revision", "effective_date", "policy", "drive_link", "review_frequency_months", "next_review_date", "evidence_frequency", "status"];
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    
    // 4. เขียนข้อมูลพนักงานใหม่ที่แปลงตามโครงสร้างใหม่แล้ว
    if (oldData.length > 0) {
      var newRows = oldData.map(function(oldRow) {
        var docId = oldRow.doc_id || "";
        var policyVal = oldRow.policy || (docToControls[docId] ? docToControls[docId].join(", ") : "");
        var ownerVal = oldRow.owner || oldRow.owner_email || "";
        var docTypeVal = oldRow.doc_type || "Procedure";
        
        // แปลงเวอร์ชันเดิมเป็น ฉบับที่/แก้ไขครั้งที่
        var editionVal = oldRow.edition || "1";
        var revisionVal = oldRow.revision || "0";
        if (oldRow.version) {
          var versionStr = String(oldRow.version).replace(/v/gi, "");
          var versionParts = versionStr.split(".");
          editionVal = versionParts[0] || "1";
          revisionVal = versionParts[1] || "0";
        }
        
        return [
          docTypeVal,
          docId,
          oldRow.title || "",
          ownerVal,
          editionVal,
          revisionVal,
          oldRow.effective_date ? safeFormatDate(oldRow.effective_date) : "",
          policyVal,
          oldRow.drive_link || "",
          oldRow.review_frequency_months || 12,
          oldRow.next_review_date ? safeFormatDate(oldRow.next_review_date) : "",
          oldRow.evidence_frequency || "Monthly",
          oldRow.status || "Approved"
        ];
      });
      
      sheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
    }
    
    // 5. เซ็ตกฎความถูกต้องข้อมูล (Data Validation) ของคอลัมน์ใหม่
    var validations = {
      "A2:A1000": { type: "list", values: ["Policy", "Procedure", "Form", "Report"] },
      "G2:G1000": { type: "date" },
      "K2:K1000": { type: "date" },
      "L2:L1000": { type: "list", values: ["Monthly", "Quarterly", "Annually", "Ad-hoc"] },
      "M2:M1000": { type: "list", values: ["Draft", "Approved", "Obsolete"] }
    };
    
    for (var rangeStr in validations) {
      try {
        var ruleInfo = validations[rangeStr];
        var cellRange = sheet.getRange(rangeStr);
        var ruleBuilder = SpreadsheetApp.newDataValidation().setAllowInvalid(false);

        if (ruleInfo.type === "list") {
          ruleBuilder.requireValueInList(ruleInfo.values, true);
          cellRange.setDataValidation(ruleBuilder.build());
        } else if (ruleInfo.type === "date") {
          ruleBuilder.requireDateOnOrAfter(new Date(1970, 0, 1));
          cellRange.setDataValidation(ruleBuilder.build());
        }
      } catch (e) {
        Logger.log("Validation error for " + rangeStr + ": " + e.message);
      }
    }
    
    SpreadsheetApp.flush();
    Logger.log("อัปเกรดโครงสร้างตาราง Master_Documents และย้ายข้อมูลเสร็จสมบูรณ์!");
  }
}

/**
 * ตรวจสอบกรณีตาราง Master_Documents ว่างเปล่า (ไม่มีคอลัมน์เลย) ให้สร้างคอลัมน์ตั้งต้นให้
 */
function ensureMasterDocumentsHeadersExist() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master_Documents");
  if (!sheet) return;
  
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    var newHeaders = ["doc_type", "doc_id", "title", "owner", "edition", "revision", "effective_date", "policy", "drive_link", "review_frequency_months", "next_review_date", "evidence_frequency", "status"];
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    sheet.setFrozenRows(1);
    for (var col = 1; col <= newHeaders.length; col++) {
      sheet.autoResizeColumn(col);
    }
    SpreadsheetApp.flush();
    Logger.log("ตาราง Master_Documents โดนเคลียร์ว่างเปล่า ได้รับการกำหนดหัวตารางตั้งต้นให้แล้ว!");
  }
}

/**
 * ฟังก์ชันแปลง Date เป็น string yyyy-MM-dd อย่างปลอดภัย ไม่พังเมื่อมีค่าขยะ
 */
function safeFormatDate(dateVal) {
  if (!dateVal) return "";
  try {
    var dateObj = new Date(dateVal);
    if (isNaN(dateObj.getTime())) {
      return String(dateVal);
    }
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } catch (e) {
    return String(dateVal);
  }
}

/**
 * แก้ไขคอลัมน์คัดกรอง (Data Validation) ของตาราง Evidence_Logs และ Document_Control_Mapping 
 * ให้ชี้ไปที่ doc_id (คอลัมน์ B) แทน doc_type (คอลัมน์ A) ป้องกันการบันทึกไม่ได้
 */
function fixValidationRules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master_Documents");
  if (!masterSheet) return;

  // 1. แก้ไข Validation ของ Evidence_Logs!C2:C1000
  var evidenceSheet = ss.getSheetByName("Evidence_Logs");
  if (evidenceSheet) {
    try {
      var cellRange = evidenceSheet.getRange("C2:C1000");
      cellRange.clearDataValidations();
      
      var sourceRange = masterSheet.getRange("B2:B1000");
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceRange, true)
        .setAllowInvalid(true) // ป้องกันการเด้งปฏิเสธกรณีผิดพลาด
        .build();
      cellRange.setDataValidation(rule);
      Logger.log("แก้ไข Data Validation ของ Evidence_Logs!C2:C1000 สำเร็จ");
    } catch(e) {
      Logger.log("ไม่สามารถแก้ Validation ของ Evidence_Logs: " + e.message);
    }
  }

  // 2. แก้ไข Validation ของ Document_Control_Mapping!A2:A1000
  var mappingSheet = ss.getSheetByName("Document_Control_Mapping");
  if (mappingSheet) {
    try {
      var cellRange = mappingSheet.getRange("A2:A1000");
      cellRange.clearDataValidations();
      
      var sourceRange = masterSheet.getRange("B2:B1000");
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceRange, true)
        .setAllowInvalid(true)
        .build();
      cellRange.setDataValidation(rule);
      Logger.log("แก้ไข Data Validation ของ Document_Control_Mapping!A2:A1000 สำเร็จ");
    } catch(e) {
      Logger.log("ไม่สามารถแก้ Validation ของ Document_Control_Mapping: " + e.message);
    }
  }
}
