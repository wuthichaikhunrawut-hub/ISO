/**
 * ระบบจัดการไฟล์มาตรฐาน ISO 27001
 * Services.js - โมดูลสำหรับจัดการ Business Logic ทั้งหมดของระบบ
 */

/**
 * ดึงสิทธิ์ของผู้ใช้ปัจจุบัน
 */
function getCurrentUserRole() {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    // กรณีทดสอบใน Script Editor หรือเรียกใช้ทั่วไปแบบไม่ระบุตัวตน
    return { email: "test@company.com", role: "Admin", name: "Developer (Local Test)", allowed: true };
  }
  
  var user = DB.getRowByField("User_Access_Matrix", "user_email", email);
  if (!user) {
    return { email: email, role: "Viewer", name: "Guest User", allowed: false };
  }
  
  return {
    email: email,
    role: user.role,
    name: user.name,
    department: user.department,
    allowed: (user.account_status === "Active")
  };
}

/**
 * ฟังก์ชันสำหรับเขียนบันทึกประวัติการเปลี่ยนแปลงข้อมูล (Audit Logs)
 */
function logAuditTrail(action, refTable, refId, changeDetails, ipAddress) {
  var user = getCurrentUserRole();
  var logData = {
    timestamp: new Date(),
    action: action,
    ref_table: refTable,
    ref_id: String(refId),
    performed_by: user.email,
    change_details: changeDetails || "",
    ip_address: ipAddress || ""
  };
  
  DB.insertRow("Audit_Trails", logData);
}

/**
 * ดึงรายการเอกสารหลักทั้งหมดพร้อมรายชื่อ Control ที่แมปไว้
 */
function getMasterDocuments() {
  var docs = DB.getTableData("Master_Documents");
  
  // แปลงค่าคอมมาในคอลัมน์ policy ให้กลับมาเป็นอาร์เรย์ controls เพื่อใช้ในหน้าบ้าน
  docs.forEach(function(d) {
    d.controls = d.policy ? String(d.policy).split(",").map(function(c) { return c.trim(); }).filter(Boolean) : [];
  });
  
  return docs;
}

/**
 * บันทึกหรือแก้ไขเอกสารหลัก (Save/Update Document)
 */
function saveMasterDocument(docData, selectedControls, ipAddress) {
  var user = getCurrentUserRole();
  if (user.role !== "Admin") {
    throw new Error("คุณไม่มีสิทธิ์ในการบันทึกเอกสารหลัก");
  }
  
  // บันทึกรหัส Control ที่เชื่อมโยงเข้าคอลัมน์ policy แบบคอมมาคั่นตรงๆ เพื่อความง่ายในการแก้ไขในชีต
  docData.policy = selectedControls ? selectedControls.join(", ") : "";
  
  var existing = DB.getRowByField("Master_Documents", "doc_id", docData.doc_id);
  
  // คำนวณวันที่ต้องทบทวนถัดไป (Next Review Date) อัตโนมัติ
  if (docData.effective_date && docData.review_frequency_months) {
    var effDate = new Date(docData.effective_date);
    effDate.setMonth(effDate.getMonth() + Number(docData.review_frequency_months));
    docData.next_review_date = Utilities.formatDate(effDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  
  if (existing) {
    // กรณีแก้ไข (Update)
    DB.updateRow("Master_Documents", "doc_id", docData.doc_id, docData);
    logAuditTrail("Update", "Master_Documents", docData.doc_id, "Updated document: " + docData.title, ipAddress);
  } else {
    // กรณีสร้างใหม่ (Insert)
    DB.insertRow("Master_Documents", docData);
    logAuditTrail("Create", "Master_Documents", docData.doc_id, "Created document: " + docData.title, ipAddress);
  }
  
  return { success: true };
}

/**
 * ลบเอกสารหลัก
 */
function deleteMasterDocument(docId, ipAddress) {
  var user = getCurrentUserRole();
  if (user.role !== "Admin") {
    throw new Error("คุณไม่มีสิทธิ์ในการลบเอกสารหลัก");
  }
  
  var doc = DB.getRowByField("Master_Documents", "doc_id", docId);
  if (!doc) throw new Error("ไม่พบเอกสารที่ต้องการลบ");
  
  DB.deleteRow("Master_Documents", "doc_id", docId);
  
  // ตรวจสอบความปลอดภัยกรณีไม่มีตาราง Mapping แยกแล้ว
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mapSheet = ss.getSheetByName("Document_Control_Mapping");
  if (mapSheet) {
    var mappings = DB.getTableData("Document_Control_Mapping");
    for (var i = mappings.length - 1; i >= 0; i--) {
      if (mappings[i].doc_id === docId) {
        mapSheet.deleteRow(mappings[i]._rowNum);
      }
    }
  }
  
  logAuditTrail("Delete", "Master_Documents", docId, "Deleted document: " + doc.title, ipAddress);
  return { success: true };
}

/**
 * จัดการหาหรือสร้างโฟลเดอร์สำหรับเก็บหลักฐานของแต่ละ Control ใน Google Drive
 */
function getOrCreateControlFolder(controlCode) {
  var rootName = "ISO27001_System_Files";
  var folderName = "Evidence_" + controlCode.replace(/\./g, "_");
  
  var rootFolder;
  var rootSearch = DriveApp.getFoldersByName(rootName);
  
  if (rootSearch.hasNext()) {
    rootFolder = rootSearch.next();
  } else {
    rootFolder = DriveApp.createFolder(rootName);
  }
  
  var subFolder;
  var subSearch = rootFolder.getFoldersByName(folderName);
  if (subSearch.hasNext()) {
    subFolder = subSearch.next();
  } else {
    subFolder = rootFolder.createFolder(folderName);
  }
  
  return subFolder;
}

/**
 * อัปโหลดไฟล์หลักฐานและบันทึกลงในระบบ
 * fileData: { data: "base64...", mimeType: "...", fileName: "..." }
 */
function uploadEvidence(docId, controlCode, period, details, fileData, ipAddress) {
  var user = getCurrentUserRole();
  if (!user.allowed || user.role === "Auditor") {
    throw new Error("คุณไม่มีสิทธิ์ในการยื่นไฟล์หลักฐาน");
  }
  
  var fileUrl = "";
  if (fileData && fileData.data) {
    // 1. ถอดรหัสไฟล์อัปโหลด
    var decoded = Utilities.base64Decode(fileData.data);
    var blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.fileName);
    
    // 2. หาโฟลเดอร์เก็บไฟล์และบันทึกลง Google Drive
    var folder = getOrCreateControlFolder(controlCode);
    var file = folder.createFile(blob);
    
    // 3. กำหนดสิทธิ์ให้ผู้มีลิงก์เปิดดูได้ (อย่างน้อยสำหรับผู้ตรวจสอบสิทธิ์)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileUrl = file.getUrl();
  } else {
    throw new Error("ไม่พบไฟล์ที่อัปโหลด");
  }
  
  // 4. บันทึกข้อมูลการยื่นหลักฐานลงตาราง
  var evidenceData = {
    recorded_date: new Date(),
    doc_id: docId,
    evidence_period: period,
    details: details,
    drive_link: fileUrl,
    recorded_by: user.email
  };
  
  var inserted = DB.insertRow("Evidence_Logs", evidenceData);
  
  // 5. บันทึก Audit Log
  logAuditTrail("Create", "Evidence_Logs", inserted.evidence_id, "Uploaded evidence file for: " + docId + " (" + fileData.fileName + ")", ipAddress);
  
  return { success: true, evidence_id: inserted.evidence_id };
}

/**
 * อัปเดตสถานะการประเมินตนเอง (Assessment Status)
 */
function updateControlStatus(controlCode, status, gapAnalysis, ipAddress) {
  var user = getCurrentUserRole();
  if (user.role !== "Admin") {
    throw new Error("คุณไม่มีสิทธิ์ในการปรับปรุงสถานะการประเมิน");
  }
  
  DB.updateRow("ISO_Controls", "control_code", controlCode, {
    assessment_status: status,
    gap_analysis: gapAnalysis
  });
  
  logAuditTrail("Update", "ISO_Controls", controlCode, "Updated assessment status to: " + status, ipAddress);
  return { success: true };
}

/**
 * ดึงความคืบหน้าภาพรวมสัดส่วนสถานะการประเมินตนเอง
 */
function getComplianceStats() {
  var controls = DB.getTableData("ISO_Controls");
  
  var stats = {
    "Ready": 0,
    "In Progress": 0,
    "Missing": 0,
    "N/A": 0,
    "Total": controls.length
  };
  
  controls.forEach(function(c) {
    var status = c.assessment_status || "Missing";
    if (status in stats) {
      stats[status]++;
    } else {
      stats["Missing"]++;
    }
  });
  
  return stats;
}

/**
 * อัปโหลดไฟล์เอกสารหลักควบคุมเก็บเข้าโฟลเดอร์หลักของ Drive อัตโนมัติ
 */
function uploadMasterDocumentFile(fileName, mimeType, base64Data) {
  var user = getCurrentUserRole();
  if (user.role !== "Admin") {
    throw new Error("คุณไม่มีสิทธิ์ในการอัปโหลดเอกสารหลัก");
  }
  
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType, fileName);
  
  var rootName = "ISO27001_System_Files";
  var rootFolder;
  var rootSearch = DriveApp.getFoldersByName(rootName);
  if (rootSearch.hasNext()) {
    rootFolder = rootSearch.next();
  } else {
    rootFolder = DriveApp.createFolder(rootName);
  }
  
  var file = rootFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}

function updateEvidenceLog(evidenceId, period, details, docId, fileData, controlCode, ipAddress, deleteOldFile) {
  var user = getCurrentUserRole();
  var log = DB.getRowByField("Evidence_Logs", "evidence_id", evidenceId);
  if (!log) throw new Error("ไม่พบรายการหลักฐานที่ต้องการแก้ไข");
  
  if (user.role !== "Admin" && log.recorded_by !== user.email) {
    throw new Error("คุณไม่มีสิทธิ์แก้ไขรายการหลักฐานนี้");
  }
  
  var updated = {
    evidence_period: period,
    details: details,
    doc_id: docId,
    drive_link: log.drive_link // ค่าเริ่มต้นใช้ของเดิม
  };
  
  // หากสั่งลบไฟล์เดิม (หรือมีการอัปโหลดไฟล์ใหม่มาแทนที่)
  if (deleteOldFile || (fileData && fileData.data)) {
    try {
      if (log.drive_link) {
        var fileId = log.drive_link.match(/id=([^&]+)/) || log.drive_link.match(/\/d\/([^/]+)/);
        if (fileId && fileId[1]) {
          DriveApp.getFileById(fileId[1]).setTrashed(true);
        }
      }
    } catch(e) {
      Logger.log("ไม่สามารถย้ายไฟล์เก่าลงถังขยะได้: " + e.message);
    }
    updated.drive_link = ""; // ล้างลิงก์เดิมออกชั่วคราว
  }
  
  // หากมีการอัปโหลดไฟล์ใหม่เข้ามา
  if (fileData && fileData.data) {
    var decoded = Utilities.base64Decode(fileData.data);
    var blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.fileName);
    var folder = getOrCreateControlFolder(controlCode || "General_Evidence");
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    updated.drive_link = file.getUrl();
  }
  
  DB.updateRow("Evidence_Logs", "evidence_id", evidenceId, updated);
  logAuditTrail("Update", "Evidence_Logs", evidenceId, "Updated evidence log (Doc: " + docId + ", Period: " + period + ")", ipAddress);
  return { success: true };
}

/**
 * ลบประวัติหลักฐานการยื่นตรวจสอบ (ย้ายไฟล์ใน Drive ลงถังขยะด้วย)
 */
function deleteEvidenceLog(evidenceId, ipAddress) {
  var user = getCurrentUserRole();
  var log = DB.getRowByField("Evidence_Logs", "evidence_id", evidenceId);
  if (!log) throw new Error("ไม่พบรายการหลักฐานที่ต้องการลบ");
  
  if (user.role !== "Admin" && log.recorded_by !== user.email) {
    throw new Error("คุณไม่มีสิทธิ์ในการลบรายการหลักฐานนี้");
  }
  
  // พยายามย้ายไฟล์ Google Drive ลงถังขยะ (Trash)
  try {
    if (log.drive_link) {
      var fileId = log.drive_link.match(/id=([^&]+)/) || log.drive_link.match(/\/d\/([^/]+)/);
      if (fileId && fileId[1]) {
        DriveApp.getFileById(fileId[1]).setTrashed(true);
      }
    }
  } catch(e) {
    Logger.log("ไม่สามารถย้ายไฟล์ลงถังขยะได้: " + e.message);
  }
  
  DB.deleteRow("Evidence_Logs", "evidence_id", evidenceId);
  logAuditTrail("Delete", "Evidence_Logs", evidenceId, "Deleted evidence log (Doc: " + log.doc_id + ", Period: " + log.evidence_period + ")", ipAddress);
  return { success: true };
}
