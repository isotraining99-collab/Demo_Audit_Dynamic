// 🌟 ใส่ SHEET_ID ของบัญชีเดโม่ของคุณที่นี่
var SHEET_ID = "1zUeF35TzE3l6pm45cMKclPGx0XGHNvaEE6xhAcPT5sw"; 

function doGet() { 
  try {
    // 🌟 ใช้ชื่อไฟล์ index สำหรับเดโม่
    var output = HtmlService.createTemplateFromFile('01.CustomerAudit').evaluate(); 
    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
    return output;
  } catch (e) {
    return HtmlService.createHtmlOutput("Error Loading System: " + e.message);
  }
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// ==========================================
// 🌟 DYNAMIC MODULES CONTROL (ระบบสร้างหน้าต่างใหม่)
// ==========================================
function getSystemModules() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var master = ss.getSheetByName("Master_Modules");
    if (!master) {
      master = ss.insertSheet("Master_Modules");
      master.appendRow(["Module_ID", "Module_Name", "Icon", "Type"]);
      if(ss.getSheetByName("01_CustomAudit")) master.appendRow(["M01", "01_CustomAudit", "bi-file-earmark-richtext", "RichText"]);
      if(ss.getSheetByName("03_AuditMaterial")) master.appendRow(["M03", "03_AuditMaterial", "bi-box-seam", "RichText"]);
    }
    var data = master.getDataRange().getValues();
    if(data.length > 0) data.shift(); 
    return data;
  } catch (e) { throw new Error(e.message); }
}

function createNewModule(moduleName, iconClass) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var master = ss.getSheetByName("Master_Modules");
    var newId = "M" + new Date().getTime();
    master.appendRow([newId, moduleName, iconClass, "RichText"]);

    var newSheet = ss.getSheetByName(moduleName);
    if (!newSheet) {
      newSheet = ss.insertSheet(moduleName);
      newSheet.appendRow(["Timestamp","Year","Customer","Audit_ID","Doc_Type","Content_HTML","Editor","Real_Name","Status","Parent_ID"]);
    }
    return "Success: โมดูลใหม่ [" + moduleName + "] ถูกสร้างและพร้อมใช้งานแล้ว!";
  } catch (e) { throw new Error(e.message); }
}

// ==========================================
// 🌟 UNIVERSAL DYNAMIC FUNCTIONS
// ==========================================
function getDynamicData(sheetName) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    if(!sheet) return [];
    var data = sheet.getDataRange().getDisplayValues(); 
    if (data.length <= 1) return []; data.shift(); return data;
  } catch (e) { throw new Error(e.message); }
}

function createDynamicSession(sheetName, year, customerName, realName, docTitle, parentId) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    var rawDate = new Date();
    var timestampStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User"; 
    var auditId = "DY-" + Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    var pId = parentId || ""; 
    sheet.appendRow([timestampStr, year, customerName, auditId, docTitle, "", userEmail, realName, "Draft", pId]);
    SpreadsheetApp.flush(); return "Success: สร้างเอกสาร [" + auditId + "] เรียบร้อยแล้ว";
  } catch (e) { throw new Error(e.message); }
}

function renameDynamicDocTitle(sheetName, auditId, newTitle) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { 
        if (data[i][3] == auditId) { sheet.getRange(i + 1, 5).setValue(newTitle); SpreadsheetApp.flush(); return "Success: อัปเดตชื่อเอกสารแล้ว"; } 
    } throw new Error("ไม่พบข้อมูล");
  } catch(e) { throw new Error(e.message); }
}

function saveDynamicDocContent(sheetName, auditId, htmlContent) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues(); 
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User";

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === String(auditId).trim()) { 
        var targetRow = i + 1;
        if (htmlContent.length > 45000) throw new Error("เนื้อหาใหญ่เกินไป กรุณาลบรูปลิงก์ภายนอกออกแล้วใช้อัปโหลดเข้า Drive");
        sheet.getRange(targetRow, 1).setValue(timestampStr); 
        sheet.getRange(targetRow, 6).setValue(htmlContent); 
        sheet.getRange(targetRow, 7).setValue(userEmail);
        sheet.getRange(targetRow, 9).setValue("Completed"); 
        SpreadsheetApp.flush(); return "บันทึกข้อมูลเรียบร้อยแล้ว!";
      }
    } throw new Error("หา ID ไม่พบ");
  } catch (e) { throw new Error(e.message); }
}

function deleteDynamicSession(sheetName, auditId) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues();
    var idsToDelete = [auditId];
    function findChildren(id) {
      for (var i = 1; i < data.length; i++) { if (data[i][9] == id) { idsToDelete.push(data[i][3]); findChildren(data[i][3]); } }
    }
    findChildren(auditId);
    for (var i = data.length - 1; i >= 1; i--) { if (idsToDelete.indexOf(data[i][3]) !== -1) { sheet.deleteRow(i + 1); } }
    SpreadsheetApp.flush(); return "Success: ลบเอกสารและหัวข้อย่อยเรียบร้อย";
  } catch(e) { throw new Error(e.message); }
}

// ==========================================
// 🌟 02 Audit Result (ระบบ Upload)
// ==========================================
function getAuditData() {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getDisplayValues(); 
    if (data.length <= 1) return []; data.shift(); return data;
  } catch (e) { throw new Error(e.message); }
}

function createNewCustomerFolders(year, customerName, realName) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var rawDate = new Date();
    var timestampStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User";
    var auditId = "AD-" + Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyMMdd-HHmmss");
    var docTypes = ["1. Plant Tour information", "2. Audit Time Table", "3. Audit Preparation Material", "4. Audit Report (Customer)", "5. Audit Report (Internal)"];
    var dataToWrite = [];
    for (var i = 0; i < docTypes.length; i++) { dataToWrite.push([timestampStr, year, customerName, auditId, docTypes[i], "", "", userEmail, realName, ""]); }
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);
    SpreadsheetApp.flush(); return "Success: Audit Session [" + auditId + "] has been created.";
  } catch (e) { throw new Error(e.message); }
}

function uploadFileToDrive(base64Data, fileName, year, customerName, auditId, docType, folderId, detail) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var splitBase = base64Data.split(','); var type = splitBase[0].split(';')[0].replace('data:', '');
    var blob = Utilities.newBlob(Utilities.base64Decode(splitBase[1]), type, fileName);
    var file = folder.createFile(blob);
    
    // 🌟 เปลี่ยนเป็น ANYONE_WITH_LINK สำหรับ Demo
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    
    var fileUrl = file.getUrl();
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    var userEmail = Session.getActiveUser().getEmail() || "Unknown User";
    var realName = ""; var emptyRowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][3] == auditId && data[i][4] == docType) { realName = data[i][8]; if (data[i][5] == "" && emptyRowIndex == -1) emptyRowIndex = i + 1; }
    }
    if (emptyRowIndex != -1) {
      sheet.getRange(emptyRowIndex, 1).setValue(timestampStr); sheet.getRange(emptyRowIndex, 6).setValue(fileName); sheet.getRange(emptyRowIndex, 7).setValue(fileUrl); sheet.getRange(emptyRowIndex, 8).setValue(userEmail); sheet.getRange(emptyRowIndex, 10).setValue(detail); 
    } else { sheet.appendRow([timestampStr, year, customerName, auditId, docType, fileName, fileUrl, userEmail, realName, detail]); }
    SpreadsheetApp.flush(); return "Success: File uploaded and details saved.";
  } catch (e) { throw new Error(e.message); }
}

function editFileDetail(auditId, docType, fileLink, newDetail) {
  try { var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { if (data[i][3] == auditId && data[i][4] == docType && data[i][6] == fileLink) { sheet.getRange(i + 1, 10).setValue(newDetail); SpreadsheetApp.flush(); return "Success: Detail updated."; } } throw new Error("File not found."); } catch(e) { throw new Error(e.message); }
}
function removeFileFromSheet(auditId, docType, fileLink) {
  try { var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues(); var matchCount = 0; var targetRow = -1;
    for (var i = 1; i < data.length; i++) { if (data[i][3] == auditId && data[i][4] == docType) { matchCount++; if (data[i][6] == fileLink) targetRow = i + 1; } }
    if (targetRow != -1) { var fLink = sheet.getRange(targetRow, 7).getValue(); if (fLink) { var fileIdMatch = fLink.match(/[-\w]{25,}/); if (fileIdMatch) { try { DriveApp.getFileById(fileIdMatch[0]).setTrashed(true); } catch(e) {} } }
      if (matchCount > 1) { sheet.deleteRow(targetRow); } else { sheet.getRange(targetRow, 6).clearContent(); sheet.getRange(targetRow, 7).clearContent(); sheet.getRange(targetRow, 10).clearContent(); } SpreadsheetApp.flush(); return "Success: File removed from system and Drive."; } } catch(e) { throw new Error(e.message); }
}
function deleteAuditSession(auditId) {
  try { var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues(); for (var i = data.length - 1; i >= 1; i--) { if (data[i][3] == auditId) sheet.deleteRow(i + 1); } SpreadsheetApp.flush(); return "Success: Session deleted."; } catch(e) { throw new Error(e.message); }
}
function renameCustomerInSheet(oldName, newName) {
  try { var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0]; var data = sheet.getDataRange().getValues(); for (var i = 1; i < data.length; i++) { if (data[i][2] == oldName) sheet.getRange(i + 1, 3).setValue(newName); } SpreadsheetApp.flush(); return "Success: Customer name updated."; } catch(e) { throw new Error(e.message); }
}

// ==========================================
// 🌟 MEDIA & EXPORT FUNCTIONS
// ==========================================
function uploadMediaToDrive(base64Data, fileName, isImage) {
  try {
    // 🌟 ใส่ FOLDER_ID ของบัญชีเดโม่
    var folderId = "1cX4x5WnHrKUg0zNWHH-8qjQ3d1iCWbwQ"; 
    var folder = DriveApp.getFolderById(folderId);
    var splitBase = base64Data.split(','); var type = splitBase[0].split(';')[0].replace('data:', '');
    var blob = Utilities.newBlob(Utilities.base64Decode(splitBase[1]), type, fileName); var file = folder.createFile(blob);
    
    // 🌟 เปลี่ยนเป็น ANYONE_WITH_LINK สำหรับ Demo
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e) {}
    
    if (isImage) { return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000"; } else { return file.getUrl(); }
  } catch (e) { throw new Error(e.message); }
}

function getBase64FromDrive(fileIds) {
  try { var result = {}; for(var i = 0; i < fileIds.length; i++) { var id = fileIds[i]; try { var file = DriveApp.getFileById(id); var blob = file.getBlob(); var b64 = Utilities.base64Encode(blob.getBytes()); var mime = blob.getContentType(); result[id] = "data:" + mime + ";base64," + b64; } catch(e) { result[id] = null; } } return result; } catch(e) { throw new Error(e.message); }
}
