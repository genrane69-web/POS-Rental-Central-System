// ⚠️ เอา ID ที่คัดลอกมาจากขั้นตอนที่ 1 มาวางแทนที่ข้อความภาษาไทยในเครื่องหมายคำพูดนะครับ
var TEMPLATE_SHEET_ID = "17D9HFfhNY1KIazj3AoGTrjMD22GVdAP5kuhQlQmu1QU";
var ADMIN_SHEET_ID    = "17gIAKqnX3Hde5J7fcBjB8wTTaE7UP-nbLOHrLh-PRK4";

function doGet(e) {
  var userEmail = Session.getActiveUser().getEmail();
  var page = e.parameter.page;
  
  if (page === "admin") {
    var adminTemplate = HtmlService.createTemplateFromFile("AdminDashboard");
    return adminTemplate.evaluate().setTitle("ระบบจัดการผู้เช่า");
  }

  var tenantInfo = getTenantInfo(userEmail);

  if (!tenantInfo) {
    var setupTemplate = HtmlService.createTemplateFromFile("SetupPage");
    setupTemplate.userEmail = userEmail;
    return setupTemplate.evaluate().setTitle("ติดตั้งระบบ POS");
  }

  if (tenantInfo.status !== "Active" || new Date() > new Date(tenantInfo.expireDate)) {
    return HtmlService.createHtmlOutput("<h1 style='color:red;text-align:center;margin-top:50px;'>⛔ สิทธิ์ใช้งานหมดอายุ กรุณาติดต่อผู้บริการ</h1>");
  }

  var posTemplate = HtmlService.createTemplateFromFile("PosPage");
  posTemplate.sheetId = tenantInfo.sheetId;
  posTemplate.userEmail = userEmail;
  return posTemplate.evaluate().setTitle("ระบบ POS ขายหน้าร้าน");
}

function setupNewStore() {
  var userEmail = Session.getActiveUser().getEmail();
  var existing = getTenantInfo(userEmail);
  if (existing) return { success: true };

  try {
    var templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
    var newFile = templateFile.makeCopy("ระบบ POS - ร้านของ " + userEmail);
    var newSheetId = newFile.getId();

    var expireDate = new Date();
    expireDate.setDate(expireDate.getDate() + 30); // ทดลองใช้ฟรี 30 วัน

    var adminSS = SpreadsheetApp.openById(ADMIN_SHEET_ID);
    adminSS.getSheets()[0].appendRow([userEmail, newSheetId, new Date(), expireDate, "Active"]);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getProducts(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName("Products") || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  data.shift();
  return data;
}

function getAllTenants() {
  var adminSS = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  var sheet = adminSS.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(function(row) {
    return { email: row[0], installDate: String(row[2]), expireDate: String(row[3]), status: row[4] };
  });
}

function updateTenantStatus(email, status, addDays) {
  var adminSS = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  var sheet = adminSS.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      if (status) sheet.getRange(i + 1, 5).setValue(status);
      if (addDays > 0) {
        var base = new Date(data[i][3]) > new Date() ? new Date(data[i][3]) : new Date();
        base.setDate(base.getDate() + parseInt(addDays));
        sheet.getRange(i + 1, 4).setValue(base);
      }
      return { success: true };
    }
  }
}

function getTenantInfo(email) {
  var adminSS = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  var sheet = adminSS.getSheets()[0];
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email) return { email: data[i][0], sheetId: data[i][1], expireDate: data[i][3], status: data[i][4] };
  }
  return null;
}
