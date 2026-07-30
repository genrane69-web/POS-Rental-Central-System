/**
 * 1. ตรวจสอบ PIN ล็อกอินเข้าใช้งาน
 */
function checkLogin(customerSheetId, pin) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('staffs');
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน 'staffs'" };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const username = String(data[i][0]).trim();
      const userPin = String(data[i][1]).trim();
      const role = String(data[i][2]).trim().toLowerCase();
      
      if (userPin === String(pin).trim()) {
        return { success: true, user: { name: username, role: role } };
      }
    }
    return { success: false, message: "รหัส PIN ไม่ถูกต้อง" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 2. ดึงข้อมูลสินค้าทั้งหมด
 */
function getProducts(customerSheetId) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Products');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const products = [];
    for (let i = 1; i < data.length; i++) {
      products.push({
        barcode: String(data[i][0]),
        name: data[i][1],
        category: data[i][2],
        cost: Number(data[i][3]) || 0,
        price: Number(data[i][4]) || 0,
        stock: Number(data[i][5]) || 0,
        minStock: Number(data[i][6]) || 5
      });
    }
    return products;
  } catch (error) {
    throw new Error("ดึงข้อมูลสินค้าไม่สำเร็จ: " + error.toString());
  }
}

/**
 * 3. บันทึกการขาย + ตัดสต็อกสินค้า + บันทึกรายการสินค้าลง Transactions
 */
function processSale(customerSheetId, saleData) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    const productSheet = ss.getSheetByName('Products');
    const txSheet = ss.getSheetByName('Transactions');
    const cashLogSheet = ss.getSheetByName('CashLogs');
    const timestamp = new Date();
    const txId = "TX-" + timestamp.getTime();
    const staff = saleData.staffName || "พนักงาน";
    const paymentMethod = saleData.paymentMethod || "เงินสด";
    const cart = saleData.cart || [];
    const totalAmount = saleData.totalAmount || 0;
    const discount = saleData.discount || 0;

    // แปลงตะกร้าสินค้าเป็นข้อความ JSON สำหรับเก็บบันทึกใบเสร็จ
    const cartJson = JSON.stringify(cart);

    if (txSheet) {
      // บันทึกลง Transactions: A=TxID, B=Timestamp, C=Staff, D=Total, E=Discount, F=Payment, G=Status, H=CartJSON
      txSheet.appendRow([txId, timestamp, staff, totalAmount, discount, paymentMethod, 'Completed', cartJson]);
    }

    const pData = productSheet ? productSheet.getDataRange().getValues() : [];

    cart.forEach(item => {
      if (cashLogSheet) {
        cashLogSheet.appendRow([
          timestamp, txId, staff, item.barcode, item.name, item.qty, item.price, item.qty * item.price, paymentMethod
        ]);
      }
      if (productSheet && pData.length > 1) {
        for (let i = 1; i < pData.length; i++) {
          if (String(pData[i][0]) === String(item.barcode)) {
            const currentStock = Number(pData[i][5]) || 0;
            const newStock = currentStock - item.qty;
            productSheet.getRange(i + 1, 6).setValue(newStock);
            
            pData[i][5] = newStock;
            break;
          }
        }
      }
    });

    return { success: true, txId: txId, message: "ชำระเงินและตัดสต็อกสำเร็จ!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
// ==========================================================
// 1. ตั้งค่า ID สำหรับระบบ SaaS
// ==========================================================
const ADMIN_SHEET_ID    = "17gIAKqnX3Hde5J7fcBjB8wTTaE7UP-nbLOHrLh-PRK4"; // ID ชีตแอดมิน
const TEMPLATE_SHEET_ID = "17D9HFfhNY1KIazj3AoGTrjMD22GVdAP5kuhQlQmu1QU"; // ID ชีตแม่แบบ POS

// ==========================================================
// 2. ฟังก์ชัน doGet(e) - ตัวรับลิงก์หลัก (ขามหน้าติดตั้งให้อัตโนมัติถ้าเคยติดตั้งแล้ว)
// ==========================================================
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  const txId = (e && e.parameter) ? (e.parameter.txId || e.parameter.receiptId || '') : '';

  // 1️⃣ ถ้าเป็นการเปิดดูใบเสร็จ (ลูกค้าสแกน QR Code)
  if (page === 'receipt' && txId) {
    const tpl = HtmlService.createTemplateFromFile('Receipt');
    tpl.txId = txId;
    tpl.customerSheetId = (e && e.parameter && e.parameter.sheetId) ? e.parameter.sheetId : '';
    
    return tpl.evaluate()
      .setTitle('ใบเสร็จรับเงิน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // 2️⃣ ดึงอีเมลของผู้ที่เปิดลิงกเข้ามาใช้งาน
  const userEmail = Session.getActiveUser().getEmail();

  // 3️⃣ ค้นหา Sheet ID ของลูกคาคนนี้จากชีตแอดมิน
  const customerInfo = getCustomerSheetIdByEmail(userEmail);

  // 4️⃣ ถ้ายังไม่มีในชีตแอดมิน -> เปิดหน้าเด้งติดตั้ง (Install.html)
  if (!customerInfo.isInstalled) {
    const tpl = HtmlService.createTemplateFromFile('Install');
    tpl.userEmail = userEmail;
    return tpl.evaluate()
      .setTitle('ติดตั้งระบบ VEGA POS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  // 5️⃣ ถ้าติดตั้งแล้ว -> เปิดหน้า POS (Index.html) พร้อมส่ง customerSheetId ไปใช้งาน
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.customerSheetId = customerInfo.sheetId; 
  tpl.webAppUrl = ScriptApp.getService().getUrl(); // ← บรรทัดใหม่ที่เพิ่มเข้ามา
  
  return tpl.evaluate()
    .setTitle('VEGA POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ==========================================================
// 3. ฟังก์ชันค้นหา Sheet ID ในชีตแอดมิน จากอีเมลผู้ใช้ (ปรับปรุงระบบเช็กไฟล์)
// ==========================================================
function getCustomerSheetIdByEmail(email) {
  if (!email || String(email).trim() === "") return { isInstalled: false };

  try {
    const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
    const sheet = ss.getSheetByName('ชีต1'); // ชื่อแท็บด้านล่างของชีตแอดมิน
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const cellEmail = String(data[i][0]).trim().toLowerCase();
      if (cellEmail !== "" && cellEmail === String(email).trim().toLowerCase()) {
        const sheetId = data[i][1];
        
        // เช็กว่าไฟล์ใน Google Drive ของลูกค้ารายนี้ยังอยู่ไหม (ถ้าโดนลบไปแล้ว จะพาไปหน้าติดตั้งใหม่)
        try {
          DriveApp.getFileById(sheetId);
          return {
            isInstalled: true,
            sheetId: sheetId, // คอลัมน์ B: UserSheetID
            status: data[i][4]  // คอลัมน์ E: Status
          };
        } catch (fileErr) {
          return { isInstalled: false };
        }
      }
    }
  } catch (err) {
    console.error("Error finding customer:", err);
  }
  return { isInstalled: false };
}

// ==========================================================
// 4. ฟังก์ชันติดตั้งระบบ 1-Click (ก๊อปปี้แม่แบบ + ลงตารางแอดมิน)
// ==========================================================
function installSystem() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) return { success: false, message: "ไม่พบอีเมลผู้ใช้งาน กรุณาลงชื่อเข้าใช้ Google" };

    // ก๊อปปี้ Google Sheets แม่แบบ ออกมาเป็นไฟล์ใหม่ให้ลูกค้า
    const templateFile = DriveApp.getFileById(TEMPLATE_SHEET_ID);
    const newFile = templateFile.makeCopy("VEGA POS - " + userEmail);
    newFile.addEditor(userEmail); // ให้สิทธิ์ลูกค้าเป็นผู้แก้ไข
    const newSheetId = newFile.getId();

    // คำนวณวันติดตั้ง และวันหมดอายุ (+30 วัน)
    const now = new Date();
    const expire = new Date();
    expire.setDate(now.getDate() + 30);
    const formatDate = (d) => Utilities.formatDate(d, "Asia/Bangkok", "d/M/yyyy, H:mm:ss");

    // บันทึกลงตารางแอดมิน
    const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
    const sheet = ss.getSheetByName('ชีต1');
    const data = sheet.getDataRange().getValues();

    // ค้นหาว่าเคยมีอีเมลนี้ในตารางไหม
    let existingRowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === String(userEmail).trim().toLowerCase()) {
        existingRowIndex = i + 1; // +1 เพราะแถวใน Sheet เริ่มที่ 1
        break;
      }
    }

    if (existingRowIndex > 0) {
      // ถ้าเคยมีอยู่แล้ว ให้ทับบรรทัดเดิม
      sheet.getRange(existingRowIndex, 2).setValue(newSheetId);
      sheet.getRange(existingRowIndex, 3).setValue(formatDate(now));
      sheet.getRange(existingRowIndex, 4).setValue(formatDate(expire));
      sheet.getRange(existingRowIndex, 5).setValue("Active");
    } else {
      // ถ้าเป็นคนใหม่ ให้เพิ่มแถวใหม่ต่อท้าย
      sheet.appendRow([
        userEmail,
        newSheetId,
        formatDate(now),
        formatDate(expire),
        "Active"
      ]);
    }

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 5. ฟังก์ชันดึงข้อมูลใบเสร็จหลัก (เรียกใช้จาก Receipt.html)
 */
function getReceiptDetail(customerSheetId, txId) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Transactions');
    if (!sheet) return { success: false, message: "ไม่พบชีตชื่อ Transactions" };
    
    const storeConfig = getStoreConfig(customerSheetId) || {};
    const data = sheet.getDataRange().getValues();
    const targetTxId = String(txId).trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === targetTxId) {
        
        let parsedItems = [];
        // อ่าน JSON รายการสินค้าจาก คอลัมน์ H (Index 7)
        if (data[i][7]) {
          try {
            parsedItems = typeof data[i][7] === 'string' ? JSON.parse(data[i][7]) : data[i][7];
          } catch (e) {
            parsedItems = [];
          }
        }

        const formattedDate = data[i][1] instanceof Date 
          ? Utilities.formatDate(data[i][1], "Asia/Bangkok", "dd/MM/yyyy HH:mm") 
          : data[i][1];

        return {
          success: true,
          data: {
            txId: data[i][0],
            date: formattedDate,
            staff: data[i][2],
            total: Number(data[i][3]) || 0,
            discount: Number(data[i][4]) || 0,
            subtotal: (Number(data[i][3]) || 0) + (Number(data[i][4]) || 0),
            paymentMethod: data[i][5] || 'เงินสด',
            status: data[i][6],
            items: parsedItems
          },
          storeConfig: storeConfig
        };
      }
    }
    
    return { success: false, message: "ไม่พบข้อมูลใบเสร็จเลขที่นี้" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function saveProduct(customerSheetId, product) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Products');
    if (!sheet) {
      return { success: false, message: "ไม่พบแผ่นงาน 'Products' ใน Google Sheet" };
    }
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(product.barcode)) {
        rowIndex = i + 1;
        break;
      }
    }

    const isEditing = !!product.isEditing;

    if (isEditing) {
      // โหมดแก้ไข: ต้องเจอแถวเดิมเท่านั้น ถ้าไม่เจอห้ามสร้างใหม่ทับมั่ว
      if (rowIndex <= 0) {
        return { success: false, message: "ไม่พบสินค้ารหัสนี้ในระบบ (อาจถูกลบไปแล้ว)" };
      }
    } else {
      // โหมดเพิ่มใหม่: ถ้าเจอบาร์โค้ดซ้ำ ให้ปฏิเสธทันที ห้ามทับ
      if (rowIndex > 0) {
        return { success: false, message: "รหัสบาร์โค้ดนี้มีสินค้าอยู่แล้วในระบบ กรุณาใช้รหัสอื่น หรือกดแก้ไขสินค้าเดิมแทน" };
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 2, 1, 6).setValues([[
        product.name,
        product.category || 'ทั่วไป',
        Number(product.cost) || 0,
        Number(product.price) || 0,
        Number(product.stock) || 0,
        Number(product.minStock) || 5
      ]]);
    } else {
      sheet.appendRow([
        String(product.barcode),
        product.name,
        product.category || 'ทั่วไป',
        Number(product.cost) || 0,
        Number(product.price) || 0,
        Number(product.stock) || 0,
        Number(product.minStock) || 5
      ]);
    }

    return { success: true, message: "บันทึกข้อมูลสินค้าเรียบร้อยแล้ว!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 7. ปรับสต็อกด่วน
 */
function adjustStock(customerSheetId, barcode, changeAmount) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Products');
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน 'Products'" };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(barcode)) {
        let currentStock = Number(data[i][5]) || 0;
        let newStock = currentStock + Number(changeAmount);
        
        // บันทึกค่าใหม่ลงระบบทันที (ยอมให้ติดลบได้ตามจริงแล้วครับ)
        sheet.getRange(i + 1, 6).setValue(newStock);
        return { success: true, newStock: newStock };
      }
    }
    return { success: false, message: "ไม่พบสินค้านี้ในระบบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 8. ระบบจัดการกะการทำงาน (Shifts)
 */
function getActiveShift(customerSheetId, staffName) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Shifts');
    if (!sheet) return { success: true, activeShift: null };
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(staffName) && String(data[i][13]) === 'Open') {
        return {
          success: true,
          activeShift: {
            rowIndex: i + 1,
            shiftId: data[i][0],
            staff: data[i][1],
            openTime: data[i][2] ? String(data[i][2]) : '',
            startCash: Number(data[i][4]) || 0
          }
        };
      }
    }
    return { success: true, activeShift: null };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function openShift(customerSheetId, staffName, startCash) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    let sheet = ss.getSheetByName('Shifts');
    
    if (!sheet) {
      sheet = ss.insertSheet('Shifts');
      sheet.appendRow([
        'ShiftID', 'Staff', 'OpenTime', 'CloseTime', 'StartCash', 
        'CashSales', 'CashIn', 'CashOut', 'TransferSales', 'TotalSales', 
        'ExpectedCash', 'ActualCash', 'Difference', 'Status'
      ]);
    }

    const now = new Date();
    const shiftId = 'SHIFT-' + now.getTime();
    const startCashNum = Number(startCash) || 0;
    const timeZone = Session.getScriptTimeZone();
    const openTimeStr = Utilities.formatDate(now, timeZone, 'dd/MM/yyyy HH:mm:ss');

    sheet.appendRow([
      shiftId, staffName, openTimeStr, '', startCashNum,
      0, 0, 0, 0, 0, startCashNum, 0, 0, 'Open'
    ]);

    return {
      success: true,
      message: "เปิดกะสำเร็จ!",
      shift: {
        shiftId: shiftId,
        staff: staffName,
        staffName: staffName,
        openTime: openTimeStr,
        startCash: startCashNum
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getShiftSummary(customerSheetId, staffName, openTime, shiftId) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    const txSheet = ss.getSheetByName('Transactions');
    const moveSheet = ss.getSheetByName('CashMovements');

    let shiftStart = parseCustomDate(customerSheetId, openTime);
    const shiftStartTime = shiftStart.getTime();

    let cashSales = 0;
    let transferSales = 0;
    let cashIn = 0;
    let cashOut = 0;

    if (txSheet) {
      const data = txSheet.getDataRange().getValues();
      const targetStaff = String(staffName || '').trim().toLowerCase();

      for (let i = 1; i < data.length; i++) {
        const txStaff = String(data[i][2] || '').trim().toLowerCase();
        let txTime = parseCustomDate(customerSheetId, data[i][1]);

        const txAmount = Number(data[i][3]) || 0;
        const paymentMethod = String(data[i][5] || '').trim();
        const status = String(data[i][6] || '').trim().toLowerCase();

        const matchStaff = !targetStaff || txStaff === targetStaff;
        const matchTime = !isNaN(shiftStartTime) && !isNaN(txTime.getTime()) && txTime.getTime() >= (shiftStartTime - 10000);
        const matchStatus = status === 'completed' || status === 'สำเร็จ' || status === '';

        if (matchStaff && matchTime && matchStatus) {
          if (paymentMethod === 'เงินสด') {
            cashSales += txAmount;
          } else {
            transferSales += txAmount;
          }
        }
      }
    }

    if (moveSheet) {
      const moveData = moveSheet.getDataRange().getValues();
      const targetShiftId = String(shiftId || '').trim();

      for (let i = 1; i < moveData.length; i++) {
        const mShiftId = String(moveData[i][1] || '').trim();
        const type = String(moveData[i][3] || '').trim();
        const amount = Number(moveData[i][4]) || 0;

        if (mShiftId === targetShiftId) {
          if (type === 'Paid In') cashIn += amount;
          if (type === 'Paid Out') cashOut += amount;
        }
      }
    }

    return {
      success: true,
      summary: {
        cashSales: cashSales,
        transferSales: transferSales,
        totalSales: cashSales + transferSales,
        cashIn: cashIn,
        cashOut: cashOut
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function closeShift(customerSheetId, data) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Shifts');
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน 'Shifts'" };

    const sheetData = sheet.getDataRange().getValues();
    const closeTime = new Date();
    const timeZone = Session.getScriptTimeZone();
    const closeTimeStr = Utilities.formatDate(closeTime, timeZone, 'dd/MM/yyyy HH:mm:ss');

    for (let i = 1; i < sheetData.length; i++) {
      if (String(sheetData[i][0]) === String(data.shiftId) && String(sheetData[i][13]) === 'Open') {
        const rowIndex = i + 1;
        
        const startCash = Number(sheetData[i][4]) || 0;
        const cashSales = Number(data.cashSales) || 0;
        const cashIn = Number(data.cashIn) || 0;
        const cashOut = Number(data.cashOut) || 0;
        const transferSales = Number(data.transferSales) || 0;
        const totalSales = cashSales + transferSales;
        
        const expectedCash = startCash + cashSales + cashIn - cashOut;
        const actualCash = Number(data.actualCash) || 0;
        const diff = actualCash - expectedCash;

        sheet.getRange(rowIndex, 4).setValue(closeTimeStr);
        sheet.getRange(rowIndex, 6).setValue(cashSales);
        sheet.getRange(rowIndex, 7).setValue(cashIn);
        sheet.getRange(rowIndex, 8).setValue(cashOut);
        sheet.getRange(rowIndex, 9).setValue(transferSales);
        sheet.getRange(rowIndex, 10).setValue(totalSales);
        sheet.getRange(rowIndex, 11).setValue(expectedCash);
        sheet.getRange(rowIndex, 12).setValue(actualCash);
        sheet.getRange(rowIndex, 13).setValue(diff);
        sheet.getRange(rowIndex, 14).setValue('Closed');

        return { success: true, message: "ปิดกะเรียบร้อยแล้ว!" };
      }
    }
    return { success: false, message: "ไม่พบข้อมูลกะที่เปิดอยู่" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 9. บันทึกเงินเข้า/ถอนเงิน
 */
function saveCashMovement(customerSheetId, data) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    let sheet = ss.getSheetByName('CashMovements');
    if (!sheet) {
      sheet = ss.insertSheet('CashMovements');
      sheet.appendRow(['Timestamp', 'ShiftID', 'Staff', 'Type', 'Amount', 'Reason', 'Detail']);
    }
    
    sheet.appendRow([
      new Date(),
      data.shiftId || '',
      data.staff || '',
      data.type || '',
      Number(data.amount) || 0,
      data.reason || '',
      data.detail || ''
    ]);

    return { success: true, message: "บันทึกรายการเงินสำเร็จ!" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * 10. ระบบตั้งค่าร้านค้า (Settings)
 */
function saveStoreConfig(customerSheetId, config) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    let sheet = ss.getSheetByName('Settings');
    
    if (!sheet) {
      sheet = ss.insertSheet('Settings');
      sheet.appendRow(['Key', 'Value']);
    }
    
    sheet.getRange("A2:B2").setValues([['storeConfig', JSON.stringify(config)]]);
    return { success: true };
  } catch (error) {
    return { error: error.toString() };
  }
}

function getStoreConfig(customerSheetId) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Settings');
    if (!sheet) return null;
    const data = sheet.getRange("A2:B2").getValues();
    if (data[0][0] === 'storeConfig' && data[0][1]) {
      return JSON.parse(data[0][1]);
    }
    return null;
  } catch (error) {
    return null;
  }
}

/**
 * ฟังก์ชันช่วยแปลงวันที่ (พ.ศ./ค.ศ.)
 */
function parseCustomDate(customerSheetId, dateVal) {
  if (dateVal instanceof Date) return dateVal;
  if (!dateVal) return new Date(0);
  
  const str = String(dateVal).trim();
  const parts = str.split(' ');
  const dateParts = parts[0].split('/');
  
  if (dateParts.length === 3) {
    let day = parseInt(dateParts[0], 10);
    let month = parseInt(dateParts[1], 10) - 1;
    let year = parseInt(dateParts[2], 10);
    if (year > 2500) year -= 543;
    
    let hours = 0, minutes = 0, seconds = 0;
    if (parts[1]) {
      const timeParts = parts[1].split(':');
      hours = parseInt(timeParts[0], 10) || 0;
      minutes = parseInt(timeParts[1], 10) || 0;
      seconds = parseInt(timeParts[2], 10) || 0;
    }
    return new Date(year, month, day, hours, minutes, seconds);
  }
  return new Date(str);
}

// ==========================================
// 📊 ฟังก์ชัน ADMIN DASHBOARD & REPORTING
// ==========================================

// 1. ดึงข้อมูลสถิติภาพรวม, ยอดขาย, ต้นทุน, กำไร
function getAdminDashboardData(customerSheetId, startDateStr, endDateStr) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    const transSheet = ss.getSheetByName('Transactions');
    const prodSheet = ss.getSheetByName('Products');
    
    const transData = transSheet ? transSheet.getDataRange().getValues() : [];
    const prodData = prodSheet ? prodSheet.getDataRange().getValues() : [];
    
    let totalSales = 0;
    let totalCost = 0;
    let totalOrders = 0;
    let cashSales = 0;
    let qrSales = 0;
    let itemSalesMap = {};
    
    const start = startDateStr ? new Date(startDateStr + 'T00:00:00') : new Date(0);
    const end = endDateStr ? new Date(endDateStr + 'T23:59:59') : new Date();

    // วนลูปอ่านรายการขายจากชีต Transactions
    if (transData.length > 1) {
      for (let i = 1; i < transData.length; i++) {
        const row = transData[i];
        const status = String(row[6] || 'COMPLETED');
        if (status === 'VOIDED' || status === 'CANCELLED') continue; // ข้ามบิลที่ยกเลิก

        const rowDate = new Date(row[1]);
        if (rowDate >= start && rowDate <= end) {
          const amount = Number(row[3]) || 0;
          const payType = String(row[4] || '').toLowerCase();
          const itemsJson = row[5];

          totalSales += amount;
          totalOrders += 1;

          if (payType.includes('cash') || payType.includes('เงินสด')) {
            cashSales += amount;
          } else {
            qrSales += amount;
          }

          // ถอดรหัสรายการสินค้าเพื่อคิดทุน และนับจำนวนขายดี
          try {
            const items = typeof itemsJson === 'string' ? JSON.parse(itemsJson) : itemsJson;
            if (Array.isArray(items)) {
              items.forEach(item => {
                const name = item.name || 'ไม่ระบุชื่อ';
                const qty = Number(item.qty) || 1;
                const cost = Number(item.cost) || 0;
                
                totalCost += (cost * qty);
                itemSalesMap[name] = (itemSalesMap[name] || 0) + qty;
              });
            }
          } catch(e) {}
        }
      }
    }

    // จัดอันดับสินค้าขายดี Top 5
    const topSellers = Object.keys(itemSalesMap)
      .map(name => ({ name: name, qty: itemSalesMap[name] }))
      .sort((a, b) => b.qty - a.qty)
      .slice(0, 5);

    // ตรวจสอบสินค้าสต็อกใกล้หมด (น้อยกว่าหรือเท่ากับ 10 ชิ้น)
    let lowStockItems = [];
    if (prodData.length > 1) {
      for (let i = 1; i < prodData.length; i++) {
        const pName = prodData[i][1];
        const pQty = Number(prodData[i][3]) || 0;
        if (pQty <= 10) {
          lowStockItems.push({ name: pName, qty: pQty });
        }
      }
    }

    const netProfit = totalSales - totalCost;
    const profitMargin = totalSales > 0 ? ((netProfit / totalSales) * 100).toFixed(1) : 0;

    return {
      success: true,
      data: {
        totalSales: totalSales,
        totalCost: totalCost,
        netProfit: netProfit,
        profitMargin: profitMargin,
        totalOrders: totalOrders,
        avgPerOrder: totalOrders > 0 ? (totalSales / totalOrders) : 0,
        cashSales: cashSales,
        qrSales: qrSales,
        topSellers: topSellers,
        lowStockItems: lowStockItems
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 2. ดึงประวัติใบเสร็จรับเงิน
function getReceiptHistory(customerSheetId, startDateStr, endDateStr) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    const sheet = ss.getSheetByName('Transactions');
    if (!sheet) return { success: true, receipts: [] };
    
    const data = sheet.getDataRange().getValues();
    
    const start = startDateStr ? new Date(startDateStr + 'T00:00:00') : new Date(0);
    const end = endDateStr ? new Date(endDateStr + 'T23:59:59') : new Date('2099-12-31T23:59:59');
    if (start.getFullYear() > 2400) start.setFullYear(start.getFullYear() - 543);
    if (end.getFullYear() > 2400) end.setFullYear(end.getFullYear() - 543);
    
    let receipts = [];
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const rowDate = parseCustomDate(customerSheetId, row[1]); 
      
      if (rowDate >= start && rowDate <= end) {
        receipts.push({
          rowIndex: i + 1,
          transId: row[0],
          timestamp: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
          staff: row[2] || '-',
          totalAmount: Number(row[3]) || 0,
          discount: Number(row[4]) || 0,
          payType: row[5], // แก้ไข: ดึงวิธีจ่ายเงินจากคอลัมน์ F ให้ถูกต้อง
          status: row[6] || 'COMPLETED',
          // แก้ไข: ดึงข้อมูลตะกร้าสินค้าจากคอลัมน์ H (Index 7) ป้องกันการ Crash
          items: typeof row[7] === 'string' ? JSON.parse(row[7] || '[]') : (row[7] || [])
        });
      }
    }
    return { success: true, receipts: receipts };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// 3. ฟังก์ชันยกเลิกใบเสร็จ (Void)
function voidReceipt(customerSheetId, transId, rowIndex) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    const sheet = ss.getSheetByName('Transactions');
    if (!sheet) return { success: false, message: 'ไม่พบชีต Transactions' };
    sheet.getRange(rowIndex, 7).setValue('VOIDED');
    return { success: true, message: 'ยกเลิกใบเสร็จเรียบร้อยแล้ว' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ฟังก์ชันดึงข้อมูล Dashboard (แก้ไขให้ดึงข้อมูลจาก 3 ชีตอย่างถูกต้อง)
 */
function getDashboardData(customerSheetId, startDateStr, endDateStr) {
  try {
    var ss = SpreadsheetApp.openById(customerSheetId);
    
    // --------------------------------------------------------
    // สเต็ปที่ 1: ไปที่หน้า "Products" เพื่อจำ "ต้นทุน" ของสินค้าแต่ละตัว
    // --------------------------------------------------------
    var productSheet = ss.getSheetByName("Products");
    var costMap = {}; 
    var lowStockItems = []; // เตรียมตะกร้าไว้ใส่สินค้าใกล้หมด
    
    if (productSheet) {
      var pData = productSheet.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        var barcode = String(pData[p][0] || "").trim(); // คอลัมน์ A: Barcode
        var pName = String(pData[p][1] || "");          // คอลัมน์ B: ชื่อสินค้า
        var cost = Number(pData[p][3]) || 0;            // คอลัมน์ D: ต้นทุน (Cost Price)
        var stock = Number(pData[p][5]) || 0;           // คอลัมน์ F: สต็อกปัจจุบัน
        var minStock = Number(pData[p][6]) || 5;        // คอลัมน์ G: จุดเตือนสต็อกหมด
        
        if (barcode !== "") {
          costMap[barcode] = cost; // จำไว้ว่าบาร์โค้ดนี้ ทุนเท่าไหร่
          
          // ถ้าสต็อกเหลือน้อยกว่ากำหนด ให้หยิบใส่ตะกร้าสินค้าใกล้หมด
          if (stock <= minStock) {
            lowStockItems.push({ name: pName, stock: stock, minStock: minStock });
          }
        }
      }
    }

    // --------------------------------------------------------
    // สเต็ปที่ 2: ไปที่หน้า "CashLogs" เพื่อดูว่าแต่ละบิลซื้ออะไรไปบ้าง
    // --------------------------------------------------------
    var logsSheet = ss.getSheetByName("CashLogs");
    var groupedItems = {}; // เตรียมสมุดจดแยกตามเลขบิล
    
    if (logsSheet) {
      var logsData = logsSheet.getDataRange().getValues();
      for (var j = 1; j < logsData.length; j++) {
        var logTxId = String(logsData[j][1] || "").trim();    // คอลัมน์ B: เลขบิล (TxID)
        var logBarcode = String(logsData[j][3] || "").trim(); // คอลัมน์ D: บาร์โค้ด
        var logName = String(logsData[j][4] || "");           // คอลัมน์ E: ชื่อสินค้า
        var logQty = Number(logsData[j][5]) || 0;             // คอลัมน์ F: จำนวนที่ซื้อ
        
        if (logTxId !== "") {
          // ถ้ายังไม่เคยจดเลขบิลนี้ ให้เปิดหน้ากระดาษใหม่
          if (!groupedItems[logTxId]) {
            groupedItems[logTxId] = [];
          }
          // จดลงไปว่าบิลนี้ ซื้ออะไร จำนวนเท่าไหร่ และทุนเท่าไหร่
          groupedItems[logTxId].push({
            name: logName,
            qty: logQty,
            cost: costMap[logBarcode] || 0 // เอาต้นทุนที่จำไว้ในสเต็ป 1 มาใส่
          });
        }
      }
    }

    // --------------------------------------------------------
    // สเต็ปที่ 3: ไปที่หน้า "Transactions" เพื่อดึงข้อมูลบิล และกรองวันที่
    // --------------------------------------------------------
    var sheet = ss.getSheetByName("Transactions"); 
    if (!sheet) {
      return { success: false, message: "หาชีตชื่อ 'Transactions' ไม่เจอครับ" };
    }
    
    var data = sheet.getDataRange().getValues();
    var sales = [];
    
    // ตั้งค่านากฬิกาสำหรับกรองวันที่
    var start = startDateStr ? new Date(startDateStr) : new Date('2000-01-01');
    var end = endDateStr ? new Date(endDateStr) : new Date('2099-12-31');
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var txId = String(row[0] || '').trim();          // คอลัมน์ A: เลขบิล
      var rawDate = row[1];                            // คอลัมน์ B: วันที่
      var totalAmount = Number(row[3]) || 0;           // คอลัมน์ D: ยอดรวม
      var payMethod = String(row[5] || 'เงินสด');       // คอลัมน์ F: วิธีจ่ายเงิน
      var status = String(row[6] || '').toLowerCase(); // คอลัมน์ G: สถานะ (Status)
      
      if (!rawDate) continue;
      
      var rowDate = new Date(rawDate);
      
      // แปลง พ.ศ. เป็น ค.ศ. (กันระบบหาไม่เจอ)
      if (rowDate.getFullYear() > 2400) {
        rowDate.setFullYear(rowDate.getFullYear() - 543);
      }

      // เช็กว่าบิลนี้สมบูรณ์ไหม (Completed)
      var isCompleted = status === '' || status.indexOf('completed') !== -1;

      // ถ้าวันที่อยู่ในช่วงที่เลือก และบิลสมบูรณ์
      if (rowDate >= start && rowDate <= end && isCompleted) {
        
        // เอาของที่จดไว้ในสเต็ปที่ 2 มาแนบใส่บิลนี้
        var itemsArr = groupedItems[txId] || [];
        
        sales.push({
          orderId: txId,
          totalAmount: totalAmount,
          paymentMethod: payMethod,
          itemsJson: JSON.stringify(itemsArr) // แปลงเป็นข้อความแพ็คส่งให้หน้าบ้าน
        });
      }
    }

    // --------------------------------------------------------
    // สเต็ปที่ 4: ส่งข้อมูลทั้งหมดกลับไปให้หน้าจอ Dashboard วาดรูป
    // --------------------------------------------------------
    return { 
      success: true, 
      sales: sales,
      lowStockItems: lowStockItems // ส่งของใกล้หมดไปด้วยเลย
    };

  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * ฟังก์ชันเสริม: เผื่อฝั่งหน้าจอ (หน้าบ้าน) มีการเรียกขอข้อมูลสต๊อกแยกต่างหาก
 */
function getLowStockProducts(customerSheetId) {
  try {
    var ss = SpreadsheetApp.openById(customerSheetId);
    var productSheet = ss.getSheetByName("Products");
    var lowStockItems = [];
    
    if (productSheet) {
      var pData = productSheet.getDataRange().getValues();
      for (var p = 1; p < pData.length; p++) {
        var pName = String(pData[p][1] || "");
        var stock = Number(pData[p][5]) || 0;
        var minStock = Number(pData[p][6]) || 5;
        
        // ถ้าสต๊อกน้อยกว่ากำหนด ให้ส่งชื่อไปเตือน
        if (pName !== "" && stock <= minStock) {
          lowStockItems.push({ name: pName, stock: stock, minStock: minStock });
        }
      }
    }
    return { success: true, products: lowStockItems };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ฟังก์ชันช่วยแปลงวันที่ (วางไว้ล่างสุดของไฟล์ Code.gs)
function parseCustomDate(customerSheetId, dateVal) {
  if (!dateVal) return new Date(0);
  if (dateVal instanceof Date) return dateVal;
  
  const str = String(dateVal).trim();
  
  // รองรับการแปลงวันที่รูปแบบ DD/MM/YYYY หรือ DD-MM-YYYY
  if (str.includes('/') || str.includes('-')) {
    const parts = str.split(/[\/\-]/);
    if (parts.length === 3) {
      let day = parseInt(parts[0], 10);
      let month = parseInt(parts[1], 10) - 1;
      let year = parseInt(parts[2], 10);
      if (year > 2500) year -= 543; // แปลง พ.ศ. ให้เป็น ค.ศ.
      return new Date(year, month, day);
    }
  }
  
  return new Date(str);
}

/**
 * ดึงรายชื่อสินค้าที่สต็อกหมดหรือติดลบ (<= 0) จากชีต "Products"
 */
function getOutOfStockProducts(customerSheetId) {
  try {
    var ss = SpreadsheetApp.openById(customerSheetId);
    var sheet = ss.getSheetByName("Products"); // ชี้ตรงไปที่ชีต Products
    
    if (!sheet) {
      return { success: false, message: "หาชีตชื่อ 'Products' ไม่เจอครับ" };
    }
    
    var data = sheet.getDataRange().getValues();
    var outOfStockList = [];
    
    // เริ่มอ่านตั้งแต่แถวที่ 2 (ข้ามหัวตารางแถวแรก)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var barcode = String(row[0] || '').trim();     // คอลัมน์ A: Barcode
      var productName = String(row[1] || '');        // คอลัมน์ B: Product Name
      var category = String(row[2] || '');           // คอลัมน์ C: Category
      var costPrice = Number(row[3]) || 0;           // คอลัมน์ D: Cost Price
      var sellingPrice = Number(row[4]) || 0;        // คอลัมน์ E: Selling Price
      var currentStock = Number(row[5]);             // คอลัมน์ F: Current Stock
      var minStock = Number(row[6]) || 5;            // คอลัมน์ G: MinStock
      
      if (productName === '') continue; // ข้ามแถวที่ไม่มีชื่อสินค้า
      
      // 🌟 เงื่อนไข: สต็อก <= 0 (จับหมดทั้งเลข 0 และค่าติดลบ เช่น -10)
      if (!isNaN(currentStock) && currentStock <= 0) {
        outOfStockList.push({
          barcode: barcode,
          name: productName,
          category: category,
          costPrice: costPrice,
          sellingPrice: sellingPrice,
          stock: currentStock,
          minStock: minStock
        });
      }
    }
    
    return { success: true, products: outOfStockList };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * 8. ดึงประวัติการขายทั้งหมด (จากหน้า Transactions)
 */
function getSalesHistory(customerSheetId, startDateStr, endDateStr) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('Transactions');
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน Transactions" };
    
    const data = sheet.getDataRange().getDisplayValues(); 
    if (data.length <= 1) return { success: true, data: [] };
    
    const start = startDateStr ? new Date(startDateStr + 'T00:00:00') : new Date(0);
    const end = endDateStr ? new Date(endDateStr + 'T23:59:59') : new Date('2099-12-31T23:59:59');
    if (start.getFullYear() > 2400) start.setFullYear(start.getFullYear() - 543);
    if (end.getFullYear() > 2400) end.setFullYear(end.getFullYear() - 543);
    
    const sales = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      
      const rowDate = parseCustomDate(customerSheetId, data[i][1]);
      
      if (rowDate >= start && rowDate <= end) {
        sales.push({
          txId: String(data[i][0]),
          timestamp: String(data[i][1]),
          staff: String(data[i][2]),
          totalAmount: Number(String(data[i][3]).replace(/,/g, '')) || 0,
          discount: Number(String(data[i][4]).replace(/,/g, '')) || 0,
          paymentMethod: String(data[i][5]),
          status: String(data[i][6]) || 'Completed'
        });
      }
    }
    return { success: true, data: sales.reverse() }; // บิลใหม่อยู่บนสุด
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 9. ดึงรายละเอียดสินค้าภายในบิล (จากหน้า CashLogs)
 */
function getTransactionItems(customerSheetId, txId) {
  try {
    const sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName('CashLogs');
    if (!sheet) return { success: false, message: "ไม่พบแผ่นงาน CashLogs" };
    
    const data = sheet.getDataRange().getDisplayValues();
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(txId)) {
        items.push({
          timestamp: String(data[i][0]),
          barcode: String(data[i][3]),
          productName: String(data[i][4]),
          qty: Number(data[i][5]) || 0,
          price: Number(data[i][6]) || 0,
          total: Number(data[i][7]) || 0,
          paymentMethod: String(data[i][8])
        });
      }
    }
    return { success: true, items: items };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 10. ยกเลิกบิล + บวกสินค้าคืนเข้าสต็อกอัตโนมัติ
 */
function cancelTransaction(customerSheetId, txId) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    const txSheet = ss.getSheetByName('Transactions');
    const logsSheet = ss.getSheetByName('CashLogs');
    const prodSheet = ss.getSheetByName('Products');
    
    if (!txSheet || !logsSheet || !prodSheet) {
      return { success: false, message: "แผ่นงานในระบบไม่ครบถ้วน" };
    }

    // 1. เปลี่ยนสถานะบิลใน Transactions เป็น Cancelled
    const txData = txSheet.getDataRange().getDisplayValues();
    let foundTx = false;
    for (let i = 1; i < txData.length; i++) {
      if (String(txData[i][0]) === String(txId)) {
        if (txData[i][6] === 'Cancelled') {
          return { success: false, message: "บิลนี้ถูกยกเลิกไปแล้ว" };
        }
        txSheet.getRange(i + 1, 7).setValue('Cancelled');
        foundTx = true;
        break;
      }
    }
    if (!foundTx) return { success: false, message: "ไม่พบบิลนี้ในระบบ" };

    // 2. รวบรวมรายการสินค้าที่ขายในบิลนี้
    const logsData = logsSheet.getDataRange().getDisplayValues();
    const itemsToReturn = [];
    for (let i = 1; i < logsData.length; i++) {
      if (String(logsData[i][1]) === String(txId)) {
        itemsToReturn.push({
          barcode: String(logsData[i][3]),
          qty: Number(logsData[i][5]) || 0
        });
      }
    }

    // 3. ปรับสต็อกในแผ่นงาน Products คืนให้อัตโนมัติ
    const prodData = prodSheet.getDataRange().getDisplayValues();
    for (let item of itemsToReturn) {
      for (let p = 1; p < prodData.length; p++) {
        if (String(prodData[p][0]) === String(item.barcode)) {
          let currentStock = Number(prodData[p][5]) || 0;
          let newStock = currentStock + item.qty;
          prodSheet.getRange(p + 1, 6).setValue(newStock);
          prodData[p][5] = newStock;
          break;
        }
      }
    }

    return { success: true, message: "ยกเลิกบิลและคืนสินค้าเข้าสต็อกเรียบร้อยแล้ว" };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * 11. ดึงข้อมูลประวัติการปิดกะ และ รายการเงินเข้า/ออก
 */
function getShiftHistory(customerSheetId, startDateStr, endDateStr) {
  try {
    const ss = SpreadsheetApp.openById(customerSheetId);
    const shiftSheet = ss.getSheetByName('Shifts');
    const moveSheet = ss.getSheetByName('CashMovements');
    
    // จัดการวันที่สำหรับกรอง และแปลงปี พ.ศ. เป็น ค.ศ.
    const start = startDateStr ? new Date(startDateStr + 'T00:00:00') : new Date(0);
    const end = endDateStr ? new Date(endDateStr + 'T23:59:59') : new Date();
    if (start.getFullYear() > 2400) start.setFullYear(start.getFullYear() - 543);
    if (end.getFullYear() > 2400) end.setFullYear(end.getFullYear() - 543);

    const shifts = [];
    if (shiftSheet) {
      const sData = shiftSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < sData.length; i++) {
        if (!sData[i][0]) continue;
        
        // เช็กคอลัมน์ OpenTime (Index 2)
        const openTimeDate = parseCustomDate(customerSheetId, sData[i][2]);
        
        if (openTimeDate >= start && openTimeDate <= end) {
          shifts.push({
            shiftId: String(sData[i][0]),
            staff: String(sData[i][1]),
            openTime: String(sData[i][2]),
            closeTime: String(sData[i][3]),
            startCash: Number(String(sData[i][4]).replace(/,/g, '')) || 0,
            cashSales: Number(String(sData[i][5]).replace(/,/g, '')) || 0,
            cashIn: Number(String(sData[i][6]).replace(/,/g, '')) || 0,
            cashOut: Number(String(sData[i][7]).replace(/,/g, '')) || 0,
            transferSales: Number(String(sData[i][8]).replace(/,/g, '')) || 0,
            totalSales: Number(String(sData[i][9]).replace(/,/g, '')) || 0,
            expectedCash: Number(String(sData[i][10]).replace(/,/g, '')) || 0,
            actualCash: Number(String(sData[i][11]).replace(/,/g, '')) || 0,
            difference: Number(String(sData[i][12]).replace(/,/g, '')) || 0,
            status: String(sData[i][13])
          });
        }
      }
    }

    const movements = [];
    if (moveSheet) {
      const mData = moveSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < mData.length; i++) {
        if (!mData[i][0]) continue;
        
        // เช็กคอลัมน์ Timestamp (Index 0)
        const moveTimeDate = parseCustomDate(customerSheetId, mData[i][0]);

        if (moveTimeDate >= start && moveTimeDate <= end) {
          movements.push({
            timestamp: String(mData[i][0]),
            shiftId: String(mData[i][1]),
            staff: String(mData[i][2]),
            type: String(mData[i][3]),
            amount: Number(String(mData[i][4]).replace(/,/g, '')) || 0,
            reason: String(mData[i][5]),
            detail: String(mData[i][6])
          });
        }
      }
    }

    return { success: true, shifts: shifts.reverse(), movements: movements.reverse() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ==========================================
// 1. ระบบจัดการข้อมูลพนักงาน (ดึงจากแผ่นงาน "staffs")
// ==========================================

// ดึงรายชื่อพนักงานทั้งหมด
function getStaffList(customerSheetId) {
  var sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName("staffs");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var staff = [];
  
  // ข้ามแถวที่ 1 (หัวตาราง)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) { // เช็คว่ามีชื่อผู้ใช้
      staff.push({
        username: data[i][0],
        pin: data[i][1],
        role: data[i][2]
      });
    }
  }
  return staff;
}

// บันทึกพนักงานใหม่ลง Sheet
function saveNewStaff(customerSheetId, username, pin, role) {
  var sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName("staffs");
  if (!sheet) {
    sheet = SpreadsheetApp.openById(customerSheetId).insertSheet("staffs");
    sheet.appendRow(["Username", "PIN", "Role"]);
  }
  sheet.appendRow([username, pin, role]);
  return "Success";
}

// ลบพนักงานออกจาก Sheet
function removeStaff(customerSheetId, username) {
  var sheet = SpreadsheetApp.openById(customerSheetId).getSheetByName("staffs");
  if (!sheet) return "Error";
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      sheet.deleteRow(i + 1); // ลบแถวที่ตรงกัน
      break;
    }
  }
  return "Success";
}

// ==========================================
// 2. ระบบจัดการข้อมูลการชำระเงิน (Payment Settings)
// ==========================================

/**
 * ดึงข้อมูลการตั้งค่าการชำระเงินจากหน้า Settings (แถวที่ Key = paymentSettings)
 */
function getPaymentSettings(customerSheetId) {
  try {
    var ss = SpreadsheetApp.openById(customerSheetId);
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) {
      return { success: false, message: "ไม่พบแผ่นงาน Settings" };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "paymentSettings") {
        var settingsJson = data[i][1];
        return {
          success: true,
          data: typeof settingsJson === 'string' ? JSON.parse(settingsJson) : settingsJson
        };
      }
    }
    return { success: false, message: "ไม่พบข้อมูล paymentSettings" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * บันทึกหรืออัปเดตข้อมูลการตั้งค่าการชำระเงินลงหน้า Settings
 */
function savePaymentSettings(customerSheetId, newSettings) {
  try {
    var ss = SpreadsheetApp.openById(customerSheetId);
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) {
      return { success: false, message: "ไม่พบแผ่นงาน Settings" };
    }
    
    var data = sheet.getDataRange().getValues();
    var foundRow = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "paymentSettings") {
        foundRow = i + 1;
        break;
      }
    }
    
    var jsonString = JSON.stringify(newSettings);
    
    if (foundRow !== -1) {
      sheet.getRange(foundRow, 2).setValue(jsonString);
    } else {
      sheet.appendRow(["paymentSettings", jsonString]);
    }
    
    return { success: true, message: "บันทึกข้อมูลเรียบร้อยแล้ว" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// ส่วนที่เพิ่ม: ระบบจัดการข้อมูลการชำระเงิน (Payment Settings)
// ==========================================

/**
 * ดึงข้อมูลการตั้งค่าการชำระเงินจากหน้า Settings (แถวที่ Key = paymentSettings)
 */
function getPaymentSettings(customerSheetId) {
  try {
    var ss = SpreadsheetApp.openById(customerSheetId);
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) {
      return { success: false, message: "ไม่พบแผ่นงาน Settings" };
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "paymentSettings") {
        var settingsJson = data[i][1];
        return {
          success: true,
          data: typeof settingsJson === 'string' ? JSON.parse(settingsJson) : settingsJson
        };
      }
    }
    return { success: false, message: "ไม่พบข้อมูล paymentSettings" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * บันทึกหรืออัปเดตข้อมูลการตั้งค่าการชำระเงินลงหน้า Settings
 */
function savePaymentSettings(customerSheetId, newSettings) {
  try {
    var ss = SpreadsheetApp.openById(customerSheetId);
    var sheet = ss.getSheetByName("Settings");
    if (!sheet) {
      return { success: false, message: "ไม่พบแผ่นงาน Settings" };
    }
    
    var data = sheet.getDataRange().getValues();
    var foundRow = -1;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "paymentSettings") {
        foundRow = i + 1;
        break;
      }
    }
    
    var jsonString = JSON.stringify(newSettings);
    
    if (foundRow !== -1) {
      sheet.getRange(foundRow, 2).setValue(jsonString);
    } else {
      sheet.appendRow(["paymentSettings", jsonString]);
    }
    
    return { success: true, message: "บันทึกข้อมูลเรียบร้อยแล้ว" };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
