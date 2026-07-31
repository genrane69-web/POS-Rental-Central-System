// ==========================================================
// 📦 ระบบต่ออายสมาชิกอัตโนมัติผ่าน Omise (Opn Payments)
// ไฟล์นี้แยกออกจาก Code.gs โดยตั้งใจ — เพอไม่ให้เสี่ยงกระทบโค้ดเดิม
// วิธีติดตั้ง: ใน Apps Script กด "+" ข้างไฟล์ > Script > ตั้งชื่อ "Billing"
// แล้ววางโค้ดทั้งหมดนี้ลงไป (ไฟล์ .gs ทุกไฟล์ในโปรเจกต์เดียวกัน
// ใช้ตัวแปร/ฟังก์ชันร่วมกันได้อัตโนมัติ ไม่ต้อง import)
// ==========================================================

// ⚠️ ยังไม่ต้องใสตอนนี้ — รอสมัคร Omise แล้วมี Public/Secret Key ค่อยแทนที่ตรงนี้
// ตอนทดสอบใช้ key ที่ขึ้นต้นด้วย pkey_test_ / skey_test_ ก่อน
// พอใช้งานจริงค่อยเปลี่ยนเป็น pkey_live_ / skey_live_
const OMISE_PUBLIC_KEY = "pkey_test_68irdufiu38kxwnn298";
const OMISE_SECRET_KEY = "skey_test_68irdug165wnf64t396";// 🔑 ใส่ Secret Key ตรงนี้

const RENEWAL_AMOUNT_BAHT = 199; // ราคาต่ออายต่อครั้ง (บาท) — ปรับตามที่ต้องการ
const RENEWAL_DAYS = 30;         // จำนวนวันที่ต่ออายุให้ต่อการชำระเงิน 1 ครั้ง

// ==========================================================
// 1. สร้างรายการชำระเงิน + QR พร้อมเพย์ (เรียกจากหน้าตออายุฝั่งลูกค้า)
// ==========================================================
function createRenewalCharge() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return { success: false, message: "ไม่พบอีเมลผู้ใช้งาน กรุณาลงชือเข้าใช้ Google ก่อน" };

    if (OMISE_SECRET_KEY.indexOf("XXXX") !== -1) {
      return { success: false, message: "ระบบชำระเงินยังไม่พร้อมใช้งาน (ยังไม่ได้ตั้งค่า API Key)" };
    }

    const amountSatang = Math.round(RENEWAL_AMOUNT_BAHT * 100);

    // 1) สร้าง Source แบบพร้อมเพย์ (ใช้ Public Key — ปลอดภัยฝั่ง client ได้)
    const sourceRes = omiseRequest_(
      "https://vault.omise.co/sources",
      OMISE_PUBLIC_KEY,
      { amount: amountSatang, currency: "thb", type: "promptpay" }
    );
    if (!sourceRes || !sourceRes.id) {
      return { success: false, message: "สร้าง QR พร้อมเพย์ไม่สำเร็จ: " + (sourceRes && sourceRes.message ? sourceRes.message : "unknown error") };
    }

    // 2) สร้าง Charge ผูกกับ Source (ต้องใช้ Secret Key ฝั่ง server เท่านั้น)
    const chargeRes = omiseRequest_(
      "https://api.omise.co/charges",
      OMISE_SECRET_KEY,
      {
        amount: amountSatang,
        currency: "thb",
        source: sourceRes.id,
        "metadata[email]": email,
        "metadata[days]": RENEWAL_DAYS,
        "metadata[purpose]": "vega_pos_renewal"
      }
    );
    if (!chargeRes || !chargeRes.id) {
      return { success: false, message: "สร้างรายการชำระเงินไม่สำเร็จ: " + (chargeRes && chargeRes.message ? chargeRes.message : "unknown error") };
    }

    // 3) บันทึกลงชีต Payments เป็นสถานะ "Pending" รอ Webhook ยืนยันการจายจริง
    logPayment_(chargeRes.id, email, RENEWAL_AMOUNT_BAHT, "Pending");

    const qrImageUrl = (chargeRes.source && chargeRes.source.scannable_code && chargeRes.source.scannable_code.image)
      ? chargeRes.source.scannable_code.image.download_uri
      : "";

    return {
      success: true,
      chargeId: chargeRes.id,
      amount: RENEWAL_AMOUNT_BAHT,
      qrImageUrl: qrImageUrl,
      status: chargeRes.status
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================================
// 2. ให้หน้าเว็บถามสถานะเป็นระยะ (Poll) ว่าจ่ายสำเร็จหรือยัง
//    ใช้คูกับ createRenewalCharge — เรียกทุก 3-5 วิ จนกว่า paid = true
// ==========================================================
function getChargeStatus(chargeId) {
  try {
    const res = omiseRequest_("https://api.omise.co/charges/" + chargeId, OMISE_SECRET_KEY, null, "get");
    return { success: true, status: res.status, paid: !!res.paid };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================================
// 3. Webhook รับแจ้งจาก Omise เมื่อมีการจ่ายเงินสำเร็จ
//    ตั้งค่า URL นี้ใน Omise Dashboard > Webhooks
//    ใช้ URL เดียวกับที่ deploy เป็น Web App (ลงทายด้วย /exec)
// ==========================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // Omise ส่ง event มารูปแบบ { key: "charge.complete", data: { id: "chrg_xxx", ... } }
    if (body.key !== "charge.complete" || !body.data || !body.data.id) {
      return ContentService.createTextOutput("ignored");
    }

    const chargeId = body.data.id;

    // ⚠️ ไม่เชื่อสถานะจาก webhook ตรงๆ — ยิงกลับไปถาม Omise เองอีกครั้งเพื่อความชัวร
    // (ป้องกันคนปลอมแปลงคำขอ webhook มาหลอกระบบ)
    const verifyRes = omiseRequest_("https://api.omise.co/charges/" + chargeId, OMISE_SECRET_KEY, null, "get");

    if (verifyRes && verifyRes.status === "successful" && verifyRes.paid) {
      const email = verifyRes.metadata && verifyRes.metadata.email;
      const days = (verifyRes.metadata && Number(verifyRes.metadata.days)) || RENEWAL_DAYS;

      if (email) {
        extendTenant_(email, days);
        updatePaymentStatus_(chargeId, "Confirmed");
      }
    }

    return ContentService.createTextOutput("ok");
  } catch (error) {
    console.error("Webhook error: " + error.toString());
    return ContentService.createTextOutput("error");
  }
}

// ==========================================================
// 4. ฟังก์ชันช่วยเหลือภายใน (ขึ้นต้นด้วย _ = ไม่ให้เรียกจากหน้าเว็บโดยตรง)
// ==========================================================

// เรียก Omise API แบบทั่วไป (รองรับทง GET และ POST)
function omiseRequest_(url, key, payloadObj, method) {
  const options = {
    method: method || "post",
    headers: { "Authorization": "Basic " + Utilities.base64Encode(key + ":") },
    muteHttpExceptions: true
  };
  if (payloadObj) options.payload = payloadObj;
  const response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText());
}

// ต่ออายุผู้เช่าในชีตแอดมิน — ถ้ายังไม่หมดอายุ จะ "ต่อจากวนเดิม" ไม่ใช่นับใหม่จากวันนี้
function extendTenant_(email, days) {
  const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  const sheet = ss.getSheetByName('ชีต1');
  const data = sheet.getDataRange().getValues();
  const formatDate = (d) => Utilities.formatDate(d, "Asia/Bangkok", "d/M/yyyy, H:mm:ss");

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      const currentExpireRaw = data[i][3];
      let baseDate = new Date();
      if (currentExpireRaw) {
        const currentExpire = parseCustomDate(null, currentExpireRaw);
        if (currentExpire > baseDate) baseDate = currentExpire; // ยังไม่หมดอายุ ต่อจากวันเดิม
      }
      const newExpire = new Date(baseDate);
      newExpire.setDate(newExpire.getDate() + Number(days));

      sheet.getRange(i + 1, 4).setValue(formatDate(newExpire)); // คอลัมน์ D: ExpireDate
      sheet.getRange(i + 1, 5).setValue("Active");              // คอลัมน์ E: Status
      return true;
    }
  }
  return false; // ไม่เจออีเมลนี้ในชีตแอดมิน
}

// ชีต Payments ไว้เก็บประวัติการจายเงินทุกครั้ง (สร้างอัตโนมัติถ้ายงไม่มี)
function getPaymentsSheet_() {
  const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  let sheet = ss.getSheetByName('Payments');
  if (!sheet) {
    sheet = ss.insertSheet('Payments');
    sheet.appendRow(['ChargeID', 'Email', 'AmountBaht', 'Status', 'CreatedAt', 'ConfirmedAt']);
  }
  return sheet;
}

function logPayment_(chargeId, email, amountBaht, status) {
  const sheet = getPaymentsSheet_();
  sheet.appendRow([chargeId, email, amountBaht, status, new Date(), '']);
}

function updatePaymentStatus_(chargeId, newStatus) {
  const sheet = getPaymentsSheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(chargeId)) {
      sheet.getRange(i + 1, 4).setValue(newStatus);
      sheet.getRange(i + 1, 6).setValue(new Date());
      return;
    }
  }
}
