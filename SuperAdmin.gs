// ==========================================================
// 👑 ระบบ Super Admin — สำหรับเจ้าของระบบ (นาย) เท่านั้น
// ==========================================================

// ⚠️ สำคัญมาก: ต้องแก้เป็น Gmail ของนายเอง ไม่งั้นนายเองจะเข้าหน้านี้ไม่ได้
const SUPER_ADMIN_EMAIL = "genrane69@gmail.com";

function requireSuperAdmin_() {
  const email = Session.getActiveUser().getEmail();
  if (!email || email.trim().toLowerCase() !== SUPER_ADMIN_EMAIL.trim().toLowerCase()) {
    throw new Error("ไม่มีสทธิ์เข้าถึงหน้านี้");
  }
}

function getAllTenants() {
  try {
    requireSuperAdmin_();
    const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
    const sheet = ss.getSheetByName('ชีต1');
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const tenants = [];

    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const expireRaw = data[i][3];
      const expireDate = expireRaw ? parseCustomDate(null, expireRaw) : null;
      const daysLeft = expireDate ? Math.ceil((expireDate - now) / 86400000) : null;

      tenants.push({
        email: data[i][0],
        sheetId: data[i][1],
        installDate: data[i][2] ? String(data[i][2]) : '',
        expireDate: expireRaw ? String(expireRaw) : '',
        status: data[i][4] || '',
        daysLeft: daysLeft,
        isExpired: daysLeft !== null && daysLeft < 0
      });
    }

    tenants.sort((a, b) => (a.daysLeft === null ? 999 : a.daysLeft) - (b.daysLeft === null ? 999 : b.daysLeft));
    return { success: true, tenants: tenants };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function manualExtendTenant(email, days) {
  try {
    requireSuperAdmin_();
    const ok = extendTenant_(email, Number(days));
    if (ok) {
      logPayment_('MANUAL-' + Date.now(), email, 0, 'Manual by Admin (+' + days + ' days)');
      return { success: true, message: 'ตออายุให้ ' + email + ' อก ' + days + ' วันเรียบร้อยแล้ว' };
    }
    return { success: false, message: 'ไม่พบอีเมลนี้ในระบบ' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function suspendTenant(email) {
  try {
    requireSuperAdmin_();
    const ok = setTenantStatus_(email, 'Suspended');
    return ok
      ? { success: true, message: 'ระงับการใช้งานของ ' + email + ' แล้ว' }
      : { success: false, message: 'ไม่พบอีเมลนี้ในระบบ' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function activateTenant(email) {
  try {
    requireSuperAdmin_();
    const ok = setTenantStatus_(email, 'Active');
    return ok
      ? { success: true, message: 'เปิดใช้งาน ' + email + ' อีกครั้งแล้ว' }
      : { success: false, message: 'ไม่พบอีเมลนี้ในระบบ' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function setTenantStatus_(email, status) {
  const ss = SpreadsheetApp.openById(ADMIN_SHEET_ID);
  const sheet = ss.getSheetByName('ชีต1');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      sheet.getRange(i + 1, 5).setValue(status);
      return true;
    }
  }
  return false;
}
