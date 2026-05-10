// ============================================================
// عالم العروض — Client.gs (FIXED VERSION)
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1DpbiARHR46jbawxMcC9Dz9VRdMnS6A5NSppcKhblPaE',
  ADMIN_EMAIL: 'mahmoudsap74@gmail.com',
  DRIVE_FOLDER_ID: '1CJPhZKqn6YQGQNn-wsWAi6clfeSzPuEe',
  SYSTEM_NAME: 'عالم العروض',
  ADMIN_URL: 'https://offersworld.github.io/offers-world/admin.html',
  SHEETS: {
    PACKAGES: '📦 Packages',
    ORDERS: '📋 Orders',
    LOGS: '📝 Logs'
  }
};

// ✅ قراءة كلمة المرور من إعدادات السكربت (PropertiesService) للحماية
// يجب إضافتها من خلال: Project Settings -> Script Properties -> ADMIN_PASSWORD
const ADMIN_PASSWORD = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '01225949346';

// ============================================================
// doGet — مع CORS
// ============================================================
function doGet(e) {
  var action = e.parameter.action || '';
  var page   = e.parameter.page   || '';

  if (page === 'admin') {
    return HtmlService
      .createHtmlOutputFromFile('admin')
      .setTitle('لوحة تحكم عالم العروض')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (action === 'checkPassword') {
    var password = e.parameter.pass;
    return respond({ success: password === ADMIN_PASSWORD });
  }

  if (action === 'packages') {
    return respond(getPackages(e.parameter.company));
  }

  if (action === 'orders') {
    return respond(getAllOrders());
  }

  if (action === 'stats') {
    return respond(getAdminStats());
  }

  if (action === 'activity') {
    return respond(getActivityLog());
  }

  return respond({ status: 'ok', system: CONFIG.SYSTEM_NAME });
}

// ============================================================
// doPost — مع جميع الإجراءات
// ============================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // 🔒 منع التزامن (Race condition) عن طريق قفل السكربت لمدة تصل لـ 10 ثوانٍ
    lock.waitLock(10000); 
    
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    if (action === 'checkPassword') {
      return respond({ success: data.pass === ADMIN_PASSWORD });
    }

    if (!action) {
      return respond(handleNewOrder(data));
    }

    if (action === 'updateOrderStatus') {
      return respond(updateOrderStatus(data.orderId, data.status, data.note));
    }

    if (action === 'addPackage') {
      return respond(addPackage(data));
    }

    if (action === 'updatePackage') {
      return respond(updatePackage(data));
    }

    if (action === 'deletePackage') {
      return respond(deletePackage(data.packageId));
    }

    return respond({ success: false, message: 'إجراء غير معروف' });

  } catch (err) {
    logError('doPost', err.toString());
    return respond({ success: false, error: err.toString() });
  } finally {
    // 🔓 تحرير القفل ليسمح للطلبات الأخرى بالمرور
    lock.releaseLock();
  }
}

// ============================================================
// جلب الباقات
// ============================================================
function getPackages(company) {
  company = company || null;
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PACKAGES);
    var data = sheet.getDataRange().getValues();
    var packages = [];

    for (var i = 2; i < data.length; i++) {
      var row = data[i];
      var pkgId      = row[0];
      var pkgCompany = String(row[1] || '').trim();
      var pkgName    = String(row[2] || '').trim();
      var pkgPrice   = Number(row[3]) || 0;
      var pkgType    = String(row[4] || '').trim();
      var status     = String(row[5] || '').trim();

      if (!pkgName || status !== 'متاح') continue;

      if (!company || pkgCompany.includes(company)) {
        packages.push({
          id:      pkgId,
          company: pkgCompany,
          name:    pkgName,
          price:   pkgPrice,
          type:    pkgType,
          status:  status
        });
      }
    }

    packages.sort(function(a, b) { return a.price - b.price; });
    return packages;

  } catch (err) {
    logError('getPackages', err.toString());
    return [];
  }
}

// ============================================================
// ✅ FIX: معالجة الطلب الجديد — أُضيف activationPhone
// ============================================================
function handleNewOrder(data) {
  var orderId = data.orderId || generateOrderId();

  // 🛡️ منع التكرار: لو الـ orderId موجود مسبقًا → ارجع بنجاح بدون حفظ تاني
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ORDERS);
    var existing = sheet.getRange('A:A').createTextFinder(String(orderId).trim()).matchEntireCell(true).findNext();
    if (existing) {
      return { success: true, orderId: orderId, duplicate: true };
    }
  } catch (e) { /* استمر في حالة خطأ في الفحص */ }

  var proofUrl = '';
  if (data.proofBase64) {
    proofUrl = saveBase64File(data.proofBase64, orderId);
  }

  var row = [
    orderId,                        // A[0]:  رقم الطلب
    formatDate(new Date()),         // B[1]:  التاريخ
    data.customerName  || '',       // C[2]:  اسم العميل
    data.phone         || '',       // D[3]:  رقم الواتساب
    data.company       || '',       // E[4]:  الشركة
    data.package       || '',       // F[5]:  الباقة
    data.price         || 0,        // G[6]:  السعر
    data.transferRef   || '',       // H[7]:  رقم التحويل
    data.payment       || '',       // I[8]:  طريقة الدفع
    'معلق',                         // J[9]:  الحالة
    '',                             // K[10]: تاريخ التفعيل
    data.notes         || '',       // L[11]: ملاحظات
    proofUrl,                       // M[12]: إيصال الدفع
    data.activationPhone || '',     // N[13]: ✅ رقم التفعيل
    data.vodafonePassword || ''     // O[14]: 🔑 باسورد أنا فودافون (فودافون فقط)
  ];

  saveOrder(row);
  addLog('طلب جديد', orderId, (data.company || '') + ' - ' + (data.package || ''));
  sendAdminNotification(orderId, data, proofUrl);

  return { success: true, orderId: orderId };
}

// ============================================================
// دوال إدارة الطلبات
// ============================================================
function getAllOrders() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ORDERS);
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 2; i < data.length; i++) {
    result.push({
      orderId:         data[i][0],
      date:            data[i][1],
      customerName:    data[i][2],
      phone:           data[i][3],
      company:         data[i][4],
      package:         data[i][5],
      price:           data[i][6],
      transferRef:     data[i][7],
      payment:         data[i][8],
      status:          data[i][9],
      activationDate:  data[i][10] || '',
      notes:           data[i][11],
      proofImage:      data[i][12],
      activationPhone: data[i][13] || ''
    });
  }
  return result;
}

function updateOrderStatus(orderId, newStatus, note) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.ORDERS);
    
    // ⚡ تسريع البحث باستخدام TextFinder بدلاً من الـ for loop
    var findResult = sheet.getRange("A:A").createTextFinder(String(orderId).trim()).matchEntireCell(true).findNext();
    if (!findResult) return { success: false, message: 'الطلب غير موجود' };
    
    var rowIndex = findResult.getRow();

    sheet.getRange(rowIndex, 10).setValue(newStatus);   // col J = status

    if (newStatus === 'تم التفعيل') {
      var currentActivation = sheet.getRange(rowIndex, 11).getValue();
      if (!currentActivation) {
        sheet.getRange(rowIndex, 11).setValue(formatDate(new Date()));
      }
    }

    if (note) {
      var oldNote = sheet.getRange(rowIndex, 12).getValue() || '';
      var newNote = oldNote + (oldNote ? '\n' : '') + '[' + formatDate(new Date()) + '] ' + note;
      sheet.getRange(rowIndex, 12).setValue(newNote);
    }

    addLog('تغيير حالة', orderId, 'إلى ' + newStatus);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// دوال إدارة الباقات
// ============================================================
function addPackage(pkg) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PACKAGES);
    var newId = generatePackageId();
    sheet.appendRow([
      newId,
      pkg.company || '',
      pkg.name,
      pkg.price,
      pkg.type || '',
      pkg.status || 'متاح',
      pkg.notes || '',
      formatDate(new Date())
    ]);
    addLog('إضافة باقة', newId, pkg.name);
    return { success: true, id: newId };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function updatePackage(pkg) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PACKAGES);
    
    // ⚡ تسريع البحث باستخدام TextFinder
    var findResult = sheet.getRange("A:A").createTextFinder(String(pkg.packageId).trim()).matchEntireCell(true).findNext();
    if (!findResult) return { success: false, message: 'الباقة غير موجودة' };
    
    var rowIndex = findResult.getRow();

    sheet.getRange(rowIndex, 2).setValue(pkg.company || '');
    sheet.getRange(rowIndex, 3).setValue(pkg.name);
    sheet.getRange(rowIndex, 4).setValue(pkg.price);
    sheet.getRange(rowIndex, 5).setValue(pkg.type || '');
    sheet.getRange(rowIndex, 6).setValue(pkg.status || 'متاح');
    sheet.getRange(rowIndex, 7).setValue(pkg.notes || '');
    addLog('تحديث باقة', pkg.packageId, pkg.name);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deletePackage(packageId) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PACKAGES);
    
    // ⚡ تسريع البحث باستخدام TextFinder
    var findResult = sheet.getRange("A:A").createTextFinder(String(packageId).trim()).matchEntireCell(true).findNext();
    if (!findResult) return { success: false, message: 'الباقة غير موجودة' };
    
    var rowIndex = findResult.getRow();
    sheet.deleteRow(rowIndex);
    addLog('حذف باقة', packageId, '');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function generatePackageId() {
  return 'PKG-' + Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyMMddHHmmss') + '-' + Math.floor(Math.random() * 1000);
}

// ============================================================
// سجل النشاطات
// ============================================================
function getActivityLog() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LOGS);
    var data = sheet.getDataRange().getValues();
    var logs = [];
    for (var i = data.length - 1; i >= 1; i--) {
      logs.push({
        timestamp: data[i][0],
        action:    data[i][1],
        orderId:   data[i][2],
        details:   data[i][3]
      });
    }
    return logs;
  } catch (e) {
    return [];
  }
}

function addLog(action, orderId, details) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.LOGS);
    sheet.appendRow([formatDate(new Date()), action, orderId, details]);
  } catch (e) {}
}

// ============================================================
// إحصائيات
// ============================================================
function getAdminStats() {
  var orders = getAllOrders();
  var total = orders.length;
  var pending = 0, active = 0, cancelled = 0, revenue = 0;

  orders.forEach(function(order) {
    var st = order.status || '';
    if (st.includes('معلق'))   pending++;
    else if (st.includes('تم')) active++;
    else if (st.includes('ملغ')) cancelled++;
    if (st.includes('تم')) revenue += Number(order.price) || 0;
  });

  return { total: total, pending: pending, active: active, cancelled: cancelled, revenue: revenue };
}

// ============================================================
// دوال مساعدة
// ============================================================
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveBase64File(base64Data, orderId) {
  try {
    var matches = base64Data.match(/^data:(.+);base64,(.*)$/);
    if (!matches) return '';
    var mimeType  = matches[1];
    var bytes     = Utilities.base64Decode(matches[2]);
    var extension = mimeType.includes('png') ? 'png' : 'jpg';
    var blob   = Utilities.newBlob(bytes, mimeType, orderId + '.' + extension);
    var folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    var file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    logError('saveBase64File', err.toString());
    return '';
  }
}

function saveOrder(row) {
  var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.ORDERS);
  sheet.appendRow(row);
}

function generateOrderId() {
  var now  = new Date();
  var date = Utilities.formatDate(now, 'Africa/Cairo', 'MMddHHmm');
  var rand = Math.floor(1000 + Math.random() * 9000);
  return 'ORD-' + date + '-' + rand;
}

function formatDate(d) {
  return Utilities.formatDate(d, 'Africa/Cairo', 'dd/MM/yyyy HH:mm:ss');
}

function logError(fn, msg) {
  console.error(fn + ' => ' + msg);
}

// ============================================================
// ✅ FIX: إرسال الإيميل — إصلاح الـ HTML المكسور (<td><td>)
// ============================================================
function sendAdminNotification(orderId, data, proofImageUrl) {
  var subject = '[عالم العروض] طلب جديد - ' + orderId;

  var plainText = [
    'طلب جديد في عالم العروض',
    '================================',
    'رقم الطلب      : ' + orderId,
    'الاسم           : ' + (data.customerName   || '-'),
    'الواتساب        : ' + (data.phone           || '-'),
    'رقم التفعيل    : ' + (data.activationPhone  || '-'),
    'الشركة          : ' + (data.company          || '-'),
    'الباقة          : ' + (data.package          || '-'),
    'السعر           : ' + (data.price            || 0)  + ' جنيه',
    'طريقة الدفع    : ' + (data.payment           || '-'),
    'رقم التحويل    : ' + (data.transferRef        || '-'),
    'ملاحظات        : ' + (data.notes             || 'لا يوجد'),
    'اثبات الدفع    : ' + (proofImageUrl           || 'لا يوجد'),
    '================================'
  ].join('\n');

  var now = Utilities.formatDate(new Date(), 'Africa/Cairo', 'dd/MM/yyyy - hh:mm a');

  var proofSection = proofImageUrl
    ? '<tr><td style="padding:12px 20px;border-bottom:1px solid #eee;width:38%;color:#888;font-size:13px;">اثبات الدفع</td><td style="padding:12px 20px;border-bottom:1px solid #eee;"><a href="' + proofImageUrl + '" style="display:inline-block;background:#7B2FBE;color:white;padding:8px 22px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:bold;">عرض الايصال</a></td></tr>'
    : '<tr><td style="padding:12px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">اثبات الدفع</td><td style="padding:12px 20px;border-bottom:1px solid #eee;color:#aaa;">لا يوجد</td></tr>';

  var htmlBody = '<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"></head><body style="margin:0;padding:0;background:#F4F6F9;font-family:Arial,sans-serif;direction:rtl;">'
    + '<table width="100%" cellpadding="0" cellspacing="0" style="background:#F4F6F9;padding:30px 0;"><tr><td align="center">'
    + '<table width="600" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,0.08);max-width:600px;">'
    + '<tr><td style="background:linear-gradient(135deg,#7B2FBE,#FF6B00);padding:32px 20px;text-align:center;">'
    + '<h1 style="color:white;margin:0;font-size:24px;font-weight:bold;">عالم العروض</h1>'
    + '<p style="color:rgba(255,255,255,0.8);margin:6px 0 0;font-size:14px;">طلب جديد بانتظار المراجعة</p></td></tr>'
    + '<tr><td style="padding:20px;text-align:center;background:#FAFAFA;border-bottom:2px solid #eee;">'
    + '<span style="display:inline-block;background:#7B2FBE;color:white;padding:9px 28px;border-radius:30px;font-size:14px;font-weight:bold;letter-spacing:1px;">' + orderId + '</span>'
    + '<p style="margin:8px 0 0;color:#999;font-size:12px;">' + now + '</p></td></tr>'
    + '<tr><td style="padding:8px 0;"><table width="100%" cellpadding="0" cellspacing="0">'
    + '<tr style="background:#F9F5FF;"><td style="padding:13px 20px;border-bottom:1px solid #eee;width:38%;color:#888;font-size:13px;">الاسم</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><strong style="color:#1a1a2e;font-size:15px;">' + (data.customerName || '-') + '</strong></td></tr>'
    + '<tr><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">رقم الواتساب</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><a href="https://wa.me/2' + (data.phone || '') + '" style="color:#25D366;font-weight:bold;text-decoration:none;font-size:15px;">' + (data.phone || '-') + '</a></td></tr>'
    + '<tr style="background:#F9F5FF;"><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">رقم التفعيل</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><strong style="color:#1a1a2e;">' + (data.activationPhone || '-') + '</strong></td></tr>'
    + '<tr><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">الشركة</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><strong style="color:#1a1a2e;">' + (data.company || '-') + '</strong></td></tr>'
    + '<tr style="background:#F9F5FF;"><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">الباقة</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><strong style="color:#7B2FBE;font-size:15px;">' + (data.package || '-') + '</strong></td></tr>'
    + '<tr><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">السعر</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><strong style="color:#22C55E;font-size:18px;">' + (data.price || 0) + ' جنيه</strong></td></tr>'
    + '<tr style="background:#F9F5FF;"><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">طريقة الدفع</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><span style="background:#EEE;padding:4px 14px;border-radius:20px;font-size:13px;">' + (data.payment || '-') + '</span></td></tr>'
    + '<tr><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">رقم التحويل</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><code style="background:#f0f0f0;padding:4px 12px;border-radius:6px;font-size:14px;">' + (data.transferRef || '-') + '</code></td></tr>'
    + '<tr style="background:#F9F5FF;"><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#888;font-size:13px;">ملاحظات</td><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#555;font-style:italic;">' + (data.notes || 'لا يوجد') + '</td></tr>'
    + (data.vodafonePassword ? '<tr style="background:#FFF5F5;"><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#DC2626;font-size:13px;font-weight:bold;">🔑 باسورد فودافون</td><td style="padding:13px 20px;border-bottom:1px solid #eee;"><code style="background:#FEE2E2;color:#DC2626;padding:5px 14px;border-radius:6px;font-size:15px;font-weight:bold;">' + data.vodafonePassword + '</code></td></tr>' : '')
    + (String(data.company || '').includes('اتصالات') ? '<tr style="background:#EFF6FF;"><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#1D4ED8;font-size:13px;font-weight:bold;">📲 آلية التفعيل</td><td style="padding:13px 20px;border-bottom:1px solid #eee;color:#1D4ED8;font-weight:bold;">ابعت للعميل كود التفعيل — ينتظر الرد على واتساب</td></tr>' : '')
    + proofSection
    + '</table></td></tr>'
    + '<tr><td style="padding:28px;text-align:center;"><a href="' + CONFIG.ADMIN_URL + '" style="display:inline-block;background:linear-gradient(135deg,#7B2FBE,#FF6B00);color:white;padding:14px 40px;border-radius:12px;text-decoration:none;font-weight:bold;font-size:15px;">فتح لوحة التحكم</a></td></tr>'
    + '<tr><td style="background:#F4F6F9;padding:16px;text-align:center;border-top:1px solid #eee;"><p style="margin:0;color:#aaa;font-size:12px;">عالم العروض &mdash; هذا البريد تلقائي لا ترد عليه</p></td></tr>'
    + '</table></td></tr></table></body></html>';

  GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, plainText, { htmlBody: htmlBody, name: 'عالم العروض' });
}

// ============================================================
// اختبارات
// ============================================================
function testEmail() {
  var fakeData = {
    customerName:    'حمدي حجاج',
    phone:           '01154620997',
    activationPhone: '01154620997',
    company:         'وي',
    package:         '10,000 ميجا + 1000 دقيقة',
    price:           210,
    payment:         'انستا باي',
    transferRef:     '7363636',
    notes:           'لا يوجد'
  };
  sendAdminNotification('ORD-TEST-0001', fakeData, '');
  console.log('تم ارسال الايميل التجريبي');
}

function testSystem() {
  Logger.log('النظام يعمل بنجاح');
}
