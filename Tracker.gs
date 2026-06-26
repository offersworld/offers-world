// ============================================================
// Tracker.gs — نظام متابعة الباقات اليدوية (صفحة عبير)
// ============================================================

const TRACKER_SHEET = '📊 PackageLogs';
const TRACKER_PASSWORD = PropertiesService.getScriptProperties().getProperty('TRACKER_PASSWORD') || 'abeer123';

// ============================================================
// doGet — إضافة page=tracker
// ============================================================
// ملاحظة: أضف السطر ده في Client.gs داخل doGet:
//   if (page === 'tracker') {
//     return HtmlService.createHtmlOutputFromFile('tracker')
//       .setTitle('متابعة الباقات — عالم العروض')
//       .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
//   }

// ============================================================
// Tracker Actions — تُضاف لـ doGet في Client.gs
// ============================================================
// أضف ده في doGet بعد action === 'packages':
//
//   if (action === 'trackerPackages') {
//     return respond(getTrackerPackages());
//   }
//   if (action === 'trackerLogs') {
//     return respond(getTrackerLogs(e.parameter));
//   }
//   if (action === 'trackerSummary') {
//     return respond(getTrackerSummary(e.parameter.period, e.parameter.month, e.parameter.year));
//   }
//
// وأضف ده في doPost داخل الـ switch/if:
//   if (action === 'addTrackerLog')    return respond(addTrackerLog(data));
//   if (action === 'updateTrackerLog') return respond(updateTrackerLog(data));
//   if (action === 'deleteTrackerLog') return respond(deleteTrackerLog(data.logId));
//   if (action === 'checkTrackerPass') return respond({ success: data.pass === TRACKER_PASSWORD });

// ============================================================
// جلب الباقات من Master Data (نفس الـ Packages sheet الموجودة)
// ============================================================
function getTrackerPackages() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PACKAGES);
    var data = sheet.getDataRange().getValues();
    // تجميع الباقات حسب الشركة
    var grouped = {};

    for (var i = 2; i < data.length; i++) {
      var row = data[i];
      var pkgId      = String(row[0] || '').trim();
      var company    = String(row[1] || '').trim();
      var pkgName    = String(row[2] || '').trim();
      var pkgPrice   = Number(row[3]) || 0;
      var status     = String(row[5] || '').trim();

      if (!pkgName || !company || status !== 'متاح') continue;

      if (!grouped[company]) grouped[company] = [];
      grouped[company].push({ id: pkgId, name: pkgName, price: pkgPrice });
    }

    // ترتيب الباقات حسب السعر داخل كل شركة
    Object.keys(grouped).forEach(function(co) {
      grouped[co].sort(function(a, b) { return a.price - b.price; });
    });

    return { success: true, data: grouped };
  } catch (err) {
    logError('getTrackerPackages', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// إضافة عملية جديدة
// ============================================================
function addTrackerLog(data) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);

    var logId = 'LOG-' + Utilities.formatDate(new Date(), 'Africa/Cairo', 'yyMMddHHmmss') + '-' + Math.floor(Math.random() * 100);

    // تاريخ العملية: إما من المدخلات أو اليوم
    var opDate = data.date ? data.date : Utilities.formatDate(new Date(), 'Africa/Cairo', 'dd/MM/yyyy');

    sheet.appendRow([
      logId,                              // A: رقم العملية
      opDate,                             // B: تاريخ العملية
      data.company     || '',             // C: الشركة
      data.packageId   || '',             // D: رقم الباقة (FK)
      data.packageName || '',             // E: اسم الباقة (مُملأ تلقائياً)
      Number(data.price) || 0,            // F: السعر (مُملأ تلقائياً)
      data.source      || '',             // G: مصدر العملية
      data.clientRef   || '',             // H: اسم/رقم العميل (اختياري)
      data.status      || 'تم',           // I: الحالة
      data.notes       || '',             // J: ملاحظات
      Utilities.formatDate(new Date(), 'Africa/Cairo', 'dd/MM/yyyy HH:mm:ss')  // K: وقت الإدخال
    ]);

    addLog('تسجيل باقة يدوي', logId, data.company + ' - ' + data.packageName);
    return { success: true, logId: logId };
  } catch (err) {
    logError('addTrackerLog', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// تعديل عملية موجودة — بحث بالصف المباشر (rowId)
// ============================================================
function editTrackerLog(data) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);

    var row;
    var rowId = parseInt(data.rowId, 10);

    if (!isNaN(rowId) && rowId > 1) {
      // استخدام رقم الصف مباشرةً (أسرع)
      row = rowId;
    } else {
      // fallback: بحث نصي بالـ logId
      var found = sheet.getRange('A:A').createTextFinder(String(data.logId || '').trim()).matchEntireCell(true).findNext();
      if (!found) return { success: false, error: 'العملية غير موجودة' };
      row = found.getRow();
    }

    // التحديث الدفعي في خطوة واحدة (أسرع من استدعاءات متعددة)
    sheet.getRange(row, 3, 1, 8).setValues([[
      data.company     || '',
      data.packageId   || '',
      data.packageName || '',
      Number(data.price) || 0,
      data.source      || '',
      data.clientRef   || '',
      data.status      || 'تم',
      data.notes       || ''
    ]]);

    addLog('تعديل باقة يدوي', data.logId || ('row#' + row), data.packageName);
    return { success: true };
  } catch (err) {
    logError('editTrackerLog', err.toString());
    return { success: false, error: err.toString() };
  }
}

// للتوافق مع النظام القديم
function updateTrackerLog(data) {
  return editTrackerLog(data);
}

// ============================================================
// حذف عملية — بحث بالصف المباشر (rowId)
// ============================================================
function deleteTrackerLog(data) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);

    var rowToDelete;
    // data قد تكون كائناً { rowId, logId } أو مجرد string (النمط القديم)
    var rowId  = parseInt(typeof data === 'object' ? data.rowId  : null, 10);
    var logId  = typeof data === 'object' ? (data.logId || '') : String(data || '');

    if (!isNaN(rowId) && rowId > 1) {
      rowToDelete = rowId;
    } else {
      var found = sheet.getRange('A:A').createTextFinder(logId.trim()).matchEntireCell(true).findNext();
      if (!found) return { success: false, error: 'العملية غير موجودة' };
      rowToDelete = found.getRow();
    }

    sheet.deleteRow(rowToDelete);
    addLog('حذف باقة يدوي', logId || ('row#' + rowToDelete), '');
    return { success: true };
  } catch (err) {
    logError('deleteTrackerLog', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// جلب السجلات مع فلترة
// ============================================================
function getTrackerLogs(params) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);
    var data  = sheet.getDataRange().getValues();
    var logs  = [];

    var filterCompany = params.company   || '';
    var filterStatus  = params.status    || '';
    var filterClient  = params.client    || '';

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var log = {
        rowId:       i + 1,                    // رقم الصف الفعلي في الشيت (1-indexed)
        logId:       String(row[0] || ''),
        date:        String(row[1] || ''),
        company:     String(row[2] || ''),
        packageId:   String(row[3] || ''),
        packageName: String(row[4] || ''),
        price:       Number(row[5]) || 0,
        source:      String(row[6] || ''),
        clientRef:   String(row[7] || ''),
        status:      String(row[8] || ''),
        notes:       String(row[9] || ''),
        createdAt:   String(row[10] || '')
      };

      if (filterCompany && log.company !== filterCompany) continue;
      if (filterStatus  && log.status  !== filterStatus)  continue;
      if (filterClient  && !log.clientRef.toLowerCase().includes(filterClient.toLowerCase())) continue;

      logs.push(log);
    }

    // ترتيب من الأحدث للأقدم
    logs.reverse();
    return { success: true, data: logs };
  } catch (err) {
    logError('getTrackerLogs', err.toString());
    return { success: false, error: err.toString(), data: [] };
  }
}

// ============================================================
// ملخص الفترة (يومي / شهري)
// ============================================================
function getTrackerSummary(period, month, year) {
  try {
    var ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = getOrCreateTrackerSheet(ss);
    var data  = sheet.getDataRange().getValues();

    var now      = new Date();
    var todayStr = Utilities.formatDate(now, 'Africa/Cairo', 'dd/MM/yyyy');
    var targetMonth = month  ? String(month)  : Utilities.formatDate(now, 'Africa/Cairo', 'MM');
    var targetYear  = year   ? String(year)   : Utilities.formatDate(now, 'Africa/Cairo', 'yyyy');

    // تجميع البيانات
    var daily   = { count: 0, revenue: 0, byCompany: {}, byStar: {} };
    var monthly = { count: 0, revenue: 0, byCompany: {}, byPackage: {}, bySource: {}, daily: {} };

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;

      var dateStr = String(row[1] || '').trim();  // dd/MM/yyyy
      var company = String(row[2] || '').trim();
      var pkgName = String(row[4] || '').trim();
      var price   = Number(row[5]) || 0;
      var source  = String(row[6] || '').trim();
      var status  = String(row[8] || '').trim();

      if (status === 'ملغي') continue; // نستثني الملغيات من الإحصائيات

      var parts = dateStr.split('/'); // [dd, MM, yyyy]
      if (parts.length < 3) continue;
      var rowDay   = parts[0];
      var rowMonth = parts[1];
      var rowYear  = parts[2];

      // ملخص اليوم
      if (dateStr === todayStr) {
        daily.count++;
        daily.revenue += price;
        daily.byCompany[company] = (daily.byCompany[company] || 0) + 1;
      }

      // ملخص الشهر
      if (rowMonth === targetMonth && rowYear === targetYear) {
        monthly.count++;
        monthly.revenue += price;
        monthly.byCompany[company]  = (monthly.byCompany[company]  || 0) + 1;
        monthly.byPackage[pkgName]  = (monthly.byPackage[pkgName]  || 0) + 1;
        monthly.bySource[source || 'غير محدد'] = (monthly.bySource[source || 'غير محدد'] || 0) + 1;
        // مبيعات يومية داخل الشهر (لمخطط اتجاه الشهر)
        var dayKey = rowDay + '/' + rowMonth;
        monthly.daily[dayKey] = (monthly.daily[dayKey] || 0) + 1;
      }
    }

    // حساب أعلى وأقل شركة هذا الشهر
    var companies = Object.keys(monthly.byCompany);
    var topCompany    = '';
    var bottomCompany = '';
    var topCount = 0, bottomCount = Infinity;
    companies.forEach(function(c) {
      if (monthly.byCompany[c] > topCount)    { topCount = monthly.byCompany[c];    topCompany = c; }
      if (monthly.byCompany[c] < bottomCount) { bottomCount = monthly.byCompany[c]; bottomCompany = c; }
    });

    // حساب أعلى وأقل باقة هذا الشهر
    var packages = Object.keys(monthly.byPackage);
    var topPkg = '', bottomPkg = '';
    var topPkgCount = 0, bottomPkgCount = Infinity;
    packages.forEach(function(p) {
      if (monthly.byPackage[p] > topPkgCount)    { topPkgCount = monthly.byPackage[p];    topPkg = p; }
      if (monthly.byPackage[p] < bottomPkgCount) { bottomPkgCount = monthly.byPackage[p]; bottomPkg = p; }
    });

    return {
      success: true,
      today: {
        date:      todayStr,
        count:     daily.count,
        revenue:   daily.revenue,
        byCompany: daily.byCompany
      },
      monthly: {
        month:        targetMonth + '/' + targetYear,
        count:        monthly.count,
        revenue:      monthly.revenue,
        byCompany:    monthly.byCompany,
        byPackage:    monthly.byPackage,
        bySource:     monthly.bySource,
        dailyTrend:   monthly.daily,
        topCompany:    { name: topCompany,    count: topCount },
        bottomCompany: { name: bottomCompany, count: bottomCount },
        topPackage:    { name: topPkg,        count: topPkgCount },
        bottomPackage: { name: bottomPkg,     count: bottomPkgCount }
      }
    };
  } catch (err) {
    logError('getTrackerSummary', err.toString());
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// دالة مساعدة: جلب/إنشاء الـ Sheet
// ============================================================
function getOrCreateTrackerSheet(ss) {
  var sheet = ss.getSheetByName(TRACKER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TRACKER_SHEET);
    // رأس الجدول
    sheet.appendRow([
      'رقم العملية', 'تاريخ العملية', 'الشركة', 'رقم الباقة', 'اسم الباقة',
      'السعر', 'مصدر العملية', 'اسم/رقم العميل', 'الحالة', 'ملاحظات', 'وقت الإدخال'
    ]);
    // تنسيق الرأس
    sheet.getRange(1, 1, 1, 11).setBackground('#8b3cf7').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 11, 150);
  }
  return sheet;
}
