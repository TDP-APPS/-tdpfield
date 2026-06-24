// ═══════════════════════════════════════════════════════════════
//  TDP Field App — Google Apps Script
//
//  انسخ هذا الكود كاملاً في Google Apps Script وانشره كـ Web App
//  يعمل مع القالب الرسمي "Workshop Locations"
//
//  طريقة الإعداد:
//  1. افتح Google Sheet الخاص بك
//  2. Extensions → Apps Script
//  3. احذف الكود القديم والصق هذا الكود
//  4. Deploy → New Deployment → Web App
//     Execute as: Me | Who has access: Anyone
//  5. انسخ الرابط وضعه في التطبيق تحت إعدادات Google Sheets
//
// ═══════════════════════════════════════════════════════════════
//
//  بنية الجدول المطلوبة (الصف الرابع هو العناوين — rows[3]):
//  ──────────────────────────────────────────────────────────────
//  A  | * (رقم تسلسلي)
//  B  | المحافظة / Area
//  C  | اسم المشروع / Project Name
//  D  | نوع المشروع / Project Type    ← drama أو anim أو مسرح أو انيميشن
//  E  | اسم الموقع / Site Name        ← اسم المدرسة
//  F  | عنوان الموقع / Site Address
//  G  | اسم المدرب 1 / Trainer 1 Name
//  H  | رقم المدرب 1 / Trainer 1 Phone
//  I  | اسم المدرب 2 / Trainer 2 Name
//  J  | رقم المدرب 2 / Trainer 2 Phone
//  K  | اسم المدرب 3 / Trainer 3 Name
//  L  | رقم المدرب 3 / Trainer 3 Phone
//  M  | رابط الموقع (Google Maps)
//  N  | الإحداثيات (Lat, Long)         ← مثال: 31.372388,34.30218
//  O  | تاريخ بداية الورشة
//  P  | تاريخ الانتهاء التقريبي
//  Q  | ملاحظات
//
// ═══════════════════════════════════════════════════════════════

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var rows  = sheet.getDataRange().getValues();

  var result = [];

  // الصف 0-2: عنوان وتصميم | الصف 3: أسماء الأعمدة | من الصف 4: البيانات
  for (var i = 4; i < rows.length; i++) {
    var row = rows[i];

    // تجاهل الصفوف الفارغة (بدون اسم موقع على الأقل)
    var siteName = String(row[4] || '').trim();
    if (!siteName) continue;

    // ── نوع المشروع → drama أو anim ──
    var rawType = String(row[3] || '').trim().toLowerCase();
    var wsType = 'drama';
    if (rawType.indexOf('anim') !== -1 || rawType.indexOf('انيم') !== -1) {
      wsType = 'anim';
    }

    // ── بناء قائمة المدربين ──
    var trainers = [];
    var trainerPhones = {};

    var t1name  = String(row[6]  || '').trim();
    var t1phone = String(row[7]  || '').trim();
    var t2name  = String(row[8]  || '').trim();
    var t2phone = String(row[9]  || '').trim();
    var t3name  = String(row[10] || '').trim();
    var t3phone = String(row[11] || '').trim();

    if (t1name) { trainers.push(t1name); if (t1phone) trainerPhones[t1name] = t1phone; }
    if (t2name) { trainers.push(t2name); if (t2phone) trainerPhones[t2name] = t2phone; }
    if (t3name) { trainers.push(t3name); if (t3phone) trainerPhones[t3name] = t3phone; }

    // ── الإحداثيات (إزالة المسافات) ──
    var coords = String(row[13] || '').trim().replace(/\s+/g, '');

    // ── التواريخ ──
    var startDate = row[14] ? Utilities.formatDate(new Date(row[14]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
    var endDate   = row[15] ? Utilities.formatDate(new Date(row[15]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';

    result.push({
      id:          'w' + i,
      name:        siteName,
      addr:        String(row[5]  || '').trim(),
      area:        String(row[1]  || '').trim(),
      project:     String(row[2]  || '').trim(),
      coords:      coords,
      mapsLink:    String(row[12] || '').trim(),
      startDate:   startDate,
      endDate:     endDate,
      notes:       String(row[16] || '').trim(),
      sessions: [
        {
          days:          '',
          type:          wsType,
          gender:        '',
          trainers:      trainers,
          trainerPhones: trainerPhones
        }
      ],
      totalVisits: 0
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
