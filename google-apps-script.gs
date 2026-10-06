/**
 * ربط دعوة الزفاف مع Google Sheets
 *
 * طريقة التركيب:
 * 1) أنشئ Google Sheet جديد (أي اسم).
 * 2) من القائمة: الإضافات (Extensions) ← Apps Script.
 * 3) احذف الكود الموجود والصق هذا الملف كاملاً، ثم احفظ.
 * 4) نشر (Deploy) ← نشر جديد (New deployment) ← النوع: تطبيق ويب (Web app)
 *      - التنفيذ باسم (Execute as): أنا (Me)
 *      - من يمكنه الوصول (Who has access): أي شخص (Anyone)
 * 5) وافق على الصلاحيات، ثم انسخ رابط تطبيق الويب (ينتهي بـ /exec)
 *    والصقه في ملف الدعوة داخل SHEET_URL.
 *
 * ملاحظة: إذا عدّلت هذا الكود لاحقًا، اختر Deploy ← Manage deployments ← Edit ← New version
 * حتى يبقى نفس الرابط ويعمل التعديل.
 *
 * إخفاء رسالة من الموقع: غيّر خانة "إظهار" في الشيت إلى "لا".
 */

const SHEET_NAME = 'الردود';
const HEADERS = ['التاريخ', 'الاسم', 'الحضور', 'عدد المرافقين', 'الرسالة', 'إظهار', 'المعرّف'];
const ID_COL = 7; // معرّف الجهاز: يُستخدم لتحديث الرد بدل تكراره

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  } else if (!sh.getRange(1, ID_COL).getValue()) {
    sh.getRange(1, ID_COL).setValue(HEADERS[ID_COL - 1]).setFontWeight('bold');
  }
  return sh;
}

// رقم الصف الذي يحمل هذا المعرّف، أو 0 إذا لم يوجد
function findRowById_(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, ID_COL, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 2;
  }
  return 0;
}

// يمنع تنفيذ الصيغ (= + - @) ويقص النص الطويل
function clean_(value, max) {
  let s = String(value || '').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// استقبال تأكيد الحضور والرسائل
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json_({ ok: true }); // حقل مخفي لصد البوتات

  const name = clean_(p.name, 80);
  if (!name) return json_({ ok: false, error: 'name required' });

  const id = String(p.id || '').replace(/[^\w-]/g, '').slice(0, 64);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const row = findRowById_(sh, id);
    const values = [
      new Date(),
      name,
      p.attend === 'no' ? 'معتذر' : 'سيحضر',
      clean_(p.guests, 40),
      clean_(p.message, 500)
    ];
    if (row) {
      // تعديل رد سابق: نحدّث الصف نفسه ونترك خانة "إظهار" كما هي
      sh.getRange(row, 1, 1, values.length).setValues([values]);
    } else {
      sh.appendRow(values.concat(['نعم', id]));
    }
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// إرجاع الرسائل الظاهرة لعرضها في الموقع (الأحدث أولاً)
function doGet() {
  const rows = getSheet_().getDataRange().getValues().slice(1);
  const messages = rows
    .filter(r => String(r[4]).trim() && String(r[5]).trim() !== 'لا')
    .map(r => ({
      id: String(r[6] || ''),
      name: String(r[1]).replace(/^'/, ''),
      message: String(r[4]).replace(/^'/, ''),
      date: r[0] instanceof Date ? r[0].toISOString() : ''
    }))
    .reverse()
    .slice(0, 150);
  return json_({ ok: true, messages: messages });
}
