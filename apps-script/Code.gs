/**
 * ============================================================================
 *  نظام إصدار الشهادات الإلكترونية التلقائي — للمدربين والمدربات والمتدربين والمتدربات
 * ============================================================================
 *  المزايا:
 *   • تهيئة كاملة بضغطة واحدة (الأوراق، التنسيق، القوائم المنسدلة، لوحة المعلومات).
 *   • إنشاء قالبَي شهادة (متدربين / مدربين) تلقائياً في Google Slides قابلين للتعديل.
 *   • صياغة لغوية حسب الجنس (المتدرب/المتدربة، اجتاز/اجتازت، له/لها ...).
 *   • شروط أهلية للمتدربين (نسبة الحضور + الدرجة) مع حساب التقدير تلقائياً.
 *   • رقم شهادة تسلسلي فريد + رمز تحقق + QR يفتح صفحة التحقق.
 *   • حفظ الشهادات PDF في Drive مرتبة حسب الدورة + إرسالها بالبريد تلقائياً.
 *   • إصدار تلقائي فور تغيير الحالة إلى «مكتمل» + فحص دوري كل ساعة.
 *   • إعادة إصدار / إعادة إرسال / إلغاء شهادة + سجل مركزي لكل الشهادات.
 *   • صفحة ويب عامة للتحقق من صحة الشهادة (ملف Verify.html).
 *
 *  التثبيت: من ملف Google Sheets ← الإضافات (Extensions) ← Apps Script
 *  الصق هذا الكود في Code.gs، وأنشئ ملف HTML باسم Verify والصق محتواه، ثم احفظ،
 *  وحدّث صفحة الجدول وستظهر قائمة «🎓 نظام الشهادات» ← «تهيئة النظام».
 * ============================================================================
 */

// ============================== الإعدادات العامة ==============================
const CFG = {
  SHEETS: {
    SETTINGS: 'الإعدادات',
    COURSES: 'الدورات',
    TRAINEES: 'المتدربون',
    TRAINERS: 'المدربون',
    LOG: 'سجل الشهادات',
    DASH: 'لوحة المعلومات'
  },
  STATUS: {
    IN_PROGRESS: 'قيد التدريب',
    READY: 'مكتمل',          // ← عند اختيار هذه الحالة تُصدر الشهادة تلقائياً
    ISSUED: 'صدرت',
    NOT_ELIGIBLE: 'غير مؤهل',
    REVOKED: 'ملغاة',
    ERROR: 'خطأ'
  },
  SENT: { PENDING: 'بانتظار الإرسال', NO_EMAIL: 'لا يوجد بريد', DISABLED: 'الإرسال معطّل' },
  TZ: 'Asia/Riyadh',
  MAX_RUNTIME_MS: 4.5 * 60 * 1000,   // حد آمن قبل مهلة جوجل (6 دقائق)
  FONT_TITLE: 'Amiri',
  FONT_BODY: 'Cairo'
};

// أسماء الأعمدة (النظام يعتمد على أسماء العناوين وليس ترتيبها)
const H = {
  ID: 'المعرف', NAME: 'الاسم الكامل', GENDER: 'الجنس', NATIONAL_ID: 'رقم الهوية',
  EMAIL: 'البريد الإلكتروني', PHONE: 'الجوال', COURSE: 'رمز الدورة',
  ATTEND: 'نسبة الحضور %', GRADE: 'الدرجة', SPECIALTY: 'التخصص',
  STATUS: 'الحالة', CERT_NO: 'رقم الشهادة', CODE: 'رمز التحقق',
  LINK: 'رابط الشهادة', DATE: 'تاريخ الإصدار', SENT: 'حالة الإرسال', NOTES: 'ملاحظات'
};

const HEADERS = {
  TRAINEES: [H.ID, H.NAME, H.GENDER, H.NATIONAL_ID, H.EMAIL, H.PHONE, H.COURSE, H.ATTEND, H.GRADE,
             H.STATUS, H.CERT_NO, H.CODE, H.LINK, H.DATE, H.SENT, H.NOTES],
  TRAINERS: [H.ID, H.NAME, H.GENDER, H.EMAIL, H.PHONE, H.COURSE, H.SPECIALTY,
             H.STATUS, H.CERT_NO, H.CODE, H.LINK, H.DATE, H.SENT, H.NOTES],
  COURSES: ['رمز الدورة', 'اسم الدورة', 'تاريخ البداية', 'تاريخ النهاية', 'عدد الساعات', 'مكان الانعقاد', 'المدرب/المدربة'],
  LOG: ['رقم الشهادة', 'رمز التحقق', 'النوع', 'الاسم', 'الجنس', 'البريد الإلكتروني', 'رمز الدورة',
        'اسم الدورة', 'عدد الساعات', 'تاريخ الإصدار', 'رابط الشهادة', 'الحالة', 'تاريخ الإلغاء']
};

// مفاتيح ورقة الإعدادات
const S = {
  ORG: 'اسم الجهة',
  INTRO_TRAINEE: 'عبارة الافتتاح - المتدربين',
  INTRO_TRAINER: 'عبارة الافتتاح - المدربين',
  SIGNER: 'اسم المعتمد',
  SIGNER_TITLE: 'منصب المعتمد',
  PREFIX: 'بادئة رقم الشهادة',
  MIN_ATTEND: 'الحد الأدنى للحضور %',
  MIN_GRADE: 'الحد الأدنى للدرجة',
  AUTO_EMAIL: 'الإرسال بالبريد تلقائياً',
  SENDER_NAME: 'اسم المرسل',
  REPLY_TO: 'بريد الرد',
  VERIFY_URL: 'رابط صفحة التحقق',
  FOLDER_ID: 'معرف مجلد الشهادات',
  TPL_TRAINEE: 'معرف قالب المتدربين',
  TPL_TRAINER: 'معرف قالب المدربين',
  PRIMARY: 'اللون الرئيسي',
  SECONDARY: 'اللون الثانوي',
  HIJRI: 'إظهار التاريخ الهجري'
};

const SETTINGS_DEFAULTS = [
  [S.ORG, 'جامعة الحضارة العالمية', 'يظهر في رأس الشهادة وفي البريد'],
  [S.INTRO_TRAINEE, 'تشهد {{ORG}} بأنّ', 'عبارة ما قبل الاسم في شهادة المتدربين ({{ORG}} = اسم الجهة)'],
  [S.INTRO_TRAINER, 'تتقدم {{ORG}} بخالص الشكر والتقدير إلى', 'عبارة ما قبل الاسم في شهادة المدربين'],
  [S.SIGNER, 'أ.د. ............', 'اسم صاحب التوقيع'],
  [S.SIGNER_TITLE, 'رئيس الجامعة', 'منصب صاحب التوقيع'],
  [S.PREFIX, 'WCU', 'بادئة رقم الشهادة (أحرف إنجليزية)'],
  [S.MIN_ATTEND, 80, 'اتركه فارغاً أو 0 لإلغاء شرط الحضور'],
  [S.MIN_GRADE, 60, 'اتركه فارغاً إذا لا يوجد اختبار'],
  [S.AUTO_EMAIL, 'نعم', 'نعم / لا'],
  [S.SENDER_NAME, '', 'اسم المرسل الظاهر في البريد (افتراضياً اسم الجهة)'],
  [S.REPLY_TO, '', 'بريد يستقبل ردود المتدربين (اختياري)'],
  [S.VERIFY_URL, '', 'رابط صفحة التحقق التي يفتحها رمز QR: رابط صفحة GitHub أو رابط تطبيق الويب (/exec)'],
  [S.FOLDER_ID, '', 'يُنشأ تلقائياً — لا تعدّله'],
  [S.TPL_TRAINEE, '', 'يُنشأ تلقائياً — يمكن استبداله بمعرف قالب Slides خاص بك'],
  [S.TPL_TRAINER, '', 'يُنشأ تلقائياً — يمكن استبداله بمعرف قالب Slides خاص بك'],
  [S.PRIMARY, '#0B4F6C', 'لون الإطار والعناوين (يُطبّق عند إنشاء القوالب)'],
  [S.SECONDARY, '#C9A227', 'لون الزخرفة والاسم (يُطبّق عند إنشاء القوالب)'],
  [S.HIJRI, 'نعم', 'نعم / لا — إضافة التاريخ الهجري بجانب الميلادي']
];

// ================================ القائمة ================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🎓 نظام الشهادات')
    .addItem('⚙️ تهيئة النظام (أول مرة)', 'setupSystem')
    .addItem('🎨 إعادة إنشاء قوالب الشهادات', 'menuRecreateTemplates')
    .addSeparator()
    .addItem('▶️ إصدار جميع الشهادات المكتملة الآن', 'menuProcessNow')
    .addItem('📄 إصدار شهادة الصف المحدد', 'menuIssueSelected')
    .addItem('🔄 إعادة إصدار شهادة الصف المحدد', 'menuReissueSelected')
    .addItem('✉️ إعادة إرسال بريد الصف المحدد', 'menuResendSelected')
    .addItem('⛔ إلغاء شهادة الصف المحدد', 'menuRevokeSelected')
    .addSeparator()
    .addItem('✅ تفعيل الإصدار التلقائي', 'enableAutomation')
    .addItem('⏸️ إيقاف الإصدار التلقائي', 'disableAutomation')
    .addSeparator()
    .addItem('🔐 تغيير رموز دخول المنصة', 'menuChangePin')
    .addSeparator()
    .addItem('📖 دليل الاستخدام', 'showHelp')
    .addToUi();
}

// ============================== تهيئة النظام ==============================
function setupSystem() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CFG.TZ);
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  ensureUsers_();

  // 1) الإعدادات
  const set = ensureSheet_(ss, CFG.SHEETS.SETTINGS, ['البند', 'القيمة', 'شرح'], '#6A1B9A');
  const existing = getSettings_();
  SETTINGS_DEFAULTS.forEach(([k, v, note]) => {
    if (!(k in existing)) set.appendRow([k, v, note]);
  });
  set.setColumnWidth(1, 220).setColumnWidth(2, 340).setColumnWidth(3, 380);
  set.getRange('A2:A').setFontWeight('bold').setBackground('#F3E5F5');
  const yesNo = SpreadsheetApp.newDataValidation().requireValueInList(['نعم', 'لا'], true).build();
  [S.AUTO_EMAIL, S.HIJRI].forEach(k => {
    const r = findSettingRow_(k); if (r) set.getRange(r, 2).setDataValidation(yesNo);
  });

  // 2) الدورات
  const courses = ensureSheet_(ss, CFG.SHEETS.COURSES, HEADERS.COURSES, '#1565C0');
  courses.getRange('C2:D').setNumberFormat('yyyy/mm/dd');
  courses.getRange('E2:E').setNumberFormat('0');
  courses.setColumnWidth(2, 260);
  if (courses.getLastRow() < 2) {
    const t = new Date();
    courses.appendRow(['C-001', 'مهارات التواصل الفعّال', new Date(t.getFullYear(), t.getMonth(), 1),
      new Date(t.getFullYear(), t.getMonth(), 5), 15, 'قاعة التدريب الرئيسية', 'د. خالد الحربي']);
    courses.appendRow(['C-002', 'القيادة وإدارة الفرق', new Date(t.getFullYear(), t.getMonth(), 10),
      new Date(t.getFullYear(), t.getMonth(), 12), 9, 'عن بُعد (Zoom)', 'أ. هند الزهراني']);
  }

  // 3) المتدربون والمدربون
  const trainees = ensureSheet_(ss, CFG.SHEETS.TRAINEES, HEADERS.TRAINEES, '#2E7D32');
  const trainers = ensureSheet_(ss, CFG.SHEETS.TRAINERS, HEADERS.TRAINERS, '#EF6C00');
  if (trainees.getLastRow() < 2) {
    trainees.appendRow(['ST-0001', 'محمد أحمد العلي', 'ذكر', '', '', '', 'C-001', 95, 88, CFG.STATUS.IN_PROGRESS]);
    trainees.appendRow(['ST-0002', 'نورة سعد القحطاني', 'أنثى', '', '', '', 'C-001', 100, 94, CFG.STATUS.IN_PROGRESS]);
  }
  if (trainers.getLastRow() < 2) {
    trainers.appendRow(['TR-0001', 'د. خالد الحربي', 'ذكر', '', '', 'C-001', 'مهارات شخصية', CFG.STATUS.IN_PROGRESS]);
    trainers.appendRow(['TR-0002', 'أ. هند الزهراني', 'أنثى', '', '', 'C-002', 'إدارة وقيادة', CFG.STATUS.IN_PROGRESS]);
  }
  formatPeopleSheet_(trainees, HEADERS.TRAINEES, true);
  formatPeopleSheet_(trainers, HEADERS.TRAINERS, false);

  // 4) السجل
  const log = ensureSheet_(ss, CFG.SHEETS.LOG, HEADERS.LOG, '#455A64');
  log.getRange('A:B').setNumberFormat('@');
  addStatusFormatting_(log, HEADERS.LOG.indexOf('الحالة') + 1, [['سارية', '#C8E6C9'], ['ملغاة', '#E0E0E0']]);

  // 5) لوحة المعلومات
  buildDashboard_(ss);

  // حذف الورقة الافتراضية الفارغة
  ['الورقة1', 'Sheet1'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  // 6) المجلد والقوالب
  const s = getSettings_();
  const folder = getRootFolder_(s);
  if (!s[S.TPL_TRAINEE] || !safeFile_(s[S.TPL_TRAINEE])) setSetting_(S.TPL_TRAINEE, buildTemplate_('trainee', s, folder));
  if (!s[S.TPL_TRAINER] || !safeFile_(s[S.TPL_TRAINER])) setSetting_(S.TPL_TRAINER, buildTemplate_('trainer', s, folder));

  ss.setActiveSheet(ss.getSheetByName(CFG.SHEETS.DASH));
  ui.alert('✅ تمت التهيئة بنجاح',
    'تم إنشاء الأوراق والقوالب ومجلد الشهادات في Google Drive.\n\n' +
    'الخطوات التالية:\n' +
    '1) عدّل ورقة «الإعدادات» (اسم الجهة، المعتمد...).\n' +
    '2) انشر صفحة التحقق: نشر ← نشر جديد ← تطبيق ويب (التنفيذ: أنا، الوصول: أي شخص)، والصق الرابط في الإعدادات.\n' +
    '3) من القائمة اختر «تفعيل الإصدار التلقائي».\n' +
    '4) للتجربة: ضع بريدك لأحد الصفوف التجريبية وغيّر الحالة إلى «مكتمل».',
    ui.ButtonSet.OK);
}

function ensureSheet_(ss, name, headers, color) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.setRightToLeft(true);
  if (headers && headers.length) {
    const r = sh.getRange(1, 1, 1, headers.length);
    if (sh.getLastRow() === 0 || String(sh.getRange(1, 1).getValue()).trim() === '') r.setValues([headers]);
    r.setFontWeight('bold').setFontColor('#FFFFFF').setBackground(color)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 36);
  }
  if (color) sh.setTabColor(color);
  return sh;
}

function formatPeopleSheet_(sh, headers, isTrainee) {
  const col = n => headers.indexOf(n) + 1;
  const rows = sh.getMaxRows() - 1;
  const v = SpreadsheetApp.newDataValidation;
  sh.getRange(2, col(H.GENDER), rows).setDataValidation(v().requireValueInList(['ذكر', 'أنثى'], true).build());
  sh.getRange(2, col(H.STATUS), rows).setDataValidation(
    v().requireValueInList(Object.values(CFG.STATUS), true).build());
  const coursesRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEETS.COURSES).getRange('A2:A');
  sh.getRange(2, col(H.COURSE), rows).setDataValidation(
    v().requireValueInRange(coursesRange, true).setAllowInvalid(false).build());
  sh.getRange(2, col(H.EMAIL), rows).setDataValidation(
    v().requireTextIsEmail().setAllowInvalid(false).setHelpText('أدخل بريداً إلكترونياً صحيحاً').build());
  if (isTrainee) {
    sh.getRange(2, col(H.ATTEND), rows).setDataValidation(v().requireNumberBetween(0, 100).build()).setNumberFormat('0');
    sh.getRange(2, col(H.GRADE), rows).setDataValidation(v().requireNumberBetween(0, 100).build()).setNumberFormat('0.##');
    sh.getRange(2, col(H.NATIONAL_ID), rows).setNumberFormat('@');
  }
  sh.getRange(2, col(H.PHONE), rows).setNumberFormat('@');
  sh.getRange(2, col(H.CERT_NO), rows, 2).setNumberFormat('@');
  sh.getRange(2, col(H.DATE), rows).setNumberFormat('@');
  sh.setColumnWidth(col(H.NAME), 220).setColumnWidth(col(H.EMAIL), 220)
    .setColumnWidth(col(H.LINK), 260).setColumnWidth(col(H.NOTES), 260).setColumnWidth(col(H.CERT_NO), 170);
  addStatusFormatting_(sh, col(H.STATUS), [
    [CFG.STATUS.READY, '#BBDEFB'], [CFG.STATUS.ISSUED, '#C8E6C9'], [CFG.STATUS.NOT_ELIGIBLE, '#FFE0B2'],
    [CFG.STATUS.ERROR, '#FFCDD2'], [CFG.STATUS.REVOKED, '#E0E0E0']
  ]);
}

function addStatusFormatting_(sh, colIdx, pairs) {
  const range = sh.getRange(2, colIdx, sh.getMaxRows() - 1);
  const rules = sh.getConditionalFormatRules().filter(r =>
    !r.getRanges().some(x => x.getColumn() === colIdx));
  pairs.forEach(([text, bg]) => rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setRanges([range]).build()));
  sh.setConditionalFormatRules(rules);
}

function buildDashboard_(ss) {
  const sh = ensureSheet_(ss, CFG.SHEETS.DASH, null, '#AD1457');
  sh.clear();
  const L = (headers, name) => colLetter_(headers.indexOf(name) + 1);
  const T = `'${CFG.SHEETS.TRAINEES}'!`, R = `'${CFG.SHEETS.TRAINERS}'!`;
  const tG = L(HEADERS.TRAINEES, H.GENDER), tS = L(HEADERS.TRAINEES, H.STATUS), tN = L(HEADERS.TRAINEES, H.NAME), tC = L(HEADERS.TRAINEES, H.COURSE);
  const rG = L(HEADERS.TRAINERS, H.GENDER), rS = L(HEADERS.TRAINERS, H.STATUS), rN = L(HEADERS.TRAINERS, H.NAME), rC = L(HEADERS.TRAINERS, H.COURSE);

  sh.getRange('A1:E1').merge().setValue('📊 لوحة معلومات الشهادات').setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setBackground('#AD1457').setFontColor('#FFFFFF');
  sh.setRowHeight(1, 44);
  const rows = [
    ['البند', 'المتدربون', 'المدربون', 'الإجمالي'],
    ['العدد الكلي', `=COUNTA(${T}${tN}2:${tN})`, `=COUNTA(${R}${rN}2:${rN})`, '=B4+C4'],
    ['ذكور', `=COUNTIF(${T}${tG}2:${tG},"ذكر")`, `=COUNTIF(${R}${rG}2:${rG},"ذكر")`, '=B5+C5'],
    ['إناث', `=COUNTIF(${T}${tG}2:${tG},"أنثى")`, `=COUNTIF(${R}${rG}2:${rG},"أنثى")`, '=B6+C6'],
    ['شهادات صادرة', `=COUNTIF(${T}${tS}2:${tS},"${CFG.STATUS.ISSUED}")`, `=COUNTIF(${R}${rS}2:${rS},"${CFG.STATUS.ISSUED}")`, '=B7+C7'],
    ['بانتظار الإصدار (مكتمل)', `=COUNTIF(${T}${tS}2:${tS},"${CFG.STATUS.READY}")`, `=COUNTIF(${R}${rS}2:${rS},"${CFG.STATUS.READY}")`, '=B8+C8'],
    ['قيد التدريب', `=COUNTIF(${T}${tS}2:${tS},"${CFG.STATUS.IN_PROGRESS}")`, `=COUNTIF(${R}${rS}2:${rS},"${CFG.STATUS.IN_PROGRESS}")`, '=B9+C9'],
    ['غير مؤهل', `=COUNTIF(${T}${tS}2:${tS},"${CFG.STATUS.NOT_ELIGIBLE}")`, '—', '=B10'],
    ['أخطاء', `=COUNTIF(${T}${tS}2:${tS},"${CFG.STATUS.ERROR}")`, `=COUNTIF(${R}${rS}2:${rS},"${CFG.STATUS.ERROR}")`, '=B11+C11'],
    ['شهادات ملغاة', `=COUNTIF(${T}${tS}2:${tS},"${CFG.STATUS.REVOKED}")`, `=COUNTIF(${R}${rS}2:${rS},"${CFG.STATUS.REVOKED}")`, '=B12+C12'],
    ['حصة البريد المتبقية اليوم', 'تظهر بعد أول إصدار', '', '']
  ];
  sh.getRange(3, 1, rows.length, 4).setValues(rows);
  sh.getRange('A3:D3').setFontWeight('bold').setBackground('#FCE4EC');
  sh.getRange('A4:A13').setFontWeight('bold');
  sh.getRange('A3:D13').setBorder(true, true, true, true, true, true, '#F8BBD0', SpreadsheetApp.BorderStyle.SOLID)
    .setHorizontalAlignment('center');

  sh.getRange('A15:E15').merge().setValue('حسب الدورة').setFontWeight('bold').setBackground('#AD1457').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  sh.getRange('A16:E16').setValues([['رمز الدورة', 'اسم الدورة', 'عدد المتدربين', 'شهادات المتدربين الصادرة', 'عدد المدربين']])
    .setFontWeight('bold').setBackground('#FCE4EC').setHorizontalAlignment('center');
  const C = `'${CFG.SHEETS.COURSES}'!`;
  sh.getRange('A17').setFormula(`=IFERROR(FILTER(${C}A2:B,${C}A2:A<>""),"")`);
  sh.getRange('C17').setFormula(`=ARRAYFORMULA(IF(A17:A="","",COUNTIF(${T}${tC}2:${tC},A17:A)))`);
  sh.getRange('D17').setFormula(`=ARRAYFORMULA(IF(A17:A="","",COUNTIFS(${T}${tC}2:${tC},A17:A,${T}${tS}2:${tS},"${CFG.STATUS.ISSUED}")))`);
  sh.getRange('E17').setFormula(`=ARRAYFORMULA(IF(A17:A="","",COUNTIF(${R}${rC}2:${rC},A17:A)))`);
  sh.setColumnWidth(1, 210).setColumnWidth(2, 240).setColumnWidths(3, 3, 170);
  ss.setActiveSheet(sh); ss.moveActiveSheet(1);
}

// ============================== قوالب الشهادات ==============================
function menuRecreateTemplates() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('إعادة إنشاء القوالب', 'سيتم إنشاء قالبين جديدين بالألوان الحالية في الإعدادات واعتمادهما. القوالب القديمة تبقى في Drive. متابعة؟', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const s = getSettings_();
  const folder = getRootFolder_(s);
  setSetting_(S.TPL_TRAINEE, buildTemplate_('trainee', s, folder));
  setSetting_(S.TPL_TRAINER, buildTemplate_('trainer', s, folder));
  ui.alert('✅ تم إنشاء القوالب الجديدة في مجلد الشهادات. يمكنك فتحها في Google Slides وتعديل التصميم مع الإبقاء على العبارات بين {{ }}.');
}

/**
 * يبني قالب شهادة في Google Slides. الحقول المتغيرة:
 * {{ORG}} {{INTRO}} {{NAME}} {{BODY}} {{WISH}} {{CERT_NO}} {{DATE}} {{SIGNER}} {{SIGNER_TITLE}} {{QR}}
 * {{COURSE}} {{HOURS}} {{START}} {{END}} {{GRADE}} {{TITLE}} متاحة أيضاً لمن يصمم قالبه الخاص.
 */
function buildTemplate_(type, s, folder) {
  const isTrainer = type === 'trainer';
  const primary = s[S.PRIMARY] || '#0B4F6C';
  const gold = s[S.SECONDARY] || '#C9A227';
  const pres = SlidesApp.create(isTrainer ? 'قالب شهادة المدربين' : 'قالب شهادة المتدربين');
  const slide = pres.getSlides()[0];
  slide.getPageElements().forEach(e => e.remove());
  const W = pres.getPageWidth(), Hh = pres.getPageHeight();

  slide.getBackground().setSolidFill('#FFFDF7');
  const outer = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 12, 12, W - 24, Hh - 24);
  outer.getFill().setTransparent(); outer.getBorder().setWeight(7).getLineFill().setSolidFill(primary);
  const inner = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 25, 25, W - 50, Hh - 50);
  inner.getFill().setTransparent(); inner.getBorder().setWeight(1.5).getLineFill().setSolidFill(gold);
  [[18, 18], [W - 32, 18], [18, Hh - 32], [W - 32, Hh - 32]].forEach(([x, y]) => {
    const d = slide.insertShape(SlidesApp.ShapeType.DIAMOND, x, y, 14, 14);
    d.getFill().setSolidFill(gold); d.getBorder().setTransparent();
  });

  addText_(slide, '{{ORG}}', 40, 36, W - 80, 24, { size: 14, bold: true, color: primary });
  addText_(slide, isTrainer ? 'شهادة شكر وتقدير' : 'شهادة إتمام دورة تدريبية', 40, 60, W - 80, 48,
    { size: 30, bold: true, color: primary, font: CFG.FONT_TITLE });
  const line = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, W / 2 - 90, 111, W / 2 + 90, 111);
  line.getLineFill().setSolidFill(gold); line.setWeight(2);
  addText_(slide, '{{INTRO}}', 40, 118, W - 80, 24, { size: 14, color: '#333333' });
  addText_(slide, '{{NAME}}', 40, 143, W - 80, 46, { size: 28, bold: true, color: gold, font: CFG.FONT_TITLE });
  addText_(slide, '{{BODY}}', 70, 190, W - 140, 60, { size: 13, color: '#333333' });
  addText_(slide, '{{WISH}}', 40, 250, W - 80, 22, { size: 12, color: primary });

  // QR + بيانات الشهادة (يسار)
  const qr = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 48, Hh - 118, 68, 68);
  qr.getFill().setSolidFill('#FFFFFF'); qr.getBorder().setWeight(0.75).getLineFill().setSolidFill('#BBBBBB');
  qr.getText().setText('{{QR}}'); qr.getText().getTextStyle().setFontSize(8).setForegroundColor('#999999');
  qr.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  addText_(slide, 'رقم الشهادة: {{CERT_NO}}\nتاريخ الإصدار: {{DATE}}', 120, Hh - 114, 190, 60, { size: 9, color: '#555555' });

  // ختم (وسط)
  const seal = slide.insertShape(SlidesApp.ShapeType.ELLIPSE, W / 2 - 34, Hh - 118, 68, 68);
  seal.getFill().setSolidFill(gold); seal.getBorder().setWeight(3).getLineFill().setSolidFill(primary);
  seal.getText().setText('معتمدة');
  seal.getText().getTextStyle().setFontFamily(CFG.FONT_BODY).setFontSize(12).setBold(true).setForegroundColor('#FFFFFF');
  seal.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  seal.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);

  // التوقيع (يمين)
  addText_(slide, '{{SIGNER_TITLE}}', W - 270, Hh - 118, 220, 22, { size: 12, bold: true, color: primary });
  const sl = slide.insertLine(SlidesApp.LineCategory.STRAIGHT, W - 240, Hh - 74, W - 80, Hh - 74);
  sl.getLineFill().setSolidFill('#999999'); sl.setWeight(1);
  addText_(slide, '{{SIGNER}}', W - 270, Hh - 70, 220, 22, { size: 12, color: '#333333' });

  pres.saveAndClose();
  const file = DriveApp.getFileById(pres.getId());
  file.moveTo(folder);
  return pres.getId();
}

function addText_(slide, text, x, y, w, h, o) {
  const box = slide.insertTextBox(text, x, y, w, h);
  const tr = box.getText();
  tr.getParagraphStyle().setParagraphAlignment(o.align || SlidesApp.ParagraphAlignment.CENTER);
  tr.getTextStyle().setFontFamily(o.font || CFG.FONT_BODY).setFontSize(o.size || 12)
    .setForegroundColor(o.color || '#333333').setBold(!!o.bold);
  box.setContentAlignment(SlidesApp.ContentAlignment.MIDDLE);
  return box;
}

// ============================== محرك الإصدار ==============================
function menuProcessNow() { processPending_({ ui: true }); }
function autoProcessTrigger() { processPending_({ ui: false }); }
function continueProcessing() { deleteTriggers_('continueProcessing'); processPending_({ ui: false }); }

function processPending_(opts) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { if (opts.ui) toast_('عملية إصدار أخرى قيد التشغيل، حاول بعد قليل.'); return; }
  const started = Date.now();
  const stats = { issued: 0, notEligible: 0, errors: 0, emailed: 0, remaining: 0 };
  try {
    const ctx = buildContext_();
    [['trainee', CFG.SHEETS.TRAINEES], ['trainer', CFG.SHEETS.TRAINERS]].forEach(([type, name]) => {
      const sh = ctx.ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return;
      const data = sh.getDataRange().getValues();
      const map = headerMap_(data[0]);
      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const status = String(row[map[H.STATUS]] || '').trim();
        const hasCert = String(row[map[H.CERT_NO]] || '').trim() !== '';
        const needsIssue = status === CFG.STATUS.READY && !hasCert;
        const needsEmail = status === CFG.STATUS.ISSUED && String(row[map[H.SENT]] || '').indexOf(CFG.SENT.PENDING) === 0;
        if (!needsIssue && !needsEmail) continue;
        if (Date.now() - started > CFG.MAX_RUNTIME_MS) { stats.remaining++; continue; }
        try {
          if (needsIssue) {
            const res = issueRow_(type, sh, r + 1, map, row, ctx, false);
            if (res === 'issued') stats.issued++; else if (res === 'not_eligible') stats.notEligible++;
            if (res === 'issued' && ctx.lastEmailOk) stats.emailed++;
          } else if (sendEmailForRow_(type, sh, r + 1, map, row, ctx)) stats.emailed++;
        } catch (err) {
          stats.errors++;
          writeRow_(sh, r + 1, map, { [H.STATUS]: CFG.STATUS.ERROR, [H.NOTES]: 'خطأ: ' + err.message });
        }
      }
    });
    if (stats.remaining > 0) scheduleContinuation_();
    updateQuotaCell_(ctx.ss);
  } finally {
    lock.releaseLock();
  }
  const msg = `صدرت: ${stats.issued} | أُرسلت بالبريد: ${stats.emailed} | غير مؤهل: ${stats.notEligible} | أخطاء: ${stats.errors}` +
    (stats.remaining ? ` | متبقٍ ${stats.remaining} سيُستكمل تلقائياً خلال دقيقة` : '');
  console.log(msg);
  if (opts.ui) SpreadsheetApp.getUi().alert('نتيجة الإصدار', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return stats;
}

function buildContext_() {
  const ss = getSS_();
  const s = getSettings_();
  if (!s[S.TPL_TRAINEE] || !s[S.TPL_TRAINER]) throw new Error('القوالب غير موجودة — شغّل «تهيئة النظام» أولاً.');
  return { ss, s, courses: getCourses_(ss), folder: getRootFolder_(s), lastEmailOk: false };
}

/** يصدر شهادة صف واحد. force = إعادة الإصدار بنفس الرقم ورمز التحقق */
function issueRow_(type, sh, rowNum, map, row, ctx, force) {
  const s = ctx.s;
  const rec = readRecord_(row, map);
  if (!rec.name) throw new Error('الاسم فارغ');
  const course = ctx.courses[rec.course];
  if (!course) throw new Error(`رمز الدورة «${rec.course || 'فارغ'}» غير موجود في ورقة الدورات`);

  let gradeLabel = '';
  if (type === 'trainee') {
    const el = checkEligibility_(rec, s);
    if (!el.ok) {
      writeRow_(sh, rowNum, map, { [H.STATUS]: CFG.STATUS.NOT_ELIGIBLE, [H.NOTES]: el.reason });
      return 'not_eligible';
    }
    gradeLabel = el.gradeLabel;
  }

  const female = isFemale_(rec.gender);
  const certNo = rec.certNo || nextCertNo_(s[S.PREFIX] || 'CERT', type);
  const code = rec.code || randomCode_(10);
  const verifyUrl = s[S.VERIFY_URL] ? `${String(s[S.VERIFY_URL]).trim()}?c=${code}` : '';
  const now = new Date();
  const issueDate = fmtDate_(now, s);

  const fields = buildTextFields_(type, female, rec, course, s, gradeLabel);
  fields['{{CERT_NO}}'] = certNo;
  fields['{{DATE}}'] = issueDate;

  // إنشاء نسخة من القالب وتعبئتها
  const tplId = type === 'trainer' ? s[S.TPL_TRAINER] : s[S.TPL_TRAINEE];
  const courseFolder = getOrCreateSubfolder_(ctx.folder, `${rec.course} - ${course.name}`);
  const copy = DriveApp.getFileById(tplId).makeCopy(`tmp_${certNo}`, courseFolder);
  const pres = SlidesApp.openById(copy.getId());
  Object.keys(fields).forEach(k => pres.replaceAllText(k, String(fields[k] == null ? '' : fields[k])));
  insertQr_(pres, verifyUrl || `${certNo} | ${code}`);
  pres.saveAndClose();

  const pdfBlob = copy.getAs(MimeType.PDF).setName(`${certNo} - ${rec.name}.pdf`);
  const pdfFile = courseFolder.createFile(pdfBlob);
  try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { console.warn('تعذر تغيير المشاركة: ' + e.message); }
  copy.setTrashed(true);

  // حذف الملف القديم عند إعادة الإصدار
  if (force && rec.link) { const oldId = fileIdFromUrl_(rec.link); if (oldId && oldId !== pdfFile.getId()) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (e) {} } }

  const link = pdfFile.getUrl();
  writeRow_(sh, rowNum, map, {
    [H.CERT_NO]: certNo, [H.CODE]: code, [H.LINK]: link, [H.DATE]: issueDate,
    [H.STATUS]: CFG.STATUS.ISSUED, [H.NOTES]: gradeLabel ? `التقدير: ${gradeLabel}` : ''
  });
  upsertLog_(ctx.ss, [certNo, code, type === 'trainer' ? (female ? 'مدربة' : 'مدرب') : (female ? 'متدربة' : 'متدرب'),
    rec.name, rec.gender, rec.email, rec.course, course.name, course.hours, issueDate, link, 'سارية', '']);

  // البريد
  ctx.lastEmailOk = false;
  const sent = deliverEmail_(type, female, rec, course, certNo, code, verifyUrl, pdfBlob, link, s);
  writeRow_(sh, rowNum, map, { [H.SENT]: sent.label });
  ctx.lastEmailOk = sent.ok;
  return 'issued';
}

function buildTextFields_(type, female, rec, course, s, gradeLabel) {
  const org = s[S.ORG] || '';
  const start = fmtDate_(course.start, s, true), end = fmtDate_(course.end, s, true);
  const hours = course.hours || '';
  const period = start && end ? ` المنعقدة خلال الفترة من ${start} إلى ${end}` : (start ? ` المنعقدة بتاريخ ${start}` : '');
  const hoursTxt = hours ? ` بواقع (${hours}) ${Number(hours) >= 3 && Number(hours) <= 10 ? 'ساعات' : 'ساعة'} تدريبية` : '';
  let intro, body, wish, title;
  if (type === 'trainer') {
    title = female ? 'المدربة' : 'المدرب';
    intro = `${String(s[S.INTRO_TRAINER] || '').replace(/{{ORG}}/g, org)} ${title}`;
    body = `وذلك تقديراً لجهود${female ? 'ها' : 'ه'} المتميزة في تقديم الدورة التدريبية «${course.name}»${period}${hoursTxt}`;
    wish = `سائلين الله ل${female ? 'ها' : 'ه'} دوام التوفيق والسداد`;
  } else {
    title = female ? 'المتدربة' : 'المتدرب';
    intro = `${String(s[S.INTRO_TRAINEE] || '').replace(/{{ORG}}/g, org)} ${title}`;
    body = `قد ${female ? 'اجتازت' : 'اجتاز'} بنجاح الدورة التدريبية «${course.name}»${period}${hoursTxt}` +
      (gradeLabel ? ` بتقدير «${gradeLabel}»` : '');
    wish = `متمنين ل${female ? 'ها' : 'ه'} دوام التوفيق والنجاح`;
  }
  return {
    '{{ORG}}': org, '{{INTRO}}': intro, '{{TITLE}}': title, '{{NAME}}': rec.name, '{{BODY}}': body, '{{WISH}}': wish,
    '{{COURSE}}': course.name, '{{HOURS}}': hours, '{{START}}': start, '{{END}}': end, '{{GRADE}}': gradeLabel,
    '{{SIGNER}}': s[S.SIGNER] || '', '{{SIGNER_TITLE}}': s[S.SIGNER_TITLE] || ''
  };
}

function insertQr_(pres, data) {
  let blob = null;
  const urls = [
    'https://quickchart.io/qr?size=300&margin=1&text=' + encodeURIComponent(data),
    'https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=4&data=' + encodeURIComponent(data)
  ];
  for (const u of urls) {
    try { const res = UrlFetchApp.fetch(u, { muteHttpExceptions: true }); if (res.getResponseCode() === 200) { blob = res.getBlob(); break; } } catch (e) {}
  }
  pres.getSlides().forEach(slide => {
    slide.getShapes().forEach(shape => {
      let txt = '';
      try { txt = shape.getText().asString(); } catch (e) { return; }
      if (txt.indexOf('{{QR}}') === -1) return;
      if (blob) {
        slide.insertImage(blob, shape.getLeft(), shape.getTop(), shape.getWidth(), shape.getHeight());
        shape.remove();
      } else {
        shape.getText().setText('');
      }
    });
  });
}

// ============================== البريد الإلكتروني ==============================
function deliverEmail_(type, female, rec, course, certNo, code, verifyUrl, pdfBlob, link, s) {
  if (String(s[S.AUTO_EMAIL]).trim() !== 'نعم') return { ok: false, label: CFG.SENT.DISABLED };
  if (!isEmail_(rec.email)) return { ok: false, label: CFG.SENT.NO_EMAIL };
  if (MailApp.getRemainingDailyQuota() < 1) return { ok: false, label: CFG.SENT.PENDING + ' (تجاوز الحد اليومي)' };
  try {
    sendCertificateEmail_(type, female, rec, course, certNo, code, verifyUrl, pdfBlob, link, s);
    return { ok: true, label: 'أُرسلت ' + Utilities.formatDate(new Date(), CFG.TZ, 'yyyy/MM/dd HH:mm') };
  } catch (e) {
    return { ok: false, label: CFG.SENT.PENDING + ' (فشل: ' + e.message + ')' };
  }
}

function sendCertificateEmail_(type, female, rec, course, certNo, code, verifyUrl, pdfBlob, link, s) {
  const org = s[S.ORG] || '';
  const primary = s[S.PRIMARY] || '#0B4F6C', gold = s[S.SECONDARY] || '#C9A227';
  const dear = female ? 'عزيزتي' : 'عزيزي';
  const you = female ? 'كِ' : 'ك';
  const msg = type === 'trainer'
    ? `يسرّنا أن نتقدم ل${you} بخالص الشكر والتقدير على تقديم${you} الدورة التدريبية <b>«${esc_(course.name)}»</b>، ونرفق ل${you} شهادة الشكر والتقدير.`
    : `يسرّنا أن نهنئ${you} على اجتياز${you} الدورة التدريبية <b>«${esc_(course.name)}»</b> بنجاح، ونرفق ل${you} شهادة إتمام الدورة.`;
  const html = `
  <div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;background:#f4f6f8;padding:24px">
    <div style="max-width:560px;margin:auto;background:#fff;border-radius:12px;overflow:hidden;border:1px solid #e3e7ea">
      <div style="background:${primary};color:#fff;padding:20px 24px;font-size:20px;font-weight:bold">${esc_(org)}</div>
      <div style="padding:24px;color:#333;line-height:1.9;font-size:15px">
        <p>${dear} <b>${esc_(rec.name)}</b>،</p>
        <p>${msg}</p>
        <table style="width:100%;border-collapse:collapse;margin:16px 0;font-size:14px">
          <tr><td style="padding:8px;background:#f7f7f7;width:40%">رقم الشهادة</td><td style="padding:8px"><b>${esc_(certNo)}</b></td></tr>
          <tr><td style="padding:8px;background:#f7f7f7">رمز التحقق</td><td style="padding:8px">${esc_(code)}</td></tr>
          ${course.hours ? `<tr><td style="padding:8px;background:#f7f7f7">عدد الساعات</td><td style="padding:8px">${esc_(course.hours)}</td></tr>` : ''}
        </table>
        <p style="text-align:center;margin:24px 0">
          <a href="${link}" style="background:${gold};color:#fff;text-decoration:none;padding:12px 28px;border-radius:8px;font-weight:bold">تحميل الشهادة</a>
          ${verifyUrl ? `&nbsp; <a href="${verifyUrl}" style="background:${primary};color:#fff;text-decoration:none;padding:12px 28px;border-radius:8px;font-weight:bold">التحقق من الشهادة</a>` : ''}
        </p>
        <p>مع أطيب التمنيات،<br>${esc_(org)}</p>
      </div>
    </div>
  </div>`;
  const opt = { to: rec.email, subject: `${type === 'trainer' ? 'شهادة شكر وتقدير' : 'شهادة إتمام'}: ${course.name} — ${org}`,
    htmlBody: html, attachments: [pdfBlob], name: s[S.SENDER_NAME] || org };
  if (isEmail_(s[S.REPLY_TO])) opt.replyTo = s[S.REPLY_TO];
  MailApp.sendEmail(opt);
}

function sendEmailForRow_(type, sh, rowNum, map, row, ctx) {
  const rec = readRecord_(row, map);
  const course = ctx.courses[rec.course] || { name: rec.course, hours: '' };
  const id = fileIdFromUrl_(rec.link);
  if (!id) throw new Error('رابط الشهادة غير صالح — أعد إصدارها');
  const blob = DriveApp.getFileById(id).getBlob();
  const verifyUrl = ctx.s[S.VERIFY_URL] ? `${String(ctx.s[S.VERIFY_URL]).trim()}?c=${rec.code}` : '';
  const res = deliverEmail_(type, isFemale_(rec.gender), rec, course, rec.certNo, rec.code, verifyUrl, blob, rec.link,
    Object.assign({}, ctx.s, { [S.AUTO_EMAIL]: 'نعم' }));
  writeRow_(sh, rowNum, map, { [H.SENT]: res.label });
  return res.ok;
}

// ============================== إجراءات الصف المحدد ==============================
function menuIssueSelected() { selectedAction_('issue'); }
function menuReissueSelected() { selectedAction_('reissue'); }
function menuResendSelected() { selectedAction_('resend'); }
function menuRevokeSelected() { selectedAction_('revoke'); }

function selectedAction_(action) {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const name = sh.getName();
  const type = name === CFG.SHEETS.TRAINEES ? 'trainee' : name === CFG.SHEETS.TRAINERS ? 'trainer' : null;
  if (!type) return ui.alert(`انتقل إلى ورقة «${CFG.SHEETS.TRAINEES}» أو «${CFG.SHEETS.TRAINERS}» وحدد صفاً أو أكثر.`);
  const range = sh.getActiveRange();
  const first = Math.max(2, range.getRow()), last = range.getLastRow();
  if (last < 2) return ui.alert('حدد صفاً يحتوي على بيانات (ليس صف العناوين).');

  if (action === 'revoke' && ui.alert('تأكيد الإلغاء', `سيتم إلغاء ${last - first + 1} شهادة، وستظهر «ملغاة» في صفحة التحقق. متابعة؟`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return ui.alert('عملية أخرى قيد التشغيل، حاول بعد قليل.');
  const out = [];
  try {
    const ctx = buildContext_();
    const data = sh.getDataRange().getValues();
    const map = headerMap_(data[0]);
    for (let r = first; r <= last && r <= data.length; r++) {
      const row = data[r - 1];
      if (!String(row[map[H.NAME]] || '').trim()) continue;
      out.push(`${String(row[map[H.NAME]]).trim()}: ${rowAction_(action, type, sh, r, map, row, ctx).msg}`);
    }
  } finally { lock.releaseLock(); }
  ui.alert('النتيجة', out.join('\n') || 'لا توجد صفوف صالحة.', ui.ButtonSet.OK);
}

/** إجراء على صف واحد — يستخدمه الجدول والمنصة معاً */
function rowAction_(action, type, sh, r, map, row, ctx) {
  const rec = readRecord_(row, map);
  try {
    if (action === 'issue') {
      if (rec.certNo) return { ok: false, msg: 'لديه شهادة مسبقاً (استخدم إعادة الإصدار)' };
      const res = issueRow_(type, sh, r, map, row, ctx, false);
      return res === 'issued' ? { ok: true, msg: 'صدرت الشهادة ✅' } : { ok: false, msg: 'غير مؤهل ⚠️ — ' + (sh.getRange(r, map[H.NOTES] + 1).getValue() || '') };
    }
    if (action === 'reissue') {
      const res = issueRow_(type, sh, r, map, row, ctx, true);
      return res === 'issued' ? { ok: true, msg: 'أعيد إصدار الشهادة ✅' } : { ok: false, msg: 'غير مؤهل ⚠️' };
    }
    if (action === 'resend') {
      if (!rec.link) return { ok: false, msg: 'لا توجد شهادة' };
      return sendEmailForRow_(type, sh, r, map, row, ctx) ? { ok: true, msg: 'أُرسلت بالبريد ✉️' } : { ok: false, msg: 'لم تُرسل (تحقق من البريد)' };
    }
    if (action === 'revoke') {
      if (!rec.certNo) return { ok: false, msg: 'لا توجد شهادة' };
      writeRow_(sh, r, map, { [H.STATUS]: CFG.STATUS.REVOKED, [H.NOTES]: 'أُلغيت ' + fmtDate_(new Date(), ctx.s, true) });
      setLogStatus_(ctx.ss, rec.certNo, 'ملغاة');
      return { ok: true, msg: 'أُلغيت الشهادة ⛔' };
    }
    return { ok: false, msg: 'إجراء غير معروف' };
  } catch (err) {
    writeRow_(sh, r, map, { [H.NOTES]: 'خطأ: ' + err.message });
    return { ok: false, msg: 'خطأ — ' + err.message };
  }
}

// ============================== الأتمتة ==============================
function enableAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  disableAutomation(true);
  ScriptApp.newTrigger('onEditInstalled').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('autoProcessTrigger').timeBased().everyHours(1).create();
  SpreadsheetApp.getUi().alert('✅ تم تفعيل الإصدار التلقائي',
    '• تصدر الشهادة فور تغيير «الحالة» إلى «مكتمل».\n• فحص احتياطي كل ساعة لأي شهادات أو رسائل معلّقة.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function disableAutomation(silent) {
  ['onEditInstalled', 'autoProcessTrigger', 'continueProcessing'].forEach(deleteTriggers_);
  if (silent !== true) SpreadsheetApp.getUi().alert('⏸️ تم إيقاف الإصدار التلقائي. يمكنك الإصدار يدوياً من القائمة.');
}

/** مشغّل مُثبّت: يصدر الشهادة عند تغيير الحالة إلى «مكتمل» */
function onEditInstalled(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (name !== CFG.SHEETS.TRAINEES && name !== CFG.SHEETS.TRAINERS) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf(H.STATUS) + 1;
  if (statusCol < e.range.getColumn() || statusCol > e.range.getLastColumn()) return;
  const vals = sh.getRange(e.range.getRow(), statusCol, e.range.getNumRows()).getValues().flat();
  if (vals.some(v => String(v).trim() === CFG.STATUS.READY)) processPending_({ ui: false });
}

/** مشغّل بسيط: ترقيم تلقائي للمعرف + حالة افتراضية عند إدخال اسم جديد */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (name !== CFG.SHEETS.TRAINEES && name !== CFG.SHEETS.TRAINERS) return;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);
  if (map[H.NAME] === undefined) return;
  const prefix = name === CFG.SHEETS.TRAINEES ? 'ST-' : 'TR-';
  for (let r = Math.max(2, e.range.getRow()); r <= e.range.getLastRow(); r++) {
    const rowVals = sh.getRange(r, 1, 1, headers.length).getValues()[0];
    if (!String(rowVals[map[H.NAME]]).trim()) continue;
    if (!String(rowVals[map[H.ID]]).trim()) sh.getRange(r, map[H.ID] + 1).setValue(prefix + pad_(r - 1, 4));
    if (!String(rowVals[map[H.STATUS]]).trim()) sh.getRange(r, map[H.STATUS] + 1).setValue(CFG.STATUS.IN_PROGRESS);
  }
}

function scheduleContinuation_() {
  deleteTriggers_('continueProcessing');
  ScriptApp.newTrigger('continueProcessing').timeBased().after(60 * 1000).create();
}

function deleteTriggers_(fn) {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t); });
}

// ============================== صفحة التحقق (تطبيق ويب) ==============================
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api) return json_(handleApi_(Object.assign({}, p, { action: p.api })));
  const q = String(p.c || p.id || '').trim();
  const t = HtmlService.createTemplateFromFile('Verify');
  let org = '';
  try { org = getSettings_()[S.ORG] || ''; } catch (err) {}
  t.org = org;
  t.initialJson = JSON.stringify({ q, result: q ? lookupCertificate_(q) : null }).replace(/</g, '\\u003c');
  return t.evaluate()
    .setTitle('التحقق من الشهادات' + (org ? ' — ' + org : ''))
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** تُستدعى من صفحة التحقق */
function verifyCertificate(q) { return lookupCertificate_(String(q || '').trim()); }

// ============================== واجهة API لمنصة GitHub ==============================
/*
 * المنصة (index.html على GitHub Pages) تتصل بهذا الكود عبر رابط تطبيق الويب.
 * رموز الدخول محفوظة مشفّرة (SHA-256 + salt) في خصائص السكربت وليست في الصفحة.
 * الافتراضي: المشرف العام 123123 — المدير التنفيذي 123456 (غيّرها من القائمة).
 */
const ROLES = {
  admin: { name: 'المشرف العام', defaultPin: '123123' },     // صلاحيات كاملة
  ceo:   { name: 'المدير التنفيذي', defaultPin: '123456' }    // اطلاع وتقارير فقط
};
const SESSION_SECONDS = 6 * 60 * 60;

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'طلب غير صالح' }); }
  return json_(handleApi_(body));
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handleApi_(b) {
  try {
    const a = String(b.action || '');
    if (a === 'info') return apiInfo_();
    if (a === 'verify') return Object.assign({ ok: true }, lookupCertificate_(String(b.c || b.q || '').trim()));
    if (a === 'login') return apiLogin_(b);

    const sess = getSession_(b.token);
    if (!sess) return { ok: false, auth: false, error: 'انتهت الجلسة، سجّل الدخول مجدداً' };
    if (a === 'logout') { CacheService.getScriptCache().remove('S_' + b.token); return { ok: true }; }
    if (a === 'dashboard') return apiDashboard_();
    if (a === 'list') return apiList_(b, sess);
    if (a === 'courses') return { ok: true, courses: apiCourses_() };

    if (sess.role !== 'admin') return { ok: false, error: 'هذا الإجراء متاح للمشرف العام فقط' };
    if (a === 'addPerson') return apiAddPerson_(b);
    if (a === 'updatePerson') return apiUpdatePerson_(b);
    if (a === 'addCourse') return apiAddCourse_(b);
    if (['issue', 'reissue', 'resend', 'revoke'].indexOf(a) > -1) return apiRowAction_(b, a);
    if (a === 'processAll') return Object.assign({ ok: true }, processPending_({ ui: false }) || {});
    if (a === 'changePin') {
      if (!ROLES[b.role]) return { ok: false, error: 'دور غير معروف' };
      if (!/^\S{6,32}$/.test(String(b.newPin || ''))) return { ok: false, error: 'الرمز يجب أن يكون 6 خانات على الأقل' };
      setPin_(b.role, String(b.newPin));
      return { ok: true, msg: 'تم تغيير رمز ' + ROLES[b.role].name };
    }
    return { ok: false, error: 'إجراء غير معروف' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ---------- المستخدمون والجلسات ----------
function ensureUsers_() {
  const p = PropertiesService.getScriptProperties();
  Object.keys(ROLES).forEach(r => { if (!p.getProperty('USER_' + r)) setPin_(r, ROLES[r].defaultPin); });
}
function setPin_(role, pin) {
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('USER_' + role, JSON.stringify({ salt, hash: sha_(salt + pin) }));
}
function sha_(s) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8));
}
function checkPin_(role, pin) {
  ensureUsers_();
  const raw = PropertiesService.getScriptProperties().getProperty('USER_' + role);
  if (!raw) return false;
  const u = JSON.parse(raw);
  return u.hash === sha_(u.salt + String(pin));
}
function apiLogin_(b) {
  const role = String(b.role || '');
  if (!ROLES[role]) return { ok: false, error: 'اختر نوع المستخدم' };
  const cache = CacheService.getScriptCache();
  const fk = 'F_' + role, fails = Number(cache.get(fk) || 0);
  if (fails >= 5) return { ok: false, error: 'تم إيقاف الدخول مؤقتاً بسبب محاولات خاطئة متكررة. حاول بعد 15 دقيقة.' };
  if (!checkPin_(role, String(b.pin || ''))) {
    cache.put(fk, String(fails + 1), 15 * 60);
    return { ok: false, error: `رمز الدخول غير صحيح (متبقٍ ${4 - fails} محاولات)` };
  }
  cache.remove(fk);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const sess = { role, name: ROLES[role].name };
  cache.put('S_' + token, JSON.stringify(sess), SESSION_SECONDS);
  return { ok: true, token, role, name: sess.name, org: getSettings_()[S.ORG] || '' };
}
function getSession_(token) {
  if (!token || String(token).length < 40) return null;
  const raw = CacheService.getScriptCache().get('S_' + token);
  if (!raw) return null;
  CacheService.getScriptCache().put('S_' + token, raw, SESSION_SECONDS); // تمديد الجلسة
  return JSON.parse(raw);
}

function menuChangePin() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('تغيير رمز الدخول', 'اكتب 1 للمشرف العام أو 2 للمدير التنفيذي:', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const role = r1.getResponseText().trim() === '2' ? 'ceo' : r1.getResponseText().trim() === '1' ? 'admin' : '';
  if (!role) return ui.alert('اختيار غير صحيح.');
  const r2 = ui.prompt('الرمز الجديد لـ ' + ROLES[role].name, '6 خانات على الأقل:', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const pin = r2.getResponseText().trim();
  if (!/^\S{6,32}$/.test(pin)) return ui.alert('الرمز يجب أن يكون 6 خانات على الأقل بدون مسافات.');
  setPin_(role, pin);
  ui.alert('✅ تم تغيير رمز دخول ' + ROLES[role].name);
}

// ---------- بيانات المنصة ----------
const API_FIELDS = {
  id: H.ID, name: H.NAME, gender: H.GENDER, nationalId: H.NATIONAL_ID, email: H.EMAIL, phone: H.PHONE,
  course: H.COURSE, attend: H.ATTEND, grade: H.GRADE, specialty: H.SPECIALTY, status: H.STATUS,
  certNo: H.CERT_NO, code: H.CODE, link: H.LINK, date: H.DATE, sent: H.SENT, notes: H.NOTES
};

function kindSheet_(kind) {
  const name = kind === 'trainers' ? CFG.SHEETS.TRAINERS : CFG.SHEETS.TRAINEES;
  const sh = getSS_().getSheetByName(name);
  if (!sh) throw new Error('شغّل «تهيئة النظام» من الجدول أولاً');
  return { sh, type: kind === 'trainers' ? 'trainer' : 'trainee', prefix: kind === 'trainers' ? 'TR-' : 'ST-' };
}

function rowsAsObjects_(sh) {
  if (sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const map = headerMap_(data[0]);
  return data.slice(1).map((r, i) => {
    const o = { _row: i + 2 };
    Object.keys(API_FIELDS).forEach(k => {
      const idx = map[API_FIELDS[k]];
      let v = idx === undefined ? '' : r[idx];
      if (v instanceof Date) v = Utilities.formatDate(v, CFG.TZ, 'yyyy/MM/dd');
      o[k] = v === null ? '' : v;
    });
    return o;
  }).filter(o => String(o.name).trim());
}

function apiInfo_() {
  const ss = getSS_();
  const s = getSettings_();
  const log = ss.getSheetByName(CFG.SHEETS.LOG);
  const issued = log && log.getLastRow() > 1
    ? log.getRange(2, 12, log.getLastRow() - 1, 1).getValues().flat().filter(v => v === 'سارية').length : 0;
  return {
    ok: true, org: s[S.ORG] || '',
    stats: {
      issued,
      trainees: rowsAsObjects_(ss.getSheetByName(CFG.SHEETS.TRAINEES)).length,
      trainers: rowsAsObjects_(ss.getSheetByName(CFG.SHEETS.TRAINERS)).length,
      courses: Object.keys(getCourses_(ss)).length
    }
  };
}

function apiCourses_() {
  const c = getCourses_(getSS_());
  return Object.keys(c).map(code => ({
    code, name: c[code].name, hours: c[code].hours, place: c[code].place, trainer: c[code].trainer,
    start: c[code].start instanceof Date ? Utilities.formatDate(c[code].start, CFG.TZ, 'yyyy/MM/dd') : String(c[code].start || ''),
    end: c[code].end instanceof Date ? Utilities.formatDate(c[code].end, CFG.TZ, 'yyyy/MM/dd') : String(c[code].end || '')
  }));
}

function apiDashboard_() {
  const ss = getSS_();
  const tr = rowsAsObjects_(ss.getSheetByName(CFG.SHEETS.TRAINEES));
  const tn = rowsAsObjects_(ss.getSheetByName(CFG.SHEETS.TRAINERS));
  const summarize = list => {
    const o = { total: list.length, male: 0, female: 0 };
    Object.values(CFG.STATUS).forEach(st => o[st] = 0);
    list.forEach(p => { isFemale_(p.gender) ? o.female++ : o.male++; if (o[p.status] !== undefined) o[p.status]++; });
    return o;
  };
  const courses = apiCourses_().map(c => ({
    code: c.code, name: c.name, hours: c.hours,
    trainees: tr.filter(p => p.course === c.code).length,
    issued: tr.filter(p => p.course === c.code && p.status === CFG.STATUS.ISSUED).length,
    trainers: tn.filter(p => p.course === c.code).length
  }));
  const log = ss.getSheetByName(CFG.SHEETS.LOG);
  let recent = [];
  if (log && log.getLastRow() > 1) {
    recent = log.getRange(2, 1, log.getLastRow() - 1, 12).getValues()
      .map(r => ({ certNo: r[0], type: r[2], name: r[3], course: r[7], date: r[9], status: r[11] }))
      .reverse().slice(0, 10);
  }
  let quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (e) {}
  return { ok: true, statuses: CFG.STATUS, trainees: summarize(tr), trainers: summarize(tn), courses, recent, quota };
}

function apiList_(b, sess) {
  const { sh } = kindSheet_(b.kind);
  const q = String(b.q || '').trim().toLowerCase();
  let rows = rowsAsObjects_(sh);
  if (b.status) rows = rows.filter(r => r.status === b.status);
  if (b.course) rows = rows.filter(r => r.course === b.course);
  if (q) rows = rows.filter(r => [r.id, r.name, r.email, r.certNo, r.phone, r.nationalId].some(v => String(v).toLowerCase().indexOf(q) > -1));
  if (sess.role !== 'admin') rows.forEach(r => { if (r.nationalId) r.nationalId = '••••' + String(r.nationalId).slice(-3); });
  return { ok: true, total: rows.length, rows: rows.slice(0, 500) };
}

function findRowById_(sh, id) {
  const data = sh.getDataRange().getValues();
  const map = headerMap_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][map[H.ID]]).trim() === String(id).trim()) return { r: i + 1, row: data[i], map };
  }
  return null;
}

function apiAddPerson_(b) {
  const { sh, prefix } = kindSheet_(b.kind);
  const d = b.data || {};
  if (!String(d.name || '').trim()) return { ok: false, error: 'الاسم مطلوب' };
  if (['ذكر', 'أنثى'].indexOf(d.gender) < 0) return { ok: false, error: 'اختر الجنس' };
  if (!getCourses_(getSS_())[String(d.course || '').trim()]) return { ok: false, error: 'رمز الدورة غير موجود' };
  if (d.email && !isEmail_(d.email)) return { ok: false, error: 'البريد الإلكتروني غير صحيح' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const map = headerMap_(headers);
    const ids = sh.getLastRow() > 1 ? sh.getRange(2, map[H.ID] + 1, sh.getLastRow() - 1, 1).getValues().flat() : [];
    const maxN = ids.reduce((m, v) => Math.max(m, parseInt(String(v).replace(/\D/g, ''), 10) || 0), 0);
    const id = prefix + pad_(maxN + 1, 4);
    const row = new Array(headers.length).fill('');
    const put = (h, v) => { if (map[h] !== undefined && v !== undefined) row[map[h]] = v; };
    put(H.ID, id); put(H.NAME, String(d.name).trim()); put(H.GENDER, d.gender); put(H.EMAIL, d.email || '');
    put(H.PHONE, d.phone || ''); put(H.COURSE, String(d.course).trim()); put(H.NATIONAL_ID, d.nationalId || '');
    put(H.SPECIALTY, d.specialty || '');
    if (d.attend !== '' && d.attend !== undefined) put(H.ATTEND, Number(d.attend));
    if (d.grade !== '' && d.grade !== undefined) put(H.GRADE, Number(d.grade));
    put(H.STATUS, CFG.STATUS.IN_PROGRESS);
    sh.appendRow(row);
    return { ok: true, id, msg: 'تمت الإضافة برقم ' + id };
  } finally { lock.releaseLock(); }
}

function apiUpdatePerson_(b) {
  const { sh, type } = kindSheet_(b.kind);
  const f = findRowById_(sh, b.id);
  if (!f) return { ok: false, error: 'السجل غير موجود' };
  const d = b.data || {};
  const upd = {};
  if (d.attend !== undefined) upd[H.ATTEND] = d.attend === '' ? '' : Number(d.attend);
  if (d.grade !== undefined) upd[H.GRADE] = d.grade === '' ? '' : Number(d.grade);
  if (d.email !== undefined) { if (d.email && !isEmail_(d.email)) return { ok: false, error: 'البريد غير صحيح' }; upd[H.EMAIL] = d.email; }
  if (d.phone !== undefined) upd[H.PHONE] = d.phone;
  if (d.status !== undefined) {
    if ([CFG.STATUS.IN_PROGRESS, CFG.STATUS.READY].indexOf(d.status) < 0) return { ok: false, error: 'حالة غير مسموحة' };
    upd[H.STATUS] = d.status;
  }
  writeRow_(sh, f.r, f.map, upd);
  if (d.status === CFG.STATUS.READY && !String(f.row[f.map[H.CERT_NO]] || '').trim()) {
    return apiRowAction_({ kind: b.kind, id: b.id }, 'issue');
  }
  return { ok: true, msg: 'تم الحفظ' };
}

function apiAddCourse_(b) {
  const d = b.data || {};
  const code = String(d.code || '').trim();
  if (!code || !String(d.name || '').trim()) return { ok: false, error: 'رمز واسم الدورة مطلوبان' };
  const ss = getSS_();
  if (getCourses_(ss)[code]) return { ok: false, error: 'رمز الدورة مستخدم مسبقاً' };
  const toDate = v => (v ? new Date(v) : '');
  ss.getSheetByName(CFG.SHEETS.COURSES).appendRow([code, String(d.name).trim(), toDate(d.start), toDate(d.end),
    d.hours ? Number(d.hours) : '', d.place || '', d.trainer || '']);
  return { ok: true, msg: 'تمت إضافة الدورة ' + code };
}

function apiRowAction_(b, action) {
  const { sh, type } = kindSheet_(b.kind);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, error: 'عملية أخرى قيد التشغيل، حاول بعد قليل' };
  try {
    const f = findRowById_(sh, b.id);
    if (!f) return { ok: false, error: 'السجل غير موجود' };
    if (action === 'issue') writeRow_(sh, f.r, f.map, { [H.STATUS]: CFG.STATUS.READY });
    const fresh = sh.getRange(f.r, 1, 1, sh.getLastColumn()).getValues()[0];
    const res = rowAction_(action, type, sh, f.r, f.map, fresh, buildContext_());
    return res.ok ? { ok: true, msg: res.msg } : { ok: false, error: res.msg };
  } finally { lock.releaseLock(); }
}

// ============================== السجل (بحث) ==============================
function lookupCertificate_(q) {
  if (!q || q.length > 60) return { found: false };
  const sh = getSS_().getSheetByName(CFG.SHEETS.LOG);
  if (!sh || sh.getLastRow() < 2) return { found: false };
  const data = sh.getDataRange().getValues();
  const key = q.toUpperCase();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[0]).toUpperCase() === key || String(r[1]).toUpperCase() === key) {
      return {
        found: true, certNo: String(r[0]), type: String(r[2]), name: String(r[3]),
        course: String(r[7]), hours: String(r[8]), date: String(r[9]),
        status: String(r[11]), valid: String(r[11]) === 'سارية',
        revokedAt: r[12] ? String(r[12]) : '', link: String(r[11]) === 'سارية' ? String(r[10]) : ''
      };
    }
  }
  return { found: false };
}

// ============================== السجل ==============================
function upsertLog_(ss, rowValues) {
  const sh = ss.getSheetByName(CFG.SHEETS.LOG);
  const last = sh.getLastRow();
  if (last >= 2) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().flat().map(String);
    const idx = ids.indexOf(String(rowValues[0]));
    if (idx > -1) { sh.getRange(idx + 2, 1, 1, rowValues.length).setValues([rowValues]); return; }
  }
  sh.appendRow(rowValues);
}

function setLogStatus_(ss, certNo, status) {
  const sh = ss.getSheetByName(CFG.SHEETS.LOG);
  if (sh.getLastRow() < 2) return;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(String);
  const idx = ids.indexOf(String(certNo));
  if (idx < 0) return;
  const c = HEADERS.LOG.indexOf('الحالة') + 1;
  sh.getRange(idx + 2, c, 1, 2).setValues([[status, status === 'ملغاة' ? Utilities.formatDate(new Date(), CFG.TZ, 'yyyy/MM/dd') : '']]);
}

// ============================== أدوات مساعدة ==============================
function getSS_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID'));
  return ss;
}

function getSettings_() {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(([k, v]) => { if (k) out[String(k).trim()] = v; });
  return out;
}

function findSettingRow_(key) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  if (!sh || sh.getLastRow() < 2) return 0;
  const keys = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(x => String(x).trim());
  const i = keys.indexOf(key);
  return i < 0 ? 0 : i + 2;
}

function setSetting_(key, value) {
  const sh = getSS_().getSheetByName(CFG.SHEETS.SETTINGS);
  const r = findSettingRow_(key);
  if (r) sh.getRange(r, 2).setValue(value); else sh.appendRow([key, value, '']);
}

function getCourses_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.COURSES);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach(r => {
    const code = String(r[0]).trim();
    if (code) out[code] = { name: String(r[1]).trim(), start: r[2], end: r[3], hours: r[4], place: r[5], trainer: r[6] };
  });
  return out;
}

function getRootFolder_(s) {
  if (s[S.FOLDER_ID]) { try { return DriveApp.getFolderById(s[S.FOLDER_ID]); } catch (e) {} }
  const ss = getSS_();
  const folder = DriveApp.createFolder('شهادات ' + (s[S.ORG] || 'الجهة') + ' - ' + ss.getName());
  setSetting_(S.FOLDER_ID, folder.getId());
  s[S.FOLDER_ID] = folder.getId();
  return folder;
}

function getOrCreateSubfolder_(parent, name) {
  const clean = name.replace(/[\/\\]/g, '-').slice(0, 120);
  const it = parent.getFoldersByName(clean);
  return it.hasNext() ? it.next() : parent.createFolder(clean);
}

function safeFile_(id) { try { return !!DriveApp.getFileById(id); } catch (e) { return false; } }

function headerMap_(headers) {
  const m = {};
  headers.forEach((h, i) => { const k = String(h).trim(); if (k) m[k] = i; });
  return m;
}

function readRecord_(row, map) {
  const g = k => (map[k] === undefined ? '' : row[map[k]]);
  return {
    id: String(g(H.ID)).trim(), name: String(g(H.NAME)).trim(), gender: String(g(H.GENDER)).trim(),
    email: String(g(H.EMAIL)).trim(), course: String(g(H.COURSE)).trim(),
    attend: g(H.ATTEND), grade: g(H.GRADE),
    certNo: String(g(H.CERT_NO)).trim(), code: String(g(H.CODE)).trim(), link: String(g(H.LINK)).trim()
  };
}

function writeRow_(sh, rowNum, map, obj) {
  Object.keys(obj).forEach(k => { if (map[k] !== undefined) sh.getRange(rowNum, map[k] + 1).setValue(obj[k]); });
}

function checkEligibility_(rec, s) {
  const minA = toNum_(s[S.MIN_ATTEND]);
  const minG = toNum_(s[S.MIN_GRADE]);
  let att = toNum_(rec.attend);
  if (att !== null && att > 0 && att <= 1) att = att * 100;   // في حال إدخالها كنسبة مئوية 0.95
  if (minA) {
    if (att === null) return { ok: false, reason: 'لم تُسجّل نسبة الحضور' };
    if (att < minA) return { ok: false, reason: `نسبة الحضور ${Math.round(att)}% أقل من الحد الأدنى ${minA}%` };
  }
  const g = toNum_(rec.grade);
  if (minG && g !== null && g < minG) return { ok: false, reason: `الدرجة ${g} أقل من الحد الأدنى ${minG}` };
  if (minG && g === null) return { ok: false, reason: 'لم تُسجّل الدرجة' };
  return { ok: true, gradeLabel: g === null ? '' : gradeLabel_(g) };
}

function gradeLabel_(g) {
  if (g >= 90) return 'ممتاز';
  if (g >= 80) return 'جيد جداً';
  if (g >= 70) return 'جيد';
  if (g >= 60) return 'مقبول';
  return '';
}

function nextCertNo_(prefix, type) {
  const props = PropertiesService.getScriptProperties();
  const year = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy');
  const key = `SEQ_${year}`;
  const n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return `${String(prefix).trim().toUpperCase()}-${year}-${type === 'trainer' ? 'TR' : 'ST'}-${pad_(n, 5)}`;
}

function randomCode_(len) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let s = '';
  for (let i = 0; i < len; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function fmtDate_(d, s, short) {
  if (!d) return '';
  const date = d instanceof Date ? d : new Date(d);
  if (isNaN(date)) return String(d);
  const greg = Utilities.formatDate(date, CFG.TZ, 'yyyy/MM/dd') + 'م';
  if (short || String(s && s[S.HIJRI]).trim() !== 'نعم') return greg;
  const h = hijri_(date);
  return h ? `${greg} الموافق ${h}` : greg;
}

function hijri_(date) {
  try {
    const parts = new Intl.DateTimeFormat('en-u-ca-islamic-umalqura-nu-latn',
      { year: 'numeric', month: '2-digit', day: '2-digit', timeZone: CFG.TZ }).formatToParts(date);
    const p = t => (parts.find(x => x.type === t) || {}).value;
    const y = parseInt(p('year'), 10);
    if (!y || y < 1400 || y > 1600) return '';
    return `${y}/${p('month')}/${p('day')}هـ`;
  } catch (e) { return ''; }
}

function isFemale_(g) { return /أنثى|انثى|انثي|female|^f$|^ف$/i.test(String(g).trim()); }
function isEmail_(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e || '').trim()); }
function toNum_(v) { if (v === '' || v === null || v === undefined) return null; const n = Number(v); return isNaN(n) ? null : n; }
function pad_(n, w) { return String(n).padStart(w, '0'); }
function esc_(t) { return String(t == null ? '' : t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function fileIdFromUrl_(u) { const m = String(u || '').match(/\/d\/([a-zA-Z0-9_-]+)/) || String(u || '').match(/[?&]id=([a-zA-Z0-9_-]+)/); return m ? m[1] : ''; }
function colLetter_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function toast_(m) { try { SpreadsheetApp.getActiveSpreadsheet().toast(m, 'نظام الشهادات', 6); } catch (e) {} }

function updateQuotaCell_(ss) {
  try {
    const sh = ss.getSheetByName(CFG.SHEETS.DASH);
    if (sh) sh.getRange('B13').setValue(MailApp.getRemainingDailyQuota());
  } catch (e) {}
}

// ============================== دليل الاستخدام ==============================
function showHelp() {
  const html = `
  <div dir="rtl" style="font-family:Tahoma,Arial;line-height:1.9;font-size:14px;padding:4px 8px">
    <h3 style="color:#0B4F6C;margin:0 0 8px">طريقة العمل</h3>
    <ol>
      <li>أضف الدورات في ورقة <b>الدورات</b> (رمز، اسم، تواريخ، ساعات).</li>
      <li>أضف المتدربين/المتدربات والمدربين/المدربات مع <b>الجنس</b> و<b>البريد</b> و<b>رمز الدورة</b>.</li>
      <li>للمتدربين: سجّل <b>نسبة الحضور</b> و<b>الدرجة</b> (الحدود الدنيا في الإعدادات).</li>
      <li>غيّر <b>الحالة</b> إلى <b>«مكتمل»</b> ← تصدر الشهادة PDF وتُرسل بالبريد تلقائياً.</li>
    </ol>
    <h3 style="color:#0B4F6C;margin:12px 0 8px">الحالات</h3>
    <ul>
      <li><b>قيد التدريب</b>: لم تكتمل الدورة.</li>
      <li><b>مكتمل</b>: جاهز للإصدار.</li>
      <li><b>صدرت</b>: الشهادة جاهزة ورابطها في الصف.</li>
      <li><b>غير مؤهل</b>: لم يحقق الشروط (السبب في الملاحظات). عدّل البيانات ثم أعد الحالة «مكتمل».</li>
      <li><b>خطأ</b>: راجع الملاحظات ثم أعد الحالة «مكتمل».</li>
      <li><b>ملغاة</b>: أُلغيت وتظهر كذلك في صفحة التحقق.</li>
    </ul>
    <h3 style="color:#0B4F6C;margin:12px 0 8px">تخصيص التصميم</h3>
    <p>افتح القالب من مجلد الشهادات في Drive وعدّل الألوان والخطوط والشعار بحرية، مع إبقاء العبارات بين <code>{{ }}</code> كما هي.</p>
    <p style="color:#888;font-size:12px">ملاحظة: حد البريد اليومي لحسابات Gmail المجانية 100 رسالة، و1500 لحسابات Google Workspace. الرسائل الزائدة تُرسل تلقائياً في اليوم التالي.</p>
  </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(560), '📖 دليل الاستخدام');
}
