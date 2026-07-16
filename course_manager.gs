// ==========================================
// course_manager.gs — 課程管理 / LIFF 報名 V5.3 單頁快取加速版
// 功能：
// 1. 首頁公開課程列表
// 2. 設定頁只限 LINE權限表 D欄 = 會長
// 3. 建立課程支援 EDM 手機上傳到 Google Drive、活動介紹、報名時間、下拉選單
// 4. 報名總表寫入 A~O；A~H 為人工主要欄位，I~O 為系統追蹤欄位
// 5. 上手白金名單與系統選項固定從 Sheet 讀取
// ==========================================

var COURSE_LIFF_BASE_URL = 'https://liff.line.me/2010580892-WjWtHyOn';

var COURSE_SHEET_NAME = '課程管理';
var REG_SHEET_NAME    = '報名總表';
var MASTER_SHEET_NAME = '雁群總名冊';
var LINK_SHEET_NAME   = '課程報名連結';
var SETTING_SHEET_NAME = '系統設定';
var EXTRA_ANSWER_SHEET_NAME = '報名附加答案';
var PLATINUM_SHEET_NAME = '上手白金名單';
var OPTION_SHEET_NAME = '系統選項';
var EDM_FOLDER_NAME = 'CYP135_EDM';
var FAST_CACHE_SECONDS = 300;
var COURSE_PURPOSE_WEBSITE = '報名網站';
var COURSE_PURPOSE_EXTERNAL = '外部匯入';

function apiJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSs_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet_(name) {
  var ss = getSs_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function cleanNameSafe_(input) {
  if (typeof cleanName === 'function') return cleanName(input);
  if (!input) return '';
  var chineseChars = input.toString().match(/[\u4e00-\u9fa5]/g);
  return chineseChars ? chineseChars.join('') : input.toString().trim();
}

function ensureHeaders_(sh, headers) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    return;
  }

  var current = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (!current[i]) {
      sh.getRange(1, i + 1).setValue(headers[i]);
    }
  }
}

function ensureCourseSheets() {
  var courseHeaders = [
    '活動代碼', '活動名稱', '日期時間', '場次', '票種', '上限人數',
    '收費模式', '白金價', '非白金價', '類別名稱', '類別價格',
    '是否開放報名', '是否啟用QR', '報名連結', '建立時間', '狀態',
    'DM圖片網址', '活動介紹', '報名開始時間', '報名結束時間', '課程用途'
  ];
  ensureHeaders_(getSheet_(COURSE_SHEET_NAME), courseHeaders);

  var regHeaders = [
    '姓名', '上手白金/雁群組長', '所屬活動', '純淨姓名', '歸屬雁群',
    '空白/取消/加報', '備註', '上課類別',
    '報名來源', 'LINE UID', '報名ID', '建立者', '報名時間', '應收金額', '活動代碼'
  ];
  var regSheet = getSheet_(REG_SHEET_NAME);
  ensureHeaders_(regSheet, regHeaders);
  regSheet.getRange(1, 2).setValue('上手白金/雁群組長');
  regSheet.getRange(1, 1, 1, regHeaders.length).setValues([regHeaders]);

  ensureHeaders_(getSheet_(EXTRA_ANSWER_SHEET_NAME), ['報名ID', '活動代碼', '活動名稱', '姓名', '問題', '答案']);
  ensureHeaders_(getSheet_(LINK_SHEET_NAME), ['活動名稱', '報名連結', '上限人數']);
  var leaderSheet = getSheet_(PLATINUM_SHEET_NAME);
  ensureHeaders_(leaderSheet, ['上手白金/雁群組長']);
  leaderSheet.getRange(1, 1).setValue('上手白金/雁群組長');

  var optionSheet = getSheet_(OPTION_SHEET_NAME);
  ensureHeaders_(optionSheet, ['場次', '票種', '課程類別常用選項', '小時', '分鐘', 'AMPM']);

  if (optionSheet.getLastRow() === 1) {
    var rows = [
      ['上午場', '一般票', '成人', '01', '00', 'AM'],
      ['下午場', '其他',   '小孩', '02', '15', 'PM'],
      ['晚上場', '',       '白金', '03', '30', ''],
      ['全日',   '',       '一般', '04', '45', ''],
      ['其他',   '',       '複訓', '05', '', ''],
      ['',       '',       '工作人員', '06', '', ''],
      ['',       '',       '', '07', '', ''],
      ['',       '',       '', '08', '', ''],
      ['',       '',       '', '09', '', ''],
      ['',       '',       '', '10', '', ''],
      ['',       '',       '', '11', '', ''],
      ['',       '',       '', '12', '', '']
    ];
    optionSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  // V5.3：不在每次讀取時整理整張課程表，避免拖慢首頁。
  // 類別 / 價格會在建立與編輯課程時統一成半形逗號；舊資料讀取時仍相容 |、｜、，、頓號與換行。
}

// ── 權限：只允許 LINE權限表 D欄 = 會長 ─────────────────────
function getAdminPermission_(userId) {
  var ss = getSs_();
  var sh = ss.getSheetByName('LINE權限表');
  if (!sh) return { allowed: false, role: '' };

  var data = sh.getDataRange().getValues();
  var uid = (userId || '').toString().trim();

  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === uid) {
      var role = (data[i][3] || '').toString().trim();
      return {
        allowed: role === '會長',
        role: role
      };
    }
  }

  return { allowed: false, role: '' };
}

// ── 公開首頁初始化：所有人可看課程，只有會長看到設定入口 ─────
function getPublicInit(userId) {
  ensureCourseSheets();
  var perm = getAdminPermission_(userId);

  return {
    status: 'ok',
    isAdmin: perm.allowed,
    role: perm.role,
    courses: getCoursesWithStats_(COURSE_PURPOSE_WEBSITE)
  };
}

// V5.3 快速首頁初始化：只讀課程管理，不掃報名總表統計。
// 目的：讓使用者 0.5~1 秒內先看到課程，名額/統計不阻塞首頁。
function getPublicInitFast(userId) {
  // 輕量首頁入口：不跑 ensureCourseSheets()，避免每次首頁都檢查/整理所有表格。
  // 正式表格已建立後，這會比完整 getPublicInit 快很多。
  var perm = getAdminPermission_(userId);

  return {
    status: 'ok',
    fast: true,
    isAdmin: perm.allowed,
    role: perm.role,
    courses: getCoursesFast_(COURSE_PURPOSE_WEBSITE),
    loadedAt: Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss')
  };
}

// ── 設定頁初始化：只給會長 ──────────────────────────────
function getCourseAdminInit(userId) {
  ensureCourseSheets();
  var perm = getAdminPermission_(userId);

  return {
    status: 'ok',
    allowed: perm.allowed,
    role: perm.role,
    courses: perm.allowed ? getCoursesWithStats_(COURSE_PURPOSE_WEBSITE) : [],
    options: getSystemOptions_(),
    flocks: getFlockOptions_(),
    platinumNames: getPlatinumOptions_(),
    leaders: getLeaderOptions_()
  };
}

function getExternalImportInit(userId) {
  ensureCourseSheets();
  var perm = getAdminPermission_(userId);

  return {
    status: 'ok',
    allowed: perm.allowed,
    role: perm.role,
    courses: perm.allowed ? getCoursesFast_(COURSE_PURPOSE_EXTERNAL) : []
  };
}

// ── 學員報名頁初始化 ───────────────────────────────────

// ── EDM 圖片上傳：只限會長，存到 Google Drive ───────────────
function uploadEdmImage(payload) {
  ensureCourseSheets();

  var perm = getAdminPermission_(payload.userId);
  if (!perm.allowed) {
    return { status: 'error', message: '沒有 EDM 上傳權限。此功能限會長使用。' };
  }

  var eventName = (payload.eventName || '').toString().trim();
  var eventDate = (payload.eventDate || '').toString().trim();
  var dataUrl = (payload.dataUrl || '').toString();

  if (!eventName) return { status: 'error', message: '請先填活動名稱，再上傳 EDM 圖片。' };
  if (!eventDate) return { status: 'error', message: '請先選活動日期，再上傳 EDM 圖片。' };
  if (!dataUrl) return { status: 'error', message: '沒有收到圖片資料。' };

  var parsed = parseDataUrl_(dataUrl);
  if (!parsed) return { status: 'error', message: '圖片格式不正確，請重新選擇圖片。' };

  var fileName = buildEdmFileName_(eventDate, eventName, parsed.mimeType, payload.fileName);
  var bytes = Utilities.base64Decode(parsed.base64);
  var blob = Utilities.newBlob(bytes, parsed.mimeType, fileName);
  var folder = getOrCreateDriveFolder_(EDM_FOLDER_NAME);
  var file = folder.createFile(blob).setName(fileName);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var fileId = file.getId();
  var displayUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1600';
  var viewUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

  return {
    status: 'ok',
    message: 'EDM 圖片已上傳',
    fileId: fileId,
    fileName: fileName,
    url: displayUrl,
    viewUrl: viewUrl
  };
}

function parseDataUrl_(dataUrl) {
  var m = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) return null;
  return {
    mimeType: m[1],
    base64: m[2]
  };
}

function buildEdmFileName_(eventDate, eventName, mimeType, originalFileName) {
  var dateText = (eventDate || '').toString().replace(/[^0-9]/g, '');
  if (dateText.length >= 8) {
    dateText = dateText.substring(0, 8);
  } else {
    dateText = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd');
  }

  var safeName = sanitizeDriveFileName_(eventName || '未命名活動');
  var ext = getImageExtension_(mimeType, originalFileName);

  return dateText + '_' + safeName + '.' + ext;
}

function sanitizeDriveFileName_(name) {
  var text = (name || '').toString().trim();
  text = text.replace(/[\\\/:*?"<>|#%&{}$!'@+=`]/g, '_');
  text = text.replace(/\s+/g, '_');
  text = text.replace(/_+/g, '_');
  if (!text) text = '未命名活動';
  return text.substring(0, 80);
}

function getImageExtension_(mimeType, originalFileName) {
  var mime = (mimeType || '').toString().toLowerCase();
  if (mime.indexOf('png') >= 0) return 'png';
  if (mime.indexOf('webp') >= 0) return 'webp';
  if (mime.indexOf('gif') >= 0) return 'gif';
  if (mime.indexOf('jpeg') >= 0 || mime.indexOf('jpg') >= 0) return 'jpg';

  var name = (originalFileName || '').toString().toLowerCase();
  var m = name.match(/\.([a-z0-9]+)$/);
  if (m && ['jpg', 'jpeg', 'png', 'webp', 'gif'].indexOf(m[1]) >= 0) {
    return m[1] === 'jpeg' ? 'jpg' : m[1];
  }
  return 'jpg';
}


// 第一次加入 Drive 上傳功能時，可在 Apps Script 編輯器手動執行一次這個函式授權。
function authorizeDriveForEdm() {
  DriveApp.getRootFolder();
  var folders = DriveApp.getFoldersByName(EDM_FOLDER_NAME);
  if (!folders.hasNext()) DriveApp.createFolder(EDM_FOLDER_NAME);
  return 'EDM Drive 授權完成';
}

function authorizeDriveForEdm_() {
  return authorizeDriveForEdm();
}

function getOrCreateDriveFolder_(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function getRegistrationInit(eventCode, userId) {
  ensureCourseSheets();

  var found = findCourseByCode_(eventCode);
  if (!found) {
    return { status: 'error', message: '找不到課程' };
  }

  var course = found.course;
  if (courseMatchesPurpose_(course, COURSE_PURPOSE_EXTERNAL)) {
    return { status: 'error', message: '此活動使用外部報名網站，不開放此報名頁。' };
  }

  var stats = getCourseStatsForCourse_(course);
  course.currentCount = stats.currentCount;
  course.remainingCount = stats.remainingCount;
  course.statusBadge = getCourseStatusBadge_(course, stats);
  course.priceSummary = buildPriceSummary_(course);
  course.classDetails = getCourseClassDetails_(course);

  var perm = getAdminPermission_(userId || '');

  return {
    status: 'ok',
    isAdmin: perm.allowed,
    role: perm.role,
    course: course,
    // 找不到名冊時，下拉來源固定抓「上手白金名單!A2:A」
    leaders: getLeaderOptions_(),
    platinumNames: getLeaderOptions_(), // 舊前端相容
    classTypes: getCourseClassOptions_(course),
    classDetails: getCourseClassDetails_(course),
    // 本課程已報名清單，給前端本地即時判斷重複，避免送出後才等 GAS 回覆。
    existingRegistrants: getExistingRegistrantsForCourse_(course),
    // V4.1：不要在打開活動詳情時載入整份雁群總名冊，避免大量資料造成頁面卡在載入中。
    // 名冊改成輸入姓名後再用 lookupMember 即時查詢。
    memberDirectory: [],
    options: getSystemOptions_()
  };
}

function getSystemOptions_() {
  return {
    sessions: readOptionColumn_(OPTION_SHEET_NAME, 1, ['上午場', '下午場', '晚上場', '全日', '其他']),
    // 票種先簡化成一般票 / 其他；其他由前端開放自訂文字
    ticketTypes: ['一般票', '其他'],
    classTypes: readOptionColumn_(OPTION_SHEET_NAME, 3, ['成人', '小孩', '白金', '一般', '複訓', '工作人員']),
    hours: readOptionColumn_(OPTION_SHEET_NAME, 4, ['01','02','03','04','05','06','07','08','09','10','11','12']),
    minutes: readOptionColumn_(OPTION_SHEET_NAME, 5, ['00','15','30','45']),
    ampm: readOptionColumn_(OPTION_SHEET_NAME, 6, ['AM', 'PM'])
  };
}

function readOptionColumn_(sheetName, col, fallback) {
  var sh = getSheet_(sheetName);
  var last = sh.getLastRow();
  if (last < 2) return fallback || [];

  var data = sh.getRange(2, col, last - 1, 1).getValues();
  var map = {};
  var arr = [];

  data.forEach(function(row) {
    var v = (row[0] || '').toString().trim();
    if (v && !map[v]) {
      map[v] = true;
      arr.push(v);
    }
  });

  return arr.length ? arr : (fallback || []);
}

function getPlatinumOptions_() {
  return getLeaderOptions_();
}

function getLeaderOptions_() {
  var cached = getCacheJson_('leader_options_v1');
  if (cached) return cached;
  var arr = readOptionColumn_(PLATINUM_SHEET_NAME, 1, []);
  setCacheJson_('leader_options_v1', arr, FAST_CACHE_SECONDS);
  return arr;
}


function normalizeCoursePurpose_(value, fallback) {
  var text = (value || '').toString().trim();
  if (text === COURSE_PURPOSE_EXTERNAL || text === 'LINE Bot' || text === '純LINE Bot' || text === '僅LINE Bot' || text === '僅外部匯入') {
    return COURSE_PURPOSE_EXTERNAL;
  }
  if (text === COURSE_PURPOSE_WEBSITE || text === '報名網站課程' || text === '網站報名') {
    return COURSE_PURPOSE_WEBSITE;
  }
  return fallback || COURSE_PURPOSE_WEBSITE;
}

function courseMatchesPurpose_(course, purpose) {
  if (!purpose) return true;
  return normalizeCoursePurpose_(course && course.coursePurpose, COURSE_PURPOSE_WEBSITE) === purpose;
}

function coursePurposeCacheKey_(purpose) {
  if (purpose === COURSE_PURPOSE_WEBSITE) return 'website';
  if (purpose === COURSE_PURPOSE_EXTERNAL) return 'external';
  return 'all';
}

// ── 取得課程列表 ──────────────────────────────────────
function getCourses_(purpose) {
  var sh = getSheet_(COURSE_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var data = sh.getRange(2, 1, last - 1, 21).getValues();
  var list = [];

  data.forEach(function(row, idx) {
    if (!row[0] && !row[1]) return;
    var course = courseRowToObj_(row, idx + 2);
    if (courseMatchesPurpose_(course, purpose)) list.push(course);
  });

  return list;
}

function getCoursesWithStats_(purpose) {
  var courses = getCourses_(purpose);
  var statsMap = getRegistrationStatsMap_();

  return courses.map(function(course) {
    var stats = getCourseStatsFromMap_(course, statsMap);
    course.currentCount = stats.currentCount;
    course.remainingCount = stats.remainingCount;
    course.statusBadge = getCourseStatusBadge_(course, stats);
    course.priceSummary = buildPriceSummary_(course);
    return course;
  }).reverse();
}

function getCoursesFast_(purpose) {
  var key = 'courses_fast_v57_' + coursePurposeCacheKey_(purpose);
  var cached = getCacheJson_(key);
  if (cached) return cached;

  var courses = getCourses_(purpose).map(function(course) {
    // 不掃報名總表，所以首頁可先快速顯示；實際已報名人數在進報名頁時再精準讀取。
    course.currentCount = '';
    course.remainingCount = '';
    course.statusBadge = getCourseStatusBadge_(course, { currentCount: 0, remainingCount: course.maxCount || 0 });
    course.priceSummary = buildPriceSummary_(course);
    course.classDetails = getCourseClassDetails_(course);
    return course;
  }).reverse();

  setCacheJson_(key, courses, 180);
  return courses;
}

function clearCourseFastCaches_() {
  try {
    CacheService.getScriptCache().remove('courses_fast_v53');
    CacheService.getScriptCache().remove('courses_fast_v56_all');
    CacheService.getScriptCache().remove('courses_fast_v56_' + COURSE_PURPOSE_WEBSITE);
    CacheService.getScriptCache().remove('courses_fast_v56_' + COURSE_PURPOSE_EXTERNAL);
    CacheService.getScriptCache().remove('courses_fast_v57_all');
    CacheService.getScriptCache().remove('courses_fast_v57_website');
    CacheService.getScriptCache().remove('courses_fast_v57_external');
  } catch (err) {}
}

// 一次掃描報名總表，避免首頁 / 設定頁每一堂課都重讀整張表。
function getRegistrationStatsMap_() {
  var sh = getSheet_(REG_SHEET_NAME);
  var last = sh.getLastRow();
  var map = { byName: {}, byCode: {} };
  if (last < 2) return map;

  var width = Math.max(15, sh.getLastColumn());
  var data = sh.getRange(2, 1, last - 1, width).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var eventName = (row[2] || '').toString().trim();
    var status = (row[5] || '').toString().trim();
    var eventCode = (row[14] || '').toString().trim().toUpperCase();
    if (status === '取消') continue;

    if (eventName) map.byName[eventName] = (map.byName[eventName] || 0) + 1;
    if (eventCode) map.byCode[eventCode] = (map.byCode[eventCode] || 0) + 1;
  }
  return map;
}

function getCourseStatsFromMap_(course, statsMap) {
  statsMap = statsMap || getRegistrationStatsMap_();
  var code = (course.eventCode || '').toString().trim().toUpperCase();
  var name = (course.eventName || '').toString().trim();
  var codeCount = code ? (statsMap.byCode[code] || 0) : 0;
  var nameCount = name ? (statsMap.byName[name] || 0) : 0;
  var count = Math.max(codeCount, nameCount);
  var maxCount = Number(course.maxCount) || 0;
  return {
    currentCount: count,
    remainingCount: maxCount > 0 ? Math.max(maxCount - count, 0) : 0
  };
}


function courseRowToObj_(row, rowIndex) {
  var course = {
    rowIndex: rowIndex,
    eventCode: (row[0] || '').toString().trim(),
    eventName: (row[1] || '').toString().trim(),
    dateTime:  displayCell_(row[2]),
    session:   (row[3] || '').toString().trim(),
    ticketType:(row[4] || '').toString().trim(),
    maxCount:  Number(row[5]) || 0,
    pricingMode: (row[6] || '').toString().trim() || '白金制',
    platinumPrice: row[7] || '',
    nonPlatinumPrice: row[8] || '',
    categoryNames: (row[9] || '').toString().trim(),
    categoryPrices:(row[10] || '').toString().trim(),
    openFlag: (row[11] || '').toString().trim(),
    qrFlag:   (row[12] || '').toString().trim(),
    registrationLink: (row[13] || '').toString().trim(),
    createdAt: displayCell_(row[14]),
    statusText: (row[15] || '').toString().trim(),
    dmUrl: (row[16] || '').toString().trim(),
    description: (row[17] || '').toString().trim(),
    applyStart: displayCell_(row[18]),
    applyEnd: displayCell_(row[19]),
    coursePurpose: normalizeCoursePurpose_(row[20], COURSE_PURPOSE_WEBSITE)
  };

  course.priceSummary = buildPriceSummary_(course);
  return course;
}

function displayCell_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, 'GMT+8', 'yyyy/MM/dd HH:mm');
  }
  return value.toString().trim();
}

function findCourseByCode_(eventCode) {
  var code = (eventCode || '').toString().trim().toUpperCase();
  if (!code) return null;

  var sh = getSheet_(COURSE_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return null;

  var data = sh.getRange(2, 1, last - 1, 21).getValues();

  for (var i = 0; i < data.length; i++) {
    var rowCode = (data[i][0] || '').toString().trim().toUpperCase();
    if (rowCode === code) {
      return {
        rowIndex: i + 2,
        course: courseRowToObj_(data[i], i + 2)
      };
    }
  }

  return null;
}

function findCourseByName_(eventName, excludeEventCode) {
  var name = (eventName || '').toString().trim();
  var exclude = (excludeEventCode || '').toString().trim().toUpperCase();
  if (!name) return null;

  var courses = getCourses_();
  for (var i = 0; i < courses.length; i++) {
    var c = courses[i];
    var code = (c.eventCode || '').toString().trim().toUpperCase();
    if ((c.eventName || '').toString().trim() === name && (!exclude || code !== exclude)) {
      return c;
    }
  }
  return null;
}

// ── 新增課程 ──────────────────────────────────────────
function createCourse(payload) {
  ensureCourseSheets();

  var perm = getAdminPermission_(payload.userId);
  if (!perm.allowed) {
    return { status: 'error', message: '沒有課程管理權限。此功能限會長使用。' };
  }

  var eventCode = (payload.eventCode || '').toString().trim().toUpperCase();
  var eventName = (payload.eventName || '').toString().trim();

  if (!eventName) return { status: 'error', message: '請填活動名稱' };

  if (!eventCode) {
    eventCode = generateNextCourseCode_();
  }

  if (findCourseByCode_(eventCode)) {
    return { status: 'error', message: '活動代碼已存在：' + eventCode };
  }

  if (findCourseByName_(eventName, eventCode)) {
    return { status: 'error', message: '活動名稱已存在，請更換名稱或直接編輯原活動。' };
  }

  var saveResult = saveCourseToSheet_(payload, eventCode, null);
  return {
    status: 'ok',
    message: '課程已建立',
    eventCode: eventCode,
    registrationLink: saveResult.registrationLink
  };
}

// ── 編輯課程：只限會長 ─────────────────────────────────
function updateCourse(payload) {
  ensureCourseSheets();

  var perm = getAdminPermission_(payload.userId);
  if (!perm.allowed) {
    return { status: 'error', message: '沒有課程管理權限。此功能限會長使用。' };
  }

  var eventCode = (payload.eventCode || '').toString().trim().toUpperCase();
  if (!eventCode) return { status: 'error', message: '缺少活動代碼，無法更新課程。' };

  var found = findCourseByCode_(eventCode);
  if (!found) return { status: 'error', message: '找不到活動代碼：' + eventCode };

  var oldEventName = found.course.eventName;
  var newEventName = (payload.eventName || '').toString().trim();
  if (findCourseByName_(newEventName, eventCode)) {
    return { status: 'error', message: '活動名稱已存在，請更換名稱。' };
  }

  var saveResult = saveCourseToSheet_(payload, eventCode, found.rowIndex, found.course);

  if (newEventName && oldEventName && newEventName !== oldEventName) {
    updateRegistrationEventName_(oldEventName, newEventName);
  }

  return {
    status: 'ok',
    message: '課程已更新',
    eventCode: eventCode,
    registrationLink: saveResult.registrationLink
  };
}

function saveCourseToSheet_(payload, eventCode, rowIndex, oldCourse) {
  var eventName = (payload.eventName || '').toString().trim();
  if (!eventName) throw new Error('請填活動名稱');

  var maxCount = Number(payload.maxCount) || 0;
  var openFlag = payload.openFlag === '是' ? '是' : '否';
  var qrFlag = payload.qrFlag === '是' ? '是' : '否';
  var statusText = openFlag === '是' ? '開放' : '關閉';
  var categoryInfo = normalizeCategoryInfo_(payload);
  var dateTime = (payload.dateTime || '').toString().trim();
  var ticketType = normalizeTicketType_(payload.ticketType, payload.ticketTypeOther);
  var coursePurpose = normalizeCoursePurpose_(payload.coursePurpose, COURSE_PURPOSE_WEBSITE);
  var registrationLink = COURSE_LIFF_BASE_URL + '?mode=detail&course=' + encodeURIComponent(eventCode);
  var createdAt = oldCourse && oldCourse.createdAt ? oldCourse.createdAt : new Date();

  var row = [
    eventCode,
    eventName,
    dateTime,
    payload.session || '',
    ticketType,
    maxCount,
    payload.pricingMode || '白金制',
    payload.platinumPrice || '',
    payload.nonPlatinumPrice || '',
    categoryInfo.names,
    categoryInfo.prices,
    openFlag,
    qrFlag,
    registrationLink,
    createdAt,
    statusText,
    payload.dmUrl || '',
    payload.description || '',
    payload.applyStart || '',
    payload.applyEnd || '',
    coursePurpose
  ];

  var sh = getSheet_(COURSE_SHEET_NAME);
  if (rowIndex) {
    sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sh.appendRow(row);
  }

  clearCourseFastCaches_();

  var settingPayload = Object.assign({}, payload, {
    ticketType: ticketType,
    categoryNames: categoryInfo.names,
    categoryPrices: categoryInfo.prices
  });

  upsertCourseLink_(eventName, registrationLink, maxCount);
  upsertSystemSetting_(settingPayload, eventCode, eventName);

  return { registrationLink: registrationLink };
}

function normalizeTicketType_(ticketType, otherText) {
  var type = (ticketType || '').toString().trim();
  var other = (otherText || '').toString().trim();
  if (type === '其他' && other) return other;
  return type || '一般票';
}

function updateRegistrationEventName_(oldName, newName) {
  var sh = getSheet_(REG_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return;

  var data = sh.getRange(2, 3, last - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === oldName) {
      sh.getRange(i + 2, 3).setValue(newName);
    }
  }
}

function normalizeCategoryInfo_(payload) {
  var items = payload.categoryItems || [];
  var names = [];
  var prices = [];

  if (Object.prototype.toString.call(items) === '[object Array]') {
    items.forEach(function(item) {
      var name = (item && item.name ? item.name : '').toString().trim();
      var price = (item && item.price !== undefined ? item.price : '').toString().trim();
      if (name) {
        names.push(name);
        prices.push(price);
      }
    });
  }

  if (!names.length) {
    names = parseCsv_(payload.categoryNames || '');
    prices = parsePriceList_(payload.categoryPrices || '', names.length);
  }

  // V5.2：類別制統一用「半形逗號」儲存，且類別與價格必須一一對應。
  // 範例：初訓,複訓,旁聽 / 500,400,300
  if ((payload.pricingMode || '').toString().trim() === '類別制') {
    if (!names.length) throw new Error('請至少新增一個課程類別');
    if (prices.length !== names.length) {
      throw new Error('類別數量與價格數量不一致。請確認例如：初訓,複訓,旁聽 / 500,400,300');
    }

    for (var i = 0; i < names.length; i++) {
      if (prices[i] === '' || prices[i] === null || prices[i] === undefined) {
        throw new Error('請填寫「' + names[i] + '」的價格');
      }
      var numeric = Number(prices[i]);
      if (isNaN(numeric) || numeric < 0) {
        throw new Error('「' + names[i] + '」的價格格式不正確');
      }
      prices[i] = String(numeric);
    }
  }

  return {
    names: names.join(','),
    prices: prices.join(',')
  };
}

function normalizeCourseCategoryColumns_() {
  try {
    var sh = getSheet_(COURSE_SHEET_NAME);
    var last = sh.getLastRow();
    if (last < 2) return;

    var data = sh.getRange(2, 1, last - 1, 11).getValues();
    var jk = sh.getRange(2, 10, last - 1, 2).getValues();
    var changed = false;

    for (var i = 0; i < data.length; i++) {
      var mode = (data[i][6] || '').toString().trim();
      if (mode !== '類別制') continue;

      var oldNames = (data[i][9] || '').toString().trim();
      var oldPrices = (data[i][10] || '').toString().trim();
      var names = parseCsv_(oldNames);
      if (!names.length) continue;
      var prices = parsePriceList_(oldPrices, names.length);

      var newNames = names.join(',');
      var newPrices = prices.join(',');
      if (newNames !== oldNames || newPrices !== oldPrices) {
        jk[i][0] = newNames;
        jk[i][1] = newPrices;
        changed = true;
      }
    }

    if (changed) sh.getRange(2, 10, last - 1, 2).setValues(jk);
  } catch (err) {
    // 不讓格式整理影響前台讀取。
  }
}

function parseCsv_(text) {
  return (text || '')
    .toString()
    .replace(/[，、｜|\n\r]/g, ',')
    .split(',')
    .map(function(x) { return x.trim(); })
    .filter(Boolean);
}

function parsePriceList_(text, expectedCount) {
  var raw = (text || '').toString().trim();
  var prices = parseCsv_(raw);

  if (prices.length === expectedCount) return prices;

  var numberMatches = raw.match(/\d+(?:\.\d+)?/g) || [];
  if (numberMatches.length === expectedCount) return numberMatches;

  // 相容舊資料：若曾被存成 500400、15001000 這種沒有分隔符的價格字串，
  // 依類別數平均切開，避免顯示成 $0。
  if (expectedCount > 1 && /^\d+$/.test(raw) && raw.length % expectedCount === 0) {
    var size = raw.length / expectedCount;
    if (size >= 2) {
      var out = [];
      for (var i = 0; i < expectedCount; i++) {
        out.push(raw.substring(i * size, (i + 1) * size));
      }
      return out;
    }
  }

  return prices;
}

// ── 寫入 / 更新 課程報名連結 ───────────────────────────
function upsertCourseLink_(eventName, link, maxCount) {
  var sh = getSheet_(LINK_SHEET_NAME);
  var last = sh.getLastRow();

  if (last < 1) {
    sh.appendRow(['活動名稱', '報名連結', '上限人數']);
    last = 1;
  }

  if (last >= 2) {
    var data = sh.getRange(2, 1, last - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim() === eventName) {
        sh.getRange(i + 2, 2).setValue(link);
        sh.getRange(i + 2, 3).setValue(maxCount);
        return;
      }
    }
  }

  sh.appendRow([eventName, link, maxCount]);
}

// ── 寫入 / 更新 系統設定 ──────────────────────────────
// 系統設定欄位：
// A活動名稱 B白金價 C非白金價 D~F其他價格 G活動代碼 H模式 I類別名稱
function upsertSystemSetting_(payload, eventCode, eventName) {
  var sh = getSheet_(SETTING_SHEET_NAME);
  var last = sh.getLastRow();

  var pricingMode = (payload.pricingMode || '白金制').toString().trim();
  var row;

  if (pricingMode === '類別制') {
    var prices = parsePriceList_(payload.categoryPrices || '', parseCsv_(payload.categoryNames || '').length);
    while (prices.length < 5) prices.push('');

    row = [
      eventName,
      prices[0],
      prices[1],
      prices[2],
      prices[3],
      prices[4],
      eventCode,
      '類別制',
      payload.categoryNames || ''
    ];
  } else {
    row = [
      eventName,
      payload.platinumPrice || '',
      payload.nonPlatinumPrice || '',
      '',
      '',
      '',
      eventCode,
      '',
      ''
    ];
  }

  if (last >= 2) {
    var data = sh.getRange(2, 1, last - 1, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowCode = (data[i][6] || '').toString().trim().toUpperCase();
      if (rowCode === eventCode) {
        sh.getRange(i + 2, 1, 1, 9).setValues([row]);
        return;
      }
    }
  }

  sh.appendRow(row);
}

// ── 取得所有「上手白金/雁群組長」選項 ───────────────────
function getFlockOptions_() {
  return getLeaderOptions_();
}

// 一次載入名冊給前端做本地快速比對，避免大量報名時每個姓名都等 GAS。
function getMemberDirectory_() {
  var cached = getCacheJson_('member_directory_v3');
  if (cached) return cached;

  var sh = getSheet_(MASTER_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var data = sh.getRange(2, 1, last - 1, Math.max(2, sh.getLastColumn())).getValues();
  var list = [];
  for (var i = 0; i < data.length; i++) {
    var rawName = (data[i][0] || '').toString().trim();
    if (!rawName) continue;
    list.push({
      name: rawName,
      cleanName: cleanNameSafe_(rawName),
      leader: (data[i][1] || '').toString().trim(),
      memberStatus: (data[i][3] || '').toString().trim()
    });
  }

  // CacheService 單筆大小有限，名冊太大時會自動略過快取，不影響功能。
  setCacheJson_('member_directory_v3', list, FAST_CACHE_SECONDS);
  return list;
}

function getExistingRegistrantsForCourse_(courseOrEventName) {
  var sh = getSheet_(REG_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var course = (typeof courseOrEventName === 'object') ? courseOrEventName : { eventName: courseOrEventName, eventCode: '' };
  var eventName = (course.eventName || '').toString().trim();
  var eventCode = (course.eventCode || '').toString().trim().toUpperCase();

  var width = Math.max(15, sh.getLastColumn());
  var data = sh.getRange(2, 1, last - 1, width).getValues();
  var list = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!registrationRowMatchesCourse_(row, eventName, eventCode)) continue;

    var status = (row[5] || '').toString().trim();
    if (status === '取消') continue;

    var rawName = (row[0] || '').toString().trim();
    var clean = cleanNameSafe_(row[3] || rawName);
    if (!clean) continue;

    list.push({
      name: rawName || clean,
      cleanName: clean,
      leader: (row[1] || row[4] || '').toString().trim(),
      classType: (row[7] || '').toString().trim(),
      status: status || '正常',
      eventCode: (row[14] || '').toString().trim()
    });
  }

  return list;
}

function registrationRowMatchesCourse_(row, eventName, eventCode) {
  var rowEvent = (row[2] || '').toString().trim();
  var rowCode = (row[14] || '').toString().trim().toUpperCase();

  // 新資料優先用 O 欄活動代碼；舊資料沒有 O 欄時，用 C 欄活動名稱相容。
  if (eventCode && rowCode && rowCode === eventCode) return true;
  if (eventName && rowEvent && rowEvent === eventName) return true;
  return false;
}


function getCacheJson_(key) {
  try {
    var text = CacheService.getScriptCache().get(key);
    return text ? JSON.parse(text) : null;
  } catch (err) {
    return null;
  }
}

function setCacheJson_(key, value, seconds) {
  try {
    var text = JSON.stringify(value);
    if (text.length > 90000) return;
    CacheService.getScriptCache().put(key, text, seconds || FAST_CACHE_SECONDS);
  } catch (err) {}
}

function clearCourseFastCache_() {
  try {
    CacheService.getScriptCache().removeAll([
      'leader_options_v1',
      'member_directory_v2',
      'member_directory_v3',
      'member_lookup_index_v5',
      'member_lookup_index_v6'
    ]);
  } catch (err) {}
}

// ── 查詢雁群總名冊 ───────────────────────────────────
function lookupMember(name) {
  return lookupMemberCore_(name);
}

// 報名頁輸入姓名時使用：先即時檢查同課程是否重複，再回傳名冊查詢結果。
// 這樣使用者不用等到按送出才知道重複，也不會同時看到「已找到名冊資料」造成誤判。
function checkRegistrationName(payload) {
  ensureCourseSheets();

  var eventCode = (payload.eventCode || payload.courseCode || '').toString().trim();
  var rawName = (payload.name || '').toString().trim();
  var clean = cleanNameSafe_(rawName);

  if (!rawName || !clean) {
    return { status: 'empty', message: '請輸入姓名' };
  }

  var found = findCourseByCode_(eventCode);
  if (!found) {
    return { status: 'error', message: '找不到課程，請重新開啟報名頁。' };
  }

  var dup = findDuplicateRegistration_(found.course.eventName, clean, found.course.eventCode);
  if (dup) {
    return {
      status: 'duplicate',
      duplicate: true,
      message: '此姓名已報名過本課程\n\n如有疑問，請洽票務調整。',
      cleanName: clean,
      registration: dup
    };
  }

  var lookup = lookupMemberCore_(rawName);
  lookup.duplicate = false;
  lookup.cleanName = clean;
  return lookup;
}

function lookupMemberCore_(name) {
  var clean = cleanNameSafe_(name);
  if (!clean) {
    return { status: 'error', message: '請輸入姓名' };
  }

  // 先用快取索引，讓連續報名時查名冊不必每次掃表。
  var index = getMemberLookupIndexCached_();
  if (index && index[clean]) {
    return normalizeLookupMatches_(clean, index[clean]);
  }

  // 快取不存在或名冊太大時，先用 TextFinder 找輸入的原始姓名，通常會比全表掃描快。
  var matches = findMemberRowsByTextFinder_(name, clean);
  if (matches.length) return normalizeLookupMatches_(clean, matches);

  return { status: 'not_found', cleanName: clean, message: '查無名冊資料' };
}

function getMemberLookupIndexCached_() {
  var cached = getCacheJson_('member_lookup_index_v6');
  if (cached) return cached;

  var sh = getSheet_(MASTER_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return {};

  var data = sh.getRange(2, 1, last - 1, Math.max(4, sh.getLastColumn())).getValues();
  var index = {};
  for (var i = 0; i < data.length; i++) {
    var rawName = (data[i][0] || '').toString().trim();
    var leader = (data[i][1] || '').toString().trim();
    if (!rawName) continue;
    var clean = cleanNameSafe_(rawName);
    if (!clean) continue;
    if (!index[clean]) index[clean] = [];
    index[clean].push({
      name: rawName,
      flock: leader,
      leader: leader,
      uplinePlatinum: leader,
      memberStatus: (data[i][3] || '').toString().trim()
    });
  }

  setCacheJson_('member_lookup_index_v6', index, FAST_CACHE_SECONDS);
  return index;
}

function findMemberRowsByTextFinder_(rawName, clean) {
  var sh = getSheet_(MASTER_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var values = [];
  try {
    var finder = sh.getRange(2, 1, last - 1, 1).createTextFinder(rawName).matchEntireCell(true);
    var ranges = finder.findAll();
    for (var i = 0; i < ranges.length; i++) {
      var r = ranges[i].getRow();
      var row = sh.getRange(r, 1, 1, Math.max(4, sh.getLastColumn())).getValues()[0];
      if (cleanNameSafe_(row[0]) === clean) {
        var leader = (row[1] || '').toString().trim();
        values.push({
          name: (row[0] || '').toString(),
          flock: leader,
          leader: leader,
          uplinePlatinum: leader,
          memberStatus: (row[3] || '').toString().trim()
        });
      }
    }
  } catch (err) {}

  // TextFinder 找不到才退回掃名冊資料。
  if (!values.length) {
    var data = sh.getRange(2, 1, last - 1, Math.max(4, sh.getLastColumn())).getValues();
    for (var j = 0; j < data.length; j++) {
      if (cleanNameSafe_(data[j][0]) === clean) {
        var leader2 = (data[j][1] || '').toString().trim();
        values.push({
          name: (data[j][0] || '').toString(),
          flock: leader2,
          leader: leader2,
          uplinePlatinum: leader2,
          memberStatus: (data[j][3] || '').toString().trim()
        });
      }
    }
  }

  return values;
}

function normalizeLookupMatches_(clean, matches) {
  matches = matches || [];
  if (matches.length === 1) {
    return { status: 'found', cleanName: clean, member: matches[0] };
  }
  if (matches.length > 1) {
    return { status: 'multiple', cleanName: clean, members: matches };
  }
  return { status: 'not_found', cleanName: clean, message: '查無名冊資料' };
}

function getMatchedMemberFromLookup_(lookup, preferredLeader) {
  if (!lookup) return null;
  if (lookup.status === 'found') return lookup.member || null;

  if (lookup.status === 'multiple') {
    var target = (preferredLeader || '').toString().trim();
    var members = lookup.members || [];
    for (var i = 0; i < members.length; i++) {
      var leader = (members[i].leader || members[i].flock || members[i].uplinePlatinum || '').toString().trim();
      if (target && leader === target) return members[i];
    }
    return members.length === 1 ? members[0] : null;
  }

  return null;
}

function isPlatinumClassPair_(course) {
  var mode = (course && course.pricingMode ? course.pricingMode : '').toString().trim();
  if (mode === '類別制') return false;

  var details = getCourseClassDetails_(course);
  var hasPlatinum = false;
  var hasNonPlatinum = false;
  for (var i = 0; i < details.length; i++) {
    var name = (details[i].name || '').toString().trim();
    if (name === '白金') hasPlatinum = true;
    if (name === '非白金') hasNonPlatinum = true;
  }
  return hasPlatinum && hasNonPlatinum;
}

function resolvePlatinumClassTypeFromMember_(course, member) {
  if (!isPlatinumClassPair_(course)) return '';

  var status = (member && member.memberStatus ? member.memberStatus : '').toString().trim();
  if (status.indexOf('非白金') !== -1) return '非白金';
  if (status.indexOf('白金') !== -1 && status.indexOf('非') === -1) return '白金';

  // 名冊沒有狀態或查無名冊時，依既有統計邏輯歸為非白金。
  return '非白金';
}


// ── 學員送出報名 ─────────────────────────────────────
function submitRegistration(payload) {
  ensureCourseSheets();

  var eventCode = payload.eventCode || payload.courseCode || '';
  var foundCourse = findCourseByCode_(eventCode);
  if (!foundCourse) return { status: 'error', message: '找不到課程' };

  var course = foundCourse.course;
  if (courseMatchesPurpose_(course, COURSE_PURPOSE_EXTERNAL)) {
    return { status: 'error', message: '此活動使用外部報名網站，不開放此報名頁。' };
  }

  var openCheck = validateCourseOpenForRegistration_(course);
  if (!openCheck.ok) return { status: openCheck.status, message: openCheck.message };

  var rawName = (payload.name || '').toString().trim();
  var clean = cleanNameSafe_(rawName);
  if (!clean) return { status: 'error', message: '請輸入姓名' };

  if (isDuplicateRegistration_(course.eventName, clean, course.eventCode)) {
    return { status: 'duplicate', message: '此姓名已報名過本課程\n\n如有疑問，請洽票務調整。' };
  }

  var currentCount = countRegistrations_(course.eventName);
  if (course.maxCount > 0 && currentCount >= course.maxCount) {
    return { status: 'full', message: '此課程已額滿' };
  }

  var lookup = lookupMemberCore_(rawName);
  var preferredLeader = (payload.leader || payload.flock || payload.uplinePlatinum || '').toString().trim();
  var member = getMatchedMemberFromLookup_(lookup, preferredLeader);
  var leader = '';

  if (member) {
    leader = member.leader || member.flock || preferredLeader || '';
  } else {
    // 查無名冊時，直接從「上手白金名單」選一個上手白金/雁群組長。
    leader = preferredLeader;
    if (!leader) {
      return {
        status: 'need_profile',
        cleanName: clean,
        message: '查無名冊資料，請選擇上手白金/雁群組長'
      };
    }
    appendNewMember_(clean, leader, payload.userId || '');
  }

  var classType = (payload.classType || '').toString().trim();
  var autoClassType = resolvePlatinumClassTypeFromMember_(course, member);
  if (autoClassType) classType = autoClassType;
  var priceInfo = calculateCoursePrice_(course, classType);
  if (!priceInfo.valid) {
    return { status: 'error', message: priceInfo.message || '請選擇上課類別' };
  }

  var regId = createRegistrationId_();
  var sh = getSheet_(REG_SHEET_NAME);
  var rowValues = [
    rawName,                         // A 姓名
    leader,                          // B 上手白金/雁群組長
    course.eventName,                // C 所屬活動
    clean,                           // D 純淨姓名
    leader,                          // E 歸屬雁群：與 B 欄一致
    '',                              // F 狀態
    payload.remark || '',            // G 備註
    classType,                       // H 上課類別
    'LIFF',                          // I 報名來源
    payload.userId || '',            // J LINE UID
    regId,                           // K 報名ID
    payload.displayName || '',       // L 建立者
    new Date(),                      // M 報名時間
    priceInfo.price,                 // N 應收金額
    course.eventCode                 // O 活動代碼
  ];

  insertRegistrationByEvent_(sh, course.eventName, rowValues, course.eventCode);

  return {
    status: 'ok',
    message: '報名完成',
    registrationId: regId,
    eventName: course.eventName,
    name: clean,
    flock: leader,
    leader: leader,
    uplinePlatinum: leader,
    classType: classType,
    price: priceInfo.price
  };
}

function generateNextCourseCode_() {
  var sh = getSheet_(COURSE_SHEET_NAME);
  var last = sh.getLastRow();
  var max = 0;

  if (last >= 2) {
    var values = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var code = (values[i][0] || '').toString().trim().toUpperCase();
      var m = code.match(/^A(\d+)$/);
      if (m) max = Math.max(max, Number(m[1]) || 0);
    }
  }

  var next = max + 1;
  var candidate;
  do {
    candidate = 'A' + String(next).padStart(4, '0');
    next++;
  } while (findCourseByCode_(candidate));

  return candidate;
}

// ── 外部 Excel 匯入：先預覽檢查，再確認寫入 ─────────────────────
function previewExternalImport(payload) {
  try {
    var prepared = prepareExternalImport_(payload, false);
    return buildExternalImportResponse_(prepared, false);
  } catch (err) {
    return { status: 'error', message: err.message || err.toString() };
  }
}

function commitExternalImport(payload) {
  var lock = null;

  try {
    lock = LockService.getScriptLock();
    lock.waitLock(20000);

    var prepared = prepareExternalImport_(payload, true);
    var sh = getSheet_(REG_SHEET_NAME);
    var imported = 0;
    var writtenRows = [];

    prepared.rows.forEach(function(item) {
      if (item.status !== 'ok') return;
      var writtenRow = insertRegistrationByEvent_(sh, item.activityName, item.rowValues, prepared.course.eventCode);
      if (writtenRow) writtenRows.push(writtenRow);
      imported++;
    });

    clearCourseFastCaches_();

    var response = buildExternalImportResponse_(prepared, true);
    response.imported = imported;
    response.sheetName = REG_SHEET_NAME;
    response.writtenRows = writtenRows;
    response.firstWrittenRow = writtenRows.length ? Math.min.apply(null, writtenRows) : '';
    response.lastWrittenRow = writtenRows.length ? Math.max.apply(null, writtenRows) : '';
    response.message = '已匯入 ' + imported + ' 筆資料。' +
      (writtenRows.length ? ' 寫入報名總表第 ' + response.firstWrittenRow + ' 到 ' + response.lastWrittenRow + ' 列。' : '');
    return response;
  } catch (err) {
    return { status: 'error', message: err.message || err.toString() };
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (e) {}
    }
  }
}

function syncImportedRowsFromMaster(payload) {
  try {
    payload = payload || {};
    ensureCourseSheets();

    var perm = getAdminPermission_(payload.userId || '');
    if (!perm.allowed) return { status: 'error', message: '沒有同步母版權限。此功能限會長使用。' };

    var startRow = Number(payload.startRow || payload.firstWrittenRow) || 0;
    var endRow = Number(payload.endRow || payload.lastWrittenRow) || startRow;
    if (startRow < 2 || endRow < startRow) {
      return { status: 'error', message: '缺少可同步的報名總表列號。' };
    }
    if (endRow - startRow > 2000) {
      return { status: 'error', message: '同步範圍過大，請分批同步。' };
    }

    var regSheet = getSheet_(REG_SHEET_NAME);
    var last = regSheet.getLastRow();
    if (startRow > last) return { status: 'error', message: '同步列號超出報名總表範圍。' };
    endRow = Math.min(endRow, last);

    var rowCount = endRow - startRow + 1;
    var data = regSheet.getRange(startRow, 1, rowCount, 5).getValues();
    var memberIndex = getMemberLookupIndexCached_();
    var output = [];
    var synced = 0;
    var notFound = 0;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rawName = (row[0] || '').toString().trim();
      var fallbackLeader = (row[1] || '').toString().trim();
      var clean = cleanNameSafe_(rawName);
      var flockInfo = resolveImportFlock_(clean, fallbackLeader, memberIndex);
      output.push([clean, flockInfo.flock]);
      if (flockInfo.source === 'master') synced++;
      else notFound++;
    }

    regSheet.getRange(startRow, 4, rowCount, 2).setValues(output);
    return {
      status: 'ok',
      message: '已同步母版：更新報名總表第 ' + startRow + ' 到 ' + endRow + ' 列。名冊命中 ' + synced + ' 筆，未命中 ' + notFound + ' 筆。',
      startRow: startRow,
      endRow: endRow,
      synced: synced,
      notFound: notFound
    };
  } catch (err) {
    return { status: 'error', message: err.message || err.toString() };
  }
}

function prepareExternalImport_(payload, forCommit) {
  payload = payload || {};
  ensureCourseSheets();

  var perm = getAdminPermission_(payload.userId || '');
  if (!perm.allowed) throw new Error('沒有外部名單匯入權限。此功能限會長使用。');

  var eventCode = (payload.eventCode || payload.courseCode || '').toString().trim();
  var found = findCourseByCode_(eventCode);
  if (!found) throw new Error('找不到要匯入的課程。');

  var course = found.course;
  var rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!rows.length) throw new Error('沒有可預覽的 Excel 資料。');
  if (rows.length > 1500) throw new Error('單次匯入最多 1500 筆，請分批匯入。');
  var defaultClassType = (payload.defaultClassType || '').toString().trim();
  var memberIndex = getMemberLookupIndexCached_();

  var seenInFile = {};
  var preparedRows = rows.map(function(raw, idx) {
    return buildExternalImportRow_(raw, idx, course, seenInFile, forCommit, defaultClassType, memberIndex);
  });

  return {
    status: 'ok',
    course: course,
    fileName: (payload.fileName || '').toString().trim(),
    rows: preparedRows
  };
}

function buildExternalImportRow_(raw, idx, course, seenInFile, forCommit, defaultClassType, memberIndex) {
  raw = raw || {};

  var rawName = (raw.name || '').toString().trim();
  var clean = cleanNameSafe_(rawName);
  var leader = (raw.leader || '').toString().trim();
  var packageName = (raw.packageName || raw.activityName || '').toString().trim();
  var activityName = packageName || course.eventName;
  var remark = (raw.remark || '').toString().trim();
  var sourceTime = parseDateTimeFlexible_(raw.registeredAt) || new Date();
  var rowNumber = Number(raw.rowNumber) || (idx + 1);
  var messages = [];
  var status = 'ok';
  var flockInfo = resolveImportFlock_(clean, leader, memberIndex);

  if (!clean) {
    status = 'error';
    messages.push('缺少姓名');
  }
  if (!leader) {
    status = 'error';
    messages.push('缺少上手');
  }
  if (!activityName) {
    status = 'error';
    messages.push('缺少套組/所屬活動');
  }
  if (activityName && course.eventName && activityName !== course.eventName) {
    messages.push('Excel 所屬活動與選取課程名稱不同，請確認 LINE Bot 查詢是否要用此活動名稱');
  }

  var key = normalizeImportKey_(course.eventCode + '|' + clean);
  if (clean) {
    if (seenInFile[key]) {
      status = 'error';
      messages.push('同一份 Excel 內姓名重複');
    } else {
      seenInFile[key] = true;
    }
  }

  if (clean && findDuplicateRegistration_(activityName || course.eventName, clean, course.eventCode)) {
    status = 'duplicate';
    messages.push('報名總表已有同課程同姓名');
  }

  var classInfo = resolveExternalClassType_(course, raw.classType || defaultClassType || '', raw.memberType);
  if (!classInfo.valid) {
    status = 'error';
    messages.push(classInfo.message);
  }

  var priceInfo = classInfo.valid ? calculateCoursePrice_(course, classInfo.classType) : { valid: false, price: 0 };
  if (classInfo.valid && !priceInfo.valid) {
    status = 'error';
    messages.push(priceInfo.message || '無法計算金額');
  }

  var regId = forCommit && status === 'ok' ? createRegistrationId_() : '';
  var rowValues = [
    rawName,
    leader,
    activityName,
    clean,
    flockInfo.flock,
    '',
    remark,
    classInfo.classType || '',
    '外部Excel',
    '',
    regId,
    '外部匯入',
    sourceTime,
    priceInfo.valid ? priceInfo.price : '',
    course.eventCode
  ];

  return {
    rowNumber: rowNumber,
    status: status,
    messages: messages,
    name: rawName,
    cleanName: clean,
    leader: leader,
    flock: flockInfo.flock,
    flockSource: flockInfo.source,
    activityName: activityName,
    classType: classInfo.classType || '',
    price: priceInfo.valid ? priceInfo.price : '',
    registeredAt: displayCell_(sourceTime),
    rowValues: rowValues
  };
}

function resolveImportFlock_(cleanName, fallbackLeader, memberIndex) {
  var fallback = (fallbackLeader || '').toString().trim();
  if (!cleanName || !memberIndex || !memberIndex[cleanName] || !memberIndex[cleanName].length) {
    return { flock: fallback, source: 'excel' };
  }

  var matches = memberIndex[cleanName];
  for (var i = 0; i < matches.length; i++) {
    var flock = (matches[i].flock || matches[i].leader || '').toString().trim();
    if (flock) return { flock: flock, source: 'master' };
  }
  return { flock: fallback, source: 'excel' };
}

function resolveExternalClassType_(course, packageName, memberType) {
  var details = getCourseClassDetails_(course);
  var target = normalizeImportKey_(packageName);

  for (var i = 0; i < details.length; i++) {
    if (normalizeImportKey_(details[i].name) === target) {
      return { valid: true, classType: details[i].name };
    }
  }

  if (details.length === 1 && !target) {
    return { valid: true, classType: details[0].name };
  }

  var mode = (course.pricingMode || '').toString().trim();
  if (mode !== '類別制') {
    var text = ((packageName || '') + ' ' + (memberType || '')).toString();
    if (text.indexOf('非白金') !== -1) return { valid: true, classType: '非白金' };
    if (text.indexOf('白金') !== -1 && text.indexOf('非') === -1) return { valid: true, classType: '白金' };
  }

  var available = details.map(function(item) { return item.name; }).join('、');
  return {
    valid: false,
    classType: '',
    message: '套組無法對應課程類別：' + (packageName || '空白') + '；可用類別：' + available
  };
}

function buildExternalImportResponse_(prepared, committed) {
  var summary = { total: prepared.rows.length, ok: 0, duplicate: 0, error: 0, warning: 0 };
  prepared.rows.forEach(function(item) {
    if (item.status === 'ok') summary.ok++;
    else if (item.status === 'duplicate') summary.duplicate++;
    else summary.error++;
    if (item.messages && item.messages.length && item.status === 'ok') summary.warning++;
  });

  return {
    status: 'ok',
    committed: !!committed,
    course: {
      eventCode: prepared.course.eventCode,
      eventName: prepared.course.eventName,
      pricingMode: prepared.course.pricingMode,
      priceSummary: prepared.course.priceSummary
    },
    fileName: prepared.fileName,
    summary: summary,
    rows: prepared.rows.map(function(item) {
      return {
        rowNumber: item.rowNumber,
        status: item.status,
        messages: item.messages,
        name: item.name,
        leader: item.leader,
        flock: item.flock,
        flockSource: item.flockSource,
        activityName: item.activityName,
        classType: item.classType,
        price: item.price,
        registeredAt: item.registeredAt
      };
    })
  };
}

function normalizeImportKey_(text) {
  return (text || '').toString().trim().replace(/\s+/g, '').toUpperCase();
}

function insertRegistrationByEvent_(sheet, eventName, rowValues, eventCode) {
  var last = sheet.getLastRow();
  var values = rowValues.slice(0, 15);
  while (values.length < 15) values.push('');

  if (last < 2) {
    sheet.getRange(2, 1, 1, 15).setValues([values]);
    return 2;
  }

  var width = Math.max(15, sheet.getLastColumn());
  var data = sheet.getRange(2, 1, last - 1, width).getValues();
  var insertAfterRow = 0;
  var targetEvent = (eventName || '').toString().trim();
  var targetCode = (eventCode || '').toString().trim().toUpperCase();

  for (var i = 0; i < data.length; i++) {
    if (registrationRowMatchesCourse_(data[i], targetEvent, targetCode)) {
      insertAfterRow = i + 2;
    }
  }

  if (insertAfterRow > 0) {
    sheet.insertRowAfter(insertAfterRow);
    sheet.getRange(insertAfterRow + 1, 1, 1, 15).setValues([values]);
    return insertAfterRow + 1;
  } else {
    sheet.getRange(last + 1, 1, 1, 15).setValues([values]);
    return last + 1;
  }
}


function appendNewMember_(name, leader, userId) {
  var sh = getSheet_(MASTER_SHEET_NAME);

  sh.appendRow([
    name,
    leader || '',
    '',
    '',
    userId || '',
    new Date(),
    'LIFF報名自動新增',
    '待確認'
  ]);
  clearCourseFastCache_();
}

function findDuplicateRegistration_(eventName, cleanName, eventCode) {
  var sh = getSheet_(REG_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return null;

  var targetEvent = (eventName || '').toString().trim();
  var targetCode = (eventCode || '').toString().trim().toUpperCase();
  var targetName = cleanNameSafe_(cleanName);
  var width = Math.max(15, sh.getLastColumn());
  var data = sh.getRange(2, 1, last - 1, width).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!registrationRowMatchesCourse_(row, targetEvent, targetCode)) continue;

    var raw = (row[0] || '').toString().trim();
    var nm = cleanNameSafe_(row[3] || raw);
    var status = (row[5] || '').toString().trim();

    if (nm === targetName && status !== '取消') {
      return {
        row: i + 2,
        name: raw || nm,
        cleanName: nm,
        leader: (row[1] || row[4] || '').toString().trim(),
        eventName: (row[2] || '').toString().trim(),
        eventCode: (row[14] || '').toString().trim(),
        status: status || '正常',
        classType: (row[7] || '').toString().trim()
      };
    }
  }

  return null;
}

function isDuplicateRegistration_(eventName, cleanName, eventCode) {
  return !!findDuplicateRegistration_(eventName, cleanName, eventCode);
}

function countRegistrations_(eventName) {
  return getCourseStats_legacy_(eventName, 0).currentCount;
}

function getCourseStatsForCourse_(course) {
  return getCourseStatsFromMap_(course, getRegistrationStatsMap_());
}

// 舊函式保留，供舊 BOT 或舊邏輯相容。新版課程管理會優先用 getCourseStatsForCourse_。
function getCourseStats_(eventName, maxCount) {
  return getCourseStats_legacy_(eventName, maxCount);
}

function getCourseStats_legacy_(eventName, maxCount) {
  var sh = getSheet_(REG_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return { currentCount: 0, remainingCount: Number(maxCount) || 0 };

  var data = sh.getRange(2, 1, last - 1, 6).getValues();
  var count = 0;

  data.forEach(function(row) {
    var ev = (row[2] || '').toString().trim();
    var status = (row[5] || '').toString().trim();

    if (ev === eventName && status !== '取消') {
      count++;
    }
  });

  maxCount = Number(maxCount) || 0;
  return {
    currentCount: count,
    remainingCount: maxCount > 0 ? Math.max(maxCount - count, 0) : 0
  };
}


function getCourseStatusBadge_(course, stats) {
  if (course.openFlag !== '是') return '尚未開放';

  var now = new Date();
  var start = parseDateTimeFlexible_(course.applyStart);
  var end = parseDateTimeFlexible_(course.applyEnd);

  if (start && now.getTime() < start.getTime()) return '尚未開放';
  if (end && now.getTime() > end.getTime()) return '已截止';
  if (course.maxCount > 0 && stats.currentCount >= course.maxCount) return '已額滿';
  return '報名中';
}

function validateCourseOpenForRegistration_(course) {
  if (course.openFlag !== '是') {
    return { ok: false, status: 'closed', message: '此課程目前未開放報名' };
  }

  var now = new Date();
  var start = parseDateTimeFlexible_(course.applyStart);
  var end = parseDateTimeFlexible_(course.applyEnd);

  if (start && now.getTime() < start.getTime()) {
    return { ok: false, status: 'not_started', message: '此課程尚未開始報名' };
  }
  if (end && now.getTime() > end.getTime()) {
    return { ok: false, status: 'ended', message: '此課程報名已截止' };
  }
  return { ok: true };
}

function parseDateTimeFlexible_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value;

  var text = value.toString().trim();
  if (!text) return null;

  // datetime-local: 2026-07-03T12:30
  var m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})T(\d{1,2}):(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), 0);

  // display: 2026/07/03 12:30 或 2026-07-03 12:30
  m = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s+(\d{1,2}):(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), 0);

  return null;
}

function getCourseClassOptions_(course) {
  return getCourseClassDetails_(course).map(function(item) { return item.name; });
}

function getCourseClassDetails_(course) {
  var mode = (course.pricingMode || '').toString().trim();
  var result = [];

  if (mode === '類別制') {
    var names = parseCsv_(course.categoryNames || '');
    var prices = parsePriceList_(course.categoryPrices || '', names.length);
    for (var i = 0; i < names.length; i++) {
      result.push({ name: names[i], price: Number(prices[i]) || 0 });
    }
    return result;
  }

  result.push({ name: '白金', price: Number(course.platinumPrice) || 0 });
  result.push({ name: '非白金', price: Number(course.nonPlatinumPrice) || 0 });
  return result;
}

function calculateCoursePrice_(course, classType) {
  var name = (classType || '').toString().trim();
  if (!name) return { valid: false, message: '請選擇上課類別' };

  var details = getCourseClassDetails_(course);
  for (var i = 0; i < details.length; i++) {
    if (details[i].name === name) {
      return { valid: true, label: name, price: Number(details[i].price) || 0 };
    }
  }

  return { valid: false, message: '上課類別不在此課程設定內：' + name };
}

function buildPriceSummary_(course) {
  var mode = (course.pricingMode || '').toString().trim();

  if (mode === '類別制') {
    var details = getCourseClassDetails_(course);
    var parts = details.map(function(item) {
      return item.name + ' $' + (Number(item.price) || 0);
    });
    return parts.join('｜') || '依類別收費';
  }

  var p = course.platinumPrice !== '' ? '$' + course.platinumPrice : '—';
  var np = course.nonPlatinumPrice !== '' ? '$' + course.nonPlatinumPrice : '—';
  return '白金 ' + p + '｜非白金 ' + np;
}

function createRegistrationId_() {
  return 'R' +
    Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMddHHmmss') +
    '-' +
    Math.floor(Math.random() * 10000);
}

// ── 選用：修正舊版 LIFF 把報名時間寫到 B欄的資料 ─────────────
// 只需要手動執行一次。功能：如果 B欄看起來是時間，會搬到 M欄，並清空 B欄。
function fixOldRegistrationTimestampB() {
  ensureCourseSheets();
  var sh = getSheet_(REG_SHEET_NAME);
  var last = sh.getLastRow();
  if (last < 2) return '沒有需要修正的資料';

  var data = sh.getRange(2, 1, last - 1, 13).getValues();
  var fixed = 0;

  for (var i = 0; i < data.length; i++) {
    var b = data[i][1];
    var m = data[i][12];
    var isDateObj = Object.prototype.toString.call(b) === '[object Date]';
    var isDateText = (!isDateObj && b && /\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}/.test(b.toString()));

    if ((isDateObj || isDateText) && !m) {
      sh.getRange(i + 2, 13).setValue(b);
      sh.getRange(i + 2, 2).clearContent();
      fixed++;
    }
  }

  return '已修正 ' + fixed + ' 筆資料';
}

// ==========================================
// V5.5 票券產生整合：報名總表 → QR管理 → 票券 ZIP
// 目的：保留報名總表與 QR 掃描資料分開，減少手動貼資料。
// ==========================================
var QR_ACTIVE_SHEET_NAME = 'QR管理';
var QR_TICKET_REGISTRY_SHEET_NAME = 'QR票券紀錄';

function ensureTicketSheets() {
  ensureCourseSheets();
  ensureQrActiveSheet_();
  ensureQrTicketRegistrySheet_();
  return { status: 'ok', message: '票券分頁已建立 / 檢查完成' };
}

function ensureQrActiveSheet_() {
  var sh = getSheet_(QR_ACTIVE_SHEET_NAME);
  if (!sh.getRange('A1').getValue()) sh.getRange('A1').setValue('活動名稱');
  if (!sh.getRange('A2').getValue()) sh.getRange('A2').setValue('場次');
  if (!sh.getRange('A3').getValue()) sh.getRange('A3').setValue('票種');
  if (!sh.getRange('A4').getValue()) sh.getRange('A4').setValue('日期');
  if (!sh.getRange('A5').getValue()) sh.getRange('A5').setValue('Token');
  var headers = ['姓名', '小組/雁群', 'CODE', '組別', '教室別', '報到時間', '掃描人員'];
  sh.getRange(8, 1, 1, headers.length).setValues([headers]);
  try {
    sh.setFrozenRows(8);
  } catch (e) {}
  return sh;
}

function ensureQrTicketRegistrySheet_() {
  var sh = getSheet_(QR_TICKET_REGISTRY_SHEET_NAME);
  var headers = [
    '活動代碼', '活動名稱', '姓名', '上手白金/雁群組長', 'CODE',
    '組別', '教室別', '報到時間', '掃描人員', '上課類別', '更新時間'
  ];
  ensureHeaders_(sh, headers);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  try {
    sh.setFrozenRows(1);
  } catch (e) {}
  return sh;
}

function getTicketGeneratorInit(payload) {
  payload = payload || {};
  ensureTicketSheets();
  var perm = getAdminPermission_(payload.userId || '');
  if (!perm.allowed) return { status: 'error', message: '沒有票券產生權限。此功能限會長使用。' };
  return {
    status: 'ok',
    allowed: true,
    courses: getCoursesFast_()
  };
}

function syncTicketQrFromRegistrations(payload) {
  payload = payload || {};
  ensureTicketSheets();
  var perm = getAdminPermission_(payload.userId || '');
  if (!perm.allowed) return { status: 'error', message: '沒有票券產生權限。此功能限會長使用。' };

  var eventCode = (payload.eventCode || payload.courseCode || '').toString().trim();
  var foundCourse = findCourseByCode_(eventCode);
  if (!foundCourse) return { status: 'error', message: '找不到課程' };

  var course = foundCourse.course;
  var regsInfo = getTicketRegistrationsForCourse_(course);
  var activeSheet = ensureQrActiveSheet_();
  var registrySheet = ensureQrTicketRegistrySheet_();

  var existingMap = buildTicketExistingMap_(course, activeSheet, registrySheet);
  var existingCodes = collectExistingTicketCodes_(existingMap);

  var members = [];
  var stats = {
    total: 0,
    existingCode: 0,
    newCode: 0,
    cancelled: regsInfo.cancelled
  };

  for (var i = 0; i < regsInfo.rows.length; i++) {
    var r = regsInfo.rows[i];
    var key = ticketPersonKey_(r.cleanName, r.leader);
    var old = existingMap[key] || {};
    var code = (old.code || '').toString().trim().toUpperCase();
    if (code) {
      stats.existingCode++;
    } else {
      code = makeTicketCodeUnique_(existingCodes);
      stats.newCode++;
    }
    existingCodes[code] = true;
    members.push({
      name: r.cleanName || r.name,
      flock: r.leader,
      code: code,
      group: old.group || '',
      room: old.room || '',
      time: old.time || '',
      scanner: old.scanner || '',
      classType: r.classType || ''
    });
  }

  stats.total = members.length;
  writeQrActiveSheetForCourse_(activeSheet, course, members);
  rewriteTicketRegistryForCourse_(registrySheet, course, members);

  return {
    status: 'ok',
    message: '已同步 QR 管理名單',
    course: course,
    members: members,
    stats: stats
  };
}

function getTicketRegistrationsForCourse_(course) {
  var sh = getSheet_(REG_SHEET_NAME);
  var last = sh.getLastRow();
  var result = { rows: [], cancelled: 0 };
  if (last < 2) return result;

  var data = sh.getRange(2, 1, last - 1, Math.max(15, sh.getLastColumn())).getValues();
  var eventName = (course.eventName || '').toString().trim();
  var eventCode = (course.eventCode || '').toString().trim().toUpperCase();
  var seen = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowEvent = (row[2] || '').toString().trim();
    var rowCode = (row[14] || '').toString().trim().toUpperCase();
    if (rowEvent !== eventName && (!eventCode || rowCode !== eventCode)) continue;

    var status = (row[5] || '').toString().trim();
    if (status === '取消') {
      result.cancelled++;
      continue;
    }

    var rawName = (row[0] || '').toString().trim();
    var clean = (row[3] || '').toString().trim() || cleanNameSafe_(rawName);
    if (!clean) continue;
    var leader = (row[4] || '').toString().trim() || (row[1] || '').toString().trim();
    var key = ticketPersonKey_(clean, leader);
    if (seen[key]) continue;
    seen[key] = true;

    result.rows.push({
      name: rawName || clean,
      cleanName: clean,
      leader: leader,
      classType: (row[7] || '').toString().trim()
    });
  }
  return result;
}

function ticketPersonKey_(cleanName, leader) {
  return cleanNameSafe_(cleanName || '').toString().trim() + '||' + (leader || '').toString().trim();
}

function buildTicketExistingMap_(course, activeSheet, registrySheet) {
  var map = {};
  var eventName = (course.eventName || '').toString().trim();
  var eventCode = (course.eventCode || '').toString().trim().toUpperCase();

  // 先讀 QR票券紀錄，確保不同活動切換後，同活動舊 CODE 可保留。
  var lastR = registrySheet.getLastRow();
  if (lastR >= 2) {
    var reg = registrySheet.getRange(2, 1, lastR - 1, 11).getValues();
    for (var i = 0; i < reg.length; i++) {
      var rc = (reg[i][0] || '').toString().trim().toUpperCase();
      var rn = (reg[i][1] || '').toString().trim();
      if (rn !== eventName && (!eventCode || rc !== eventCode)) continue;
      var name = (reg[i][2] || '').toString().trim();
      var leader = (reg[i][3] || '').toString().trim();
      var key = ticketPersonKey_(name, leader);
      map[key] = {
        code: (reg[i][4] || '').toString().trim().toUpperCase(),
        group: (reg[i][5] || '').toString().trim(),
        room: (reg[i][6] || '').toString().trim(),
        time: reg[i][7] || '',
        scanner: (reg[i][8] || '').toString().trim(),
        classType: (reg[i][9] || '').toString().trim()
      };
    }
  }

  // 如果 QR管理 目前就是同一活動，優先用 QR管理 的最新報到時間 / 掃描人員。
  var activeEventName = (activeSheet.getRange('B1').getValue() || '').toString().trim();
  if (activeEventName === eventName && activeSheet.getLastRow() >= 9) {
    var cur = activeSheet.getRange(9, 1, activeSheet.getLastRow() - 8, 7).getValues();
    for (var j = 0; j < cur.length; j++) {
      if (!cur[j][0]) continue;
      var k = ticketPersonKey_(cur[j][0], cur[j][1]);
      map[k] = {
        code: (cur[j][2] || '').toString().trim().toUpperCase() || (map[k] ? map[k].code : ''),
        group: (cur[j][3] || '').toString().trim() || (map[k] ? map[k].group : ''),
        room: (cur[j][4] || '').toString().trim() || (map[k] ? map[k].room : ''),
        time: cur[j][5] || (map[k] ? map[k].time : ''),
        scanner: (cur[j][6] || '').toString().trim() || (map[k] ? map[k].scanner : '')
      };
    }
  }

  return map;
}

function collectExistingTicketCodes_(map) {
  var codes = {};
  for (var k in map) {
    if (map[k] && map[k].code) codes[map[k].code] = true;
  }
  return codes;
}

function makeTicketCodeUnique_(used) {
  used = used || {};
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (var attempt = 0; attempt < 2000; attempt++) {
    var code = '';
    for (var i = 0; i < 6; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
    if (!used[code]) return code;
  }
  throw new Error('無法產生不重複 CODE，請稍後再試。');
}

function writeQrActiveSheetForCourse_(sh, course, members) {
  sh.getRange('B1').setValue(course.eventName || '');
  sh.getRange('B2').setValue(course.session || '');
  sh.getRange('B3').setValue(course.ticketType || '');
  sh.getRange('B4').setValue(course.dateTime || '');
  if (!sh.getRange('B5').getValue()) sh.getRange('B5').setValue(makeTicketCodeUnique_({}));

  var last = sh.getLastRow();
  if (last >= 9) sh.getRange(9, 1, last - 8, 7).clearContent();
  if (!members.length) return;

  var values = members.map(function(m) {
    return [m.name || '', m.flock || '', m.code || '', m.group || '', m.room || '', m.time || '', m.scanner || ''];
  });
  sh.getRange(9, 1, values.length, 7).setValues(values);
}

function rewriteTicketRegistryForCourse_(sh, course, members) {
  var eventName = (course.eventName || '').toString().trim();
  var eventCode = (course.eventCode || '').toString().trim().toUpperCase();
  var last = sh.getLastRow();
  var keep = [];
  if (last >= 2) {
    var data = sh.getRange(2, 1, last - 1, 11).getValues();
    for (var i = 0; i < data.length; i++) {
      var rc = (data[i][0] || '').toString().trim().toUpperCase();
      var rn = (data[i][1] || '').toString().trim();
      if (rn === eventName || (eventCode && rc === eventCode)) continue;
      keep.push(data[i]);
    }
    sh.getRange(2, 1, last - 1, 11).clearContent();
  }

  var now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
  var rows = keep.concat(members.map(function(m) {
    return [
      course.eventCode || '', course.eventName || '', m.name || '', m.flock || '', m.code || '',
      m.group || '', m.room || '', m.time || '', m.scanner || '', m.classType || '', now
    ];
  }));
  if (rows.length) sh.getRange(2, 1, rows.length, 11).setValues(rows);
}
