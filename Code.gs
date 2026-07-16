// ==========================================
// Code.gs - V5.5.1 Ticket Integration + Original Bot Registration Links
// ==========================================
//
// Column Reference - Registration Sheet:
//   A[0] Raw Name   B[1] 上手白金/雁群組長   C[2] Event Name
//   D[3] Clean Name E[4] Flock
//   F[5] Status     <- blank=normal / 取消 / 新增
//   G[6] Remark     <- free text, record only
//   H[7] Class Type
//
// Column Reference - Settings Sheet:
//   [0]Event [1]Platinum [2]Non-Platinum [3][4][5]Other prices
//   [6]Code [7]mode [8]catStr
//
// Column Reference - 課程報名連結 Sheet:
//   A[0] 活動代碼   B[1] 活動名稱   C[2] 報名連結
// ==========================================

var LINE_TOKEN = '7AqkhWYfHk0jqK3P+hbZxnZNqoSMkoqYSYwlPFqvPXMMjOfy7HkhSijlVUk6EQOD4aDn4nPTtNhN3/mJjdiARzCOGXGrwflDlPx0d59rt38jBYbFpzgF9N/OfPKvqQfwyRzZazYlMEEBu13ofKw2BAdB04t89/1O/w1cDnyilFU=';
var SHEET_ID   = '1Hcno7Qhx_m76npbqXl5GqQKyHbD1xS-n5kTfIaudjG4';

// 新課程報名 LIFF 測試入口。
// 注意：LINE Bot「報名課程」正式指令已恢復讀取「課程報名連結」分頁，
// 不會直接導到這個 LIFF；LIFF 入口仍供測試與網站內部使用。
var COURSE_LIFF_URL = 'https://liff.line.me/2010580892-WjWtHyOn';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠️ 系統選單')
    .addItem('🔄 立即同步母版與報名資料', 'processRegistrations')
    .addItem('🎫 產生 CODE',              'generateCodes')
    .addItem('🖼️ 授權 EDM 上傳到雲端硬碟', 'authorizeDriveForEdm_')
    .addItem('🎟️ 初始化票券分頁', 'ensureTicketSheets')
    .addToUi();
}

function cleanName(input) {
  if (!input) return '';
  var chineseChars = input.toString().match(/[\u4e00-\u9fa5]/g);
  return chineseChars ? chineseChars.join('') : input;
}

function processRegistrations() {
  var ss         = SpreadsheetApp.openById(SHEET_ID);
  var regSheet   = ss.getSheetByName('報名總表');
  var mapSheet   = ss.getSheetByName('雁群總名冊');
  var setSheet   = ss.getSheetByName('系統設定');
  var regData    = regSheet.getDataRange().getValues();
  var masterData = mapSheet.getDataRange().getValues();
  var masterIndex = {};

  for (var i = 1; i < masterData.length; i++) {
    if (masterData[i][0]) {
      masterIndex[masterData[i][0]] = {
        flock:  (masterData[i][1] || '').toString().trim(),
        status: (masterData[i][3] || '').toString().trim()
      };
    }
  }

  for (var i = 1; i < regData.length; i++) {
    var rawName = regData[i][0];
    if (!rawName) continue;
    var clean = cleanName(rawName);
    regSheet.getRange(i + 1, 4).setValue(clean);
    var flock = masterIndex[clean] ? masterIndex[clean].flock : '未歸類(母版無紀錄)';
    regSheet.getRange(i + 1, 5).setValue(flock);
  }

  var now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm');
  setSheet.getRange('J1').setValue("'" + now);
  return now;
}
function doGet(e) {
  e = e || { parameter: {} };
  var params = e.parameter || {};
  var action = (params.action || '').toString();

  // 課程前端已改放 GitHub Pages / LIFF。
  // 如果有人還開舊的 Apps Script 課程網址，就直接導回 LIFF，避免在 GAS 頁面內跑 LIFF SDK。
  if (params.page === 'courseAdmin' || params.mode || params.course || params['liff.state']) {
    var target = COURSE_LIFF_URL;
    var query = [];
    if (params.mode) query.push('mode=' + encodeURIComponent(params.mode));
    if (params.course) query.push('course=' + encodeURIComponent(params.course));
    if (query.length) target += '?' + query.join('&');

    return HtmlService
      .createHtmlOutput(
        '<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<title>開啟課程系統</title></head><body style="font-family:sans-serif;padding:24px">' +
        '<p>正在開啟 CYP135 課程系統...</p>' +
        '<p><a href="' + target + '">如果沒有自動開啟，請點這裡</a></p>' +
        '<script>window.top.location.href="' + target + '";</script>' +
        '</body></html>'
      )
      .setTitle('開啟課程系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }

  var result = {};

  if (action === 'getActiveEvent') {
    result = getActiveEvent();

  } else if (action === 'getMembers') {
    result = { members: getMembers() };

  } else if (action === 'checkPermission') {
    return checkPermission(e);

  } else if (action === 'getInitData') {
    var uid  = (params.userId || '').toString().trim();
    var info = getScannerInfo(uid);

    result = {
      event:       getActiveEvent(),
      allowed:     info.allowed,
      scannerName: info.name,
      members:     getMembers()
    };

  } else if (action === 'checkin') {
    result = handleCheckin({
      code:    params.code    || '',
      scanner: params.scanner || ''
    });

  } else {
    result = {
      status: 'ok',
      message: 'CYP135 Apps Script API is running. Course frontend is on LINE LIFF / GitHub Pages.'
    };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {

  // ★ LIFF / 課程前端 API 入口 ★
  // GitHub Pages 的 index.html 會用 JSON POST 呼叫這裡。
  var body = {};
  var rawContents = '';

  try {
    if (e && e.postData && e.postData.contents) {
      rawContents = e.postData.contents;
      body = JSON.parse(rawContents);
    }
  } catch(err) {
    body = {};
  }

  if (body && body.action) {
    var result = null;

    if (body.action === 'checkin') {
      result = handleCheckin(body);

    } else if (body.action === 'getPublicInit') {
      result = getPublicInit(body.userId || '');

    } else if (body.action === 'getPublicInitFast') {
      result = getPublicInitFast(body.userId || '');

    } else if (body.action === 'getCourseAdminInit') {
      result = getCourseAdminInit(body.userId || '');

    } else if (body.action === 'getExternalImportInit') {
      result = getExternalImportInit(body.userId || '');

    } else if (body.action === 'getRegistrationInit') {
      result = getRegistrationInit(body.eventCode || body.courseCode || '', body.userId || '');

    } else if (body.action === 'createCourse') {
      result = createCourse(body);

    } else if (body.action === 'updateCourse') {
      result = updateCourse(body);

    } else if (body.action === 'uploadEdmImage') {
      result = uploadEdmImage(body);

    } else if (body.action === 'lookupMember') {
      result = lookupMember(body.name || '');

    } else if (body.action === 'checkRegistrationName') {
      result = checkRegistrationName(body);

    } else if (body.action === 'submitRegistration') {
      result = submitRegistration(body);

    } else if (body.action === 'previewExternalImport') {
      result = previewExternalImport(body);

    } else if (body.action === 'commitExternalImport') {
      result = commitExternalImport(body);

    } else if (body.action === 'syncImportedRowsFromMaster') {
      result = syncImportedRowsFromMaster(body);

    } else if (body.action === 'getTicketGeneratorInit') {
      result = getTicketGeneratorInit(body);

    } else if (body.action === 'syncTicketQrFromRegistrations') {
      result = syncTicketQrFromRegistrations(body);

    } else if (body.action === 'ensureTicketSheets') {
      result = ensureTicketSheets();

    } else {
      result = { status: 'error', message: '未知的 API action：' + body.action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── 原有 LINE Webhook 邏輯 ──
  var msg = {};
  try {
    msg = JSON.parse(rawContents || (e && e.postData ? e.postData.contents : '{}'));
  } catch(err2) {
    return ContentService.createTextOutput('OK');
  }

  var event = msg.events && msg.events.length ? msg.events[0] : null;
  if (!event || !event.message || event.message.type !== 'text') return ContentService.createTextOutput('OK');

  var userMessage     = event.message.text.trim().toUpperCase();
  var userMessageRaw  = event.message.text.trim();
  var userId          = event.source.userId;
  var ss              = SpreadsheetApp.openById(SHEET_ID);
  var permSheet       = ss.getSheetByName('LINE權限表');
  var permData        = permSheet.getDataRange().getValues();
  var userFlocksArray = [];
  var userRole        = '';
  var isExist         = false;

  for (var i = 0; i < permData.length; i++) {
    if (permData[i][0] === userId) {
      isExist = true;
      var rawFlocks = permData[i][2].toString();
      userFlocksArray = rawFlocks.split(/[，,]+/).map(function(s) { return s.trim(); }).filter(Boolean);
      userRole = permData[i][3];
      break;
    }
  }

  if (!isExist) {
    permSheet.appendRow([userId, getUserProfileName(userId), '', '']);
    return ContentService.createTextOutput('OK');
  }

  var setSheet   = ss.getSheetByName('系統設定');
  var eventsData = getBotEventsData_(ss);
  var lastUpdate = setSheet ? setSheet.getRange('J1').getDisplayValue() : '';

  var masterData2  = ss.getSheetByName('雁群總名冊').getDataRange().getValues();
  var masterIndex2 = {};
  for (var i = 1; i < masterData2.length; i++) {
    if (masterData2[i][0]) {
      masterIndex2[masterData2[i][0]] = {
        flock:  (masterData2[i][1] || '').toString().trim(),
        status: (masterData2[i][3] || '').toString().trim()
      };
    }
  }

  var subSheet      = ss.getSheetByName('包月資訊');
  var monthlyPeriod = (subSheet.getRange('E1').getDisplayValue() || '本月');

  if (userMessage === '查詢') return sendFlexMenu(event.replyToken, eventsData, userRole);

  if (userMessage === '報名課程') return sendCourseRegistration(event.replyToken);

  if (userMessage === '總表' && userRole === '會長') return sendQuickReply(event.replyToken, eventsData);

  if (userMessage === 'QR選單') return sendLinkQuickReply(event.replyToken);

  if (userMessage === '立即同步' && userRole === '會長') {
    var newTime = processRegistrations();
    sendLineMessage(event.replyToken, '✅ 同步完成！\n🕒 最新資料時間：\n' + newTime);
    return ContentService.createTextOutput('OK');
  }

  if (userMessage === '總包月名單' && userRole === '會長') return sendMonthlySummary(event.replyToken, ss, masterIndex2, [], lastUpdate, monthlyPeriod);

  if (userMessage === '小組包月查詢' && userFlocksArray.length > 0) return sendMonthlySummary(event.replyToken, ss, masterIndex2, userFlocksArray, lastUpdate, monthlyPeriod);

  if (userMessageRaw.indexOf('圖表') === 0 && userRole === '會長') {
    var parts     = userMessageRaw.trim().split(/\s+/);
    var chartCode = parts[1] ? parts[1].toUpperCase() : '';
    var colorKey  = parts[2] || '';

    if (colorKey) {
      var chartEvent = null;
      for (var i = 1; i < eventsData.length; i++) {
        if (eventsData[i][6] && eventsData[i][6].toString().toUpperCase() === chartCode) {
          chartEvent = eventsData[i]; break;
        }
      }
      if (chartEvent) {
        return sendChartFlex(event.replyToken, ss, chartEvent, masterIndex2, lastUpdate, colorKey);
      } else {
        sendLineMessage(event.replyToken, '⚠️ 找不到活動代碼：' + chartCode);
        return ContentService.createTextOutput('OK');
      }
    }

    var found = false;
    for (var i = 1; i < eventsData.length; i++) {
      if (eventsData[i][6] && eventsData[i][6].toString().toUpperCase() === chartCode) {
        found = true; break;
      }
    }
    if (!found) {
      sendLineMessage(event.replyToken, '⚠️ 找不到活動代碼：' + chartCode + '\n請輸入正確代碼，例如：圖表 A1');
      return ContentService.createTextOutput('OK');
    }
    sendColorQuickReply(event.replyToken, chartCode);
    return ContentService.createTextOutput('OK');
  }

  var targetEvent = null;
  for (var i = 1; i < eventsData.length; i++) {
    var code = eventsData[i][6] ? eventsData[i][6].toString().toUpperCase() : '';
    if (code !== '' && userMessage === code) { targetEvent = eventsData[i]; break; }
  }

  if (userMessage.indexOf('總計') !== -1 && userRole === '會長') {
    var code2 = userMessage.split(' ')[1];
    for (var i = 1; i < eventsData.length; i++) {
      if (eventsData[i][6] && eventsData[i][6].toString().toUpperCase() === (code2 || '').toString().toUpperCase()) { targetEvent = eventsData[i]; break; }
    }
    if (targetEvent) return sendRegistrationDetail(event.replyToken, ss, targetEvent, masterIndex2, [], lastUpdate);
  }

  if (targetEvent) {
    if (userFlocksArray.length > 0) return sendRegistrationDetail(event.replyToken, ss, targetEvent, masterIndex2, userFlocksArray, lastUpdate);
    else sendLineMessage(event.replyToken, '⚠️ 您尚未分配雁群權限。');
  }

  return ContentService.createTextOutput('OK');
}


function getBotEventsData_(ss) {
  // 如果同一個 Apps Script 專案裡有 course_manager.gs，先順手把舊格式 | 統一整理成逗號。
  try { if (typeof normalizeCourseCategoryColumns_ === 'function') normalizeCourseCategoryColumns_(); } catch (err) {}

  var result = [['活動名稱', '白金價', '非白金價', '', '', '', '活動代碼', '收費模式', '類別名稱']];
  var seenCodes = {};
  var seenNames = {};

  // 新版課程：優先從「課程管理」讀取，讓 LIFF 建立的活動可以直接被 BOT 查詢。
  // V5.2：類別名稱 / 類別價格支援任意數量，儲存統一用半形逗號。
  var courseSheet = ss.getSheetByName('課程管理');
  if (courseSheet && courseSheet.getLastRow() >= 2) {
    var courseData = courseSheet.getRange(2, 1, courseSheet.getLastRow() - 1, Math.min(20, courseSheet.getLastColumn())).getValues();
    for (var i = 0; i < courseData.length; i++) {
      var r = courseData[i];
      var eventCode = (r[0] || '').toString().trim();
      var eventName = (r[1] || '').toString().trim();
      if (!eventName || !eventCode) continue;

      var mode = (r[6] || '').toString().trim();
      var botRow;

      if (mode === '類別制') {
        var names = splitBotCsv_(r[9] || '');
        var prices = splitBotPriceList_(r[10] || '', names.length);
        var legacyPrices = prices.slice(0, 5);
        while (legacyPrices.length < 5) legacyPrices.push('');

        botRow = [
          eventName,
          legacyPrices[0], legacyPrices[1], legacyPrices[2], legacyPrices[3], legacyPrices[4],
          eventCode,
          '類別制',
          names.join(',')
        ];
        // 保留完整類別 / 價格，避免超過 5 類時 BOT 統計漏算。
        botRow._categoryNames = names;
        botRow._categoryPrices = prices;
      } else {
        botRow = [
          eventName,
          r[7] || '',
          r[8] || '',
          '', '', '',
          eventCode,
          '',
          ''
        ];
      }

      result.push(botRow);
      seenCodes[eventCode.toUpperCase()] = true;
      seenNames[eventName] = true;
    }
  }

  // 舊資料 / 手動資料：保留「系統設定」中沒有被課程管理覆蓋的活動。
  var setSheet = ss.getSheetByName('系統設定');
  if (setSheet && setSheet.getLastRow() >= 2) {
    var setData = setSheet.getDataRange().getValues();
    for (var j = 1; j < setData.length; j++) {
      var row = setData[j];
      var name = (row[0] || '').toString().trim();
      var code = (row[6] || '').toString().trim();
      if (!name || !code) continue;
      if (seenCodes[code.toUpperCase()] || seenNames[name]) continue;

      var out = row.slice(0, 9);
      while (out.length < 9) out.push('');
      if ((out[7] || '').toString().trim() === '類別制') {
        out._categoryNames = splitBotCsv_(out[8] || '');
        out._categoryPrices = [out[1], out[2], out[3], out[4], out[5]].map(function(x) { return (x || '').toString().trim(); });
      }
      result.push(out);
      seenCodes[code.toUpperCase()] = true;
      seenNames[name] = true;
    }
  }

  return result;
}

function splitBotCsv_(text) {
  return (text || '')
    .toString()
    .replace(/[，、｜|\n\r]/g, ',')
    .split(',')
    .map(function(x) { return x.trim(); })
    .filter(Boolean);
}

function splitBotPriceList_(text, expectedCount) {
  var raw = (text || '').toString().trim();
  var prices = splitBotCsv_(raw);
  if (!expectedCount || prices.length === expectedCount) return prices;

  var numberMatches = raw.match(/\d+(?:\.\d+)?/g) || [];
  if (numberMatches.length === expectedCount) return numberMatches;

  // 相容舊資料：500400 + 兩個類別 => 500,400。
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

function sendFlexMenu(replyToken, eventsData, userRole) {
  var contents = [];
  for (var i = 1; i < eventsData.length; i++) {
    var eventName = eventsData[i][0];
    var eventCode = eventsData[i][6] ? eventsData[i][6].toString() : '';
    if (eventName && eventCode) {
      contents.push({
        'type': 'button',
        'action': { 'type': 'message', 'label': '📅 ' + eventName, 'text': eventCode },
        'style': 'primary', 'margin': 'md', 'height': 'sm', 'color': '#1DB446'
      });
    }
  }
  if (contents.length === 0) { sendLineMessage(replyToken, '📅 目前沒有活動。'); return; }

  var flexPayload = {
    'type': 'bubble',
    'header': { 'type': 'box', 'layout': 'vertical', 'contents': [{ 'type': 'text', 'text': '請點選課程', 'weight': 'bold', 'size': 'lg', 'align': 'center' }] },
    'body':   { 'type': 'box', 'layout': 'vertical', 'contents': contents }
  };

  if (userRole === '會長') {
    flexPayload.footer = {
      'type': 'box', 'layout': 'vertical',
      'contents': [{ 'type': 'button', 'action': { 'type': 'message', 'label': '🏆 會長全場總表', 'text': '總表' }, 'style': 'link', 'color': '#000000' }]
    };
  }

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    'method': 'post',
    'payload': JSON.stringify({ 'replyToken': replyToken, 'messages': [{ 'type': 'flex', 'altText': '活動查詢', 'contents': flexPayload }] })
  });
}

// ── sendRegistrationDetail ───────────────────────────────────────
// F[5] 狀態：空白=正常 / 取消 / 新增
// G[6] 備註：自由文字（若為日期，僅顯示 yyyy/MM/dd）
// H[7] 上課類別
// ────────────────────────────────────────────────────────────────
function sendRegistrationDetail(replyToken, ss, eventRow, masterIndex, targetFlocksArray, lastUpdate) {
  var regData   = ss.getSheetByName('報名總表').getDataRange().getValues();
  var eventName = eventRow[0];
  var prices    = [Number(eventRow[1])||0, Number(eventRow[2])||0, Number(eventRow[3])||0, Number(eventRow[4])||0, Number(eventRow[5])||0];
  var mode      = (eventRow[7] || '').toString().trim();
  var catStr    = (eventRow[8] || '').toString().trim();
  var isDynamic = (mode === '類別制');
  var catMap = {}, catOrder = [];

  if (isDynamic) {
    var dynamicNames = eventRow._categoryNames || splitBotCsv_(catStr);
    var dynamicPrices = eventRow._categoryPrices || prices;
    dynamicNames.forEach(function(c, j) {
      c = (c || '').toString().trim();
      if (c) { catMap[c] = Number(dynamicPrices[j]) || 0; catOrder.push(c); }
    });
  }

  var summary = {};
  var grandTotalCount = 0;
  var grandTotalMoney = 0;
  targetFlocksArray.forEach(function(f) {
    summary[f] = { groups: {}, addNames: [], cancelNames: [], totalMoney: 0, totalCount: 0 };
  });

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][2] !== eventName) continue;

    var name    = regData[i][3];
    var flock   = regData[i][4].toString().trim();
    var status  = (regData[i][5] || '').toString().trim();

    // G[6] 備註：若為日期物件，僅格式化為 yyyy/MM/dd；否則維持原始文字
    var remarkRaw = regData[i][6];
    var remark    = (remarkRaw instanceof Date)
                      ? Utilities.formatDate(remarkRaw, 'GMT+8', 'yyyy/MM/dd')
                      : (remarkRaw || '').toString().trim();

    var userCat = (regData[i][7] || '').toString().trim();

    var mInfo = masterIndex[name] || { flock: flock, status: '非白金' };
    if (targetFlocksArray.length > 0 && targetFlocksArray.indexOf(mInfo.flock) === -1) continue;
    if (!summary[mInfo.flock]) {
      summary[mInfo.flock] = { groups: {}, addNames: [], cancelNames: [], totalMoney: 0, totalCount: 0 };
    }

    var price, label;
    if (isDynamic) {
      price = catMap.hasOwnProperty(userCat) ? catMap[userCat] : 0;
      label = catMap.hasOwnProperty(userCat) ? userCat : (userCat ? userCat + '(未定義)' : '未填寫類別');
    } else {
      var isPlat = mInfo.status.indexOf('白金') !== -1 && mInfo.status.indexOf('非') === -1;
      price = isPlat ? prices[0] : prices[1];
      label = isPlat ? '白金' : '非白金';
    }

    var displayName = name;
    if (remark) displayName += ' (' + remark + ')';

    if (status === '取消') {
      summary[mInfo.flock].cancelNames.push(displayName + ' [取消] -$' + price);
    } else if (status === '新增') {
      summary[mInfo.flock].addNames.push(displayName + ' [新增]');
      summary[mInfo.flock].totalMoney += price;
      summary[mInfo.flock].totalCount++;
      grandTotalCount++;
      grandTotalMoney += price;
    } else {
      if (!summary[mInfo.flock].groups[label]) summary[mInfo.flock].groups[label] = [];
      summary[mInfo.flock].groups[label].push(displayName);
      summary[mInfo.flock].totalMoney += price;
      summary[mInfo.flock].totalCount++;
      grandTotalCount++;
      grandTotalMoney += price;
    }
  }

  var title  = targetFlocksArray.length > 0 ? '雁群組長報名表' : '🏆 【全教室報名總覽】';
  var msgOut = title + '\n活動：' + eventName + '\n更新：' + lastUpdate + '\n------------------\n';
  msgOut += '統計\n應收總計：$' + grandTotalMoney + ' (出席 ' + grandTotalCount + ' 人)\n\n';

  Object.keys(summary).sort().forEach(function(key) {
    var fData = summary[key];
    if (!fData.totalCount && !fData.addNames.length && !fData.cancelNames.length) return;

    msgOut += '📍 【' + key + '小組】 ($' + fData.totalMoney + ')\n';

    var lbls = isDynamic ? catOrder.concat(['未填寫類別']) : ['白金', '非白金'];
    lbls.forEach(function(lbl) {
      if (fData.groups[lbl] && fData.groups[lbl].length)
        msgOut += lbl + '：\n' + fData.groups[lbl].map(function(n, i) { return (i+1) + '. ' + n; }).join('\n') + '\n';
    });
    Object.keys(fData.groups).forEach(function(lbl) {
      if (lbls.indexOf(lbl) === -1 && fData.groups[lbl].length)
        msgOut += lbl + '：\n' + fData.groups[lbl].map(function(n, i) { return (i+1) + '. ' + n; }).join('\n') + '\n';
    });

    if (fData.addNames.length)
      msgOut += '新增：\n' + fData.addNames.map(function(n, i) { return (i+1) + '. ' + n; }).join('\n') + '\n';

    if (fData.cancelNames.length)
      msgOut += '取消：\n' + fData.cancelNames.map(function(n, i) { return (i+1) + '. ' + n; }).join('\n') + '\n';

    msgOut += '\n';
  });

  sendLineMessage(replyToken, msgOut);
}

function sendMonthlySummary(replyToken, ss, masterIndex, targetFlocksArray, lastUpdate, monthlyPeriod) {
  var subData = ss.getSheetByName('包月資訊').getDataRange().getValues();
  var summary = {}, grandTotalMoney = 0;
  targetFlocksArray.forEach(function(f) { summary[f] = { AA:[], A:[], B:[], C:[], D:[], total:0 }; });

  for (var i = 1; i < subData.length; i++) {
    var name = subData[i][0]; if (!name) continue;
    var plan  = subData[i][3].toString().toUpperCase().trim();
    var mInfo = masterIndex[name] || { flock: '未歸類', status: '非白金' };
    if (targetFlocksArray.length > 0 && targetFlocksArray.indexOf(mInfo.flock) === -1) continue;
    var price = plan==='AA'?800 : plan==='A'?1000 : plan==='B'?800 : plan==='C'?200 : plan==='D'?500 : 0;
    grandTotalMoney += price;
    if (!summary[mInfo.flock]) summary[mInfo.flock] = { AA:[], A:[], B:[], C:[], D:[], total:0 };
    if (summary[mInfo.flock][plan] !== undefined) { summary[mInfo.flock][plan].push(name); summary[mInfo.flock].total += price; }
  }

  var msgOut = '🏆 包月清單 (' + monthlyPeriod + ')\n應收：$' + grandTotalMoney + '\n更新：' + lastUpdate + '\n------------------\n';
  Object.keys(summary).sort().forEach(function(flock) {
    msgOut += '📍 【' + flock + '小組】 ($' + summary[flock].total + ')\n';
    ['AA','A','B','C','D'].forEach(function(p) {
      if (summary[flock][p].length) msgOut += '【' + p + ' 方案】\n' + summary[flock][p].map(function(n,i){ return (i+1)+'. '+n; }).join('\n') + '\n';
    });
    msgOut += '\n';
  });
  sendLineMessage(replyToken, msgOut);
}

function sendQuickReply(replyToken, eventsData) {
  var rows = [];

  for (var i = 1; i < eventsData.length; i++) {
    var eventName = eventsData[i][0];
    var eventCode = eventsData[i][6] ? eventsData[i][6].toString() : '';
    if (!eventName || !eventCode) continue;

    rows.push({
      type: 'box', layout: 'vertical',
      margin: i === 1 ? 'none' : 'md',
      contents: [{
        type: 'text',
        text: eventCode + '. ' + eventName,
        size: 'sm', weight: 'bold', color: '#333333', wrap: true
      }]
    });

    rows.push({
      type: 'box', layout: 'horizontal',
      margin: 'sm', spacing: 'sm',
      contents: [
        {
          type: 'button',
          action: { type: 'message', label: '📋 查總表', text: '總計 ' + eventCode },
          style: 'primary', color: '#4C956C', height: 'sm', flex: 1
        },
        {
          type: 'button',
          action: { type: 'message', label: '📊 看圖表', text: '圖表 ' + eventCode },
          style: 'primary', color: '#2E7BB5', height: 'sm', flex: 1
        }
      ]
    });
  }

  if (rows.length === 0) { sendLineMessage(replyToken, '目前沒有活動。'); return; }

  var flexPayload = {
    type: 'bubble',
    header: {
      type: 'box', layout: 'vertical',
      backgroundColor: '#F5F5F5', paddingAll: '12px',
      contents: [{
        type: 'text', text: '🏆 會長專區　選擇活動',
        weight: 'bold', size: 'md', color: '#333333', align: 'center'
      }]
    },
    body: {
      type: 'box', layout: 'vertical',
      paddingAll: '12px', spacing: 'none',
      contents: rows
    }
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    method: 'post',
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'flex', altText: '會長專區：請選擇活動', contents: flexPayload }]
    })
  });
}

function getUserProfileName(userId) {
  try {
    return JSON.parse(UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId, {
      'headers': { 'Authorization': 'Bearer ' + LINE_TOKEN }
    }).getContentText()).displayName;
  } catch(err) { return '未知使用者'; }
}

function sendLineMessage(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    'method': 'post',
    'payload': JSON.stringify({ 'replyToken': replyToken, 'messages': [{ 'type': 'text', 'text': text }] })
  });
}

function sendLinkQuickReply(replyToken) {
  var quickReplyMessage = {
    'type': 'text',
    'text': '請選擇您要執行的操作：',
    'quickReply': {
      'items': [
        { 'type': 'action', 'action': { 'type': 'uri', 'label': '📷 開啟掃碼器', 'uri': 'https://liff.line.me/2009773305-tsZtDMOz' } },
        { 'type': 'action', 'action': { 'type': 'uri', 'label': '📁 領取 QR Code', 'uri': 'https://drive.google.com/drive/folders/1tEpBMCYticeTo1aRT3RYpUCyqioHdrkM' } }
      ]
    }
  };
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    'headers': { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    'method': 'post',
    'payload': JSON.stringify({ 'replyToken': replyToken, 'messages': [quickReplyMessage] })
  });
}

function sendColorQuickReply(replyToken, chartCode) {
  var colors = [
    { label: '🟢 自然綠', key: '綠' },
    { label: '🔵 清新藍', key: '藍' },
    { label: '🟡 陽光黃', key: '黃' },
    { label: '🟠 活力橘', key: '橘' },
    { label: '🔴 熱情紅', key: '紅' }
  ];
  var items = colors.map(function(c) {
    return { type: 'action', action: { type: 'message', label: c.label, text: '圖表 ' + chartCode + ' ' + c.key } };
  });
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    method: 'post',
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'text', text: '請選擇圖表配色：', quickReply: { items: items } }]
    })
  });
}

// ── sendChartFlex ────────────────────────────────────────────────
// 圖表統計：正常＋新增都計入，只排除取消
// F[5] 狀態
// ────────────────────────────────────────────────────────────────
function sendChartFlex(replyToken, ss, eventRow, masterIndex, lastUpdate, colorKey) {
  var regData   = ss.getSheetByName('報名總表').getDataRange().getValues();
  var eventName = eventRow[0];

  var flockCount = {};
  var grandTotal = 0;

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][2] !== eventName) continue;
    var name   = regData[i][3];
    var flock  = regData[i][4].toString().trim();
    var status = (regData[i][5] || '').toString().trim();

    if (status === '取消') continue;

    var mInfo     = masterIndex[name] || { flock: flock, status: '非白金' };
    var flockName = mInfo.flock || '未歸類';
    if (!flockCount[flockName]) flockCount[flockName] = 0;
    flockCount[flockName]++;
    grandTotal++;
  }

  var sortedFlocks = Object.keys(flockCount).sort();

  var colorThemes = {
    '綠': { header:'#3A8C5C', headerText:'#FFFFFF', headerSub:'#B7DFC9', rowHead:'#D6EFE0', rowHeadText:'#1F6B3E', rowEven:'#FFFFFF', rowOdd:'#F2FAF5', total:'#A8D8BB', totalText:'#145230', footer:'#EBF7F0', footerText:'#5A9E78' },
    '藍': { header:'#2E7BB5', headerText:'#FFFFFF', headerSub:'#B3D6F0', rowHead:'#D6EBFA', rowHeadText:'#1A5A8A', rowEven:'#FFFFFF', rowOdd:'#F0F7FD', total:'#A0CCE8', totalText:'#0D3F6A', footer:'#E8F3FB', footerText:'#4A87B8' },
    '黃': { header:'#C8960C', headerText:'#FFFFFF', headerSub:'#FAE8A0', rowHead:'#FDF3CC', rowHeadText:'#8A6200', rowEven:'#FFFFFF', rowOdd:'#FFFDF0', total:'#F5DC80', totalText:'#6B4900', footer:'#FDFAEC', footerText:'#B08020' },
    '橘': { header:'#D4621A', headerText:'#FFFFFF', headerSub:'#FAD0A8', rowHead:'#FDE8D4', rowHeadText:'#8C3A00', rowEven:'#FFFFFF', rowOdd:'#FFF6F0', total:'#F5BF98', totalText:'#6B2A00', footer:'#FDF2EB', footerText:'#C06030' },
    '紅': { header:'#B03535', headerText:'#FFFFFF', headerSub:'#F5BBBB', rowHead:'#FADDDD', rowHeadText:'#7A1515', rowEven:'#FFFFFF', rowOdd:'#FFF5F5', total:'#EFA8A8', totalText:'#5C0000', footer:'#FDF0F0', footerText:'#B05050' }
  };

  var theme = colorThemes[colorKey] || colorThemes['綠'];

  var tableRows = [];

  tableRows.push({
    type: 'box', layout: 'horizontal',
    contents: [
      { type: 'text', text: '雁群', size: 'sm', weight: 'bold', color: theme.rowHeadText, flex: 3, align: 'center' },
      { type: 'text', text: '人數', size: 'sm', weight: 'bold', color: theme.rowHeadText, flex: 1, align: 'center' }
    ],
    backgroundColor: theme.rowHead, paddingAll: '8px'
  });

  sortedFlocks.forEach(function(flock, idx) {
    var bgColor = idx % 2 === 0 ? theme.rowEven : theme.rowOdd;
    tableRows.push({
      type: 'box', layout: 'horizontal',
      contents: [
        { type: 'text', text: flock,                     size: 'sm', color: '#333333', flex: 3, align: 'center' },
        { type: 'text', text: String(flockCount[flock]), size: 'sm', color: '#333333', flex: 1, align: 'center' }
      ],
      backgroundColor: bgColor, paddingAll: '8px'
    });
  });

  tableRows.push({
    type: 'box', layout: 'horizontal',
    contents: [
      { type: 'text', text: '合計',            size: 'sm', weight: 'bold', color: theme.totalText, flex: 3, align: 'center' },
      { type: 'text', text: String(grandTotal), size: 'sm', weight: 'bold', color: theme.totalText, flex: 1, align: 'center' }
    ],
    backgroundColor: theme.total, paddingAll: '8px'
  });

  var flexPayload = {
    type: 'bubble', size: 'kilo',
    header: {
      type: 'box', layout: 'vertical',
      backgroundColor: theme.header, paddingAll: '12px',
      contents: [
        { type: 'text', text: eventName + ' 報名統計', size: 'md', weight: 'bold', color: theme.headerText, align: 'center', wrap: true },
        { type: 'text', text: '更新：' + lastUpdate, size: 'xxs', color: theme.headerSub, align: 'center', margin: 'xs' }
      ]
    },
    body: {
      type: 'box', layout: 'vertical',
      paddingAll: '0px', spacing: 'none',
      contents: tableRows
    },
    footer: {
      type: 'box', layout: 'vertical',
      backgroundColor: theme.footer, paddingAll: '8px',
      contents: [{ type: 'text', text: '共 ' + sortedFlocks.length + ' 個小組　總計 ' + grandTotal + ' 人', size: 'xxs', color: theme.footerText, align: 'center' }]
    }
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    method: 'post',
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'flex', altText: eventName + ' 報名統計（共 ' + grandTotal + ' 人）', contents: flexPayload }]
    })
  });
}
// ── sendCourseRegistration ────────────────────────────────────────
// LINE Bot 輸入「報名課程」
// 只讀取「課程報名連結」分頁，單純提供外部表單連結
//
// 課程報名連結：
// A欄：活動名稱
// B欄：報名連結
// C欄：可有可無，這版不讀人數、不顯示人數
// ─────────────────────────────────────────────────────────────────
function sendCourseRegistration(replyToken) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var linkSheet = ss.getSheetByName('課程報名連結');

  if (!linkSheet) {
    sendLineMessage(replyToken, '⚠️ 找不到「課程報名連結」工作表，請聯絡管理員。');
    return;
  }

  var data = linkSheet.getDataRange().getValues();
  var bubbles = [];

  for (var i = 1; i < data.length; i++) {
    var eventName = data[i][0] ? data[i][0].toString().trim() : '';
    var formUrl   = data[i][1] ? data[i][1].toString().trim() : '';

    if (!eventName) continue;

    var hasLink = formUrl !== '';

    var bodyContents = [
      {
        type: 'text',
        text: hasLink ? '點選下方按鈕前往報名表單' : '此課程尚未開放報名',
        size: 'sm',
        color: '#666666',
        wrap: true
      },
      hasLink
        ? {
            type: 'button',
            action: {
              type: 'uri',
              label: '✍️ 立即報名',
              uri: formUrl
            },
            style: 'primary',
            color: '#1DB446',
            height: 'sm',
            margin: 'md'
          }
        : {
            type: 'button',
            action: {
              type: 'message',
              label: '🔒 尚未開放',
              text: '報名課程'
            },
            style: 'secondary',
            height: 'sm',
            margin: 'md'
          }
    ];

    bubbles.push({
      type: 'bubble',
      size: 'kilo',
      header: {
        type: 'box',
        layout: 'vertical',
        backgroundColor: '#1DB446',
        paddingAll: '14px',
        contents: [
          {
            type: 'text',
            text: '📅 ' + eventName,
            weight: 'bold',
            size: 'sm',
            color: '#FFFFFF',
            wrap: true
          }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '12px',
        spacing: 'sm',
        contents: bodyContents
      }
    });
  }

  if (bubbles.length === 0) {
    sendLineMessage(replyToken, '📅 目前沒有可報名的課程。');
    return;
  }

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_TOKEN
    },
    method: 'post',
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [
        {
          type: 'flex',
          altText: '📋 課程報名連結',
          contents: {
            type: 'carousel',
            contents: bubbles
          }
        }
      ]
    })
  });
}
