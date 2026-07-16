// ==========================================
// qr_manager.gs — 報到時間於 F 欄，組別(D)/教室別(E)/掃描人員(G)
// ==========================================

// ── 產生 CODE（選單按鈕呼叫）────────────────────────────────────
function generateCodes() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('QR管理');
  var ui    = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('找不到「QR管理」工作表，請先建立。'); return; }
  if (!sheet.getRange('B1').getValue()) { ui.alert('請先在 B1 填入活動名稱。'); return; }

  // Token（B5 空白才產生）
  if (!sheet.getRange('B5').getValue()) {
    sheet.getRange('B5').setValue(makeCode());
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 9) { ui.alert('請先從第 9 列貼入報名名單（A欄姓名、B欄小組）。'); return; }

  var data = sheet.getRange(9, 1, lastRow - 8, 3).getValues();
  var generated = 0, skipped = 0;

  data.forEach(function(row, i) {
    if (!(row[0] || '').toString().trim()) return; // A欄空白跳過
    if (!row[2]) {
      sheet.getRange(9 + i, 3).setValue(makeCode());
      generated++;
    } else {
      skipped++;
    }
  });

  ui.alert('完成！\n新產生：' + generated + ' 人\n已有 CODE（跳過）：' + skipped + ' 人');
}

// ── LIFF 開啟時讀取活動資訊 ─────────────────────────────────────
function getActiveEvent() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('QR管理');

  if (!sheet) return { status: 'error', message: '找不到QR管理工作表' };

  return {
    status:     'ok',
    eventName:  (sheet.getRange('B1').getValue() || '').toString(),
    session:    (sheet.getRange('B2').getValue() || '').toString(),
    ticketType: (sheet.getRange('B3').getValue() || '').toString(),
    dateStr:    (sheet.getRange('B4').getValue() || '').toString()
  };
}

// ── 掃碼報到核心邏輯 ────────────────────────────────────────────
function handleCheckin(payload) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('QR管理');

  if (!sheet) return { status: 'error', message: '找不到QR管理工作表' };

  var code    = (payload.code || '').toString().trim().toUpperCase();
  var lastRow = sheet.getLastRow();
  if (lastRow < 9) return { status: 'error', message: '名單是空的' };

  // 從第 9 列搜尋 C 欄 CODE，讀取到 G 欄（共 7 欄：A~G）
  var data  = sheet.getRange(9, 1, lastRow - 8, 7).getValues();
  var found = -1;

  for (var i = 0; i < data.length; i++) {
    if ((data[i][2] || '').toString().trim().toUpperCase() === code) {
      found = i; break;
    }
  }

  if (found === -1) return { status: 'error', message: '查無報名記錄' };

  var name  = data[found][0].toString();
  var flock = data[found][1].toString();
  var group = data[found][3] ? data[found][3].toString() : '';  // D欄 組別
  var room  = data[found][4] ? data[found][4].toString() : '';  // E欄 教室別

  // F欄（index 5）已有值 = 已報到
  if (data[found][5] && data[found][5] !== '') {
    return {
      status:       'duplicate',
      name:         name,
      flock:        flock,
      group:        group,
      room:         room,
      firstCheckin: data[found][5].toString(),
      scanner:      data[found][6] ? data[found][6].toString() : ''  // G欄 掃描人員
    };
  }

  // 寫入報到時間到 F 欄（第 6 欄）
  var now = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm:ss');
  sheet.getRange(9 + found, 6).setValue(now);

  // 寫入掃描人員到 G 欄（第 7 欄），有傳入才寫
  var scanner = (payload.scanner || '').toString().trim();
  if (scanner) {
    sheet.getRange(9 + found, 7).setValue(scanner);
  }

  return { status: 'ok', name: name, flock: flock, group: group, room: room, time: now, scanner: scanner };
}

// ── 取得完整名單（LIFF 名單/統計/手動頁用）─────────────────────
function getMembers() {
  var ss      = SpreadsheetApp.openById(SHEET_ID);
  var sheet   = ss.getSheetByName('QR管理');
  var members = [];

  if (sheet && sheet.getLastRow() >= 9) {
    // 讀取到 G 欄（共 7 欄：A~G）
    var data = sheet.getRange(9, 1, sheet.getLastRow() - 8, 7).getValues();
    data.forEach(function(row) {
      if (row[0]) {
        members.push({
          name:    row[0].toString(),
          flock:   row[1].toString(),
          code:    row[2].toString(),
          group:   row[3] ? row[3].toString() : '',   // D欄 組別
          room:    row[4] ? row[4].toString() : '',   // E欄 教室別
          checked: !!row[5],
          time:    row[5] ? row[5].toString() : '',
          scanner: row[6] ? row[6].toString() : ''    // G欄 掃描人員
        });
      }
    });
  }

  return members;
}

// ── 查詢 UID 是否有掃碼權限＋對應姓名（A欄=UID，B欄=姓名）─────────
function getScannerInfo(userId) {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('掃碼權限表');
  if (!sheet) return { allowed: false, name: '' };

  var data = sheet.getDataRange().getValues();
  var uid  = userId.toString().trim();

  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim() === uid) {
      return {
        allowed: true,
        name: (data[i][1] || '').toString().trim()  // B欄 姓名
      };
    }
  }
  return { allowed: false, name: '' };
}

// 相容舊用法：僅回傳布林值
function checkScanPermission(userId) {
  return getScannerInfo(userId).allowed;
}

function checkPermission(e) {
  var userId = (e.parameter.userId || '').toString().trim();
  var info   = getScannerInfo(userId);
  return ContentService
    .createTextOutput(JSON.stringify({ allowed: info.allowed, scannerName: info.name }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 6 碼隨機 CODE（排除易混淆字元 0/O/1/I）─────────────────────
function makeCode() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var code  = '';
  for (var i = 0; i < 6; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}