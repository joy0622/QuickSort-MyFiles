/**
 * 신청자 관리 대시보드 — Google Apps Script 백엔드
 * ---------------------------------------------------------------
 * 구글 폼 응답 시트를 읽어 JSON으로 내려주고,
 * 대시보드에서 기록한 출석 / 문자발송 / 상태 / 메모를 시트에 저장합니다.
 *
 * 설치 방법은 ../README.md 를 참고하세요.
 */

var VERSION = '1.0.0';

var CONFIG = {
  // 폼 응답이 쌓이는 시트 이름. 비워두면 첫 번째 시트를 사용합니다.
  RESPONSE_SHEET: '',

  // 관리 데이터(출석/문자/메모)가 저장될 시트 이름. 없으면 자동 생성됩니다.
  ADMIN_SHEET: '_관리데이터',

  // 응답 시트 맨 끝에 추가되는 고유 ID 열. 행이 삭제/정렬돼도 메모가 따라갑니다.
  UID_HEADER: '_관리ID',

  // 선택: 아무나 URL을 알아도 못 읽게 하려면 아래에 아무 문자열이나 넣고
  // 대시보드 설정에도 똑같이 입력하세요. 비워두면 토큰 검사를 하지 않습니다.
  API_TOKEN: ''
};

var ADMIN_HEADERS = [CONFIG.UID_HEADER, '출석', '문자발송', '상태', '메모', '최종수정'];

/* ------------------------------------------------------------------ */
/* 엔드포인트                                                           */
/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    checkToken_(p.token);

    var action = p.action || 'list';
    if (action === 'ping') {
      return json_({ ok: true, version: VERSION, time: new Date().toISOString() });
    }
    if (action === 'list') {
      return json_(listData_());
    }
    return json_({ ok: false, error: '알 수 없는 action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: errMsg_(err) });
  }
}

function doPost(e) {
  try {
    var body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    checkToken_(body.token || (e && e.parameter && e.parameter.token));

    var action = body.action || 'save';

    // 단건 저장: { action:'save', uid:'...', patch:{ attendance:true, memo:'...' } }
    if (action === 'save') {
      var n = saveMany_([{ uid: body.uid, patch: body.patch || {} }]);
      return json_({ ok: true, saved: n });
    }

    // 여러 건 저장: { action:'saveMany', items:[{uid, patch}, ...] }
    if (action === 'saveMany') {
      var n2 = saveMany_(body.items || []);
      return json_({ ok: true, saved: n2 });
    }

    // 저장 후 최신 데이터까지 한 번에: { action:'saveAndList', items:[...] }
    if (action === 'saveAndList') {
      saveMany_(body.items || []);
      return json_(listData_());
    }

    return json_({ ok: false, error: '알 수 없는 action: ' + action });
  } catch (err) {
    return json_({ ok: false, error: errMsg_(err) });
  }
}

/* ------------------------------------------------------------------ */
/* 읽기                                                                 */
/* ------------------------------------------------------------------ */

function listData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = responseSheet_(ss);
  ensureUids_(sheet);

  var values = sheet.getDataRange().getDisplayValues();
  if (!values.length) {
    return { ok: true, version: VERSION, headers: [], rows: [], admin: {}, syncedAt: new Date().toISOString() };
  }

  var headers = values[0].map(function (h) { return String(h).trim(); });
  var uidIdx = headers.indexOf(CONFIG.UID_HEADER);

  // 응답 시트의 원본 날짜값(정렬용)을 함께 확보합니다.
  var raw = sheet.getDataRange().getValues();

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var line = values[r];
    if (isBlankRow_(line, uidIdx)) continue;

    var cells = [];
    for (var c = 0; c < headers.length; c++) {
      if (c === uidIdx) continue;
      cells.push(line[c] == null ? '' : String(line[c]));
    }

    // 잠금 경합으로 아직 ID 가 없는 행은 이번 응답에서만 행 번호를 임시 키로 씁니다.
    var uid = uidIdx >= 0 ? String(line[uidIdx] || '').trim() : '';
    if (!uid) uid = 'row-' + (r + 1);

    var tsRaw = raw[r][0];
    rows.push({
      uid: uid,
      row: r + 1,
      cells: cells,
      ts: (tsRaw instanceof Date) ? tsRaw.toISOString() : ''
    });
  }

  var publicHeaders = headers.filter(function (h, i) { return i !== uidIdx; });

  return {
    ok: true,
    version: VERSION,
    sheetName: sheet.getName(),
    spreadsheetName: ss.getName(),
    headers: publicHeaders,
    rows: rows,
    admin: readAdmin_(ss),
    syncedAt: new Date().toISOString()
  };
}

function readAdmin_(ss) {
  var sheet = adminSheet_(ss);
  var values = sheet.getDataRange().getValues();
  var out = {};
  for (var r = 1; r < values.length; r++) {
    var uid = String(values[r][0] || '').trim();
    if (!uid) continue;
    out[uid] = {
      attendance: truthy_(values[r][1]),
      sms: truthy_(values[r][2]),
      status: String(values[r][3] || ''),
      memo: String(values[r][4] || ''),
      updatedAt: values[r][5] instanceof Date ? values[r][5].toISOString() : String(values[r][5] || '')
    };
  }
  return out;
}

/* ------------------------------------------------------------------ */
/* 쓰기                                                                 */
/* ------------------------------------------------------------------ */

function saveMany_(items) {
  if (!items || !items.length) return 0;

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = adminSheet_(ss);
    var values = sheet.getDataRange().getValues();

    // uid -> 시트 행 번호
    var index = {};
    for (var r = 1; r < values.length; r++) {
      var uid = String(values[r][0] || '').trim();
      if (uid) index[uid] = r + 1;
    }

    var now = new Date();
    var appends = [];
    var saved = 0;

    for (var i = 0; i < items.length; i++) {
      var it = items[i] || {};
      var id = String(it.uid || '').trim();
      if (!id) continue;
      var patch = it.patch || {};

      var rowNum = index[id];
      var current = rowNum
        ? values[rowNum - 1]
        : [id, false, false, '', '', ''];

      var next = [
        id,
        patch.hasOwnProperty('attendance') ? !!patch.attendance : truthy_(current[1]),
        patch.hasOwnProperty('sms') ? !!patch.sms : truthy_(current[2]),
        patch.hasOwnProperty('status') ? String(patch.status || '') : String(current[3] || ''),
        patch.hasOwnProperty('memo') ? String(patch.memo || '') : String(current[4] || ''),
        now
      ];

      if (rowNum) {
        sheet.getRange(rowNum, 1, 1, ADMIN_HEADERS.length).setValues([next]);
      } else {
        appends.push(next);
        index[id] = values.length + appends.length; // 대략적인 자리 예약
      }
      saved++;
    }

    if (appends.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, ADMIN_HEADERS.length).setValues(appends);
    }

    SpreadsheetApp.flush();
    return saved;
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------------ */
/* 시트 유틸                                                            */
/* ------------------------------------------------------------------ */

function responseSheet_(ss) {
  var sheet = CONFIG.RESPONSE_SHEET
    ? ss.getSheetByName(CONFIG.RESPONSE_SHEET)
    : ss.getSheets()[0];
  if (!sheet) throw new Error('응답 시트를 찾을 수 없습니다: ' + CONFIG.RESPONSE_SHEET);
  if (sheet.getName() === CONFIG.ADMIN_SHEET) {
    throw new Error('응답 시트와 관리 시트가 같습니다. CONFIG.RESPONSE_SHEET 를 확인하세요.');
  }
  return sheet;
}

function adminSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.ADMIN_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ADMIN_SHEET);
    sheet.getRange(1, 1, 1, ADMIN_HEADERS.length).setValues([ADMIN_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, ADMIN_HEADERS.length).setValues([ADMIN_HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * 응답 시트 맨 끝에 _관리ID 열을 만들고, 비어 있는 행에 UUID를 채웁니다.
 * 폼으로 새 응답이 들어오면 이 열은 비어 있으므로 다음 조회 때 자동으로 채워집니다.
 *
 * 담당자 여러 명이 대시보드를 동시에 열어두면 같은 순간에 조회가 겹칠 수 있습니다.
 * 그때 두 요청이 같은 행에 서로 다른 ID 를 써 넣으면 먼저 기록한 메모가 떨어져 나가므로,
 * 실제로 쓸 내용이 있을 때만 잠금을 잡고 잠금 안에서 한 번 더 확인합니다.
 */
function ensureUids_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  // 1차 확인: 채울 것이 없으면 잠금 없이 그대로 끝냅니다 (대부분의 조회가 여기서 끝납니다).
  var uidCol = findUidCol_(sheet, lastCol);
  var needsWrite = !uidCol;
  if (uidCol && lastRow >= 2) {
    needsWrite = sheet.getRange(2, uidCol, lastRow - 1, 1).getValues().some(function (r) {
      return !String(r[0] || '').trim();
    });
  }
  if (!needsWrite) return;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return;   // 다른 요청이 처리 중입니다. 다음 조회 때 채워집니다.
  }

  try {
    // 2차 확인: 잠금을 기다리는 동안 다른 요청이 이미 채웠을 수 있습니다.
    lastRow = sheet.getLastRow();
    lastCol = sheet.getLastColumn();
    uidCol = findUidCol_(sheet, lastCol);

    if (!uidCol) {
      uidCol = lastCol + 1;
      sheet.getRange(1, uidCol).setValue(CONFIG.UID_HEADER);
    }
    if (lastRow < 2) return;

    var existing = sheet.getRange(2, uidCol, lastRow - 1, 1).getValues();
    var changed = false;
    for (var i = 0; i < existing.length; i++) {
      if (!String(existing[i][0] || '').trim()) {
        existing[i][0] = Utilities.getUuid();
        changed = true;
      }
    }
    if (changed) {
      sheet.getRange(2, uidCol, existing.length, 1).setValues(existing);
      SpreadsheetApp.flush();
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * _관리ID 열의 위치를 머리글 이름으로 찾습니다 (없으면 0).
 * 폼에 문항이 추가되어 열 위치가 밀려도 이름으로 찾으므로 영향받지 않습니다.
 */
function findUidCol_(sheet, lastCol) {
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
  return headers.indexOf(CONFIG.UID_HEADER) + 1;
}

/* ------------------------------------------------------------------ */
/* 헬퍼                                                                 */
/* ------------------------------------------------------------------ */

function checkToken_(token) {
  if (!CONFIG.API_TOKEN) return;
  if (String(token || '') !== CONFIG.API_TOKEN) {
    throw new Error('접근 토큰이 올바르지 않습니다.');
  }
}

function truthy_(v) {
  if (v === true) return true;
  var s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === 'y' || s === 'yes' || s === '1' || s === 'o' || s === '출석' || s === '완료';
}

function isBlankRow_(line, uidIdx) {
  for (var i = 0; i < line.length; i++) {
    if (i === uidIdx) continue;
    if (String(line[i] || '').trim() !== '') return false;
  }
  return true;
}

function errMsg_(err) {
  return String((err && err.message) || err);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* 스프레드시트 메뉴 (선택)                                              */
/* ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('신청자 관리')
    .addItem('관리ID 채우기 / 점검', 'menuEnsureUids')
    .addItem('관리 시트 보이기', 'menuShowAdminSheet')
    .addToUi();
}

function menuEnsureUids() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureUids_(responseSheet_(ss));
  adminSheet_(ss);
  SpreadsheetApp.getUi().alert('완료되었습니다.');
}

function menuShowAdminSheet() {
  var sheet = adminSheet_(SpreadsheetApp.getActiveSpreadsheet());
  sheet.showSheet();
  sheet.activate();
}
