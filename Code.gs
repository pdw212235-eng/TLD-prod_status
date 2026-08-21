// ============================================================
// TLD Product Status — Google Apps Script Backend v1.3
// ============================================================
// 설정: 아래 CONFIG 값을 환경에 맞게 수정하세요.
// ADMIN_PASSWORD 는 PropertiesService 로 관리 (스크립트 속성에서 설정).
// ============================================================

const CONFIG = {
  SPREADSHEET_ID: '1e5jRRL8Szm2yYDyeKvqj4QrPuNmPQ2VS-3hiKDlTFAs',
  SHEET_NAME: 'Installations',
  LISTS_SHEET: 'Lists',
  AUDIT_SHEET: 'AuditLog',
  AS_SHEET: 'AS_History',
  DRIVE_FOLDER_NAME: 'TLD_Product_Photos',
  TOKEN_EXPIRY_HOURS: 8,
};

// ============================================================
// 진입점
// ============================================================

function doGet(e) {
  const params = e.parameter || {};
  const action = params.action || 'list';
  const token  = params.token  || '';
  const isAdmin = validateToken(token);

  try {
    switch (action) {
      case 'list':    return jsonResponse(handleList(params, isAdmin));
      case 'get':     return jsonResponse(handleGet(params, isAdmin));
      case 'lists':   return jsonResponse(handleLists());
      case 'summary': return jsonResponse(handleSummary(isAdmin));
      case 'export':  return handleExport(params, isAdmin);
      case 'as-list': return jsonResponse(handleASList(params));
      default:        return jsonResponse({ error: '알 수 없는 action' }, 400);
    }
  } catch (err) {
    logError(err);
    return jsonResponse({ error: err.message }, 500);
  }
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (_) {
    return jsonResponse({ error: '잘못된 요청 본문' }, 400);
  }

  const action = body.action || '';
  const token  = body.token  || '';

  try {
    switch (action) {
      case 'google-auth': return jsonResponse(handleGoogleAuth(body));
      case 'logout':  return jsonResponse(handleLogout(body));
      case 'create':  requireAdmin(token); return jsonResponse(handleCreate(body));
      case 'update':  requireAdmin(token); return jsonResponse(handleUpdate(body));
      case 'retire':  requireAdmin(token); return jsonResponse(handleRetire(body));
      case 'upload':  requireAdmin(token); return jsonResponse(handleUpload(body));
      case 'as-add':  requireAdmin(token); return jsonResponse(handleASAdd(body));
      case 'as-delete': requireAdmin(token); return jsonResponse(handleASDelete(body));
      default:        return jsonResponse({ error: '알 수 없는 action' }, 400);
    }
  } catch (err) {
    logError(err);
    return jsonResponse({ error: err.message }, err.code || 500);
  }
}

// ============================================================
// 인증
// ============================================================

const GOOGLE_CLIENT_ID = '273196196166-r92mjb63o4dm3tns42j6k94738aedir8.apps.googleusercontent.com';

function handleGoogleAuth(body) {
  const idToken = body.idToken || '';
  if (!idToken) throw { message: 'idToken이 필요합니다.', code: 400 };

  // 구글 토큰 검증
  const res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) throw { message: '구글 토큰 검증 실패', code: 401 };

  const payload = JSON.parse(res.getContentText());
  if (payload.aud !== GOOGLE_CLIENT_ID) throw { message: '잘못된 클라이언트 ID', code: 401 };

  const email = (payload.email || '').toLowerCase();
  if (!email) throw { message: '이메일을 가져올 수 없습니다.', code: 401 };
  if (!isAdminEmail(email)) throw { message: '접근 권한이 없습니다: ' + email, code: 403 };

  const token = Utilities.getUuid();
  const expiry = new Date(Date.now() + CONFIG.TOKEN_EXPIRY_HOURS * 3600 * 1000).toISOString();
  const tokens = getTokenStore();
  tokens[token] = { expiry, email };
  saveTokenStore(tokens);
  writeAuditLog('LOGIN', '', email, '구글 로그인');
  return { token, expiry, email };
}

function isAdminEmail(email) {
  const sh = getListsSheet();
  if (!sh) return false;
  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const aIdx = headers.indexOf('Admins');
  if (aIdx < 0) return false;
  const admins = data.slice(1).map(r => String(r[aIdx]).trim().toLowerCase()).filter(Boolean);
  return admins.includes(email);
}

function handleLogout(body) {
  const tokens = getTokenStore();
  delete tokens[body.token];
  saveTokenStore(tokens);
  return { ok: true };
}

function validateToken(token) {
  if (!token) return false;
  const tokens = getTokenStore();
  const entry = tokens[token];
  if (!entry) return false;
  if (new Date(entry.expiry) < new Date()) {
    delete tokens[token];
    saveTokenStore(tokens);
    return false;
  }
  return true;
}

function requireAdmin(token) {
  if (!validateToken(token)) throw { message: '관리자 권한이 필요합니다.', code: 401 };
}

function getTokenStore() {
  const raw = PropertiesService.getScriptProperties().getProperty('TOKENS') || '{}';
  try { return JSON.parse(raw); } catch (_) { return {}; }
}

function saveTokenStore(tokens) {
  // 만료된 토큰 정리
  const now = new Date();
  Object.keys(tokens).forEach(k => {
    if (new Date(tokens[k].expiry) < now) delete tokens[k];
  });
  PropertiesService.getScriptProperties().setProperty('TOKENS', JSON.stringify(tokens));
}

// ============================================================
// 헤더 매핑
// ============================================================

const COLUMNS = [
  { key: 'id',           header: 'ID',              adminOnly: false },
  { key: 'manager',      header: '관리 주체',        adminOnly: false },
  { key: 'site',         header: 'Site 및 제품명',   adminOnly: false },
  { key: 'status',       header: '상태',             adminOnly: false },
  { key: 'panelModel',   header: 'Panel Model',      adminOnly: false },
  { key: 'panelArray',   header: 'Panel수량',        adminOnly: false },
  { key: 'panelTotal',   header: '총 모듈수',        adminOnly: false },
  { key: 'resX',         header: '해상도 X',         adminOnly: false },
  { key: 'resY',         header: '해상도 Y',         adminOnly: false },
  { key: 'brightness',   header: '휘도',             adminOnly: false },
  { key: 'installDate',  header: '설치 날짜',        adminOnly: false },
  { key: 'player',       header: 'MCU/Player',       adminOnly: false },
  { key: 'hubBoard',     header: 'Hub Board',        adminOnly: false },
  { key: 'smpsVoltage',  header: 'SMPS Output Voltage', adminOnly: false },
  { key: 'cmsId',        header: 'CMS ID',           adminOnly: false },
  { key: 'cmsPw',        header: 'CMS PW',           adminOnly: true  },
  { key: 'tvId',         header: 'Teamviewer ID',    adminOnly: true  },
  { key: 'tvPw',         header: 'Teamviewer PW',    adminOnly: true  },
  { key: 'inboundDate',  header: '입고일',           adminOnly: false },
  { key: 'outboundDate', header: '반출일',           adminOnly: false },
  { key: 'note1',        header: '비고1',            adminOnly: false },
  { key: 'note2',        header: '비고2',            adminOnly: false },
  { key: 'note3',        header: '비고3',            adminOnly: false },
  { key: 'note4',        header: '비고4',            adminOnly: false },
  { key: 'photo1',       header: '사진1',            adminOnly: false },
  { key: 'photo2',       header: '사진2',            adminOnly: false },
  { key: 'photo3',       header: '사진3',            adminOnly: false },
  { key: 'updatedAt',    header: '최종수정일시',     adminOnly: false },
  { key: 'updatedBy',    header: '최종수정자',       adminOnly: false },
];

const ADMIN_KEYS = new Set(COLUMNS.filter(c => c.adminOnly).map(c => c.key));

// 글자수 제한이 있는 필드. 프론트의 maxlength 를 우회한 요청도 여기서 막는다.
const FIELD_MAX = { hubBoard: 40, smpsVoltage: 10 };

function validateFieldLengths(data) {
  Object.keys(FIELD_MAX).forEach(k => {
    const v = data[k];
    if (v === undefined || v === null || v === '') return;
    if (String(v).length > FIELD_MAX[k]) {
      throw { message: k + ' 는 ' + FIELD_MAX[k] + '자 이내여야 합니다.', code: 400 };
    }
  });
}

function getSheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
}

function getListsSheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LISTS_SHEET);
}

function getAuditSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.AUDIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.AUDIT_SHEET);
    sh.appendRow(['타임스탬프', '액션', 'ID', '사용자', '메모']);
  }
  return sh;
}

// 시트 헤더를 읽어 컬럼 인덱스 맵 생성
function buildColIndex(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[h.toString().trim()] = i; });
  return map;
}

function rowToObj(row, colIndex, isAdmin) {
  const obj = {};
  COLUMNS.forEach(col => {
    if (!isAdmin && col.adminOnly) return;
    const i = colIndex[col.header];
    obj[col.key] = (i !== undefined) ? row[i] : '';
  });
  // 날짜는 문자열로
  ['installDate','inboundDate','outboundDate','updatedAt'].forEach(k => {
    if (obj[k] instanceof Date) {
      obj[k] = obj[k] ? Utilities.formatDate(obj[k], Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    }
  });
  // 숫자
  ['resX','resY','brightness','panelTotal'].forEach(k => {
    if (obj[k] !== undefined && obj[k] !== '') obj[k] = Number(obj[k]) || obj[k];
  });
  return obj;
}

// ============================================================
// 목록 / 단건 조회
// ============================================================

function handleList(params, isAdmin) {
  const sheet = getSheet();
  if (sheet.getLastRow() < 2) return { items: [] };
  const colIndex = buildColIndex(sheet);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  let items = data.map(row => rowToObj(row, colIndex, isAdmin));

  // 반출 건: 비관리자는 제외
  if (!isAdmin) {
    items = items.filter(it => it.status !== '반출');
  }

  // 필터
  if (params.manager)    items = items.filter(it => it.manager === params.manager);
  if (params.status)     items = items.filter(it => it.status  === params.status);
  if (params.panelModel) items = items.filter(it => (it.panelModel||'').includes(params.panelModel));
  if (params.site) {
    const q = params.site.toLowerCase();
    items = items.filter(it => (it.site||'').toLowerCase().includes(q));
  }

  return { items };
}

function handleGet(params, isAdmin) {
  const id = params.id || '';
  if (!id) throw { message: 'id 파라미터가 필요합니다.', code: 400 };

  const sheet = getSheet();
  const colIndex = buildColIndex(sheet);
  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), sheet.getLastColumn()).getValues();

  const idCol = colIndex['ID'];
  const row = data.find(r => String(r[idCol]).trim() === id);
  if (!row) throw { message: '항목을 찾을 수 없습니다.', code: 404 };

  const item = rowToObj(row, colIndex, isAdmin);
  if (!isAdmin && item.status === '반출') throw { message: '항목을 찾을 수 없습니다.', code: 404 };
  return { item };
}

// ============================================================
// 드롭다운 옵션 (Lists 시트)
// ============================================================

function handleLists() {
  const sh = getListsSheet();
  const statuses = ['현장설치', '렌트', '창고', '반출'];

  if (!sh) return { managers: [], statuses, panelModels: [] };

  const data = sh.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());
  const mIdx = headers.indexOf('Managers');
  const pIdx = headers.indexOf('PanelModels');

  const col = (idx) => idx < 0 ? [] : data.slice(1).map(r => r[idx]).filter(v => v !== '').map(String);

  return {
    managers: col(mIdx),
    statuses,
    panelModels: col(pIdx),
  };
}

// ============================================================
// 요약 (대시보드)
// ============================================================

function handleSummary(isAdmin) {
  const { items } = handleList({}, isAdmin);

  const byStatus = {};
  const byManager = {};
  const byModel = {};
  let total = 0;

  items.forEach(it => {
    const st = it.status || '미입력';
    byStatus[st] = (byStatus[st] || 0) + 1;
    const mg = it.manager || '미입력';
    byManager[mg] = (byManager[mg] || 0) + 1;
    const pm = it.panelModel || '미입력';
    byModel[pm] = (byModel[pm] || 0) + 1;
    total++;
  });

  return { byStatus, byManager, byModel, total };
}

// ============================================================
// 생성 / 수정 / 반출
// ============================================================

function handleCreate(body) {
  const data = body.data || {};
  const sheet = getSheet();
  ensureHeaders(sheet);
  const colIndex = buildColIndex(sheet);

  // ID 생성
  const newId = generateId(sheet, colIndex);
  data.id = newId;
  data.updatedAt = new Date().toISOString();
  data.updatedBy = body.updatedBy || 'admin';

  validateFieldLengths(data);

  // panelTotal 자동계산
  if (data.panelArray) data.panelTotal = calcPanelTotal(data.panelArray);

  const row = buildRow(data, colIndex, sheet.getLastColumn());
  sheet.appendRow(row);
  writeAuditLog('CREATE', newId, data.updatedBy, 'Site: ' + (data.site || ''));
  return { id: newId };
}

function handleUpdate(body) {
  const id = body.id || '';
  if (!id) throw { message: 'id가 필요합니다.', code: 400 };

  const sheet = getSheet();
  const colIndex = buildColIndex(sheet);
  const { rowIdx, existing } = findRowById(sheet, colIndex, id);

  // 낙관적 잠금
  if (body.updatedAt && existing.updatedAt && body.updatedAt !== existing.updatedAt) {
    throw { message: '다른 사람이 이미 수정했습니다. 새로고침 후 다시 시도하세요.', code: 409 };
  }

  const data = Object.assign({}, existing, body.data || {});
  data.id = id;
  data.updatedAt = new Date().toISOString();
  data.updatedBy = body.updatedBy || 'admin';
  validateFieldLengths(data);
  if (data.panelArray) data.panelTotal = calcPanelTotal(data.panelArray);

  const row = buildRow(data, colIndex, sheet.getLastColumn());
  sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  writeAuditLog('UPDATE', id, data.updatedBy, '');
  return { ok: true };
}

function handleRetire(body) {
  const id = body.id || '';
  if (!id) throw { message: 'id가 필요합니다.', code: 400 };

  const sheet = getSheet();
  const colIndex = buildColIndex(sheet);
  const { rowIdx, existing } = findRowById(sheet, colIndex, id);

  existing.status = '반출';
  existing.outboundDate = body.outboundDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  existing.updatedAt = new Date().toISOString();
  existing.updatedBy = body.updatedBy || 'admin';

  const row = buildRow(existing, colIndex, sheet.getLastColumn());
  sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  writeAuditLog('RETIRE', id, existing.updatedBy, '반출일: ' + existing.outboundDate);
  return { ok: true };
}

// ============================================================
// 사진 업로드
// ============================================================

function handleUpload(body) {
  const id   = body.id   || '';
  const slot = parseInt(body.slot || '1');
  const b64  = body.fileBase64 || '';
  const mime = body.mime || 'image/jpeg';

  if (!id || !b64) throw { message: 'id, fileBase64 필수', code: 400 };
  if (slot < 1 || slot > 3) throw { message: 'slot은 1~3', code: 400 };

  // 크기 검증 (base64 → 바이트 수 근사)
  const byteSize = Math.ceil(b64.length * 0.75);
  if (byteSize > 5 * 1024 * 1024) throw { message: '사진 1매 최대 5MB', code: 400 };

  const folder = getOrCreateDriveFolder();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const filename = `${id}_${slot}_${ts}.jpg`;

  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = `https://drive.google.com/uc?id=${file.getId()}`;

  // 시트에 기록
  const sheet = getSheet();
  const colIndex = buildColIndex(sheet);
  const { rowIdx, existing } = findRowById(sheet, colIndex, id);
  existing[`photo${slot}`] = url;
  existing.updatedAt = new Date().toISOString();
  existing.updatedBy = body.updatedBy || 'admin';
  const row = buildRow(existing, colIndex, sheet.getLastColumn());
  sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);

  writeAuditLog('UPLOAD', id, existing.updatedBy, `사진${slot}: ${file.getId()}`);
  return { url };
}

function getOrCreateDriveFolder() {
  const folders = DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
}

// ============================================================
// CSV 내보내기
// ============================================================

function handleExport(params, isAdmin) {
  if (!isAdmin) {
    return ContentService.createTextOutput('권한 없음').setMimeType(ContentService.MimeType.TEXT);
  }

  const includeSecret = params.includeSecret === 'true';
  const { items } = handleList(params, true); // 관리자로 전체 읽기

  const visibleCols = COLUMNS.filter(c => isAdmin && includeSecret ? true : !c.adminOnly);
  const headers = visibleCols.map(c => c.header);

  const rows = items.map(item =>
    visibleCols.map(c => {
      const v = item[c.key];
      if (v === null || v === undefined) return '';
      const s = String(v);
      return s.includes(',') || s.includes('"') || s.includes('\n')
        ? `"${s.replace(/"/g, '""')}"`
        : s;
    }).join(',')
  );

  const bom = '﻿';
  const csv = bom + [headers.join(','), ...rows].join('\r\n');
  writeAuditLog('EXPORT', '', 'admin', `${items.length}건, includeSecret=${includeSecret}`);

  return ContentService.createTextOutput(csv).setMimeType(ContentService.MimeType.CSV);
}

// ============================================================
// 유틸리티
// ============================================================

// 관리번호 채번 규칙: Prod + 3자리 (Prod001, Prod002, …)
// 시트에 있는 가장 큰 번호 + 1. 반출 건도 행이 남으므로 번호는 재사용되지 않습니다.
const ID_PREFIX = 'Prod';
const ID_PAD    = 3;
const ID_RE     = /^Prod(\d+)$/i;

function generateId(sheet, colIndex) {
  const first = ID_PREFIX + '1'.padStart(ID_PAD, '0');
  if (sheet.getLastRow() < 2) return first;

  const idColIdx = colIndex['ID'];
  if (idColIdx === undefined) return first;

  const nums = sheet.getRange(2, idColIdx + 1, sheet.getLastRow() - 1, 1).getValues()
    .flat()
    .map(v => {
      const m = String(v).trim().match(ID_RE);
      return m ? parseInt(m[1], 10) : NaN;
    })
    .filter(n => !isNaN(n));

  if (nums.length === 0) return first;
  return ID_PREFIX + String(Math.max(...nums) + 1).padStart(ID_PAD, '0');
}

function calcPanelTotal(panelArray) {
  if (!panelArray) return '';
  const parts = String(panelArray).split('*').map(Number);
  if (parts.some(isNaN)) return '';
  return parts.reduce((a, b) => a * b, 1);
}

function findRowById(sheet, colIndex, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw { message: '데이터 없음', code: 404 };
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const idColIdx = colIndex['ID'];
  const idx = data.findIndex(r => String(r[idColIdx]).trim() === id);
  if (idx < 0) throw { message: `ID '${id}' 를 찾을 수 없습니다.`, code: 404 };
  const existing = rowToObj(data[idx], colIndex, true);
  return { rowIdx: idx + 2, existing };
}

function buildRow(data, colIndex, colCount) {
  const row = new Array(colCount).fill('');
  COLUMNS.forEach(col => {
    const i = colIndex[col.header];
    if (i !== undefined && data[col.key] !== undefined) {
      row[i] = data[col.key] === null ? '' : data[col.key];
    }
  });
  return row;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.appendRow(COLUMNS.map(c => c.header));
  }
}

// ============================================================
// A/S 이력
// ============================================================

function getASSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.AS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.AS_SHEET);
    sh.appendRow(['rowId', 'ID', '날짜', '내용', '등록일시', '등록자']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function handleASList(params) {
  const id = params.id || '';
  if (!id) throw { message: 'id가 필요합니다.', code: 400 };
  const sh = getASSheet();
  if (sh.getLastRow() < 2) return { items: [] };
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  const items = data
    .filter(r => String(r[1]).trim() === id)
    .map(r => ({
      rowId: String(r[0]),
      id: String(r[1]),
      date: r[2] instanceof Date ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[2]),
      text: String(r[3]),
      createdAt: String(r[4]),
      createdBy: String(r[5]),
    }))
    .sort((a, b) => b.date.localeCompare(a.date));
  return { items };
}

function handleASAdd(body) {
  const id   = body.id   || '';
  const date = body.date || '';
  const text = body.text || '';
  if (!id || !date || !text) throw { message: 'id, date, text 필수', code: 400 };
  if (text.length > 200) throw { message: '내용은 200자 이내', code: 400 };

  const sh = getASSheet();
  const rowId = Utilities.getUuid();
  sh.appendRow([rowId, id, date, text, new Date().toISOString(), body.updatedBy || 'admin']);
  writeAuditLog('AS_ADD', id, body.updatedBy || 'admin', date + ': ' + text.slice(0, 30));
  return { rowId };
}

function handleASDelete(body) {
  const rowId = body.rowId || '';
  if (!rowId) throw { message: 'rowId가 필요합니다.', code: 400 };

  const sh = getASSheet();
  if (sh.getLastRow() < 2) throw { message: '데이터 없음', code: 404 };
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const idx = data.findIndex(r => String(r[0]) === rowId);
  if (idx < 0) throw { message: '이력을 찾을 수 없습니다.', code: 404 };
  sh.deleteRow(idx + 2);
  writeAuditLog('AS_DELETE', '', body.updatedBy || 'admin', 'rowId: ' + rowId);
  return { ok: true };
}

function writeAuditLog(action, id, user, memo) {
  try {
    const sh = getAuditSheet();
    sh.appendRow([new Date().toISOString(), action, id, user, memo]);
  } catch (_) {}
}

function jsonResponse(obj, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

function logError(err) {
  console.error('[TLD Error]', err.message || err);
}

// ============================================================
// 초기 설정 함수 (한 번만 실행)
// ============================================================

/**
 * 스크립트 편집기에서 이 함수를 한 번 실행하세요.
 * 실행 전: 스크립트 속성(Project Settings > Script Properties)에
 *   ADMIN_PASSWORD = 원하는비밀번호
 * 를 추가하세요.
 */
/**
 * Hub Board / SMPS Output Voltage 열 추가. 편집기에서 한 번만 실행하세요.
 * 이미 있으면 아무것도 하지 않으므로 여러 번 실행해도 안전합니다.
 */
function addComponentColumns() {
  const sheet = getSheet();
  const targets = ['Hub Board', 'SMPS Output Voltage'];
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = targets.filter(h => headers.indexOf(h) < 0);

  if (missing.length === 0) {
    Logger.log('이미 두 열이 모두 있습니다. 변경 없음.');
    return;
  }
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  SpreadsheetApp.flush();
  Logger.log('추가된 열: ' + missing.join(', '));
}

function setupSheetHeaders() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Installations 시트
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) sh = ss.insertSheet(CONFIG.SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNS.map(c => c.header));
    sh.setFrozenRows(1);
  }

  // Lists 시트
  let lists = ss.getSheetByName(CONFIG.LISTS_SHEET);
  if (!lists) {
    lists = ss.insertSheet(CONFIG.LISTS_SHEET);
    lists.appendRow(['Managers', 'Statuses', 'PanelModels']);
    lists.getRange('A2:A10').setValues([
      ['DW'],['엘리트전자'],['레브컴'],['디플럭스'],[''],[''],[''],[''],['']
    ]);
    lists.getRange('B2:B5').setValues([
      ['현장설치'],['렌트'],['창고'],['반출']
    ]);
  }

  // AuditLog 시트
  let audit = ss.getSheetByName(CONFIG.AUDIT_SHEET);
  if (!audit) {
    audit = ss.insertSheet(CONFIG.AUDIT_SHEET);
    audit.appendRow(['타임스탬프', '액션', 'ID', '사용자', '메모']);
    audit.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  Logger.log('시트 초기화 완료');
}
