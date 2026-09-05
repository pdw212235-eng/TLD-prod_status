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
  REPORT_SHEET: 'Reports',
  PROC_PROJECT_SHEET: 'ProcessProjects',
  PROC_ITEM_SHEET: 'ProcessItems',
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
      // 표준 프로세스 진행 관리 (process.html) — 조회는 로그인 없이 가능
      case 'proc-projects': return jsonResponse(handleProcProjects());
      case 'proc-items':    return jsonResponse(handleProcItems(params));
      // 접수 내용 조회는 관리자만 (접수 자체는 로그인 없이 가능)
      case 'report-list': requireAdmin(token); return jsonResponse(handleReportList());
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
      // 수정 요청 접수는 누구나 가능 (로그인 불필요)
      case 'report-add': return jsonResponse(handleReportAdd(body));
      case 'report-resolve': requireAdmin(token); return jsonResponse(handleReportResolve(body));
      // 표준 프로세스 진행 관리 — 등록/수정은 누구나, 프로젝트 삭제만 관리자
      case 'proc-project-save':   return jsonResponse(handleProcProjectSave(body));
      case 'proc-item-save':      return jsonResponse(handleProcItemSave(body));
      case 'proc-project-delete': requireAdmin(token); return jsonResponse(handleProcProjectDelete(body));
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
  { key: 'smpsModel',    header: 'SMPS Model',       adminOnly: false },
  { key: 'smpsVoltage',  header: 'SMPS Output Voltage', adminOnly: false },
  { key: 'smpsQty',      header: 'SMPS Quantity',     adminOnly: false },
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
const FIELD_MAX = { hubBoard: 40, smpsModel: 40, smpsVoltage: 10 };

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
  ['resX','resY','brightness','panelTotal','smpsQty'].forEach(k => {
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
  // drive.google.com/uc?id= 형식은 리다이렉트 때문에 브라우저 <img> 에서 로드에 실패한다.
  // (curl -L 로는 image 가 반환되어 정상처럼 보이므로 curl 로 검증하지 말 것)
  const url = `https://lh3.googleusercontent.com/d/${file.getId()}`;

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

// 설치 및 수리 이력. 시트 이름(AS_History)은 그대로 두고 화면 표기만 바꿨다.
const HISTORY_COLS = ['rowId', 'ID', 'Site/제품명', '날짜', '구분', '내용', '등록일시', '등록자'];
const HISTORY_KINDS = ['설치', '수리'];

function getASSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.AS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.AS_SHEET);
    sh.appendRow(HISTORY_COLS);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 위치가 아니라 헤더 이름으로 찾는다. 열을 옮겨도 동작한다.
function asColIndex(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i; });
  return map;
}

// 관리번호로 Site/제품명 찾기. 이력 등록 시 시트에 함께 적어 둔다.
function siteOfId(id) {
  try {
    const sheet = getSheet();
    const { existing } = findRowById(sheet, buildColIndex(sheet), id);
    return existing.site || '';
  } catch (_) {
    return '';
  }
}

function handleASList(params) {
  const id = params.id || '';
  if (!id) throw { message: 'id가 필요합니다.', code: 400 };

  const sh = getASSheet();
  if (sh.getLastRow() < 2) return { items: [] };

  const ci = asColIndex(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const at = (r, name) => {
    const i = ci[name];
    return i === undefined ? '' : r[i];
  };
  const asDate = v => v instanceof Date
    ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(v || '');

  const items = data
    .filter(r => String(at(r, 'ID')).trim() === id)
    .map(r => ({
      rowId:     String(at(r, 'rowId')),
      id:        String(at(r, 'ID')),
      site:      String(at(r, 'Site/제품명') || ''),
      date:      asDate(at(r, '날짜')),
      kind:      String(at(r, '구분') || ''),
      text:      String(at(r, '내용')),
      createdAt: String(at(r, '등록일시') || ''),
      createdBy: String(at(r, '등록자') || ''),
    }))
    .sort((a, b) => b.date.localeCompare(a.date));

  return { items };
}

function handleASAdd(body) {
  const id   = body.id   || '';
  const date = body.date || '';
  const text = body.text || '';
  const kind = body.kind || '';

  if (!id || !date || !text) throw { message: 'id, date, text 필수', code: 400 };
  if (text.length > 200) throw { message: '내용은 200자 이내', code: 400 };
  if (kind && HISTORY_KINDS.indexOf(kind) < 0) {
    throw { message: '구분은 ' + HISTORY_KINDS.join(' 또는 ') + ' 만 가능합니다.', code: 400 };
  }

  const sh = getASSheet();
  const ci = asColIndex(sh);
  const rowId = Utilities.getUuid();

  const row = new Array(sh.getLastColumn()).fill('');
  const put = (name, v) => { if (ci[name] !== undefined) row[ci[name]] = v; };
  put('rowId', rowId);
  put('ID', id);
  put('Site/제품명', siteOfId(id));
  put('날짜', date);
  put('구분', kind);
  put('내용', text);
  put('등록일시', new Date().toISOString());
  put('등록자', body.updatedBy || 'admin');

  sh.appendRow(row);
  writeAuditLog('AS_ADD', id, body.updatedBy || 'admin', (kind ? '[' + kind + '] ' : '') + date + ': ' + text.slice(0, 30));
  return { rowId };
}

function handleASDelete(body) {
  const rowId = body.rowId || '';
  if (!rowId) throw { message: 'rowId가 필요합니다.', code: 400 };

  const sh = getASSheet();
  if (sh.getLastRow() < 2) throw { message: '데이터 없음', code: 404 };
  const ci = asColIndex(sh);
  const data = sh.getRange(2, ci['rowId'] + 1, sh.getLastRow() - 1, 1).getValues();
  const idx = data.findIndex(r => String(r[0]) === rowId);
  if (idx < 0) throw { message: '이력을 찾을 수 없습니다.', code: 404 };
  sh.deleteRow(idx + 2);
  writeAuditLog('AS_DELETE', '', body.updatedBy || 'admin', 'rowId: ' + rowId);
  return { ok: true };
}

// ============================================================
// 수정 요청 / 버그 리포트
// ============================================================

const REPORT_COLS = ['rowId', '접수일시', '작성자', '내용', '처리완료', '처리자', '처리일시'];

// 시트가 없으면 만들고, 있는데 열이 빠져 있으면 뒤에 채운다.
// 처리완료 열을 나중에 추가했기 때문에 구버전 시트도 자동으로 보정된다.
function getReportSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.REPORT_SHEET);

  if (!sh) {
    sh = ss.insertSheet(CONFIG.REPORT_SHEET);
    sh.appendRow(REPORT_COLS);
    sh.setFrozenRows(1);
    sh.setColumnWidth(4, 520);
    return sh;
  }

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = REPORT_COLS.filter(h => headers.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    SpreadsheetApp.flush();
  }
  return sh;
}

// Installations 와 마찬가지로 위치가 아니라 헤더 이름으로 찾는다
function reportColIndex(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i; });
  return map;
}

function isTrue(v) {
  return v === true || String(v).trim().toUpperCase() === 'TRUE';
}

const REPORT_MAX = { author: 40, text: 1000 };

function handleReportAdd(body) {
  const author = String(body.author || '').trim();
  const text   = String(body.text   || '').trim();

  if (!text) throw { message: '내용을 입력해 주세요.', code: 400 };
  if (author.length > REPORT_MAX.author) {
    throw { message: '작성자는 ' + REPORT_MAX.author + '자 이내여야 합니다.', code: 400 };
  }
  if (text.length > REPORT_MAX.text) {
    throw { message: '내용은 ' + REPORT_MAX.text + '자 이내여야 합니다.', code: 400 };
  }

  const sh = getReportSheet();
  const ci = reportColIndex(sh);
  const rowId = Utilities.getUuid();

  const row = new Array(sh.getLastColumn()).fill('');
  row[ci['rowId']]   = rowId;
  row[ci['접수일시']] = new Date().toISOString();
  row[ci['작성자']]   = author || '(익명)';
  row[ci['내용']]     = text;

  sh.appendRow(row);
  writeAuditLog('REPORT', '', author || '(익명)', text.slice(0, 50));
  return { ok: true, rowId };
}

// 최신순 100건. 접수는 누구나 가능하지만 조회는 관리자만.
function handleReportList() {
  const sh = getReportSheet();
  if (sh.getLastRow() < 2) return { items: [], total: 0, open: 0 };

  const ci = reportColIndex(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const at = (r, name) => {
    const i = ci[name];
    return i === undefined ? '' : r[i];
  };

  const items = data.map(r => ({
    rowId:      String(at(r, 'rowId')),
    createdAt:  at(r, '접수일시') instanceof Date ? at(r, '접수일시').toISOString() : String(at(r, '접수일시')),
    author:     String(at(r, '작성자')),
    text:       String(at(r, '내용')),
    resolved:   isTrue(at(r, '처리완료')),
    resolvedBy: String(at(r, '처리자') || ''),
    resolvedAt: at(r, '처리일시') instanceof Date ? at(r, '처리일시').toISOString() : String(at(r, '처리일시') || ''),
  }));

  const open = items.filter(i => !i.resolved).length;
  items.reverse();
  return { items: items.slice(0, 100), total: items.length, open };
}

// 처리완료 표시 토글 (관리자 전용)
function handleReportResolve(body) {
  const rowId = String(body.rowId || '');
  if (!rowId) throw { message: 'rowId가 필요합니다.', code: 400 };
  const resolved = body.resolved !== false;   // 기본값 true

  const sh = getReportSheet();
  if (sh.getLastRow() < 2) throw { message: '데이터 없음', code: 404 };

  const ci = reportColIndex(sh);
  const idCol = ci['rowId'];
  const ids = sh.getRange(2, idCol + 1, sh.getLastRow() - 1, 1).getValues();
  const idx = ids.findIndex(r => String(r[0]) === rowId);
  if (idx < 0) throw { message: '접수 내역을 찾을 수 없습니다.', code: 404 };

  const rowNum = idx + 2;
  const who = body.updatedBy || 'admin';
  sh.getRange(rowNum, ci['처리완료'] + 1).setValue(resolved ? true : '');
  sh.getRange(rowNum, ci['처리자']   + 1).setValue(resolved ? who : '');
  sh.getRange(rowNum, ci['처리일시'] + 1).setValue(resolved ? new Date().toISOString() : '');

  writeAuditLog(resolved ? 'REPORT_RESOLVE' : 'REPORT_REOPEN', '', who, 'rowId: ' + rowId);
  return { ok: true };
}

// ============================================================
// 전기공사 표준 프로세스 진행 관리 (process.html)
// ------------------------------------------------------------
// ProcessProjects : 프로젝트 1행
// ProcessItems    : (프로젝트 × 항목) 1행. 항목ID '__checks' 는
//                   그 프로젝트의 체크리스트 상태를 링크JSON 에 배열로 담는다.
// 조회·등록은 로그인 없이 가능하고, 프로젝트 삭제만 관리자 전용이다.
// ============================================================

const PROC_PROJECT_COLS = ['id', '프로젝트명', '발주처', '현장', '착수일', '준공목표일', 'PM', '등록일시', '등록자'];
const PROC_ITEM_COLS = ['rowId', '프로젝트', '항목ID', '단계', '항목명', '상태',
                        '담당자', '기한', '완료일', '링크JSON', '비고', '수정일시', '수정자'];

const PROC_MAX = { name: 120, text: 200, json: 8000, note: 500 };

// Reports 시트와 같은 방식: 없으면 만들고, 열이 빠졌으면 뒤에 채운다.
function getProcSheet_(sheetName, cols){
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(cols);
    sh.setFrozenRows(1);
    return sh;
  }
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = cols.filter(h => headers.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    SpreadsheetApp.flush();
  }
  return sh;
}

function getProcProjectSheet(){ return getProcSheet_(CONFIG.PROC_PROJECT_SHEET, PROC_PROJECT_COLS); }
function getProcItemSheet()   { return getProcSheet_(CONFIG.PROC_ITEM_SHEET,    PROC_ITEM_COLS); }

function procColIndex_(sh){
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h).trim()] = i; });
  return map;
}

// 시트가 날짜로 바꿔 놓은 값도 화면에서 쓰는 문자열로 되돌린다.
function procStr_(v){
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}

function procTrim_(v, max){
  const s = String(v == null ? '' : v).trim();
  if (s.length > max) throw { message: '입력이 너무 깁니다 (' + max + '자 이내).', code: 400 };
  return s;
}

// ------------------------------------------------------------
// 조회
// ------------------------------------------------------------
function handleProcProjects(){
  const sh = getProcProjectSheet();
  if (sh.getLastRow() < 2) return { projects: [] };

  const ci = procColIndex_(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const at = (r, n) => (ci[n] === undefined ? '' : procStr_(r[ci[n]]));

  const projects = data
    .filter(r => at(r, 'id'))
    .map(r => ({
      id:         at(r, 'id'),
      name:       at(r, '프로젝트명'),
      client:     at(r, '발주처'),
      site:       at(r, '현장'),
      startDate:  at(r, '착수일'),
      targetDate: at(r, '준공목표일'),
      pm:         at(r, 'PM'),
      createdAt:  at(r, '등록일시'),
      createdBy:  at(r, '등록자'),
    }));

  return { projects: projects, total: projects.length };
}

// project 파라미터가 없으면 전체(전체 현황 화면에서 쓴다).
function handleProcItems(params){
  const filter = String((params && params.project) || '').trim();
  const sh = getProcItemSheet();
  if (sh.getLastRow() < 2) return { items: [] };

  const ci = procColIndex_(sh);
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const at = (r, n) => (ci[n] === undefined ? '' : procStr_(r[ci[n]]));

  const items = data
    .filter(r => at(r, '프로젝트') && at(r, '항목ID'))
    .filter(r => !filter || at(r, '프로젝트') === filter)
    .map(r => ({
      project:   at(r, '프로젝트'),
      itemId:    at(r, '항목ID'),
      stage:     at(r, '단계'),
      name:      at(r, '항목명'),
      status:    at(r, '상태'),
      owner:     at(r, '담당자'),
      due:       at(r, '기한'),
      done:      at(r, '완료일'),
      links:     at(r, '링크JSON'),
      note:      at(r, '비고'),
      updatedAt: at(r, '수정일시'),
      updatedBy: at(r, '수정자'),
    }));

  return { items: items, total: items.length };
}

// ------------------------------------------------------------
// 등록 · 수정
// ------------------------------------------------------------
function handleProcProjectSave(body){
  const p = body.project || {};
  const name = procTrim_(p.name, PROC_MAX.name);
  if (!name) throw { message: '프로젝트명을 입력해 주세요.', code: 400 };

  const author = procTrim_(body.author, PROC_MAX.text) || '(익명)';
  const id = procTrim_(p.id, 64) || Utilities.getUuid();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getProcProjectSheet();
    const ci = procColIndex_(sh);
    const row = new Array(sh.getLastColumn()).fill('');
    row[ci['id']]         = id;
    row[ci['프로젝트명']]   = name;
    row[ci['발주처']]      = procTrim_(p.client, PROC_MAX.text);
    row[ci['현장']]        = procTrim_(p.site, PROC_MAX.text);
    row[ci['착수일']]      = procTrim_(p.startDate, 20);
    row[ci['준공목표일']]   = procTrim_(p.targetDate, 20);
    row[ci['PM']]         = procTrim_(p.pm, PROC_MAX.text);
    row[ci['등록일시']]     = new Date().toISOString();
    row[ci['등록자']]      = author;

    const rowNum = procFindRow_(sh, ci['id'], id);
    if (rowNum > 0) sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    else            sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  writeAuditLog('PROC_PROJECT', id, author, name);
  return { ok: true, id: id };
}

// (프로젝트, 항목ID) 조합으로 upsert 한다.
function handleProcItemSave(body){
  const project = procTrim_(body.project, 64);
  const itemId  = procTrim_(body.itemId, 64);
  if (!project) throw { message: '프로젝트가 필요합니다.', code: 400 };
  if (!itemId)  throw { message: '항목ID가 필요합니다.', code: 400 };

  const links = String(body.links == null ? '[]' : body.links);
  if (links.length > PROC_MAX.json) {
    throw { message: '링크가 너무 많습니다. 일부를 정리해 주세요.', code: 400 };
  }
  try { JSON.parse(links); } catch (_) { throw { message: '링크 형식이 잘못되었습니다.', code: 400 }; }

  const author = procTrim_(body.author, PROC_MAX.text) || '(익명)';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getProcItemSheet();
    const ci = procColIndex_(sh);
    const rowNum = procFindItemRow_(sh, ci, project, itemId);

    const row = new Array(sh.getLastColumn()).fill('');
    row[ci['rowId']]    = rowNum > 0
      ? procStr_(sh.getRange(rowNum, ci['rowId'] + 1).getValue()) || Utilities.getUuid()
      : Utilities.getUuid();
    row[ci['프로젝트']]  = project;
    row[ci['항목ID']]   = itemId;
    row[ci['단계']]     = procTrim_(body.stage, 10);
    row[ci['항목명']]    = procTrim_(body.name, PROC_MAX.name);
    row[ci['상태']]     = procTrim_(body.status, 20);
    row[ci['담당자']]    = procTrim_(body.owner, PROC_MAX.text);
    row[ci['기한']]     = procTrim_(body.due, 20);
    row[ci['완료일']]    = procTrim_(body.done, 20);
    row[ci['링크JSON']] = links;
    row[ci['비고']]     = procTrim_(body.note, PROC_MAX.note);
    row[ci['수정일시']]  = new Date().toISOString();
    row[ci['수정자']]    = author;

    if (rowNum > 0) sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    else            sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  return { ok: true };
}

// 프로젝트와 그 항목을 함께 지운다. 되돌릴 수 없으므로 관리자 전용.
function handleProcProjectDelete(body){
  const id = procTrim_(body.id, 64);
  if (!id) throw { message: 'id가 필요합니다.', code: 400 };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let removed = 0;
  try {
    const psh = getProcProjectSheet();
    const pci = procColIndex_(psh);
    const rowNum = procFindRow_(psh, pci['id'], id);
    if (rowNum < 0) throw { message: '프로젝트를 찾을 수 없습니다.', code: 404 };
    psh.deleteRow(rowNum);

    const ish = getProcItemSheet();
    if (ish.getLastRow() >= 2) {
      const ici = procColIndex_(ish);
      const col = ish.getRange(2, ici['프로젝트'] + 1, ish.getLastRow() - 1, 1).getValues();
      for (let i = col.length - 1; i >= 0; i--) {
        if (procStr_(col[i][0]) === id) { ish.deleteRow(i + 2); removed++; }
      }
    }
  } finally {
    lock.releaseLock();
  }

  writeAuditLog('PROC_PROJECT_DELETE', id, body.updatedBy || 'admin', '항목 ' + removed + '건 삭제');
  return { ok: true, removed: removed };
}

// ------------------------------------------------------------
// 행 찾기
// ------------------------------------------------------------
function procFindRow_(sh, colIdx, value){
  if (sh.getLastRow() < 2) return -1;
  const vals = sh.getRange(2, colIdx + 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (procStr_(vals[i][0]) === value) return i + 2;
  }
  return -1;
}

function procFindItemRow_(sh, ci, project, itemId){
  if (sh.getLastRow() < 2) return -1;
  const n = sh.getLastRow() - 1;
  const pCol = sh.getRange(2, ci['프로젝트'] + 1, n, 1).getValues();
  const iCol = sh.getRange(2, ci['항목ID'] + 1, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    if (procStr_(pCol[i][0]) === project && procStr_(iCol[i][0]) === itemId) return i + 2;
  }
  return -1;
}

// 시트를 미리 만들어 두고 싶을 때 스크립트 편집기에서 한 번 실행한다.
function setupProcessSheets(){
  getProcProjectSheet();
  getProcItemSheet();
  Logger.log('ProcessProjects / ProcessItems 시트 준비 완료');
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
 * COLUMNS 에는 있는데 시트에 없는 헤더를 찾아 맨 뒤에 추가합니다.
 * 편집기에서 실행하세요. 빠진 열이 없으면 아무것도 하지 않으므로 여러 번 실행해도 안전합니다.
 * 앞으로 항목을 추가할 때도 COLUMNS 에만 넣고 이 함수를 실행하면 됩니다.
 */
/**
 * 설치 및 수리 이력 시트에 'Site/제품명'(ID 뒤)과 '구분'(날짜 뒤) 열을 끼워 넣고,
 * 기존 행의 Site/제품명 을 관리번호로 찾아 채웁니다. 편집기에서 실행하세요.
 * 이미 있으면 건너뛰므로 여러 번 실행해도 안전합니다.
 */
function migrateHistorySheet() {
  const sh = getASSheet();
  const added = [];

  // 지정한 열 바로 뒤에 삽입한다 (기존 값은 오른쪽으로 밀리며 보존된다)
  const ensureAfter = (name, afterName) => {
    const ci = asColIndex(sh);
    if (ci[name] !== undefined) return;
    const anchor = ci[afterName];
    if (anchor === undefined) {
      sh.getRange(1, sh.getLastColumn() + 1).setValue(name);
    } else {
      sh.insertColumnAfter(anchor + 1);
      sh.getRange(1, anchor + 2).setValue(name);
    }
    SpreadsheetApp.flush();
    added.push(name);
  };

  ensureAfter('Site/제품명', 'ID');
  ensureAfter('구분', '날짜');

  // 기존 행의 Site/제품명 백필
  let filled = 0;
  const ci = asColIndex(sh);
  const last = sh.getLastRow();
  if (last >= 2 && ci['Site/제품명'] !== undefined && ci['ID'] !== undefined) {
    const ids   = sh.getRange(2, ci['ID'] + 1, last - 1, 1).getValues();
    const sites = sh.getRange(2, ci['Site/제품명'] + 1, last - 1, 1).getValues();

    const main = getSheet();
    const mc = buildColIndex(main);
    const rows = main.getLastRow() < 2 ? []
      : main.getRange(2, 1, main.getLastRow() - 1, main.getLastColumn()).getValues();
    const siteById = {};
    rows.forEach(r => { siteById[String(r[mc['ID']]).trim()] = String(r[mc['Site 및 제품명']] || ''); });

    for (let i = 0; i < ids.length; i++) {
      if (String(sites[i][0]).trim()) continue;
      const found = siteById[String(ids[i][0]).trim()];
      if (found) { sites[i][0] = found; filled++; }
    }
    if (filled) sh.getRange(2, ci['Site/제품명'] + 1, last - 1, 1).setValues(sites);
  }

  Logger.log(added.length ? '추가된 열: ' + added.join(', ') : '추가할 열 없음 (이미 적용됨)');
  Logger.log('Site/제품명 채운 행: ' + filled);
}

function addMissingColumns() {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = COLUMNS.map(c => c.header).filter(h => headers.indexOf(h) < 0);

  if (missing.length === 0) {
    Logger.log('빠진 열이 없습니다. 변경 없음.');
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
