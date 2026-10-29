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

  // 기록자는 화면이 보낸 값을 믿지 않고 로그인 토큰의 주인으로 정한다
  const who = emailOfToken(token);
  if (who) body.updatedBy = who;

  try {
    switch (action) {
      case 'google-auth': return jsonResponse(handleGoogleAuth(body));
      case 'logout':  return jsonResponse(handleLogout(body));
      case 'create':  requireAdmin(token); return jsonResponse(handleCreate(body));
      case 'update':  requireAdmin(token); return jsonResponse(handleUpdate(body));
      case 'retire':  requireAdmin(token); return jsonResponse(handleRetire(body));
      case 'delete':  requireAdmin(token); return jsonResponse(handleDelete(body));
      case 'upload':  requireAdmin(token); return jsonResponse(handleUpload(body));
      case 'as-add':  requireAdmin(token); return jsonResponse(handleASAdd(body));
      case 'as-delete': requireAdmin(token); return jsonResponse(handleASDelete(body));
      // 수정 요청 접수는 누구나 가능 (로그인 불필요)
      case 'report-add': return jsonResponse(handleReportAdd(body));
      case 'report-resolve': requireAdmin(token); return jsonResponse(handleReportResolve(body));
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

function emailOfToken(token) {
  if (!validateToken(token)) return '';
  const entry = getTokenStore()[token];
  return (entry && entry.email) || '';
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
  // 2026-10 'Site 및 제품명' 을 Site 와 제품명 두 열로 나눴다. 옛 이름도 알아본다.
  { key: 'site',         header: 'Site',             adminOnly: false, aliases: ['Site 및 제품명'] },
  { key: 'productName',  header: '제품명',            adminOnly: false },
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
  { key: 'cmsId',        header: 'CMS ID',           adminOnly: true  },
  { key: 'cmsPw',        header: 'CMS PW',           adminOnly: true, secret: true },
  { key: 'tvId',         header: 'Teamviewer ID',    adminOnly: true, secret: true },
  { key: 'tvPw',         header: 'Teamviewer PW',    adminOnly: true, secret: true },
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
  { key: 'updatedBy',    header: '최종수정자',       adminOnly: true  },
];

const ADMIN_KEYS = new Set(COLUMNS.filter(c => c.adminOnly).map(c => c.key));

const STATUSES = ['현장설치', '렌트', '창고', '반출'];

function validateStatus(data) {
  if (data.status && STATUSES.indexOf(data.status) < 0) {
    throw { message: '상태는 ' + STATUSES.join(', ') + ' 중 하나여야 합니다.', code: 400 };
  }
}

// 배열식으로 계산이 되면 그 값을, 안 되면 기존 값을 그대로 둔다.
// (예전에는 계산이 안 되면 빈칸으로 덮어써서 x 표기 행의 총 모듈수가 지워졌다)
function applyPanelTotal(data) {
  if (!data.panelArray) return;
  const t = calcPanelTotal(data.panelArray);
  if (t !== '') data.panelTotal = t;
}

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
  // 열 이름을 바꾸는 사이에도 값을 잃지 않도록, 시트에 옛 이름만 있으면 새 이름으로 연결한다.
  // (연결이 없으면 buildRow 가 그 열을 모르는 열로 보고 빈칸으로 덮어쓴다)
  COLUMNS.forEach(c => {
    if (map[c.header] !== undefined || !c.aliases) return;
    const old = c.aliases.find(a => map[a] !== undefined);
    if (old !== undefined) map[c.header] = map[old];
  });
  return map;
}

// 이력·기록에 쓰는 표시 이름: "Site · 제품명" (빈 쪽은 뺀다)
function siteLabel(o) {
  return [o.site, o.productName].map(v => String(v || '').trim()).filter(Boolean).join(' · ');
}

// 1회차 시트 정리 전이면 새 열이 없어 값이 조용히 버려진다. 그럴 바에는 저장을 막고 알린다.
const NEED_STEP1 = '시트 정리가 아직 안 됐습니다. 편집기에서 step1SplitColumns 를 먼저 실행해 주세요.';
function requireProductColumn(data, colIndex) {
  if (String(data.productName || '').trim() && colIndex['제품명'] === undefined) {
    throw { message: NEED_STEP1, code: 409 };
  }
}

// Site 와 제품명 중 하나는 있어야 한다
function validateSiteProduct(data) {
  if (!String(data.site || '').trim() && !String(data.productName || '').trim()) {
    throw { message: 'Site 와 제품명 중 하나는 입력해야 합니다.', code: 400 };
  }
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
  const statuses = STATUSES.slice();

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
  validateFieldLengths(data);
  validateStatus(data);
  validateSiteProduct(data);
  applyPanelTotal(data);

  // 두 사람이 같은 순간 등록해도 관리번호가 겹치지 않도록 채번~기록을 잠근다
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet();
    ensureHeaders(sheet);
    const colIndex = buildColIndex(sheet);
    requireProductColumn(data, colIndex);

    const newId = generateId(sheet, colIndex);
    data.id = newId;
    data.updatedAt = new Date().toISOString();
    data.updatedBy = body.updatedBy || 'admin';

    const row = buildRow(data, colIndex, sheet.getLastColumn());
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    rememberIdNumber(newId);
    writeAuditLog('CREATE', newId, data.updatedBy, siteLabel(data));
    return { id: newId, updatedAt: data.updatedAt };
  } finally {
    lock.releaseLock();
  }
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
  validateStatus(data);
  validateSiteProduct(data);
  requireProductColumn(data, colIndex);
  applyPanelTotal(data);

  const row = buildRow(data, colIndex, sheet.getLastColumn());
  sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  writeAuditLog('UPDATE', id, data.updatedBy, '');
  return { ok: true, updatedAt: data.updatedAt };
}

// 항목 삭제 (관리자). 시트에서 행을 지운다. 되돌릴 수 없으므로
// 지운 내용(비밀번호류 제외)을 AuditLog 에 남겨 필요하면 손으로 복구할 수 있게 한다.
// 설치·수리 이력 시트의 행은 지우지 않는다. 번호를 다시 쓰지 않으므로 새 항목에 섞이지 않는다.
function handleDelete(body) {
  const id = String(body.id || '').trim();
  if (!id) throw { message: 'id가 필요합니다.', code: 400 };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet();
    const colIndex = buildColIndex(sheet);
    const { rowIdx, existing } = findRowById(sheet, colIndex, id);

    // 수정 창을 연 뒤 다른 사람이 고쳤다면 지우지 않는다
    if (body.updatedAt && existing.updatedAt && body.updatedAt !== existing.updatedAt) {
      throw { message: '그사이 다른 사람이 이 항목을 수정했습니다. 새로고침한 뒤 다시 확인해 주세요.', code: 409 };
    }

    const snap = {};
    COLUMNS.forEach(c => {
      const v = existing[c.key];
      if (!c.secret && v !== '' && v !== null && v !== undefined) snap[c.header] = v;
    });

    rememberIdNumber(id);
    sheet.deleteRow(rowIdx);
    SpreadsheetApp.flush();
    writeAuditLog('DELETE', id, body.updatedBy || 'admin', JSON.stringify(snap));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
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
  // 업로드도 수정시각을 바꾸므로, 화면이 다음 저장 때 쓸 새 값을 함께 돌려준다
  return { url, updatedAt: existing.updatedAt };
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

  const visibleCols = COLUMNS.filter(c => includeSecret || !c.secret);
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

// 지금까지 쓴 가장 큰 번호를 따로 기억한다. 맨 끝 항목을 지워도 그 번호를 다시 쓰지 않아,
// 지운 항목의 설치·수리 이력이 새 항목에 붙는 일이 없다.
const ID_MAX_KEY = 'MAX_ID_NUM';

function idNumber(v) {
  const m = String(v).trim().match(ID_RE);
  return m ? parseInt(m[1], 10) : NaN;
}

function rememberIdNumber(id) {
  const n = idNumber(id);
  if (isNaN(n)) return;
  const props = PropertiesService.getScriptProperties();
  const cur = parseInt(props.getProperty(ID_MAX_KEY) || '0', 10) || 0;
  if (n > cur) props.setProperty(ID_MAX_KEY, String(n));
}

function generateId(sheet, colIndex) {
  let max = parseInt(PropertiesService.getScriptProperties().getProperty(ID_MAX_KEY) || '0', 10) || 0;

  const idColIdx = colIndex['ID'];
  if (sheet.getLastRow() >= 2 && idColIdx !== undefined) {
    sheet.getRange(2, idColIdx + 1, sheet.getLastRow() - 1, 1).getValues()
      .forEach(r => {
        const n = idNumber(r[0]);
        if (!isNaN(n) && n > max) max = n;
      });
  }
  return ID_PREFIX + String(max + 1).padStart(ID_PAD, '0');
}

function calcPanelTotal(panelArray) {
  if (!panelArray) return '';
  const parts = String(panelArray).trim().split(/\s*[*xX×]\s*/);
  if (parts.some(p => p === '')) return '';
  const nums = parts.map(Number);
  if (nums.some(n => !isFinite(n))) return '';
  return nums.reduce((a, b) => a * b, 1);
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
// 2026-10: 내용 한 칸을 작업내역·불량원인으로 나누고, 구분 뒤에 대분류를 넣었다.
// 2026-10-30: 'Site/제품명' 한 칸도 Installations 처럼 Site 와 제품명 두 칸으로 나눴다.
const HISTORY_COLS = ['rowId', 'ID', 'Site', '제품명', '날짜', '구분', '대분류', '작업내역', '불량원인', '등록일시', '등록자'];
const HISTORY_KINDS = ['설치', '수리', '기타'];
const HISTORY_CATS = ['Panel', '구동부', 'System'];
const HISTORY_MAX = 200;
// 시트에 옛 이름만 있으면 새 이름으로 연결한다 (1회차 함수 실행 전에도 동작하게)
const HISTORY_ALIASES = { '작업내역': ['내용'], 'Site': ['Site/제품명'] };

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
  Object.keys(HISTORY_ALIASES).forEach(name => {
    if (map[name] !== undefined) return;
    const old = HISTORY_ALIASES[name].find(a => map[a] !== undefined);
    if (old !== undefined) map[name] = map[old];
  });
  return map;
}

// 관리번호로 Site·제품명 찾기. 이력 등록 시 시트에 함께 적어 둔다.
function namesOfId(id) {
  try {
    const sheet = getSheet();
    const { existing } = findRowById(sheet, buildColIndex(sheet), id);
    return { site: String(existing.site || ''), productName: String(existing.productName || '') };
  } catch (_) {
    return { site: '', productName: '' };
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
      site:        String(at(r, 'Site') || ''),
      productName: String(at(r, '제품명') || ''),
      date:      asDate(at(r, '날짜')),
      kind:      String(at(r, '구분') || ''),
      category:  String(at(r, '대분류') || ''),
      work:      String(at(r, '작업내역') || ''),
      cause:     String(at(r, '불량원인') || ''),
      text:      String(at(r, '작업내역') || ''),   // 옛 화면 호환 (작업내역과 같은 값)
      createdAt: String(at(r, '등록일시') || ''),
      createdBy: String(at(r, '등록자') || ''),
    }))
    .sort((a, b) => b.date.localeCompare(a.date));

  return { items };
}

// 시트의 '구분' 열에 선택 목록 규칙이 걸려 있는데 HISTORY_KINDS 중 빠진 값이 있으면(예: '기타')
// 앱이 그 값을 저장할 때 시트가 거절한다. 그 경우에만 규칙의 선택지를 HISTORY_KINDS 로 넓힌다.
// 규칙이 없거나, 목록이 아닌 다른 종류(범위 참조 등)이면 손대지 않는다.
function ensureKindRule_(sh, ci) {
  if (ci['구분'] === undefined || sh.getMaxRows() < 2) return;
  const dv = sh.getRange(2, ci['구분'] + 1).getDataValidation();
  if (!dv) return;
  const vals = (dv.getCriteriaValues() || [])[0];
  if (!Array.isArray(vals)) return;
  if (HISTORY_KINDS.every(k => vals.indexOf(k) >= 0)) return;
  const merged = vals.concat(HISTORY_KINDS.filter(k => vals.indexOf(k) < 0));
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(merged, true).setAllowInvalid(dv.getAllowInvalid()).build();
  sh.getRange(2, ci['구분'] + 1, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  SpreadsheetApp.flush();
  writeAuditLog('KIND_RULE', '', 'admin', '구분 선택지: ' + merged.join(', '));
}

function handleASAdd(body) {
  const id       = body.id   || '';
  const date     = body.date || '';
  const kind     = body.kind || '';
  const category = String(body.category || '').trim();
  const work     = String(body.work || body.text || '').trim();   // 옛 화면은 text 로 보낸다
  const cause    = String(body.cause || '').trim();

  if (!id || !date || !work) throw { message: '날짜와 작업내역은 꼭 입력해 주세요.', code: 400 };
  if (work.length > HISTORY_MAX)  throw { message: '작업내역은 ' + HISTORY_MAX + '자 이내여야 합니다.', code: 400 };
  if (cause.length > HISTORY_MAX) throw { message: '불량 원인은 ' + HISTORY_MAX + '자 이내여야 합니다.', code: 400 };
  if (kind && HISTORY_KINDS.indexOf(kind) < 0) {
    throw { message: '구분은 ' + HISTORY_KINDS.join(', ') + ' 중 하나여야 합니다.', code: 400 };
  }
  if (category && HISTORY_CATS.indexOf(category) < 0) {
    throw { message: '대분류는 ' + HISTORY_CATS.join(', ') + ' 중 하나여야 합니다.', code: 400 };
  }

  const sh = getASSheet();
  const ci = asColIndex(sh);
  if ((category || cause) && (ci['대분류'] === undefined || ci['불량원인'] === undefined)) {
    throw { message: NEED_STEP1, code: 409 };
  }
  if (kind) ensureKindRule_(sh, ci);
  const rowId = Utilities.getUuid();

  const row = new Array(sh.getLastColumn()).fill('');
  const put = (name, v) => { if (ci[name] !== undefined) row[ci[name]] = v; };
  put('rowId', rowId);
  put('ID', id);
  const names = namesOfId(id);
  if (ci['제품명'] !== undefined) { put('Site', names.site); put('제품명', names.productName); }
  else put('Site', siteLabel(names));          // 2회차 전: 'Site/제품명' 한 칸에 "Site · 제품명"
  put('날짜', date);
  put('구분', kind);
  put('대분류', category);
  put('작업내역', work);
  put('불량원인', cause);
  put('등록일시', new Date().toISOString());
  put('등록자', body.updatedBy || 'admin');

  sh.appendRow(row);
  const tag = [kind, category].filter(Boolean).map(t => '[' + t + ']').join('');
  writeAuditLog('AS_ADD', id, body.updatedBy || 'admin', (tag ? tag + ' ' : '') + date + ': ' + work.slice(0, 30));
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
// (migrateHistorySheet: 2026-08 이관 함수. 할 일을 마쳤고, 지금 실행하면 옛 'Site/제품명' 열을 다시 만들어서 지웠다)

/**
 * 데이터 정리 (2026-10). 편집기에서 한 번 실행하세요. 여러 번 실행해도 안전합니다.
 *  1) 상태가 '로비' / 'ISB' 인 행 → 상태 '현장설치', 비고1 에 '위치: 로비' 처럼 기록
 *  2) 총 모듈수가 비어 있고 패널 배열로 계산되는 행 → 총 모듈수 채움
 *     (예전 계산이 x 표기를 몰라 앱에서 저장할 때 지워진 값 복구)
 */
function cleanupData() {
  const sheet = getSheet();
  const ci = buildColIndex(sheet);
  const last = sheet.getLastRow();
  if (last < 2) { Logger.log('데이터 없음'); return; }

  ['상태', '비고1', 'Panel수량', '총 모듈수'].forEach(h => {
    if (ci[h] === undefined) throw new Error('시트에 열이 없습니다: ' + h);
  });

  const n = last - 1;
  const col = h => sheet.getRange(2, ci[h] + 1, n, 1);
  const status = col('상태').getValues();
  const note1  = col('비고1').getValues();
  const arr    = col('Panel수량').getValues();
  const total  = col('총 모듈수').getValues();

  const LOCATIONS = ['로비', 'ISB'];
  let moved = 0, filled = 0;

  for (let i = 0; i < n; i++) {
    const s = String(status[i][0]).trim();
    if (LOCATIONS.indexOf(s) >= 0) {
      const tag = '위치: ' + s;
      const cur = String(note1[i][0]).trim();
      if (cur.indexOf(tag) < 0) note1[i][0] = cur ? cur + ' / ' + tag : tag;
      status[i][0] = '현장설치';
      moved++;
    }
    if (String(total[i][0]).trim() === '') {
      const t = calcPanelTotal(arr[i][0]);
      if (t !== '') { total[i][0] = t; filled++; }
    }
  }

  if (moved)  { col('상태').setValues(status); col('비고1').setValues(note1); }
  if (filled) col('총 모듈수').setValues(total);
  SpreadsheetApp.flush();

  writeAuditLog('CLEANUP', '', 'admin', '로비/ISB→현장설치 ' + moved + '행, 총 모듈수 채움 ' + filled + '행');
  Logger.log('로비/ISB → 현장설치: ' + moved + '행');
  Logger.log('총 모듈수 채움: ' + filled + '행');
}

/**
 * [1회차] step1SplitColumns (2026-10). 편집기에서 실행하세요. 여러 번 실행해도 안전합니다.
 * ── Installations: Site 와 제품명 나누기
 *  1) C열 'Site 및 제품명' 이름을 'Site' 로 바꾼다
 *  2) 그 바로 오른쪽에 '제품명' 열을 끼워 넣는다 (기존 열은 한 칸씩 오른쪽으로)
 *  3) 제품명으로 볼 행은 C열 값을 제품명으로 옮기고 Site 는 비운다
 *     - 제품 일련번호 형태(예: SLD-27S-XXX-B-001)
 *     - 사용자가 정한 7행: Prod005, 006, 007, 008, 009, 014, 040
 *     옮긴 행 중 비고1 에 '위치: ○○' 이 있으면 그 위치를 Site 로 넣는다 (비고1 은 그대로)
 *  제품명이 이미 채워진 행은 건드리지 않는다.
 */
const PRODUCT_IDS_2026_10 = ['Prod005', 'Prod006', 'Prod007', 'Prod008', 'Prod009', 'Prod014', 'Prod040'];
const SERIAL_RE = /^[A-Z]{2,}-[\w-]+$/;

function step1SplitColumns() {
  splitSiteProduct_();
  splitHistoryColumns_();
}

// 이름 끝의 _ 는 편집기 실행 목록에서 숨긴다 (step1SplitColumns 로만 실행)
function splitSiteProduct_() {
  const sheet = getSheet();
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

  // 1) 이름 바꾸기
  let siteCol = headers.indexOf('Site');
  if (siteCol < 0) {
    siteCol = headers.indexOf('Site 및 제품명');
    if (siteCol < 0) throw new Error("시트에 'Site 및 제품명' 또는 'Site' 열이 없습니다.");
    sheet.getRange(1, siteCol + 1).setValue('Site');
    Logger.log("C열 이름: 'Site 및 제품명' → 'Site'");
  }

  // 2) 제품명 열 끼워 넣기
  if (headers.indexOf('제품명') < 0) {
    sheet.insertColumnAfter(siteCol + 1);
    sheet.getRange(1, siteCol + 2).setValue('제품명');
    Logger.log("'제품명' 열을 Site 오른쪽에 추가");
  }
  SpreadsheetApp.flush();

  // 3) 값 옮기기
  const ci = buildColIndex(sheet);
  const last = sheet.getLastRow();
  if (last < 2) { Logger.log('데이터 없음'); return; }
  const n = last - 1;
  const get = h => sheet.getRange(2, ci[h] + 1, n, 1).getValues();
  const ids = get('ID'), site = get('Site'), prod = get('제품명'), note1 = get('비고1');

  let moved = 0, located = 0;
  for (let i = 0; i < n; i++) {
    const id = String(ids[i][0]).trim();
    const s  = String(site[i][0]).trim();
    if (!s || String(prod[i][0]).trim()) continue;
    if (!SERIAL_RE.test(s) && PRODUCT_IDS_2026_10.indexOf(id) < 0) continue;

    prod[i][0] = s;
    const m = String(note1[i][0]).match(/위치:\s*([^/]+?)\s*(?:\/|$)/);
    site[i][0] = m ? m[1].trim() : '';
    if (m) located++;
    moved++;
  }

  if (moved) {
    sheet.getRange(2, ci['Site'] + 1, n, 1).setValues(site);
    sheet.getRange(2, ci['제품명'] + 1, n, 1).setValues(prod);
  }
  SpreadsheetApp.flush();
  writeAuditLog('SPLIT', '', 'admin', '제품명으로 옮김 ' + moved + '행, 그중 위치를 Site 로 ' + located + '행');
  Logger.log('제품명으로 옮김: ' + moved + '행 (그중 비고1 위치를 Site 로: ' + located + '행)');
}

/**
 * ── AS_History: 구분 뒤에 '대분류', 내용 → '작업내역' 이름 변경, 그 뒤에 '불량원인'
 * 기존 내용은 작업내역으로 그대로 남고, 대분류·불량원인은 빈칸으로 시작한다.
 */
// name 열을 after 열 바로 오른쪽에 끼운다 (기존 값은 오른쪽으로 밀리며 보존). 끼웠으면 true.
// 구글 시트는 끼운 열에 왼쪽 열의 '데이터 확인' 규칙을 복사한다. 구분 열의 설치·수리 규칙이
// 대분류 열에 붙으면 머리글부터 거절되므로, 끼운 열의 규칙을 먼저 지운다.
// 바로 오른쪽 열의 머리글이 비어 있으면(이전 실행이 도중에 멈춘 흔적) 새로 끼우지 않고 그 열을 쓴다.
function ensureColAfter_(sh, name, after, rule) {
  const cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  if (cur.indexOf(name) >= 0) return false;
  const a = cur.indexOf(after);
  let col;
  if (a < 0) {
    col = cur.length + 1;
  } else if (a + 1 < cur.length && cur[a + 1] === '') {
    col = a + 2;
  } else {
    sh.insertColumnAfter(a + 1);
    col = a + 2;
  }
  sh.getRange(1, col, sh.getMaxRows(), 1).clearDataValidations();
  sh.getRange(1, col).setValue(name);
  if (rule && sh.getMaxRows() > 1) sh.getRange(2, col, sh.getMaxRows() - 1, 1).setDataValidation(rule);
  SpreadsheetApp.flush();
  return true;
}

function splitHistoryColumns_() {
  const sh = getASSheet();
  const hdr = () => sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
  const done = [];

  let h = hdr();
  if (h.indexOf('작업내역') < 0 && h.indexOf('내용') >= 0) {
    sh.getRange(1, h.indexOf('내용') + 1).setValue('작업내역');
    done.push("'내용' → '작업내역'");
  }
  const ensureAfter = (name, after, rule) => { if (ensureColAfter_(sh, name, after, rule)) done.push("'" + name + "' 추가"); };
  // 대분류에는 구분처럼 선택 목록 규칙을 건다 (빈칸은 허용)
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(HISTORY_CATS, true).setAllowInvalid(false).build();
  ensureAfter('대분류', '구분', catRule);
  ensureAfter('불량원인', '작업내역', null);
  SpreadsheetApp.flush();
  Logger.log('이력 시트: ' + (done.length ? done.join(', ') : '바꿀 것 없음 (이미 적용됨)'));
}

/**
 * [2회차] step2SyncHistory: 설치·수리 이력의 Site·제품명을 지금 메인 시트 값에 맞춘다.
 * 이력 시트에 'Site/제품명' 한 칸만 있으면 먼저 'Site' + '제품명' 두 칸으로 나눈다.
 * 메인 시트 정리를 다시 하셨을 때도 이 함수만 다시 실행하면 된다. 여러 번 실행해도 안전합니다.
 *  - 항목이 있는 이력: Site·제품명을 메인 시트 값으로 다시 쓴다
 *  - 항목이 지워진 이력: 'Site · 제품명' 처럼 합쳐 적혀 있으면 가운뎃점 앞뒤로 나누고, 아니면 Site 에 둔다
 */
function step2SyncHistory() {
  const main = getSheet();
  const mc = buildColIndex(main);
  const rows = main.getLastRow() < 2 ? []
    : main.getRange(2, 1, main.getLastRow() - 1, main.getLastColumn()).getValues();
  const byId = {};
  rows.forEach(r => {
    const id = String(r[mc['ID']]).trim();
    if (id) byId[id] = {
      site: String(r[mc['Site']] || '').trim(),
      productName: mc['제품명'] === undefined ? '' : String(r[mc['제품명']] || '').trim(),
    };
  });

  const sh = getASSheet();
  const done = [];
  let h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(x => String(x).trim());
  if (h.indexOf('Site') < 0 && h.indexOf('Site/제품명') >= 0) {
    sh.getRange(1, h.indexOf('Site/제품명') + 1).setValue('Site');
    SpreadsheetApp.flush();
    done.push("'Site/제품명' → 'Site'");
  }
  if (ensureColAfter_(sh, '제품명', 'Site', null)) done.push("'제품명' 추가");

  const ci = asColIndex(sh);
  const last = sh.getLastRow();
  if (last < 2 || ci['ID'] === undefined || ci['Site'] === undefined || ci['제품명'] === undefined) {
    Logger.log('이력 시트: ' + (done.join(', ') || '바꿀 것 없음') + ' / 바꿀 이력이 없습니다.');
    return;
  }
  const n = last - 1;
  const ids  = sh.getRange(2, ci['ID'] + 1, n, 1).getValues();
  const siteCol = sh.getRange(2, ci['Site'] + 1, n, 1), prodCol = sh.getRange(2, ci['제품명'] + 1, n, 1);
  const site = siteCol.getValues(), prod = prodCol.getValues();

  let changed = 0, orphan = 0;
  for (let i = 0; i < n; i++) {
    const want = byId[String(ids[i][0]).trim()];
    let s2, p2;
    if (want) {
      s2 = want.site; p2 = want.productName;
    } else {
      orphan++;
      const cur = String(site[i][0]).trim();
      if (String(prod[i][0]).trim() || cur.indexOf(' · ') < 0) continue;
      const k = cur.indexOf(' · ');
      s2 = cur.slice(0, k).trim(); p2 = cur.slice(k + 3).trim();
    }
    if (String(site[i][0]) !== s2 || String(prod[i][0]) !== p2) { site[i][0] = s2; prod[i][0] = p2; changed++; }
  }
  if (changed) { siteCol.setValues(site); prodCol.setValues(prod); }
  SpreadsheetApp.flush();
  writeAuditLog('SYNC_HISTORY', '', 'admin', '이력 Site·제품명 갱신 ' + changed + '행');
  Logger.log('이력 시트: ' + (done.length ? done.join(', ') : '열 구성 그대로'));
  Logger.log('이력 Site·제품명 갱신: ' + changed + '행 (항목이 지워진 이력 ' + orphan + '행)');
}

function addMissingColumns() {
  const sheet = getSheet();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const has = c => headers.indexOf(c.header) >= 0 || (c.aliases || []).some(a => headers.indexOf(a) >= 0);
  const missing = COLUMNS.filter(c => !has(c)).map(c => c.header);

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
