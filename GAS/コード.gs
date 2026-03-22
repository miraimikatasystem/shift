const AI_API_KEY = PropertiesService.getScriptProperties().getProperty('AI_API_KEY');
const LEGACY_BASE_SS_ID = '1P7nz8hmV0MtG6lg88JlivMBAwzglV64_PQh-Jlw2FfQ';
const BASE_SS_ID = '1aS_u3Fyj2jwY_Obcrb19iCFcrOF1k2Iy92lcqXkaLzA';
const TEMPLATE_SS_ID = '1CDL3xOVAvdxrcX30bNyEPdUVqeri83OcIt0wIaA2QOY';
const FOLDER_ID = '1AWL9D8mbZXeIQIEIHT-jlpiXKRFJxgAb';
const ADMIN_PASSWORD = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
const GOOGLE_TOKENINFO_URL = 'https://oauth2.googleapis.com/tokeninfo';
const GOOGLE_OAUTH_CLIENT_ID = PropertiesService.getScriptProperties().getProperty('GOOGLE_OAUTH_CLIENT_ID') || '';
const SESSION_TTL_SEC = 60 * 60;
const BASE_MASTER_SHEET_NAMES = ['スタッフ一覧', '募集履歴一覧', 'お知らせ'];
const ALLOWED_EMAILS_SHEET_NAME = 'アクセス許可メール';
const ALLOWED_EMAILS_SHEET_HEADER = ['有効', 'メールアドレス', 'メモ', '更新日時', '更新者'];

function doPost(e) {
  const bootstrapSessionToken = e && e.parameter ? String(e.parameter.bootstrapSessionToken || '') : '';
  if (bootstrapSessionToken) {
    const templateName = (e.parameter && e.parameter.view === 'admin') ? 'AdminView' : 'StaffView';
    const session = _getSession_(bootstrapSessionToken, true);
    if (!session) return createUnauthorizedHtml_();
    if (templateName === 'AdminView' && session.role !== 'admin') return createForbiddenHtml_();
    return buildAppHtml_(templateName, bootstrapSessionToken, session);
  }

  const request = _parseJsonRequest_(e);
  if (!request.ok) return _jsonResponse_({ status: 400, ok: false, code: 'invalid_json', message: 'invalid request body' });

  const action = String(request.data.action || '');
  const payload = request.data.payload && typeof request.data.payload === 'object' ? request.data.payload : {};

  if (action === 'revokeSession') {
    _revokeSession_(String(payload.sessionToken || ''));
    return _jsonResponse_({
      status: 200,
      ok: true,
      data: { revoked: true }
    });
  }

  const idToken = String(request.data.idToken || '');
  if (!idToken) {
    return _jsonResponse_({
      status: 401,
      ok: false,
      code: 'missing_token',
      message: 'id token is required'
    });
  }

  const verifyResult = _verifyGoogleIdToken_(idToken);
  if (!verifyResult.ok) {
    return _jsonResponse_({
      status: 401,
      ok: false,
      code: verifyResult.code || 'unauthorized',
      message: _authFailureMessage_(verifyResult.code)
    });
  }

  if (!_isAllowedEmail_(verifyResult.email)) {
    return _jsonResponse_({
      status: 403,
      ok: false,
      code: 'forbidden',
      message: 'このGoogleアカウントはアクセス許可メールに登録されていません'
    });
  }

  if (action === 'createSession') {
    const session = _createSession_(verifyResult.email, 'staff');
    return _jsonResponse_({
      status: 200,
      ok: true,
      email: verifyResult.email,
      data: {
        now: new Date().toISOString(),
        sessionToken: session.token,
        expiresIn: session.expiresIn
      }
    });
  }

  if (action === 'ping') {
    return _jsonResponse_({
      status: 200,
      ok: true,
      email: verifyResult.email,
      data: {
        now: new Date().toISOString()
      }
    });
  }

  return _jsonResponse_({
    status: 400,
    ok: false,
    code: 'unknown_action',
    message: 'unknown action'
  });
}

function _parseJsonRequest_(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
    const parsed = JSON.parse(raw || '{}');
    if (!parsed || typeof parsed !== 'object') return { ok: false };
    return { ok: true, data: parsed };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : 'parse_error' };
  }
}

function _authFailureMessage_(code) {
  switch (String(code || '')) {
    case 'server_misconfigured':
      return 'GAS の GOOGLE_OAUTH_CLIENT_ID が未設定です';
    case 'tokeninfo_unreachable':
      return 'Google の認証確認サーバーに接続できませんでした';
    case 'token_invalid':
      return 'Google ID トークンが無効です';
    case 'token_parse_failed':
      return 'Google 認証レスポンスの解析に失敗しました';
    case 'aud_mismatch':
      return 'NEXT_PUBLIC_GOOGLE_CLIENT_ID と GAS の GOOGLE_OAUTH_CLIENT_ID が一致していません';
    case 'iss_invalid':
      return 'Google の発行元情報を確認できませんでした';
    case 'token_expired':
      return 'Google ログインのトークン期限が切れています。再ログインしてください';
    case 'email_not_verified':
      return 'Google アカウントのメール確認が完了していません';
    case 'email_missing':
      return 'Google アカウントのメールアドレスを取得できませんでした';
    default:
      return '認証に失敗しました';
  }
}

function _verifyGoogleIdToken_(idToken) {
  if (!GOOGLE_OAUTH_CLIENT_ID) return { ok: false, code: 'server_misconfigured' };

  let res;
  try {
    res = UrlFetchApp.fetch(GOOGLE_TOKENINFO_URL + '?id_token=' + encodeURIComponent(idToken), {
      method: 'get',
      muteHttpExceptions: true
    });
  } catch (err) {
    return { ok: false, code: 'tokeninfo_unreachable' };
  }

  if (res.getResponseCode() !== 200) return { ok: false, code: 'token_invalid' };

  let tokenInfo;
  try {
    tokenInfo = JSON.parse(res.getContentText() || '{}');
  } catch (err) {
    return { ok: false, code: 'token_parse_failed' };
  }

  const aud = String(tokenInfo.aud || '');
  if (aud !== GOOGLE_OAUTH_CLIENT_ID) return { ok: false, code: 'aud_mismatch' };

  const iss = String(tokenInfo.iss || '');
  if (iss !== 'accounts.google.com' && iss !== 'https://accounts.google.com') {
    return { ok: false, code: 'iss_invalid' };
  }

  const exp = Number(tokenInfo.exp || 0);
  const nowSec = Math.floor(Date.now() / 1000);
  if (!isFinite(exp) || exp <= nowSec) return { ok: false, code: 'token_expired' };

  const emailVerified = String(tokenInfo.email_verified || '').toLowerCase() === 'true';
  if (!emailVerified) return { ok: false, code: 'email_not_verified' };

  const email = String(tokenInfo.email || '').toLowerCase();
  if (!email) return { ok: false, code: 'email_missing' };

  return { ok: true, email: email };
}

function _isAllowedEmail_(email) {
  return getAllowedEmailEntries_().some(function(entry) {
    return entry.enabled && entry.email === _normalizeEmail_(email);
  });
}

function _jsonResponse_(body) {
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function _createSession_(email, role) {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const session = {
    email: String(email || '').toLowerCase(),
    role: String(role || 'staff'),
    issuedAt: Date.now(),
    expiresAt: Date.now() + (SESSION_TTL_SEC * 1000)
  };
  _saveSession_(token, session);
  return { token: token, expiresIn: SESSION_TTL_SEC };
}

function _saveSession_(token, session) {
  if (!token || !session) return;
  const cacheKey = 'sess:' + String(token);
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(session), SESSION_TTL_SEC);
}

function _getSession_(token, touch) {
  if (!token) return null;
  const cacheKey = 'sess:' + String(token);
  const raw = CacheService.getScriptCache().get(cacheKey);
  if (!raw) return null;
  try {
    const session = JSON.parse(raw);
    if (!session || !session.email) return null;
    if (Number(session.expiresAt || 0) <= Date.now()) {
      CacheService.getScriptCache().remove(cacheKey);
      return null;
    }
    if (touch) {
      session.expiresAt = Date.now() + (SESSION_TTL_SEC * 1000);
      _saveSession_(token, session);
    }
    return session;
  } catch (err) {
    return null;
  }
}

function _revokeSession_(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('sess:' + String(token));
}

function _requireSession_(sessionToken) {
  const session = _getSession_(sessionToken, true);
  if (!session) throw new Error('セッションが切れました。再ログインしてください。');
  return session;
}

function _requireAdminSession_(sessionToken) {
  const session = _requireSession_(sessionToken);
  if (session.role !== 'admin') throw new Error('管理者権限が必要です。');
  return session;
}

function _normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function _isEmailEnabledValue_(value) {
  const text = String(value == null ? '' : value).trim().toLowerCase();
  if (!text) return true;
  return ['false', '0', 'off', 'no', 'ng', '無効'].indexOf(text) === -1;
}

function ensureAllowedEmailsSheet_() {
  const ss = SpreadsheetApp.openById(BASE_SS_ID);
  let sheet = ss.getSheetByName(ALLOWED_EMAILS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ALLOWED_EMAILS_SHEET_NAME);

  const headerRange = sheet.getRange(1, 1, 1, ALLOWED_EMAILS_SHEET_HEADER.length);
  const currentHeader = headerRange.getDisplayValues()[0];
  if (currentHeader.join('') !== ALLOWED_EMAILS_SHEET_HEADER.join('')) {
    sheet.clearContents();
    headerRange.setValues([ALLOWED_EMAILS_SHEET_HEADER]);
  }

  return sheet;
}

function getAllowedEmailEntries_() {
  const sheet = ensureAllowedEmailsSheet_();
  const data = sheet.getDataRange().getValues();
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const email = _normalizeEmail_(data[i][1]);
    if (!email) continue;
    entries.push({
      enabled: _isEmailEnabledValue_(data[i][0]),
      email: email,
      note: String(data[i][2] || ''),
      updatedAt: data[i][3] || '',
      updatedBy: String(data[i][4] || '')
    });
  }
  return entries;
}

function _copyBaseMasterSheet_(sourceSs, targetSs, sheetName) {
  const sourceSheet = sourceSs.getSheetByName(sheetName);
  if (!sourceSheet) throw new Error('コピー元にシートがありません: ' + sheetName);

  const tempName = '__tmp__' + sheetName + '__' + Utilities.getUuid().slice(0, 8);
  const copiedSheet = sourceSheet.copyTo(targetSs).setName(tempName);
  const existingSheet = targetSs.getSheetByName(sheetName);
  if (existingSheet) targetSs.deleteSheet(existingSheet);
  copiedSheet.setName(sheetName);
  return sheetName;
}

function migrateLegacyBaseToCurrentBase() {
  if (!LEGACY_BASE_SS_ID) return { success: false, message: '旧ベースシートIDが未設定です' };
  if (LEGACY_BASE_SS_ID === BASE_SS_ID) return { success: false, message: '旧IDと新IDが同じため移行不要です' };

  const sourceSs = SpreadsheetApp.openById(LEGACY_BASE_SS_ID);
  const targetSs = SpreadsheetApp.openById(BASE_SS_ID);
  const copied = [];

  BASE_MASTER_SHEET_NAMES.forEach(function(sheetName) {
    copied.push(_copyBaseMasterSheet_(sourceSs, targetSs, sheetName));
  });

  return {
    success: true,
    message: 'ベースデータを移行しました: ' + copied.join(', '),
    sourceId: LEGACY_BASE_SS_ID,
    targetId: BASE_SS_ID
  };
}

function setupPermissions() {
  DriveApp.getFolderById(FOLDER_ID);
  DriveApp.getFileById(TEMPLATE_SS_ID);
  SpreadsheetApp.openById(BASE_SS_ID);
  ensureAllowedEmailsSheet_();
}

function doGet(e) {
  let templateName = (e.parameter && e.parameter.view === 'admin') ? 'AdminView' : 'StaffView';
  const sessionToken = e && e.parameter ? String(e.parameter.st || '') : '';
  const session = sessionToken ? _getSession_(sessionToken, true) : null;
  if (!session) return createUnauthorizedHtml_();
  if (templateName === 'AdminView' && session.role !== 'admin') return createForbiddenHtml_();
  return buildAppHtml_(templateName, sessionToken, session);
}

function buildAppHtml_(templateName, sessionToken, session) {
  const template = HtmlService.createTemplateFromFile(templateName);
  template.sessionToken = String(sessionToken || '');
  template.sessionEmail = String((session && session.email) || '');
  template.sessionRole = String((session && session.role) || '');

  const html = template.evaluate();
  html.setTitle(templateName === 'AdminView' ? '管理者 | CAFE SHIFT' : 'スタッフ | CAFE SHIFT');
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

function createUnauthorizedHtml_() {
  const deny = HtmlService.createHtmlOutput(
    '<!doctype html><html><body><h3>401 Unauthorized</h3><p>Session is missing or expired.</p><script>try{window.parent&&window.parent.postMessage({type:"gas-session-expired"}, "*");}catch(e){}</script></body></html>'
  );
  deny.setTitle('Unauthorized');
  deny.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return deny;
}

function createForbiddenHtml_() {
  const deny = HtmlService.createHtmlOutput('<!doctype html><html><body><h3>403 Forbidden</h3><p>Admin session is required.</p></body></html>');
  deny.setTitle('Forbidden');
  deny.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return deny;
}

/**
 * 募集履歴一覧（BASE）
 * A: status(募集中/調整中/確定済)
 * B: createdAt
 * C: recruitId
 * D: periodLabel
 * E: deadline
 * F: fileUrl
 */

function getRecruitSS(targetId) {
  const sheet = SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('募集履歴一覧');
  if(!sheet) return null;
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return null;

  for (let i = data.length - 1; i >= 1; i--) {
    if (targetId ? data[i][2] === targetId : (data[i][0] === '募集中' || data[i][0] === '調整中' || data[i][0] === '確定済')) {
      if (!data[i][5]) continue;
      const match = String(data[i][5]).match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) return SpreadsheetApp.openById(match[1]);
    }
  }
  return null;
}

function generateDateList(startStr, endStr) {
  const dates = []; if (!startStr || !endStr) return [];
  let current = new Date(startStr); const end = new Date(endStr);
  const days = ["(日)","(月)","(火)","(水)","(木)","(金)","(土)"];
  while (current <= end) {
    dates.push(`${current.getMonth()+1}/${current.getDate()}${days[current.getDay()]}`);
    current.setDate(current.getDate() + 1);
  }
  return dates;
}

function safeGetHour(val, fallback) {
  if (val == null || val === '') return fallback;
  if (typeof val === 'object' && val instanceof Date) return parseInt(Utilities.formatDate(val, "JST", "HH"), 10);
  let str = String(val); if (str.includes(':')) return parseInt(str.split(':')[0], 10);
  return fallback;
}

function parseRecruitSettings(setVals) {
  const dates = generateDateList(setVals[1][0], setVals[1][1]);
  let defStart = setVals[1][3] || "8:00"; let defEnd = setVals[1][4] || "20:00";
  let minH = safeGetHour(defStart, 8); let maxH = safeGetHour(defEnd, 20);
  let ex = []; try { ex = JSON.parse(setVals[1][5] || "[]"); } catch(e){}
  const days = ["(日)","(月)","(火)","(水)","(木)","(金)","(土)"];
  ex.forEach(e => {
    let d = new Date(e.date); e.formattedDate = `${d.getMonth()+1}/${d.getDate()}${days[d.getDay()]}`;
    if(e.type === 'short' && e.start && e.end) {
      let s = safeGetHour(e.start, minH); let e_hr = safeGetHour(e.end, maxH);
      if(s < minH) minH = s; if(e_hr > maxH) maxH = e_hr;
    }
  });
  const times = []; for(let i=minH; i<maxH; i++) times.push(i + ":00");
  return { start: setVals[1][0], end: setVals[1][1], deadline: setVals[1][2], defStart: defStart, defEnd: defEnd, exceptions: ex, dates: dates, times: times };
}

function getEndTimeStr(timeStr) { try { return (safeGetHour(timeStr, 8) + 1) + ":00"; } catch(e) { return ""; } }
function fixZero(str, isPin = false) { let s = String(str).replace(/'/g, ""); if (!s) return s; if (isPin && s.length === 3) return '0' + s; if (!isPin && s.length >= 9 && !s.startsWith('0')) return '0' + s; return s; }
function parseTargetHoursValue(v) { const n = Number(v); return (isFinite(n) && n > 0 && n <= 300) ? n : null; }

function parseMonthlyTargetHours(note) {
  if (note == null) return null;
  const toHalf = s => String(s).replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 65248));
  const text = toHalf(String(note)).toLowerCase();
  if (!text.trim()) return null;

  let m = text.match(/月\s*([0-9]+(?:\.[0-9]+)?)\s*(?:時間|h|hr|hrs)?/i);
  if (!m) m = text.match(/([0-9]+(?:\.[0-9]+)?)\s*(?:時間|h|hr|hrs)\s*(?:\/?\s*月)?/i);
  if (!m) return null;

  const v = Number(m[1]);
  if (!isFinite(v) || v <= 0 || v > 300) return null;
  return v;
}

function normalizeNoteText(note) {
  return String(note || "")
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 65248))
    .replace(/[．。]/g, ".")
    .replace(/\s+/g, " ")
    .trim();
}

function parseWeeklyTargetHours(note) {
  const text = normalizeNoteText(note);
  if (!text) return null;
  let m = text.match(/週\s*([0-9]+(?:\.[0-9]+)?)\s*(?:時間|h|hr|hrs)?/i);
  if (!m) m = text.match(/([0-9]+(?:\.[0-9]+)?)\s*(?:時間|h|hr|hrs)\s*(?:\/?\s*週)?/i);
  if (!m) return null;
  const v = Number(m[1]);
  if (!isFinite(v) || v <= 0 || v > 80) return null;
  return v;
}

function estimateMonthlyHoursFromWeekly(weeklyHours) {
  const v = Number(weeklyHours);
  if (!isFinite(v) || v <= 0) return null;
  return Math.round(v * 4.5 * 10) / 10;
}

function _extractJsonFromText(text) {
  if (!text) return null;
  const raw = String(text).trim();
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  const candidate = fenced ? fenced[1] : raw;
  const start = candidate.indexOf("{");
  const end = candidate.lastIndexOf("}");
  if (start < 0 || end < 0 || end <= start) return null;
  return candidate.slice(start, end + 1);
}

function extractStaffPrefsByGemini(note) {
  const empty = { monthlyTargetHours: null, weeklyTargetHours: null, unavailableWeekdays: [], unavailableTimeRanges: [] };
  const text = normalizeNoteText(note);
  if (!text || !AI_API_KEY) return empty;

  const cache = CacheService.getScriptCache();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text);
  const key = "note_pref_v1_" + Utilities.base64EncodeWebSafe(digest).replace(/=+$/, "");
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=" + encodeURIComponent(AI_API_KEY);
  const prompt = [
    "Extract scheduling constraints from this Japanese staff note.",
    "Return JSON only with this exact schema:",
    "{\"monthly_target_hours\": number|null, \"weekly_target_hours\": number|null, \"unavailable_weekdays\": string[], \"unavailable_time_ranges\": [{\"start\":\"HH:MM\",\"end\":\"HH:MM\"}]}",
    "Rules:",
    "- unavailable_weekdays: each item must be one-char Japanese weekday: 日,月,火,水,木,金,土",
    "- unknown values must be null or []",
    "- no extra keys",
    "",
    "NOTE:",
    text
  ].join("\n");

  try {
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        generationConfig: { responseMimeType: "application/json", temperature: 0.1 }
      }),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) return empty;

    const body = JSON.parse(res.getContentText() || "{}");
    const parts = (((body.candidates || [])[0] || {}).content || {}).parts || [];
    const mergedText = parts.map(p => p.text || "").join("\n");
    const jsonText = _extractJsonFromText(mergedText);
    if (!jsonText) return empty;

    const parsed = JSON.parse(jsonText);
    const monthly = parsed.monthly_target_hours == null ? null : Number(parsed.monthly_target_hours);
    const weekly = parsed.weekly_target_hours == null ? null : Number(parsed.weekly_target_hours);
    const unavailableWeekdays = Array.isArray(parsed.unavailable_weekdays)
      ? parsed.unavailable_weekdays.map(v => String(v).trim()).filter(v => /^[日月火水木金土]$/.test(v))
      : [];
    const unavailableTimeRanges = Array.isArray(parsed.unavailable_time_ranges)
      ? parsed.unavailable_time_ranges
          .map(r => ({ start: String((r || {}).start || "").trim(), end: String((r || {}).end || "").trim() }))
          .filter(r => /^\d{1,2}:\d{2}$/.test(r.start) && /^\d{1,2}:\d{2}$/.test(r.end))
      : [];

    const result = {
      monthlyTargetHours: isFinite(monthly) && monthly > 0 && monthly <= 300 ? monthly : null,
      weeklyTargetHours: isFinite(weekly) && weekly > 0 && weekly <= 80 ? weekly : null,
      unavailableWeekdays: unavailableWeekdays,
      unavailableTimeRanges: unavailableTimeRanges
    };
    cache.put(key, JSON.stringify(result), 21600);
    return result;
  } catch (e) {
    return empty;
  }
}

function sendWNotice_(staffName, title, body) {
  try {
    const ss = SpreadsheetApp.openById(BASE_SS_ID);
    const noticeSheet = ss.getSheetByName('お知らせ');
    if(noticeSheet) noticeSheet.appendRow([new Date(), staffName, title, body]);

    const staffData = ss.getSheetByName('スタッフ一覧').getDataRange().getDisplayValues();
    let targets = [];
    if (staffName === '全員') targets = staffData.filter((r, i) => i > 0 && r[0] === '稼働中' && r[4]).map(r => r[4]);
    else { let s = staffData.find((r, i) => i > 0 && r[1] === staffName && r[4]); if (s) targets.push(s[4]); }

    if (targets.length > 0) {
      MailApp.sendEmail({
        bcc: targets.join(','),
        subject: title,
        body: `【CAFE SHIFT お知らせ】\n\n${staffName}さん\n\n${body}\n\n※アプリから詳細をご確認ください。`
      });
    }
  } catch(e) { console.log(e); }
}

function verifyStaffPin(name, pin, sessionToken) {
  _requireSession_(sessionToken);
  const data = SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('スタッフ一覧').getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    let dbPin = fixZero(data[i][2], true);
    if (data[i][1] == name && dbPin === String(pin)) return { success: true, staffData: { name: data[i][1], pin: dbPin, phone: fixZero(data[i][3]), email: data[i][4], note: data[i][5], targetHours: data[i][7] || '' } };
  }
  return { success: false, message: "名前かPINが間違っています" };
}

function authenticateAdmin(password, sessionToken) {
  const session = _requireSession_(sessionToken);
  if (!ADMIN_PASSWORD) return { success: false, message: '管理者パスワードが未設定です。Script Properties に ADMIN_PASSWORD を設定してください。' };
  if (String(password || '') !== ADMIN_PASSWORD) return { success: false, message: 'パスワードが違います' };

  session.role = 'admin';
  session.expiresAt = Date.now() + (SESSION_TTL_SEC * 1000);
  _saveSession_(sessionToken, session);
  const baseUrl = ScriptApp.getService().getUrl();
  return {
    success: true,
    adminSessionToken: sessionToken,
    redirectUrl: baseUrl ? (baseUrl + '?view=admin&st=' + encodeURIComponent(sessionToken)) : ''
  };
}

function registerStaff(req, sessionToken) {
  _requireSession_(sessionToken);
  const targetHours = parseTargetHoursValue(req.targetHours);
  SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('スタッフ一覧')
    .appendRow(['申請中', req.name, "'" + req.pin, "'" + req.phone, req.email, req.note, '', targetHours || '']);
  return { success: true, message: "店長に申請を送りました" };
}

function updateStaffProfile(oldName, req, sessionToken) {
  _requireSession_(sessionToken);
  const sheet = SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('スタッフ一覧');
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == oldName) {
      sheet.getRange(i + 1, 2).setValue(req.name);
      sheet.getRange(i + 1, 3).setValue("'" + req.pin);
      sheet.getRange(i + 1, 4).setValue("'" + req.phone);
      sheet.getRange(i + 1, 5).setValue(req.email);
      sheet.getRange(i + 1, 6).setValue(req.note);
      sheet.getRange(i + 1, 8).setValue(parseTargetHoursValue(req.targetHours) || '');
      return { success: true, message: "更新しました！", newName: req.name };
    }
  }
  return { success: false, message: "エラーが発生しました" };
}

/** ====== 募集リストをフェーズ別に返す ====== */
function _todayEndJST() {
  const now = new Date();
  const end = new Date(now);
  end.setHours(23,59,59,999);
  return end;
}
function _parseDeadlineToDate(deadlineStr) {
  if (!deadlineStr) return null;
  const s = String(deadlineStr).replace(/\//g,'-');
  const d = new Date(s + "T23:59:59");
  return isNaN(d.getTime()) ? null : d;
}

function getRecruitListsForAdmin() {
  const res = { all: [], open: [], edit: [], confirmed: [] };
  const ss = SpreadsheetApp.openById(BASE_SS_ID);
  const hSheet = ss.getSheetByName('募集履歴一覧');
  if(!hSheet) return res;

  const hData = hSheet.getDataRange().getDisplayValues();
  const todayEnd = _todayEndJST();

  for(let i=1; i<hData.length; i++){
    const status = hData[i][0];
    const id = hData[i][2];
    const label = `[${status}] ${hData[i][3]}`;
    const deadline = _parseDeadlineToDate(hData[i][4]);
    const isExpired = deadline ? (deadline < todayEnd) : false;

    const item = { id, label, status, deadline: hData[i][4] };
    res.all.push(item);

    if(status === '募集中' && !isExpired) res.open.push(item);
    if(status === '調整中' || (status === '募集中' && isExpired)) res.edit.push(item);
    if(status === '確定済') res.confirmed.push(item);
  }

  res.all.reverse();
  res.open.reverse();
  res.edit.reverse();
  res.confirmed.reverse();
  return res;
}

/** ====== 必要枠（required）関連 ====== */
function ensureReqSheet(recruitSs) {
  const sheet = recruitSs.getSheetByName('必要枠') || recruitSs.insertSheet('必要枠');
  const header = ['日付', '開始時間', '終了時間', '必要人数'];
  const first = sheet.getRange(1,1,1,header.length).getDisplayValues()[0];
  if(first.join('') !== header.join('')) {
    sheet.clearContents();
    sheet.getRange(1,1,1,header.length).setValues([header]);
  }
  return sheet;
}

function ensureReqDefaultSheet(recruitSs) {
  const sheet = recruitSs.getSheetByName('必要枠デフォルト') || recruitSs.insertSheet('必要枠デフォルト');
  const header = ['曜日', '開始時間', '必要人数'];
  const first = sheet.getRange(1, 1, 1, header.length).getDisplayValues()[0];
  if (first.join('') !== header.join('')) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sheet;
}

function buildReqDefaultRows_(recruitSettings, defaultNum) {
  const rows = [];
  const weekdays = ['月', '火', '水', '木', '金', '土', '日'];
  const defStartH = safeGetHour(recruitSettings.defStart, 8);
  const defEndH = safeGetHour(recruitSettings.defEnd, 20);

  weekdays.forEach(function(weekday) {
    for (let h = defStartH; h < defEndH; h++) {
      rows.push([weekday, h + ':00', Number(defaultNum || 0)]);
    }
  });
  return rows;
}

function buildDefaultReqDefaultsFromSettings(recruitSs, recruitSettings, defaultNum) {
  const sheet = ensureReqDefaultSheet(recruitSs);
  const rows = buildReqDefaultRows_(recruitSettings, defaultNum);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['曜日', '開始時間', '必要人数']]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  return true;
}

function getReqDefaultMap(recruitSs) {
  const sheet = recruitSs.getSheetByName('必要枠デフォルト');
  const map = {};
  if (!sheet) return map;
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    map[data[i][0] + '|' + data[i][1]] = parseInt(data[i][2], 10) || 0;
  }
  return map;
}

function _weekdayFromDateLabel_(dateLabel) {
  const m = String(dateLabel || '').match(/[(（]([日月火水木金土])[)）]/);
  return m ? m[1] : '';
}

function applyReqDefaultsToSheet_(recruitSs, recruitSettings, reqDefaultMap) {
  const sheet = ensureReqSheet(recruitSs);
  const rows = [];
  const defStartH = safeGetHour(recruitSettings.defStart, 8);
  const defEndH = safeGetHour(recruitSettings.defEnd, 20);

  recruitSettings.times.forEach(function(t) {
    const h = parseInt(String(t).split(':')[0], 10);
    recruitSettings.dates.forEach(function(d) {
      let isClosed = false;
      let isOut = false;
      const ex = (recruitSettings.exceptions || []).find(function(e) { return e.formattedDate === d; });
      if (ex) {
        if (ex.type === 'close') isClosed = true;
        else if (ex.type === 'short') {
          const sH = safeGetHour(ex.start, defStartH);
          const eH = safeGetHour(ex.end, defEndH);
          if (h < sH || h >= eH) isOut = true;
        }
      } else if (h < defStartH || h >= defEndH) {
        isOut = true;
      }
      if (isClosed || isOut) return;

      const weekday = _weekdayFromDateLabel_(d);
      const num = parseInt(reqDefaultMap[weekday + '|' + t], 10) || 0;
      rows.push([d, t, getEndTimeStr(t), num]);
    });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, 4).setValues([['日付', '開始時間', '終了時間', '必要人数']]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return true;
}

function buildDefaultReqSlotsFromSettings(recruitSs, recruitSettings, defaultNum) {
  buildDefaultReqDefaultsFromSettings(recruitSs, recruitSettings, defaultNum);
  applyReqDefaultsToSheet_(recruitSs, recruitSettings, getReqDefaultMap(recruitSs));
  return true;
}

function getReqMap(recruitSs) {
  const sheet = recruitSs.getSheetByName('必要枠');
  const map = {};
  if(!sheet) return map;
  const data = sheet.getDataRange().getDisplayValues();
  for(let i=1;i<data.length;i++){
    const d = data[i][0], t = data[i][1];
    const num = parseInt(data[i][3],10) || 0;
    map[`${d} ${t}`] = num;
  }
  return map;
}

/** ====== Staff Init ====== */
function getStaffInitData(staffName, recruitId, sessionToken) {
  _requireSession_(sessionToken);
  const res = {
    recruit: null,
    shifts: [],
    recruitList: [],
    activeStaffs: [],
    recruitStatus: '募集中',
    notices: [],
    weeklyShifts: [],
    reqMap: {},
    confirmedHours: 0,
    submitComment: ''
  };
  const ss = SpreadsheetApp.openById(BASE_SS_ID);

  // notices
  const noticeSheet = ss.getSheetByName('お知らせ');
  if (noticeSheet) {
    const nData = noticeSheet.getDataRange().getDisplayValues();
    for (let i = nData.length - 1; i >= 1; i--) {
      if (nData[i][1] === staffName || nData[i][1] === '全員') {
        res.notices.push({ date: nData[i][0], title: nData[i][2], msg: nData[i][3] });
        if (res.notices.length >= 10) break;
      }
    }
  }

  // recruit list (+status)
  const hSheet = ss.getSheetByName('募集履歴一覧');
  let confirmedFiles = [];
  if(hSheet) {
    const hData = hSheet.getDataRange().getDisplayValues();
    for(let i = 1; i < hData.length; i++) {
      res.recruitList.push({ id: hData[i][2], label: `[${hData[i][0]}] ${hData[i][3]}`, status: hData[i][0] });
      if (hData[i][2] === recruitId || (!recruitId && i === hData.length - 1)) res.recruitStatus = hData[i][0];
      if (hData[i][0] === '確定済') {
        let match = String(hData[i][5]).match(/\/d\/([a-zA-Z0-9-_]+)/);
        if(match) confirmedFiles.push(match[1]);
      }
    }
    res.recruitList.reverse();
  }

  // weekly (latest 2 confirmed)
  try {
    for(let i = confirmedFiles.length - 1; i >= Math.max(0, confirmedFiles.length - 2); i--) {
      let sSheet = SpreadsheetApp.openById(confirmedFiles[i]).getSheetByName('提出情報');
      if(sSheet) {
        let sData = sSheet.getDataRange().getDisplayValues();
        for(let j=1; j<sData.length; j++) {
          if(sData[j][0] === '確定' && sData[j][1] === staffName) res.weeklyShifts.push({date: sData[j][2], time: sData[j][3]});
        }
      }
    }
  } catch(e){}

  // active staff names
  const staffSheet = ss.getSheetByName('スタッフ一覧');
  if(staffSheet) res.activeStaffs = staffSheet.getDataRange().getDisplayValues().filter(r => r[0] === '稼働中').map(r => r[1]);

  const recruitSs = getRecruitSS(recruitId);
  if (!recruitSs) return res;

  const setting = recruitSs.getSheetByName('基本設定');
  if (setting) {
    const data = setting.getDataRange().getDisplayValues();
    if(data.length > 1 && data[1][0]) res.recruit = parseRecruitSettings(data);
  }

  if(res.recruit){
    res.reqMap = getReqMap(recruitSs);
  }

  const submitSheet = recruitSs.getSheetByName('提出情報');
  if (submitSheet) {
    const submitData = submitSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < submitData.length; i++) {
      res.shifts.push({ status: submitData[i][0], name: submitData[i][1], date: submitData[i][2], time: submitData[i][3], note: submitData[i][6] });
      if (submitData[i][1] === staffName && submitData[i][0] === '希望' && !res.submitComment) res.submitComment = submitData[i][6] || '';
      if (submitData[i][1] === staffName && submitData[i][0] === '確定') res.confirmedHours += 1;
    }
  }
  return res;
}

function submitShiftData(recruitId, name, slots, comment, sessionToken) {
  _requireSession_(sessionToken);
  const recruitSs = getRecruitSS(recruitId);
  if (!recruitSs) return { success: false, message: "対象のシフトが見つかりません" };

  const sheet = recruitSs.getSheetByName('提出情報');
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === '希望' && data[i][1] === name) sheet.deleteRow(i + 1);
  }

  if (slots.length > 0) {
    const commentText = String(comment || '');
    const newRows = slots.map(s => ['希望', name, s.date, s.time, getEndTimeStr(s.time), new Date(), commentText]);
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  return { success: true, message: "希望シフトを送信しました！" };
}

function sendChangeRequest(recruitId, dt, name, reason, sessionToken) {
  _requireSession_(sessionToken);
  const recruitSs = getRecruitSS(recruitId);
  if (!recruitSs) return { success: false, message: "エラー" };

  const sheet = recruitSs.getSheetByName('提出情報');
  const parts = dt.split(" ");
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] === '確定' || data[i][0] === '仮決定') && data[i][1] === name && data[i][2] == parts[0] && data[i][3] == parts[1]) {
      sheet.getRange(i + 1, 1).setValue('変更申請');
      sheet.getRange(i + 1, 7).setValue("理由: " + reason);
      return { success: true, message: "店長に申請を送りました" };
    }
  }
  return { success: false, message: "該当のシフトが見つかりません" };
}

function applyForVacancy(recruitId, dt, name, sessionToken) {
  _requireSession_(sessionToken);
  const recruitSs = getRecruitSS(recruitId);
  if (!recruitSs) return { success: false, message: "エラー" };

  const parts = dt.split(" ");
  recruitSs.getSheetByName('提出情報').appendRow(['立候補', name, parts[0], parts[1], getEndTimeStr(parts[1]), new Date(), '']);
  return { success: true, message: "空き枠に立候補しました！" };
}

/** ====== Admin Init ====== */
function getAdminInitData(recruitId, sessionToken) {
  _requireAdminSession_(sessionToken);
  const res = {
    alerts: [],
    staffs: [],
    currentRecruit: null,
    currentRecruitId: null,
    shifts: [],
    recruitLists: { all: [], open: [], edit: [], confirmed: [] },
    recruitList: [],
    todayShifts: [],
    todayDateStr: '',
    reqMap: {},
    reqDefaultMap: {},
    allowedEmails: []
  };

  const ss = SpreadsheetApp.openById(BASE_SS_ID);

  const staffSheet = ss.getSheetByName('スタッフ一覧');
  if (staffSheet) {
    const staffData = staffSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < staffData.length; i++) {
      res.staffs.push({ status: staffData[i][0], name: staffData[i][1], phone: fixZero(staffData[i][3]), email: staffData[i][4], note: staffData[i][5], role: staffData[i][6], targetHours: parseTargetHoursValue(staffData[i][7]) });
      if (staffData[i][0] === '申請中') res.alerts.push({ type: '新規登録', name: staffData[i][1], msg: 'スタッフ登録の申請があります' });
    }
  }

  res.recruitLists = getRecruitListsForAdmin();
  res.recruitList = (res.recruitLists && res.recruitLists.all) ? res.recruitLists.all.slice() : [];
  res.allowedEmails = getAllowedEmailEntries_();

  // today shifts from last 2 confirmed
  let confirmedFiles = [];
  const hSheet = ss.getSheetByName('募集履歴一覧');
  if(hSheet){
    const hData = hSheet.getDataRange().getDisplayValues();
    for(let i=1;i<hData.length;i++){
      if(hData[i][0] === '確定済'){
        let match = String(hData[i][5]).match(/\/d\/([a-zA-Z0-9-_]+)/);
        if(match) confirmedFiles.push(match[1]);
      }
    }
  }

  let today = new Date(); const days = ["(日)","(月)","(火)","(水)","(木)","(金)","(土)"];
  let todayStr = `${today.getMonth()+1}/${today.getDate()}${days[today.getDay()]}`;
  res.todayDateStr = todayStr;

  try {
    for(let i = confirmedFiles.length - 1; i >= Math.max(0, confirmedFiles.length - 2); i--) {
      let sSheet = SpreadsheetApp.openById(confirmedFiles[i]).getSheetByName('提出情報');
      if(sSheet) {
        let sData = sSheet.getDataRange().getDisplayValues();
        for(let j=1; j<sData.length; j++) {
          if(sData[j][0] === '確定' && sData[j][2] === todayStr) res.todayShifts.push({name: sData[j][1], time: sData[j][3]});
        }
      }
    }
  } catch(e){}

  // recruitId fallback
  if(!recruitId && res.recruitList.length > 0) recruitId = res.recruitList[0].id;
  res.currentRecruitId = recruitId || null;

  const recruitSs = getRecruitSS(recruitId);
  if (recruitSs) {
    const setSheet = recruitSs.getSheetByName('基本設定');
    if (setSheet) {
      const setVals = setSheet.getDataRange().getDisplayValues();
      if(setVals.length > 1 && setVals[1][0] && setVals[1][1]) res.currentRecruit = parseRecruitSettings(setVals);
    }

    if(res.currentRecruit){
      res.reqMap = getReqMap(recruitSs);
      res.reqDefaultMap = getReqDefaultMap(recruitSs);
    }

    const subSheet = recruitSs.getSheetByName('提出情報');
    if (subSheet) {
      const subData = subSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < subData.length; i++) {
        res.shifts.push({ status: subData[i][0], name: subData[i][1], date: subData[i][2], time: subData[i][3], note: subData[i][6] });

        if (subData[i][0] === '変更申請') res.alerts.push({ type: '変更申請', name: subData[i][1], msg: `${subData[i][2]} ${subData[i][3]}`, reason: subData[i][6] });
        if (subData[i][0] === '立候補') res.alerts.push({ type: '立候補', name: subData[i][1], msg: `${subData[i][2]} ${subData[i][3]}` });
      }
    }
  }
  return res;
}

/** ====== 新規募集作成 ====== */
function startNewRecruit(data, sessionToken) {
  _requireAdminSession_(sessionToken);
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const newFile = DriveApp.getFileById(TEMPLATE_SS_ID).makeCopy(`シフト募集_${data.startDate}〜${data.endDate}`, folder);

    const recruitSs = SpreadsheetApp.openById(newFile.getId());
    recruitSs.getSheetByName('基本設定').getRange('A2:F2')
      .setValues([[data.startDate, data.endDate, data.deadline, data.defaultStart, data.defaultEnd, data.exceptions]]);

    const setVals = recruitSs.getSheetByName('基本設定').getDataRange().getDisplayValues();
    const settings = parseRecruitSettings(setVals);
    const defaultNeed = Number(data.basicNeed || 0);
    buildDefaultReqSlotsFromSettings(recruitSs, settings, defaultNeed);

    const recruitId = 'REC'+Date.now();
    SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('募集履歴一覧')
      .appendRow(['募集中', new Date(), recruitId, `${data.startDate}〜${data.endDate}`, data.deadline, newFile.getUrl()]);

    sendWNotice_('全員', '📢 新しいシフト募集', `${data.startDate}〜のシフト募集が開始されました。\n締切は ${data.deadline} です。`);
    return { success: true, message: "募集を開始しました！" };
  } catch(err) {
    return { success: false, message: "処理エラー: " + err.message };
  }
}

function updateAdminStaffStatus(name, status, role, note, sessionToken) {
  _requireAdminSession_(sessionToken);
  const sheet = SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('スタッフ一覧');
  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == name) {
      sheet.getRange(i + 1, 1).setValue(status);
      sheet.getRange(i + 1, 6).setValue(note);
      sheet.getRange(i + 1, 7).setValue(role);
      return { success: true, message: "更新しました" };
    }
  }
  return { success: false, message: "エラー" };
}

function saveReqDefaults(recruitId, defaults, sessionToken) {
  _requireAdminSession_(sessionToken);
  const recruitSs = getRecruitSS(recruitId);
  if (!recruitSs) return { success: false, message: '対象の募集が見つかりません' };

  const sheet = ensureReqDefaultSheet(recruitSs);
  const header = ['曜日', '開始時間', '必要人数'];
  const rows = (defaults || []).map(function(row) {
    return [row.weekday, row.time, Number(row.num || 0)];
  });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([header]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  return { success: true, message: 'デフォルト設定を保存しました' };
}

function applyReqDefaultsToRecruit(recruitId, sessionToken) {
  _requireAdminSession_(sessionToken);
  const recruitSs = getRecruitSS(recruitId);
  if (!recruitSs) return { success: false, message: '対象の募集が見つかりません' };
  const setSheet = recruitSs.getSheetByName('基本設定');
  if (!setSheet) return { success: false, message: '基本設定が見つかりません' };

  const setVals = setSheet.getDataRange().getDisplayValues();
  if (setVals.length <= 1) return { success: false, message: '基本設定が未入力です' };
  const settings = parseRecruitSettings(setVals);
  applyReqDefaultsToSheet_(recruitSs, settings, getReqDefaultMap(recruitSs));
  return { success: true, message: 'デフォルト設定を実表へ反映しました' };
}

function saveReqSlots(recruitId, slots, sessionToken) {
  _requireAdminSession_(sessionToken);
  const recruitSs = getRecruitSS(recruitId);
  if (!recruitSs) return { success: false, message: "対象の募集が見つかりません" };

  const sheet = ensureReqSheet(recruitSs);
  const header = ['日付', '開始時間', '終了時間', '必要人数'];
  const rows = (slots || []).map(s => [s.date, s.time, getEndTimeStr(s.time), Number(s.num || 0)]);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);

  return { success: true, message: "保存しました" };
}

/** ====== 管理者：各種アクション ====== */
function executeShiftAction(recruitId, action, dt, staffName, replyMsg = "", changeToStaff = "", sessionToken) {
  _requireAdminSession_(sessionToken);
  const recruitSs = getRecruitSS(recruitId); if (!recruitSs) return { success: false, message: "エラー" };
  const sheet = recruitSs.getSheetByName('提出情報'); const parts = dt.split(" "); const tDate = parts[0]; const tTime = parts[1];
  const ts = new Date();
  const data = sheet.getDataRange().getDisplayValues();

  if (action === 'delete') action = 'approve_change';

  if (action === 'approve_change') {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === '変更申請' && data[i][1] === staffName && data[i][2] == tDate && data[i][3] == tTime) {
        sheet.deleteRow(i + 1);
        sendWNotice_(staffName, "✅ 休み申請【承認】", `${dt} の休みを承認しました。\n\n[店長より]\n${replyMsg||'了解です'}`);
        return { success: true, message: "承認しました" };
      }
    }
  } else if (action === 'reject_change') {
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === '変更申請' && data[i][1] === staffName && data[i][2] == tDate && data[i][3] == tTime) {
        sheet.getRange(i + 1, 1).setValue('確定');
        sendWNotice_(staffName, "❌ 休み申請【見送り】", `${dt} の休みは見送らせてください。\n\n[店長より]\n${replyMsg||'お願いします'}`);
        return { success: true, message: "却下しました" };
      }
    }
  } else if (action === 'approve_volunteer') {
    const reqMap = getReqMap(recruitSs);
    const required = reqMap[`${tDate} ${tTime}`] || 0;

    let assigned = 0;
    for(let i=1;i<data.length;i++){
      if(data[i][2]==tDate && data[i][3]==tTime){
        if(data[i][0]==='確定' || data[i][0]==='変更申請') assigned++;
      }
    }
    if(required > 0 && assigned >= required){
      return { success:false, message:"⚠️ すでに必要人数を満たしています（欠員なし）" };
    }

    let found = false;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === '立候補' && data[i][2] == tDate && data[i][3] == tTime) {
        if (data[i][1] === staffName) {
          sheet.getRange(i + 1, 1).setValue('確定');
          sheet.getRange(i + 1, 7).setValue('立候補承認');
          sendWNotice_(staffName, "🎉 立候補【承認】", `${dt} の立候補を承認しました！\n\n[店長より]\n${replyMsg||'助かります'}`);
          found = true;
        } else {
          let other = data[i][1];
          sheet.deleteRow(i + 1);
          sendWNotice_(other, "🙏 立候補【見送り】", `${dt} は他の方にお願いしました。`);
        }
      }
    }
    return { success: true, message: found ? "立候補を承認しました" : "対象の立候補が見つかりません" };

  } else if (action === 'reject_volunteer') {
    let rejected = [];
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === '立候補' && data[i][2] == tDate && data[i][3] == tTime) {
        rejected.push(data[i][1]);
        sheet.deleteRow(i + 1);
      }
    }
    rejected.forEach(name => sendWNotice_(name, "🙏 立候補【見送り】", `${dt} の立候補は見送らせてください。\n\n[店長より]\n${replyMsg||'またお願いします'}`));
    return { success: true, message: "全員を却下しました" };

  } else if (action === 'manual_delete') {
    for (let i = data.length - 1; i >= 1; i--) {
      if ((data[i][0] === '仮決定' || data[i][0] === '確定') && data[i][1] === staffName && data[i][2] == tDate && data[i][3] == tTime) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "削除しました" };
      }
    }
  } else if (action === 'manual_change') {
    for (let i = data.length - 1; i >= 1; i--) {
      if ((data[i][0] === '仮決定' || data[i][0] === '確定') && data[i][1] === staffName && data[i][2] == tDate && data[i][3] == tTime) {
        sheet.getRange(i + 1, 2).setValue(changeToStaff);
        sheet.getRange(i + 1, 7).setValue('手動変更');
        return { success: true, message: "変更しました" };
      }
    }
  } else if (action === 'manual_add') {
    if(!staffName || staffName === '欠員') return { success:false, message:"欠員は追加できません（必要人数との差で自動表示されます）" };
    sheet.appendRow(["仮決定", staffName, tDate, tTime, getEndTimeStr(tTime), ts, '手動追加']);
    return { success: true, message: "追加しました" };
  }
  return { success: false, message: "失敗しました" };
}

/** ====== AI下書き（必要枠を確定ソースとして使用） ====== */
function generateShiftByAI(options, sessionToken) {
  _requireAdminSession_(sessionToken);
  try {
    const recruitSs = getRecruitSS(options.recruitId); if (!recruitSs) return { success: false, message: "エラー" };
    const normalize = str => String(str).replace(/[\s　]+/g, "").trim();
    const hourOf = t => parseInt(String(t).split(':')[0], 10);

    const reqSheet = recruitSs.getSheetByName('必要枠');
    if(!reqSheet) return { success:false, message:"⚠️ 必要人数（必要枠）が未設定です" };
    const reqData = reqSheet.getDataRange().getDisplayValues().slice(1);
    const requiredSlots = reqData.map(r => ({ date: r[0], time: r[1], num: parseInt(r[3],10)||0 })).filter(s => s.num > 0);
    if(requiredSlots.length === 0) return { success:false, message:"⚠️ 必要人数が0のままです（新規作成の基本必要人数を設定してください）" };

    const staffInfo = {};
    SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('スタッフ一覧').getDataRange().getDisplayValues()
      .forEach((r, i) => {
        if (i > 0 && r[0] === '稼働中') {
          const key = normalize(r[1]);
          const monthlyTarget = parseTargetHoursValue(r[7]);
          staffInfo[key] = { active: true, monthlyTarget: monthlyTarget };
        }
      });

    const submitSheet = recruitSs.getSheetByName('提出情報');
    const submitData = submitSheet.getDataRange().getDisplayValues();

    const requests = [];
    submitData.forEach((r, i) => {
      if(i > 0 && r[0] === '希望') requests.push({ name: normalize(r[1]), date: String(r[2]).trim(), time: String(r[3]).trim() });
    });
    if (requests.length === 0) return { success: false, message: "⚠️ 提出希望がありません。" };
    const slotKey = (date, time) => `${date}|${time}`;
    const remainingNeedBySlot = {};
    requiredSlots.forEach(s => { remainingNeedBySlot[slotKey(s.date, s.time)] = Number(s.num || 0); });
    const requestSlotKeysByName = {};
    requests.forEach(r => {
      if (!requestSlotKeysByName[r.name]) requestSlotKeysByName[r.name] = new Set();
      requestSlotKeysByName[r.name].add(slotKey(r.date, r.time));
    });
    const countRemainingOptions = name => {
      const set = requestSlotKeysByName[name];
      if (!set) return 0;
      let cnt = 0;
      set.forEach(k => { if ((remainingNeedBySlot[k] || 0) > 0) cnt++; });
      return cnt;
    };

    for (let i = submitData.length - 1; i >= 1; i--) { if (submitData[i][0] === '仮決定') submitSheet.deleteRow(i + 1); }

    // Track assigned hours for cap/break rules and monthly-target balancing.
    const assignedHoursByStaffDay = {};
    const assignedHoursByStaffTotal = {};
    const keyOf = (name, date) => `${name}__${date}`;
    const ensureTotal = name => {
      if (!Object.prototype.hasOwnProperty.call(assignedHoursByStaffTotal, name)) assignedHoursByStaffTotal[name] = 0;
      return assignedHoursByStaffTotal[name];
    };
    const getAssignedSet = (name, date) => {
      const key = keyOf(name, date);
      if (!assignedHoursByStaffDay[key]) assignedHoursByStaffDay[key] = new Set();
      return assignedHoursByStaffDay[key];
    };
    const hasGap = sortedHours => {
      for (let i = 1; i < sortedHours.length; i++) {
        if (sortedHours[i] - sortedHours[i - 1] > 1) return true;
      }
      return false;
    };
    const crossesLunch = sortedHours => {
      if (sortedHours.length === 0) return false;
      const minH = sortedHours[0];
      const maxH = sortedHours[sortedHours.length - 1];
      return minH < 12 && maxH >= 13;
    };
    const canAssignByRule = (name, date, timeStr) => {
      const h = hourOf(timeStr);
      if (isNaN(h)) return false;
      const set = getAssignedSet(name, date);
      if (set.has(h)) return false;
      if (set.size >= 8) return false; // 1日8時間上限

      const trial = Array.from(set);
      trial.push(h);
      trial.sort((a, b) => a - b);

      // 昼跨ぎなら 12:00-13:00 を休憩にする（12時台を勤務にしない）
      if (crossesLunch(trial) && trial.includes(12)) return false;

      // 昼跨ぎなし かつ 6時間以下なら休憩不要
      if (!crossesLunch(trial) && trial.length <= 6) return true;

      // それ以外は、連続勤務を避けるため1時間以上の休憩（ギャップ）を要求
      return hasGap(trial);
    };
    const weekdayFromDateLabel = dateLabel => {
      const m = String(dateLabel || "").match(/[(（]([日月火水木金土])[)）]/);
      return m ? m[1] : null;
    };
    const hourInRange = (hour, range) => {
      const start = parseInt(String(range.start || "").split(":")[0], 10);
      const end = parseInt(String(range.end || "").split(":")[0], 10);
      if (!isFinite(start) || !isFinite(end)) return false;
      return hour >= start && hour < end;
    };
    const canAssignByTargetUpper = name => {
      const info = staffInfo[name];
      if (!info || info.monthlyTarget == null) return true;
      return (ensureTotal(name) + 1) <= (Number(info.monthlyTarget) * 1.1);
    };
    const commitAssign = (name, date, timeStr) => {
      const h = hourOf(timeStr);
      getAssignedSet(name, date).add(h);
      assignedHoursByStaffTotal[name] = ensureTotal(name) + 1;
    };
    const candidateCompare = (a, b) => {
      const optA = countRemainingOptions(a.name);
      const optB = countRemainingOptions(b.name);
      if (optA !== optB) return optA - optB;

      const infoA = staffInfo[a.name];
      const infoB = staffInfo[b.name];
      const targetA = infoA && infoA.monthlyTarget != null ? Number(infoA.monthlyTarget) : null;
      const targetB = infoB && infoB.monthlyTarget != null ? Number(infoB.monthlyTarget) : null;
      const assignedA = ensureTotal(a.name);
      const assignedB = ensureTotal(b.name);
      const remainA = targetA == null ? null : (targetA - assignedA);
      const remainB = targetB == null ? null : (targetB - assignedB);

      if (remainA != null && remainB != null && remainA !== remainB) return remainB - remainA;
      if (remainA != null && remainB == null && remainA > 0) return -1;
      if (remainB != null && remainA == null && remainB > 0) return 1;
      if (assignedA !== assignedB) return assignedA - assignedB;
      return String(a.name).localeCompare(String(b.name));
    };

    const newRows = [];
    const ts = new Date();

    requiredSlots.sort((a, b) => {
      if (a.date !== b.date) return String(a.date).localeCompare(String(b.date));
      return hourOf(a.time) - hourOf(b.time);
    });

    requiredSlots.forEach(req => {
      const cands = requests.filter(r => r.date === req.date && r.time === req.time).sort(candidateCompare);
      let assignedCount = 0;
      for (let i = 0; i < cands.length && assignedCount < req.num; i++) {
        const cand = cands[i];
        if (!staffInfo[cand.name] || !staffInfo[cand.name].active) continue;
        if (!canAssignByTargetUpper(cand.name)) continue;
        if (!canAssignByRule(cand.name, cand.date, cand.time)) continue;
        newRows.push(['仮決定', cand.name, cand.date, cand.time, getEndTimeStr(cand.time), ts, "下書き"]);
        commitAssign(cand.name, cand.date, cand.time);
        remainingNeedBySlot[slotKey(req.date, req.time)] = Math.max(0, (remainingNeedBySlot[slotKey(req.date, req.time)] || 0) - 1);
        assignedCount++;
      }
    });

    if(newRows.length > 0) {
      submitSheet.getRange(submitSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);

      SpreadsheetApp.flush();
      return { success: true, message: `📝 ${newRows.length}枠を仮決定しました！` };
    } else {
      return { success: false, message: "⚠️ 割り当て可能なシフトがありません。" };
    }
  } catch (err) {
    return { success: false, message: "システムエラー: " + err.message };
  }
}

function publishRecruit(recruitId, sessionToken) {
  _requireAdminSession_(sessionToken);
  try {
    const sheet = SpreadsheetApp.openById(BASE_SS_ID).getSheetByName('募集履歴一覧');
    const data = sheet.getDataRange().getDisplayValues();
    let fileUrl = ""; let period = "";

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][2] === recruitId) {
        sheet.getRange(i + 1, 1).setValue('確定済');
        period = data[i][3];
        fileUrl = data[i][5];
        break;
      }
    }

    if (fileUrl) {
      const match = String(fileUrl).match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (match) {
        const tSheet = SpreadsheetApp.openById(match[1]).getSheetByName('提出情報');
        if (tSheet) {
          const tData = tSheet.getDataRange().getValues();
          for(let j=1; j<tData.length; j++){
            if(tData[j][0] === '仮決定') tSheet.getRange(j+1, 1).setValue('確定');
          }
        }
      }
      sendWNotice_('全員', '✅ シフト公開のお知らせ', `【${period}】のシフトが確定・公開されました！\n「シフト＞確定版」タブからご確認ください。`);
      return { success: true, message: "🎉 公開し、全員に通知を送りました！" };
    }
    return { success: false, message: "見つかりませんでした" };
  } catch (e) { return { success: false, message: "エラー: " + e.message }; }
}

function sendRemind(staffName, sessionToken) {
  _requireAdminSession_(sessionToken);
  sendWNotice_(staffName, "⚠️ シフト提出のお願い", "シフトの提出期限が迫っています。アプリから希望シフトを提出してください。");
  return { success: true, message: "催促通知を送りました！" };
}

function saveAllowedEmails(entries, sessionToken) {
  const session = _requireAdminSession_(sessionToken);
  const sheet = ensureAllowedEmailsSheet_();
  const deduped = [];
  const seen = {};

  (entries || []).forEach(function(entry) {
    const email = _normalizeEmail_((entry || {}).email);
    if (!email || seen[email]) return;
    seen[email] = true;
    deduped.push([
      (entry && entry.enabled === false) ? '無効' : '有効',
      email,
      String((entry || {}).note || ''),
      new Date(),
      session.email
    ]);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, ALLOWED_EMAILS_SHEET_HEADER.length).setValues([ALLOWED_EMAILS_SHEET_HEADER]);
  if (deduped.length > 0) {
    sheet.getRange(2, 1, deduped.length, deduped[0].length).setValues(deduped);
  }

  return {
    success: true,
    message: 'アクセス許可メールを保存しました',
    emails: getAllowedEmailEntries_()
  };
}
