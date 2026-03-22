//==========================================
// WebApp.gs (sample)
// Minimal sample for the template in this repository.
//==========================================

const GOOGLE_TOKENINFO_URL = 'https://oauth2.googleapis.com/tokeninfo';
const GOOGLE_OAUTH_CLIENT_ID =
  PropertiesService.getScriptProperties().getProperty('GOOGLE_OAUTH_CLIENT_ID') || '';
const SESSION_TTL_SEC = 60 * 60;

function doGet(e) {
  const sessionToken = e && e.parameter ? String(e.parameter.st || '') : '';
  const session = sessionToken ? getSessionFromToken_(sessionToken, true) : null;
  if (!session) return createUnauthorizedHtml_();
  return buildAppHtml_(sessionToken);
}

function doPost(e) {
  const bootstrapSessionToken = e && e.parameter ? String(e.parameter.bootstrapSessionToken || '') : '';
  if (bootstrapSessionToken) {
    const bootstrapSession = getSessionFromToken_(bootstrapSessionToken, true);
    if (!bootstrapSession) return createUnauthorizedHtml_();
    return buildAppHtml_(bootstrapSessionToken);
  }

  const request = parseJsonRequest_(e);
  if (!request.ok) {
    return jsonResponse_({
      status: 400,
      ok: false,
      code: 'invalid_json',
      message: 'invalid request body'
    });
  }

  const idToken = String(request.data.idToken || '');
  const action = String(request.data.action || '');
  const payload = request.data.payload && typeof request.data.payload === 'object' ? request.data.payload : {};

  if (action === 'revokeSession') {
    revokeSession_(String(payload.sessionToken || ''));
    return jsonResponse_({
      status: 200,
      ok: true,
      data: { revoked: true }
    });
  }

  if (!idToken) {
    return jsonResponse_({
      status: 401,
      ok: false,
      code: 'missing_token',
      message: 'id token is required'
    });
  }

  const verifyResult = verifyGoogleIdToken_(idToken);
  if (!verifyResult.ok) {
    return jsonResponse_({
      status: 401,
      ok: false,
      code: verifyResult.code || 'unauthorized',
      message: 'authentication failed'
    });
  }

  if (!isAllowedEmail_(verifyResult.email)) {
    return jsonResponse_({
      status: 403,
      ok: false,
      code: 'forbidden',
      message: 'permission denied'
    });
  }

  if (action === 'createSession') {
    const session = createSession_(verifyResult.email);
    return jsonResponse_({
      status: 200,
      ok: true,
      email: verifyResult.email,
      data: {
        sessionToken: session.token,
        expiresIn: session.expiresIn
      }
    });
  }

  if (action === 'ping') {
    return jsonResponse_({
      status: 200,
      ok: true,
      email: verifyResult.email,
      data: { now: new Date().toISOString() }
    });
  }

  return jsonResponse_({
    status: 400,
    ok: false,
    code: 'unknown_action',
    message: 'unknown action'
  });
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function buildAppHtml_(sessionToken) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.sessionToken = String(sessionToken || '');
  return template
    .evaluate()
    .setTitle('Embedded GAS App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createUnauthorizedHtml_() {
  const deny = HtmlService.createHtmlOutput(
    '<!doctype html><html><body><h3>401 Unauthorized</h3><p>Session is missing or expired.</p></body></html>'
  );
  deny.setTitle('Unauthorized');
  deny.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return deny;
}

function parseJsonRequest_(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
    const parsed = JSON.parse(raw || '{}');
    if (!parsed || typeof parsed !== 'object') return { ok: false };
    return { ok: true, data: parsed };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : 'parse_error'
    };
  }
}

function verifyGoogleIdToken_(idToken) {
  if (!GOOGLE_OAUTH_CLIENT_ID) return { ok: false, code: 'server_misconfigured' };

  let res;
  try {
    res = UrlFetchApp.fetch(
      GOOGLE_TOKENINFO_URL + '?id_token=' + encodeURIComponent(idToken),
      {
        method: 'get',
        muteHttpExceptions: true
      }
    );
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

function readAllowedEmails_() {
  const raw = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS') || '';
  return raw
    .split(',')
    .map(function(email) { return String(email || '').trim().toLowerCase(); })
    .filter(function(email) { return email; });
}

function isAllowedEmail_(email) {
  const allowList = readAllowedEmails_();
  if (allowList.length === 0) return false;
  return allowList.indexOf(String(email || '').toLowerCase()) !== -1;
}

function jsonResponse_(body) {
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function createSession_(email) {
  const token =
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '');
  const cacheKey = 'sess:' + token;
  const session = {
    email: String(email || '').toLowerCase(),
    issuedAt: Date.now(),
    expiresAt: Date.now() + (SESSION_TTL_SEC * 1000)
  };
  CacheService.getScriptCache().put(cacheKey, JSON.stringify(session), SESSION_TTL_SEC);
  return { token: token, expiresIn: SESSION_TTL_SEC };
}

function getSessionFromToken_(token, touch) {
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
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(session), SESSION_TTL_SEC);
    }
    return session;
  } catch (err) {
    return null;
  }
}

function revokeSession_(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('sess:' + String(token));
}

function requireSession_(sessionToken) {
  const session = getSessionFromToken_(sessionToken, true);
  if (!session) {
    throw new Error('セッションが切れました。再ログインしてください。');
  }
  return session;
}

// Example public GAS function.
function getInitialData(sessionToken) {
  const session = requireSession_(sessionToken);
  return {
    user: session.email,
    now: new Date().toISOString()
  };
}
