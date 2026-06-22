// ═══════════════════════════════════════════════════════════════
// BPC CLIENT HUB — Google Apps Script Backend
// ═══════════════════════════════════════════════════════════════
// SETUP:
//   1. Create a Google Sheet with two tabs: "Codes" and "Events"
//   2. Paste this script in Extensions > Apps Script
//   3. Set SHEET_ID to your Sheet's ID (from the URL)
//   4. Set ADMIN_KEY to a secret string of your choice
//   5. Deploy > New deployment > Web app
//      - Execute as: Me
//      - Who has access: Anyone
//   6. Copy the deployment URL into index.html and admin.html
// ═══════════════════════════════════════════════════════════════

const SHEET_ID  = '1j6Z5asITjHty6mr_4-3D2baV7xh0uJSX76HMvPoik3E';
const ADMIN_KEY = 'Hagbierez1!';

// ── Sheet column indexes (0-based) ───────────────────────────
const C = {
  // Codes sheet
  CODE:       0,
  LABEL:      1,
  EDIT:       2,  // TRUE / FALSE
  ACTIVE:     3,
  CREATED:    4,
  USE_COUNT:  5,
  LAST_USED:  6,

  // Events sheet
  EV_TIME:    0,
  EV_SESSION: 1,
  EV_EVENT:   2,
  EV_CODE:    3,
  EV_IP:      4,
  EV_CITY:    5,
  EV_REGION:  6,
  EV_COUNTRY: 7,
  EV_ORG:     8,
  EV_DEVICE:  9,
  EV_OS:      10,
  EV_BROWSER: 11,
  EV_SCREEN:  12,
  EV_LANG:    13,
  EV_TZ:      14,
  EV_REF:     15,
  EV_UTM:     16,
  EV_RETURN:  17,
  EV_DURATION:18,
  EV_SCROLL:  19,
};

function getSheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

// ── Response helper (supports JSONP via ?callback=fn) ────────
function respond(data, cb) {
  const json = JSON.stringify(data);
  const out  = cb ? `${cb}(${json})` : json;
  const mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(out).setMimeType(mime);
}

// ═══════════════════════════════════════════════════════════════
// GET handler — handles reads AND admin write actions
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  const action = e.parameter.action;
  const key    = e.parameter.key;
  const cb     = e.parameter.callback || null; // JSONP callback

  if (action === 'codes') {
    return serveCodes(cb);
  }

  // Tracking actions — open endpoints (no admin key)
  if (action === 'login')      return logLoginGet(e.parameter, cb);
  if (action === 'exit')       return logExitGet(e.parameter, cb);
  if (action === 'heartbeat')  return logExitGet(e.parameter, cb); // periodic update, same logic as exit
  if (action === 'event')      return logEventGet(e.parameter, cb);

  if (key !== ADMIN_KEY) return respond({ error: 'unauthorized' }, cb);

  if (action === 'analytics')  return serveAnalytics(cb);
  if (action === 'addCode')    return addCode(e.parameter, cb);
  if (action === 'toggleCode') return toggleCode(e.parameter, cb);
  if (action === 'deleteCode') return deleteCode(e.parameter, cb);

  return respond({ error: 'unknown action' }, cb);
}

// ── Serve active codes to the client page ────────────────────
function serveCodes(cb) {
  const sheet = getSheet('Codes');
  const rows  = sheet.getDataRange().getValues().slice(1);
  const codes = {};
  rows.forEach(row => {
    const code   = String(row[C.CODE]).trim();
    const active = row[C.ACTIVE] === true || row[C.ACTIVE] === 'TRUE';
    if (code && active) {
      codes[code] = { edit: row[C.EDIT] === true || row[C.EDIT] === 'TRUE', label: row[C.LABEL] || '' };
    }
  });
  return respond({ codes }, cb);
}

// ── Serve analytics data to admin page ───────────────────────
function serveAnalytics(cb) {
  const codeSheet  = getSheet('Codes');
  const eventSheet = getSheet('Events');

  const codeRows  = codeSheet.getDataRange().getValues().slice(1);
  const eventRows = eventSheet.getDataRange().getValues().slice(1);

  const codes = codeRows.map(row => ({
    code:      String(row[C.CODE]).trim(),
    label:     row[C.LABEL],
    edit:      row[C.EDIT] === true || row[C.EDIT] === 'TRUE',
    active:    row[C.ACTIVE] === true || row[C.ACTIVE] === 'TRUE',
    created:   row[C.CREATED] ? new Date(row[C.CREATED]).toISOString() : '',
    useCount:  row[C.USE_COUNT] || 0,
    lastUsed:  row[C.LAST_USED] ? new Date(row[C.LAST_USED]).toISOString() : '',
  }));

  const events = eventRows.map(row => ({
    time:      row[C.EV_TIME] ? new Date(row[C.EV_TIME]).toISOString() : '',
    session:   row[C.EV_SESSION],
    event:     row[C.EV_EVENT],
    code:      row[C.EV_CODE],
    city:      row[C.EV_CITY],
    region:    row[C.EV_REGION],
    country:   row[C.EV_COUNTRY],
    device:    row[C.EV_DEVICE],
    os:        row[C.EV_OS],
    browser:   row[C.EV_BROWSER],
    screen:    row[C.EV_SCREEN],
    lang:      row[C.EV_LANG],
    tz:        row[C.EV_TZ],
    referrer:  row[C.EV_REF],
    returning: row[C.EV_RETURN],
    duration:  row[C.EV_DURATION],
    scroll:    row[C.EV_SCROLL],
  })).reverse(); // newest first

  return respond({ codes, events }, cb);
}

// ═══════════════════════════════════════════════════════════════
// POST handler — tracking only (login/exit events)
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  let data;
  const ct = (e.postData.type || '').toLowerCase();

  if (ct.includes('application/x-www-form-urlencoded')) {
    // sendBeacon with URLSearchParams body
    const p = e.parameter;
    if (p.action === 'exit') return logExitGet(p, null);
    return respond({ error: 'unknown action' });
  }

  try { data = JSON.parse(e.postData.contents); }
  catch(err) { return respond({ error: 'bad json' }); }

  if (data.event === 'login') return logLogin(data);
  if (data.event === 'exit')  return logExit(data);

  return respond({ error: 'unknown action' });
}

// ── Log login event ───────────────────────────────────────────
function logLogin(d) {
  const sheet = getSheet('Events');
  sheet.appendRow([
    new Date(), d.sessionId, 'login', d.code,
    d.ip, d.city, d.region, d.country, d.org,
    d.device, d.os, d.browser, d.screen, d.lang, d.tz,
    d.referrer, d.utm, d.returning ? 'returning' : 'new',
    '', '', // duration & scroll filled on exit
  ]);

  // Update use count + last used on Codes sheet
  const codeSheet = getSheet('Codes');
  const rows = codeSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][C.CODE]).trim() === String(d.code).trim()) {
      codeSheet.getRange(i + 1, C.USE_COUNT + 1).setValue((rows[i][C.USE_COUNT] || 0) + 1);
      codeSheet.getRange(i + 1, C.LAST_USED + 1).setValue(new Date());
      break;
    }
  }
  return respond({ ok: true });
}

// ── Log exit / update duration & scroll ──────────────────────
function logExit(d) {
  const sheet = getSheet('Events');
  const data  = sheet.getDataRange().getValues();
  // Find the matching login row for this session (search from bottom)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][C.EV_SESSION] === d.sessionId && data[i][C.EV_EVENT] === 'login') {
      sheet.getRange(i + 1, C.EV_DURATION + 1).setValue(d.duration);
      sheet.getRange(i + 1, C.EV_SCROLL + 1).setValue(d.scrollDepth);
      break;
    }
  }
  return respond({ ok: true });
}

// ── Log login via GET (image pixel) ──────────────────────────
function logLoginGet(p, cb) {
  const sheet = getSheet('Events');
  sheet.appendRow([
    new Date(), p.sessionId, 'login', p.code,
    p.ip||'', p.city||'', p.region||'', p.country||'', p.org||'',
    p.device, p.os, p.browser, p.screen, p.lang, p.tz,
    p.ref || '', p.utm || '', p.returning === '1' ? 'returning' : 'new',
    '', '',                      // duration & scroll — filled on exit
  ]);
  // Update use count + last used on Codes sheet
  const codeSheet = getSheet('Codes');
  const rows = codeSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][C.CODE]).trim() === String(p.code).trim()) {
      codeSheet.getRange(i + 1, C.USE_COUNT + 1).setValue((rows[i][C.USE_COUNT] || 0) + 1);
      codeSheet.getRange(i + 1, C.LAST_USED + 1).setValue(new Date());
      break;
    }
  }
  return respond({ ok: true }, cb);
}

// ── Log exit via GET (fetch keepalive) ───────────────────────
function logExitGet(p, cb) {
  const sheet = getSheet('Events');
  const data  = sheet.getDataRange().getValues();
  const sid   = String(p.sessionId || '');
  let found = false;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][C.EV_SESSION]) === sid && data[i][C.EV_EVENT] === 'login') {
      sheet.getRange(i + 1, C.EV_DURATION + 1).setValue(Number(p.duration) || 0);
      sheet.getRange(i + 1, C.EV_SCROLL   + 1).setValue(Number(p.scrollDepth) || 0);
      found = true;
      break;
    }
  }
  // Fallback: write a standalone exit row if login row wasn't found yet
  if (!found && sid) {
    const row = new Array(20).fill('');
    row[C.EV_TIME]     = new Date();
    row[C.EV_SESSION]  = sid;
    row[C.EV_EVENT]    = 'exit';
    row[C.EV_CODE]     = p.code || '';
    row[C.EV_DURATION] = Number(p.duration) || 0;
    row[C.EV_SCROLL]   = Number(p.scrollDepth) || 0;
    sheet.appendRow(row);
  }
  return respond({ ok: true }, cb);
}

// ── Log arbitrary event (quote_request, etc.) via GET ─────────
function logEventGet(p, cb) {
  const sheet = getSheet('Events');
  const row   = new Array(20).fill('');
  row[C.EV_TIME]    = new Date();
  row[C.EV_SESSION] = p.sessionId || '';
  row[C.EV_EVENT]   = p.type || 'event';
  row[C.EV_CODE]    = p.code || '';
  sheet.appendRow(row);
  return respond({ ok: true }, cb);
}

// ── Add new code ──────────────────────────────────────────────
function addCode(d, cb) {
  const sheet = getSheet('Codes');
  const code  = String(d.code || '').trim();
  if (!code) return respond({ error: 'empty code' }, cb);

  const existing = sheet.getDataRange().getValues().slice(1);
  if (existing.some(r => String(r[C.CODE]).trim() === code)) {
    return respond({ error: 'duplicate' }, cb);
  }

  sheet.appendRow([
    code,
    d.label || '',
    d.edit === 'true' || d.edit === true,
    true,
    new Date(),
    0,
    '',
  ]);
  return respond({ ok: true }, cb);
}

// ── Toggle code active/inactive ───────────────────────────────
function toggleCode(d, cb) {
  const sheet = getSheet('Codes');
  const rows  = sheet.getDataRange().getValues();
  const code  = String(d.code || '').trim();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][C.CODE]).trim() === code) {
      const current = rows[i][C.ACTIVE] === true || rows[i][C.ACTIVE] === 'TRUE';
      sheet.getRange(i + 1, C.ACTIVE + 1).setValue(!current);
      return respond({ ok: true, active: !current }, cb);
    }
  }
  return respond({ error: 'not found' }, cb);
}

// ── Delete code ───────────────────────────────────────────────
function deleteCode(d, cb) {
  const sheet = getSheet('Codes');
  const rows  = sheet.getDataRange().getValues();
  const code  = String(d.code || '').trim();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][C.CODE]).trim() === code) {
      sheet.deleteRow(i + 1);
      return respond({ ok: true }, cb);
    }
  }
  return respond({ error: 'not found' }, cb);
}

// ── One-time setup: create sheet headers ─────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let codes = ss.getSheetByName('Codes');
  if (!codes) codes = ss.insertSheet('Codes');
  if (codes.getLastRow() === 0) {
    codes.appendRow(['Code', 'Label', 'Edit', 'Active', 'Created', 'UseCount', 'LastUsed']);
    codes.getRange('1:1').setFontWeight('bold');
  }

  let events = ss.getSheetByName('Events');
  if (!events) events = ss.insertSheet('Events');
  if (events.getLastRow() === 0) {
    events.appendRow(['Time','SessionID','Event','Code','IP','City','Region','Country','Org','Device','OS','Browser','Screen','Lang','TZ','Referrer','UTM','Returning','Duration(s)','ScrollDepth(%)']);
    events.getRange('1:1').setFontWeight('bold');
  }
}
