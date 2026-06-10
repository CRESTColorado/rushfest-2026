// ════════════════════════════════════════════════════════════════
// USYS PRESIDENTS CUP 2026 . BCC APPS SCRIPT (MERGED)
// ════════════════════════════════════════════════════════════════
// This single script serves THREE apps. Do not delete sections.
//
//   1. REF COACH APP      → actions: getGame (GET), submit (POST)
//      File: coach-app/index.html (existing, deployed)
//
//   2. COMMAND CENTER     → actions: auth, rulesStatus (POST/GET)
//      File: command-center/index.html (new)
//
//   3. RULES WALKTHROUGH  → POSTs name + sections on completion
//      File: rules/index.html (new)
//
// DEPLOYMENT:
//   1. Open the BCC Sheet → Extensions → Apps Script
//   2. Replace Code.gs with this file
//   3. Manage deployments → Edit current deployment
//      → Version: New version → Deploy
//      (Same URL keeps working for the existing Ref Coach App)
//   4. If this is a first deploy: Deploy → New deployment → Web app
//      Execute as: Me. Who has access: Anyone. Copy /exec URL.
//      Paste that URL into:
//        - coach-app/index.html        (if not already wired)
//        - command-center/index.html   (SCRIPT_URL)
//        - rules/index.html            (VALIDATION_URL)
// ════════════════════════════════════════════════════════════════


// ════════════════════════════════════════════════════════════════
// CONFIG
// ════════════════════════════════════════════════════════════════
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Ref Coach App tabs
const SCHEDULE_SHEET = 'GAME ASSIGNMENTS';
const SUBMIT_SHEET   = 'SUBMISSIONS';

// Command Center / Rules tabs
const STAFF_TAB         = 'Staff PINs';        // A=PIN, B=Name, C=Role, D=Greeting
const ROSTER_TAB        = 'Roster';            // existing Roster tab
const ROSTER_NAME_COL   = 18;                  // column R = Full Name (helper col)
const ROSTER_FLAG_COL   = 20;                  // column T = Rules Complete (new)
const LOG_TAB           = 'Rules Completion';  // A=Timestamp, B=Name as typed, C=Matched Roster Name, D=Sections


// ── Column positions in GAME ASSIGNMENTS (1-indexed). Ref Coach App.
const COL = {
  GAME_NUM:   1,
  DATE:       2,
  TIME:       3,
  FIELD:      4,
  AGE_GROUP:  5,
  GENDER:     6,
  ROUND:      7,
  HOME_TEAM:  8,
  AWAY_TEAM:  9,
  REFEREE:   10,
  AR1:       11,
  AR2:       12,
  FOURTH:    13,
  REF_COACH: 17,
};


// ════════════════════════════════════════════════════════════════
// ENTRY POINTS . Routed by action parameter.
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  // ── Ref Coach App ──
  if (action === 'getGame') {
    const field = parseInt(e.parameter.field);
    const date  = e.parameter.date;
    const time  = e.parameter.time;
    const result = getGame(field, date, time);
    return jsonRaw(result);
  }

  // ── Command Center ──
  if (action === 'rulesStatus') {
    try { return jsonOk(getRulesStatus()); }
    catch (err) { return jsonErr(err.toString()); }
  }

  if (action === 'ping') {
    return jsonOk({ message: 'pong' });
  }

  // ── Ref Coach App food alerts ──
  if (action === 'foodStatus') {
    return jsonRaw(getFoodStatus());
  }

  return jsonRaw({ error: 'Unknown action: ' + action });
}


function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); }
  catch (err) { return jsonRaw({ error: 'Bad JSON: ' + err.message }); }

  const action = body.action || '';

  // ── Ref Coach App ──
  if (action === 'submit') {
    try {
      writeSubmission(body.data);
      return jsonRaw({ success: true });
    } catch (err) {
      return jsonRaw({ error: err.message });
    }
  }

  // ── Ref Coach App food alerts ──
  if (action === 'foodCoach') {
    try {
      setFoodTime('foodCoachTime', body.timestamp);
      return jsonRaw({ success: true });
    } catch (err) {
      return jsonRaw({ error: err.message });
    }
  }

  if (action === 'foodRefs') {
    try {
      setFoodTime('foodRefTime', body.timestamp);
      return jsonRaw({ success: true });
    } catch (err) {
      return jsonRaw({ error: err.message });
    }
  }

  // ── Command Center ──
  if (action === 'auth') {
    try { return jsonOk(handleAuth(body.pin)); }
    catch (err) { return jsonErr(err.toString()); }
  }

  if (action === 'rulesComplete') {
    try { return jsonOk(handleRulesComplete(body)); }
    catch (err) { return jsonErr(err.toString()); }
  }

  // ── Rules Walkthrough bare payload (no action field) ──
  // rules/index.html POSTs {name, timestamp, sections} without an action.
  if (!action && body.name && body.sections) {
    try { return jsonOk(handleRulesComplete(body)); }
    catch (err) { return jsonErr(err.toString()); }
  }

  return jsonRaw({ error: 'Unknown action: ' + action });
}


// ════════════════════════════════════════════════════════════════
// REF COACH APP . getGame, writeSubmission, formatTime
// (Original code preserved verbatim. Do not modify without testing
// the coach-app/index.html flow end to end.)
// ════════════════════════════════════════════════════════════════
function getGame(fieldNum, dateStr, timeStr) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SCHEDULE_SHEET);
  const data  = sheet.getDataRange().getValues();

  const HEADER_ROWS = 3;

  const inputDate = new Date(dateStr);
  const [inputH, inputM] = timeStr.split(':').map(Number);
  const inputMins = inputH * 60 + inputM;

  const GAME_WINDOW_MINS = 90;

  for (let i = HEADER_ROWS; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.GAME_NUM - 1]) continue;

    const rowField = parseInt(row[COL.FIELD - 1]);
    if (rowField !== fieldNum) continue;

    const rowDate = row[COL.DATE - 1];
    if (!rowDate) continue;
    const rd = new Date(rowDate);
    if (rd.toDateString() !== inputDate.toDateString()) continue;

    const rowTimeRaw = row[COL.TIME - 1];
    let rowMins = 0;
    if (rowTimeRaw instanceof Date) {
      rowMins = rowTimeRaw.getHours() * 60 + rowTimeRaw.getMinutes();
    } else if (typeof rowTimeRaw === 'string') {
      const clean = rowTimeRaw.replace(/[^0-9:AaPpMm\s]/g,'').trim();
      const parts = clean.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
      if (parts) {
        let h = parseInt(parts[1]);
        const m = parseInt(parts[2]);
        const ampm = (parts[3] || '').toUpperCase();
        if (ampm === 'PM' && h < 12) h += 12;
        if (ampm === 'AM' && h === 12) h = 0;
        rowMins = h * 60 + m;
      }
    } else if (typeof rowTimeRaw === 'number') {
      rowMins = Math.round(rowTimeRaw * 24 * 60);
    }

    if (inputMins >= rowMins && inputMins <= rowMins + GAME_WINDOW_MINS) {
      const game = {
        gameNum:  row[COL.GAME_NUM - 1],
        field:    rowField,
        date:     dateStr,
        time:     formatTime(rowMins),
        ageGroup: (row[COL.AGE_GROUP - 1] || '') + ' ' + (row[COL.GENDER - 1] || ''),
        round:    row[COL.ROUND - 1]     || '',
        homeTeam: row[COL.HOME_TEAM - 1] || '',
        awayTeam: row[COL.AWAY_TEAM - 1] || '',
        referee:  row[COL.REFEREE - 1]   || '',
        ar1:      row[COL.AR1 - 1]       || '',
        ar2:      row[COL.AR2 - 1]       || '',
        fourth:   row[COL.FOURTH - 1]    || '',
        refCoach: row[COL.REF_COACH - 1] || '',
      };
      return { game };
    }
  }

  return { game: null };
}

function formatTime(totalMins) {
  const h = Math.floor(totalMins / 60);
  const m = totalMins % 60;
  const ampm = h < 12 ? 'AM' : 'PM';
  const h12 = h === 0 ? 12 : h > 12 ? h - 12 : h;
  return `${h12}:${String(m).padStart(2,'0')} ${ampm}`;
}

function writeSubmission(data) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SUBMIT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(SUBMIT_SHEET);
    const headers = [
      'Timestamp', 'Coach', 'Field', 'Game #', 'Age Group', 'Round', 'Game Time',
      'Game Difficulty',
      'Role', 'Official Name',
      'Performance', 'Professionalism', 'Coachability', 'Presence',
      'Public Notes', 'Private Notes',
      'Flag: Needs Obs', 'Flag: Standout', 'Flag: Not Ready', 'Flag: Concern',
      'Flag: Attitude', 'Flag: Injury', 'Flag: Diff Up',
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  const rows = (data.officials || []).map(o => [
    data.timestamp,
    data.coach,
    data.field,
    data.gameNum,
    data.ageGroup,
    data.round,
    data.gameTime,
    data.difficulty,
    o.role,
    o.name,
    o.perf  || 3,
    o.prof  || 3,
    o.coach || 3,
    o.pres  || 3,
    o.pub   || '',
    o.priv  || '',
    o.flags?.needs_obs  ? 'YES' : '',
    o.flags?.standout   ? 'YES' : '',
    o.flags?.not_ready  ? 'YES' : '',
    o.flags?.concern    ? 'YES' : '',
    o.flags?.attitude   ? 'YES' : '',
    o.flags?.injury     ? 'YES' : '',
    o.flags?.diff_up    ? 'YES' : '',
  ]);

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
}


// ════════════════════════════════════════════════════════════════
// COMMAND CENTER . AUTH
// Validates 4-digit PIN against the Staff PINs tab.
// ════════════════════════════════════════════════════════════════
function handleAuth(pin) {
  if (!pin) { throw new Error('Missing PIN'); }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tab = ss.getSheetByName(STAFF_TAB);
  if (!tab) {
    throw new Error('Staff PINs tab missing. Create it with columns: PIN | Name | Role | Greeting');
  }
  const last = tab.getLastRow();
  if (last < 2) { throw new Error('No PINs configured.'); }
  const rows = tab.getRange(2, 1, last - 1, 4).getValues();
  const target = String(pin).trim();
  for (let i = 0; i < rows.length; i++) {
    const rowPin = String(rows[i][0] || '').trim();
    if (rowPin && rowPin === target) {
      return {
        name:     String(rows[i][1] || '').trim(),
        role:     String(rows[i][2] || 'admin').trim().toLowerCase(),
        greeting: String(rows[i][3] || '').trim()
      };
    }
  }
  throw new Error('Incorrect PIN');
}


// ════════════════════════════════════════════════════════════════
// RULES COMPLETION . WRITE (from rules/index.html)
// ════════════════════════════════════════════════════════════════
function handleRulesComplete(payload) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let log = ss.getSheetByName(LOG_TAB);
  if (!log) {
    log = ss.insertSheet(LOG_TAB);
    const headers = ['Timestamp', 'Name (as typed)', 'Matched Roster Name', 'Sections'];
    log.getRange(1, 1, 1, headers.length).setValues([headers]);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  const roster = ss.getSheetByName(ROSTER_TAB);
  const submitted = String(payload.name || '').trim();
  const submittedLower = submitted.toLowerCase();
  if (!submitted) { throw new Error('Missing name'); }

  let matched = '';
  let matchRow = -1;

  if (roster && roster.getLastRow() > 1) {
    const names = roster.getRange(2, ROSTER_NAME_COL, roster.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      const rosterName = String(names[i][0] || '').trim();
      if (!rosterName) { continue; }
      const rosterLower = rosterName.toLowerCase();
      if (rosterLower === submittedLower ||
          rosterLower.indexOf(submittedLower) !== -1 ||
          submittedLower.indexOf(rosterLower) !== -1) {
        matched = rosterName;
        matchRow = i + 2;
        break;
      }
    }
  }

  log.appendRow([new Date(), submitted, matched, payload.sections || '']);

  if (roster && matchRow > 0) {
    roster.getRange(matchRow, ROSTER_FLAG_COL).setValue(true);
  }

  return { matched: matched, recorded: true };
}


// ════════════════════════════════════════════════════════════════
// RULES COMPLETION . READ (for Command Center dashboard)
// Returns { total, done, refs: [{ name, status, timestamp }] }
// ════════════════════════════════════════════════════════════════
function getRulesStatus() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const roster = ss.getSheetByName(ROSTER_TAB);
  if (!roster) { throw new Error('Roster tab not found'); }
  const lastRow = roster.getLastRow();
  if (lastRow < 2) { return { total: 0, done: 0, refs: [] }; }

  const range = roster.getRange(2, ROSTER_NAME_COL, lastRow - 1, ROSTER_FLAG_COL - ROSTER_NAME_COL + 1).getValues();

  const log = ss.getSheetByName(LOG_TAB);
  const tsByName = {};
  if (log && log.getLastRow() > 1) {
    const logRows = log.getRange(2, 1, log.getLastRow() - 1, 4).getValues();
    logRows.forEach(function(r) {
      const matched = String(r[2] || '').trim();
      if (matched) {
        const ts = r[0] instanceof Date ? r[0].toISOString() : String(r[0] || '');
        if (!tsByName[matched] || ts > tsByName[matched]) { tsByName[matched] = ts; }
      }
    });
  }

  const refs = [];
  let done = 0;
  range.forEach(function(row) {
    const name = String(row[0] || '').trim();
    if (!name) { return; }
    const flag = row[ROSTER_FLAG_COL - ROSTER_NAME_COL] === true;
    const status = flag ? 'done' : 'pending';
    if (flag) { done++; }
    refs.push({
      name: name,
      status: status,
      timestamp: flag ? (tsByName[name] || '') : ''
    });
  });

  return { total: refs.length, done: done, refs: refs };
}


// ════════════════════════════════════════════════════════════════
// REF COACH APP . FOOD ALERTS
// Cross-device broadcast of food-ready notifications.
// Stored in Script Properties so all coach app instances see updates.
// Stale alerts older than FOOD_CLEAR_MS are returned as 0.
// ════════════════════════════════════════════════════════════════
const FOOD_CLEAR_MS = 90 * 60 * 1000; // 90 minutes, matches APPindex.html FOOD_CLEAR

function setFoodTime(key, timestamp) {
  const ts = parseInt(timestamp) || Date.now();
  PropertiesService.getScriptProperties().setProperty(key, String(ts));
}

function getFoodStatus() {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  const coachRaw = parseInt(props.getProperty('foodCoachTime') || '0') || 0;
  const refRaw   = parseInt(props.getProperty('foodRefTime')   || '0') || 0;
  return {
    foodCoachTime: (coachRaw && (now - coachRaw) < FOOD_CLEAR_MS) ? coachRaw : 0,
    foodRefTime:   (refRaw   && (now - refRaw)   < FOOD_CLEAR_MS) ? refRaw   : 0
  };
}


// ════════════════════════════════════════════════════════════════
// JSON HELPERS
// jsonRaw . unwrapped JSON (preserves Ref Coach App response shape)
// jsonOk / jsonErr . wrapped JSON with status field (Command Center)
// ════════════════════════════════════════════════════════════════
function jsonRaw(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonErr(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
