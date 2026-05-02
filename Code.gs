// ════════════════════════════════════════════════════════════════
// USYS PRESIDENTS CUP 2026 — REF COACH APP · Google Apps Script
// ════════════════════════════════════════════════════════════════
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet → Extensions → Apps Script
// 2. Paste this entire file, replacing any existing code
// 3. Click Deploy → New Deployment → Web App
//    - Execute as: Me
//    - Who has access: Anyone
// 4. Copy the deployment URL → paste into coach-app/index.html as SCRIPT_URL
// 5. Re-deploy after any code changes (create new version)
// ════════════════════════════════════════════════════════════════

const SPREADSHEET_ID  = SpreadsheetApp.getActiveSpreadsheet().getId();
const SCHEDULE_SHEET  = 'GAME ASSIGNMENTS';
const SUBMIT_SHEET    = 'SUBMISSIONS';

// ── Column positions in GAME ASSIGNMENTS (1-indexed, adjust if sheet changes) ──
const COL = {
  GAME_NUM:   1,   // A - Game #
  DATE:       2,   // B - Date
  TIME:       3,   // C - Time
  FIELD:      4,   // D - Field
  AGE_GROUP:  5,   // E - Age Group
  GENDER:     6,   // F - Gender
  ROUND:      7,   // G - Round
  HOME_TEAM:  8,   // H - Home Team
  AWAY_TEAM:  9,   // I - Away Team
  REFEREE:   10,   // J - Referee
  AR1:       11,   // K - AR1
  AR2:       12,   // L - AR2
  FOURTH:    13,   // M - 4th Official
  REF_COACH: 17,   // Q - Ref Coach Assigned
};

// ─────────────────────────────────────────
// HTTP GET — fetch current game for a field
// ─────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  const origin = e.parameter.origin || '*';

  let result;
  if (action === 'getGame') {
    const field = parseInt(e.parameter.field);
    const date  = e.parameter.date;   // YYYY-MM-DD
    const time  = e.parameter.time;   // HH:MM
    result = getGame(field, date, time);
  } else {
    result = { error: 'Unknown action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────
// HTTP POST — write submission to sheet
// ─────────────────────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'submit') {
      writeSubmission(body.data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─────────────────────────────────────────
// GET GAME — find active game by field + time
// ─────────────────────────────────────────
function getGame(fieldNum, dateStr, timeStr) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SCHEDULE_SHEET);
  const data  = sheet.getDataRange().getValues();

  // Skip header rows (first 3 rows are headers in your sheet)
  const HEADER_ROWS = 3;

  const inputDate = new Date(dateStr);
  const [inputH, inputM] = timeStr.split(':').map(Number);
  const inputMins = inputH * 60 + inputM;

  // Game window: we consider a game "active" from kickoff to +90 min
  const GAME_WINDOW_MINS = 90;

  for (let i = HEADER_ROWS; i < data.length; i++) {
    const row = data[i];
    if (!row[COL.GAME_NUM - 1]) continue;   // skip empty rows

    const rowField = parseInt(row[COL.FIELD - 1]);
    if (rowField !== fieldNum) continue;

    // Parse date from sheet
    const rowDate = row[COL.DATE - 1];
    if (!rowDate) continue;
    const rd = new Date(rowDate);
    if (rd.toDateString() !== inputDate.toDateString()) continue;

    // Parse time from sheet (handles "09:00", "9:00 AM", or Date objects)
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
      // Fraction of day
      rowMins = Math.round(rowTimeRaw * 24 * 60);
    }

    // Is current time within game window?
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

// ─────────────────────────────────────────
// WRITE SUBMISSION
// ─────────────────────────────────────────
function writeSubmission(data) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet   = ss.getSheetByName(SUBMIT_SHEET);

  // Create the SUBMISSIONS sheet if it doesn't exist yet
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

  // Write one row per official
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
