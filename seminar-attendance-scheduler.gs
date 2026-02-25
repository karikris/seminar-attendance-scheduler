/**
 * Minimal single-form scheduler
 * Opens/publishes form during any active Start_AEST..End_AEST window
 * ONLY for rows where Title is non-empty
 * Closes/unpublishes form outside those windows
 *
 * Bind this script to the production spreadsheet.
 */

const CONFIG = Object.freeze({
  SPREADSHEET_ID: 'insert spreadsheet id found between d/ and /edit',
  SHEET_NAME: 'insert the name of the sheet',
  FORM_ID: 'insert form id found between d/ and /edit',
  TZ: 'Australia/Melbourne',
  TITLE_HEADER: 'insert a column header with a non-empty value as condition',
  START_HEADER: 'Start_AEST',
  END_HEADER: 'End_AEST',

  CLOSED_MESSAGE:
    'This attendance form is only open during the scheduled seminar time. Please try again during the seminar session.',
});

/**
 * Run once to install/refresh the recurring trigger.
 */
function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || ss.getId() !== CONFIG.SPREADSHEET_ID) {
    throw new Error('Run this from the Apps Script project bound to the production spreadsheet.');
  }

  const handler = 'syncFormVisibility';
  const triggers = ScriptApp.getProjectTriggers();

  for (const t of triggers) {
    if (typeof t.getHandlerFunction === 'function' && t.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(t);
    }
  }

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(5)
    .create();

  syncFormVisibility();
}

/**
 * Time-driven trigger target.
 */
function syncFormVisibility() {
  const form = FormApp.openById(CONFIG.FORM_ID);

  if (isWithinAnyEligibleWindow_()) {
    safeSetPublished_(form, true);
    safeCall_(() => form.setAcceptingResponses(true));
  } else {
    safeCall_(() => form.setAcceptingResponses(false));
    safeCall_(() => form.setCustomClosedFormMessage(CONFIG.CLOSED_MESSAGE));
    safeSetPublished_(form, false);
  }
}

/* -------------------- Core logic -------------------- */

function isWithinAnyEligibleWindow_() {
  const sh = getSheetOrThrow_();
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return false;

  const header = values[0];
  const titleIdx = findHeaderIndexStrict_(header, CONFIG.TITLE_HEADER);
  const startIdx = findHeaderIndexStrict_(header, CONFIG.START_HEADER);
  const endIdx = findHeaderIndexStrict_(header, CONFIG.END_HEADER);

  const now = new Date();

  for (let r = 1; r < values.length; r++) {
    const title = values[r][titleIdx];
    if (!hasText_(title)) continue; // hard gate

    const start = asDateOrNull_(values[r][startIdx]);
    const end = asDateOrNull_(values[r][endIdx]);

    if (!start || !end) continue;
    if (end < start) continue;

    if (now >= start && now <= end) return true;
  }

  return false;
}

/* -------------------- Helpers -------------------- */

function getSheetOrThrow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No active spreadsheet.');
  if (ss.getId() !== CONFIG.SPREADSHEET_ID) throw new Error('Wrong spreadsheet bound to script.');

  const sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) throw new Error('Missing worksheet: ' + CONFIG.SHEET_NAME);
  return sh;
}

function findHeaderIndexStrict_(headerRow, wantedHeader) {
  let match = -1;
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || '').trim() === wantedHeader) {
      if (match !== -1) {
        throw new Error('Duplicate header "' + wantedHeader + '" found. Keep only one.');
      }
      match = i;
    }
  }
  if (match === -1) throw new Error('Missing required header: ' + wantedHeader);
  return match;
}

function hasText_(v) {
  return String(v == null ? '' : v).trim() !== '';
}

function asDateOrNull_(v) {
  if (v === null || v === undefined || v === '') return null;

  // Native Date object from Sheets datetime cells
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return v;
  }

  // Fallback parser for text like "23.02.2026 11:00"
  const s = String(v).trim();
  if (!s) return null;

  const m = s.match(/^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
  if (m) {
    const day = Number(m[1]);
    const month = Number(m[2]);
    const year = Number(m[3]);
    const hour = m[4] !== undefined ? Number(m[4]) : 0;
    const minute = m[5] !== undefined ? Number(m[5]) : 0;

    const d = new Date(year, month - 1, day, hour, minute, 0, 0);
    if (
      d.getFullYear() === year &&
      d.getMonth() === month - 1 &&
      d.getDate() === day &&
      d.getHours() === hour &&
      d.getMinutes() === minute
    ) {
      return d;
    }
    return null;
  }

  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function safeSetPublished_(form, shouldPublish) {
  try {
    if (typeof form.supportsAdvancedResponderPermissions === 'function' &&
        form.supportsAdvancedResponderPermissions() &&
        typeof form.setPublished === 'function') {
      form.setPublished(shouldPublish);
      return;
    }
    if (typeof form.setPublished === 'function') {
      form.setPublished(shouldPublish);
    }
  } catch (e) {
    // silent by design (minimal script)
  }
}

function safeCall_(fn) {
  try {
    return fn();
  } catch (e) {
    return null;
  }
}
