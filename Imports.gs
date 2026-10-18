/***********************
 * Imports.gs (PATH)
 * - Imports Scouts (dedup by ScoutID)
 * - Imports Roster rows (OfferingID resolved via OfferingAliases)
 ***********************/

/** =========================
 *  PUBLIC: Import Scouts
 *  ========================= */
function importScoutsFromRosterExport() {
  const ui = SpreadsheetApp.getUi();
  const year = promptNumber_(ui, 'Import Scouts', 'Enter Year (e.g., 2026):');
  const session = promptNumber_(ui, 'Import Scouts', 'Enter Session number (1–6):');

  const st = openStaging_();
  const raw = readTable_(st, CFG.STAGING_SHEETS.ROSTER_RAW);

  const required = [
    'Last Name', 'First Name', 'Unit Type', 'Unit Number', 'Unit Qualifier', 'City', 'State'
  ];
  assertHeaders_(raw.headers, required);

  const main = SpreadsheetApp.openByUrl(CFG.SS_URL);
  const scoutsSh = main.getSheetByName(CFG.SHEETS.SCOUTS);

  const scoutsTable = readSheetTable_(scoutsSh);
  const outHeaders = scoutsTable.headers;
  const col = headerIndex_(outHeaders);

  // Map: ScoutID -> sheet row number (1-based)
  const scoutRowById = new Map();
  for (let i = 0; i < scoutsTable.rows.length; i++) {
    const sid = String(scoutsTable.rows[i][col['ScoutID']] || '').trim();
    if (sid) scoutRowById.set(sid, i + 2); // +2 because headers are row 1
  }

  let added = 0;
  let updatedSession = 0;

  const toAppend = [];
  const pendingAdds = new Set(); // prevent duplicate appends within same run

  raw.rows.forEach(r => {
    const last = cleanName_(r['Last Name']);
    const first = cleanName_(r['First Name']);
    if (!last || !first) return;

    const unitNumber = String(r['Unit Number'] || '').trim();
    const unitTypeRaw = String(r['Unit Type'] || '').trim();
    const unitQualifier = String(r['Unit Qualifier'] || '').trim();

    const unitSuffix = unitSuffix_(unitTypeRaw, unitQualifier); // B/G/P/V/S/etc
    const scoutId = makeScoutId_(last, first, unitNumber, unitSuffix);
    if (!scoutId) return;

    // If already exists, update Session if needed
    const existingRow = scoutRowById.get(scoutId);
    if (existingRow) {
      if (col['Session'] != null) {
        const current = String(
          scoutsSh.getRange(existingRow, col['Session'] + 1).getValue() || ''
        ).trim();

        if (String(session) !== current) {
          scoutsSh.getRange(existingRow, col['Session'] + 1).setValue(session);
          updatedSession++;
        }
      }
      return;
    }

    // Avoid duplicates within the same import run
    if (pendingAdds.has(scoutId)) return;
    pendingAdds.add(scoutId);

    const unitCity = String(r['City'] || '').trim();
    const unitState = String(r['State'] || '').trim();

    const newRow = new Array(outHeaders.length).fill('');
    if (col['ScoutID'] != null) newRow[col['ScoutID']] = scoutId;
    if (col['Session'] != null) newRow[col['Session']] = session;
    if (col['FirstName'] != null) newRow[col['FirstName']] = first;
    if (col['LastName'] != null) newRow[col['LastName']] = last;
    if (col['UnitNumber'] != null) newRow[col['UnitNumber']] = unitNumber;

    // Store B/G/etc in UnitType (per your preference)
    if (col['UnitType'] != null) newRow[col['UnitType']] = unitSuffix;

    if (col['UnitCity'] != null) newRow[col['UnitCity']] = unitCity;
    if (col['UnitState'] != null) newRow[col['UnitState']] = unitState;

    toAppend.push(newRow);
    added++;
  });

  if (toAppend.length) {
    scoutsSh
      .getRange(scoutsSh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length)
      .setValues(toAppend);
  }

  log_(st, `Import Scouts done. Added ${added}, updated Session for ${updatedSession}. (Year ${year}, Session ${session})`);
  ui.alert(
    `Import Scouts complete.\n\nAdded: ${added}\nUpdated Session: ${updatedSession}\nTotal rows scanned: ${raw.rows.length}`
  );
}

/** =========================
 *  PUBLIC: Import Roster
 *  ========================= */
function importRosterFromRosterExport() {
  const ui = SpreadsheetApp.getUi();
  const year = promptNumber_(ui, 'Import Roster', 'Enter Year (e.g., 2026):');
  const session = promptNumber_(ui, 'Import Roster', 'Enter Session number (1–6):');

  const st = openStaging_();
  const badgeAliasMap = loadBadgeAliases_(st);
  const raw = readTable_(st, CFG.STAGING_SHEETS.ROSTER_RAW);

  const required = [
    'Class Name','Period Start Time',
    'Last Name','First Name','Unit Type','Unit Number','Unit Qualifier'
  ];
  assertHeaders_(raw.headers, required);

  const offeringMap = loadOfferingAliases_(st); // key = period|classname -> offeringId

  const main = SpreadsheetApp.openByUrl(CFG.SS_URL);
  const rosterSh = main.getSheetByName(CFG.SHEETS.ROSTER);
  const rosterTable = readSheetTable_(rosterSh);
  const rosterHeaders = rosterTable.headers;
  const rcol = headerIndex_(rosterHeaders);

  // Build existing key set to prevent duplicates
  const existingKeys = new Set();
  rosterTable.rows.forEach(row => {
    const y = String(row[ rcol[CFG.COLS.ROSTER_YEAR] ] || '').trim();
    const s = String(row[ rcol[CFG.COLS.ROSTER_SESSION] ] || '').trim();
    const off = String(row[ rcol[CFG.COLS.ROSTER_OFFERINGID] ] || '').trim();
    const sid = String(row[ rcol[CFG.COLS.ROSTER_SCOUTID] ] || '').trim();
    if (y && s && off && sid) existingKeys.add(`${y}|${s}|${off}|${sid}`);
  });

  let added = 0;
  let missingOffering = 0;

  const toAppend = [];

  raw.rows.forEach(r => {
    const last = cleanName_(r['Last Name']);
    const first = cleanName_(r['First Name']);
    if (!last || !first) return;

    const unitNumber = String(r['Unit Number'] || '').trim();
    const unitType = String(r['Unit Type'] || '').trim();
    const unitQualifier = String(r['Unit Qualifier'] || '').trim();
    const unitSuffix = unitSuffix_(unitType, unitQualifier);
    const scoutId = makeScoutId_(last, first, unitNumber, unitSuffix);

    const className = normalizeClassName_(String(r['Class Name'] || ''));
    const period = normalizeTime_(r['Period Start Time']); // from 09:00:00 -> 9:00

    const offeringKey = `${period}|${className}`;
    let offeringId = offeringMap.get(offeringKey);

    if (!offeringId) {
      offeringId = buildOfferingId_(className, period, badgeAliasMap);

      if (offeringId) {
        appendNewOfferingAlias_(st, period, className, offeringId);
      } else {
        appendMissingOfferingAlias_(st, period, className);
        missingOffering++;
        return;
      }
    }

    function appendNewOfferingAlias_(stSS, period, className, offeringId) {
      const sh = stSS.getSheetByName(CFG.STAGING_SHEETS.OFFERING_ALIASES);
      if (!sh) return;

      const values = sh.getDataRange().getValues();
      if (values.length < 2) {
        sh.appendRow([period, className, offeringId]);
        return;
      }

      const headers = values[0];
      const idx = headerIndex_(headers);

      const pNorm = normalizeTime_(period);
      const cNorm = normalizeClassName_(className);

      for (let i = 1; i < values.length; i++) {
        const rowPeriod = normalizeTime_(values[i][idx['Period']]);
        const rowClass = normalizeClassName_(values[i][idx['ExportClassName']]);
        const rowId = String(values[i][idx['OfferingID']] || '').trim();

        if (rowPeriod === pNorm && rowClass === cNorm) {
          // Row exists
          if (!rowId) {
            // Fill missing OfferingID
            sh.getRange(i + 1, idx['OfferingID'] + 1).setValue(offeringId);
          }
          return;
        }
      }

      // If no match found, append new row
      sh.appendRow([period, className, offeringId]);
    }

    const key = `${year}|${session}|${offeringId}|${scoutId}`;
    if (existingKeys.has(key)) return;

    const newRow = new Array(rosterHeaders.length).fill('');

    // Required roster fields
    newRow[rcol[CFG.COLS.ROSTERID]] = String(Date.now()) + '_' + Math.random().toString(36).slice(2, 8);
    newRow[rcol[CFG.COLS.ROSTER_YEAR]] = year;
    newRow[rcol[CFG.COLS.ROSTER_SESSION]] = session;
    newRow[rcol[CFG.COLS.ROSTER_OFFERINGID]] = offeringId;
    newRow[rcol[CFG.COLS.ROSTER_SCOUTID]] = scoutId;

    // Attendance + req lists left blank; earned blank
    // (Your sheet already has these columns: M T W TH F ReqComplete ReqOptional ReqIncomplete Comments Earned_1 Earned_2 etc.)

    toAppend.push(newRow);
    existingKeys.add(key);
    added++;
  });

  if (toAppend.length) {
    rosterSh.getRange(rosterSh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
  }

  log_(st, `Import Roster done. Added ${added} roster rows. Missing offering aliases: ${missingOffering}. (Year ${year}, Session ${session})`);
  ui.alert(`Import Roster complete.\n\nAdded roster rows: ${added}\nMissing OfferingAliases rows created: ${missingOffering}`);
}

/** =========================
 *  OfferingAliases helpers
 *  ========================= */

function loadOfferingAliases_(stSS) {
  const sh = stSS.getSheetByName(CFG.STAGING_SHEETS.OFFERING_ALIASES);
  if (!sh) throw new Error(`Missing staging sheet: ${CFG.STAGING_SHEETS.OFFERING_ALIASES}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return new Map();

  const headers = values[0].map(String);
  const idx = headerIndex_(headers);

  // Expect: Period | ExportClassName | OfferingID
  const map = new Map();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const period = normalizeTime_(row[idx['Period']]);
    const className = normalizeClassName_(row[idx['ExportClassName']]);
    const offeringId = String(row[idx['OfferingID']] || '').trim();
    if (!period || !className || !offeringId) continue;
    map.set(`${period}|${className}`, offeringId);
  }
  return map;
}

function getPeriodCode_(timeStr) {
  const map = {
    '9:00': 'P1',
    '10:00': 'P2',
    '11:00': 'P3',
    '14:30': 'P4',
    '2:30': 'P4',
    '15:30': 'P5',
    '3:30': 'P5',
    '16:30': 'P6',
    '4:30': 'P6'
  };

  return map[timeStr] || '';
}

function appendMissingOfferingAlias_(stSS, period, className) {
  const sh = stSS.getSheetByName(CFG.STAGING_SHEETS.OFFERING_ALIASES);
  if (!sh) return;

  // avoid duplicates
  const existing = sh.getDataRange().getValues().map(r => `${normalizeTime_(r[0])}|${normalizeClassName_(r[1])}`);
  const key = `${period}|${className}`;
  if (existing.includes(key)) return;

  sh.appendRow([period, className, '']);
}

function buildOfferingId_(className, period, badgeAliasMap) {
  const normalized = normalizeClassName_(className).toLowerCase();

  const periodCode = getPeriodCode_(period);
  if (!periodCode) return '';

  const badgeIds = [];

  // Try exact match first
  const exact = badgeAliasMap.get(normalized);
  if (exact) {
    badgeIds.push(exact);
  } else {
    // Try to detect combos (split on common separators)
    const parts = normalized.split(/\/|,| or | and /);

    parts.forEach(p => {
      const trimmed = p.trim();
      const id = badgeAliasMap.get(trimmed);
      if (id && !badgeIds.includes(id)) {
        badgeIds.push(id);
      }
    });
  }

  if (!badgeIds.length) return '';

  // Sort to keep consistency (AQ02+AQ05 not AQ05+AQ02)
  badgeIds.sort();

  const combo = badgeIds.join('+');

  return `${combo} | ${periodCode}`;
}

function loadBadgeAliases_(stSS) {
  const sh = stSS.getSheetByName('BadgeAliases');
  if (!sh) return new Map();

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return new Map();

  const headers = values[0];
  const idx = headerIndex_(headers);

  const map = new Map();

  for (let i = 1; i < values.length; i++) {
    const name = normalizeClassName_(values[i][idx['ExportBadgeName']]).toLowerCase();
    const badgeId = String(values[i][idx['BadgeID']] || '').trim();

    if (name && badgeId) {
      map.set(name, badgeId);
    }
  }

  return map;
}

/** =========================
 *  Staging read helpers
 *  ========================= */

function openStaging_() {
  if (!CFG.STAGING_SS_URL) throw new Error('CFG.STAGING_SS_URL is missing.');
  return SpreadsheetApp.openByUrl(CFG.STAGING_SS_URL);
}

function readTable_(stSS, sheetName) {
  const sh = stSS.getSheetByName(sheetName);
  if (!sh) throw new Error(`Missing staging sheet: ${sheetName}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [] };

  const headers = values[0].map(h => String(h).trim());
  const idx = headerIndex_(headers);

  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const rowArr = values[i];
    if (rowArr.join('').trim() === '') continue;
    const obj = {};
    headers.forEach((h, c) => (obj[h] = rowArr[c]));
    rows.push(obj);
  }

  return { headers, idx, rows };
}

/** =========================
 *  Main sheet helpers
 *  ========================= */

function readSheetTable_(sh) {
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);
  return { headers, rows };
}

function indexBy_(rows, headers, headerName) {
  const idx = headerIndex_(headers);
  const c = idx[headerName];
  const set = new Set();
  rows.forEach(r => {
    const v = String(r[c] || '').trim();
    if (v) set.add(v);
  });
  // Return a Set-like object with .has and .add
  return set;
}

/** =========================
 *  Normalization
 *  ========================= */

function normalizeClassName_(s) {
  return String(s || '')
    .replace(/\s+/g, ' ')
    .replace(/\s*,\s*/g, ', ') // normalize commas
    .trim();
}

function normalizeTime_(v) {
  // Accepts Date, "09:00:00", "9:00", etc. returns "9:00" or "10:00" etc.
  if (v instanceof Date) {
    const h = v.getHours();
    const m = v.getMinutes();
    return `${h}:${String(m).padStart(2, '0')}`;
  }
  const s = String(v || '').trim();
  if (!s) return '';
  // "16:30:00" -> "16:30"
  const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!m) return s;
  return `${Number(m[1])}:${m[2]}`;
}

function cleanName_(s) {
  return String(s || '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/,+$/g, '');
}

function unitSuffix_(unitType, unitQualifier) {
  const ut = String(unitType || '').toLowerCase().trim();
  const uq = String(unitQualifier || '').toLowerCase().trim();

  // Most reliable: Unit Qualifier
  if (uq.includes('girl')) return 'G';
  if (uq.includes('boy')) return 'B';
  if (uq.includes('family')) return 'F';

  // Fallbacks by unit type
  if (ut.includes('troop')) return 'B'; // default troop to Boy if qualifier missing
  if (ut.includes('pack')) return 'P';
  if (ut.includes('crew')) return 'V';
  if (ut.includes('ship')) return 'S';

  return ut ? ut[0].toUpperCase() : '';
}


function makeScoutId_(last, first, unitNumber, unitSuffix) {
  const L = scrubIdPart_(last);
  const F = scrubIdPart_(first);
  const U = scrubIdPart_(unitNumber);
  const S = scrubIdPart_(unitSuffix);
  if (!L || !F || !U) return '';
  return `${L}_${F}_${U}${S}`;
}

function scrubIdPart_(s) {
  return String(s || '')
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

/** =========================
 *  Validation / prompts / logging
 *  ========================= */

function assertHeaders_(headers, required) {
  const set = new Set(headers);
  const missing = required.filter(h => !set.has(h));
  if (missing.length) throw new Error(`Missing required columns in export: ${missing.join(', ')}`);
}

function headerIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => (idx[String(h).trim()] = i));
  return idx;
}

function promptNumber_(ui, title, message) {
  const res = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) throw new Error('Canceled.');
  const n = Number(String(res.getResponseText()).trim());
  if (!Number.isFinite(n) || n <= 0) throw new Error('Invalid number.');
  return n;
}

function log_(stSS, message) {
  try {
    const sh = stSS.getSheetByName(CFG.STAGING_SHEETS.IMPORT_LOG);
    if (!sh) return;
    sh.appendRow([new Date(), 'INFO', message]);
  } catch (e) {
    // ignore logging failures
  }
}
