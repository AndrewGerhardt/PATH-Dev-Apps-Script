function openSheet_(name) {
  const ss = getSS_();

  if (!ss) {
    throw new Error("Spreadsheet is undefined (check CFG.SS_URL)");
  }

  const sh = ss.getSheetByName(name);

  if (!sh) {
    throw new Error("Sheet not found: " + name);
  }

  return sh;
}

function readHeaders_(sheet) {
  const header = sheet.getRange(CFG.HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = header.map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => (idx[h] = i));
  return { headers, idx };
}

function readSheet_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= CFG.HEADER_ROW) return { values: [], idx: {}, headers: [] };

  const headers = data[CFG.HEADER_ROW - 1].map(h => String(h).trim());
  const idx = {};
  headers.forEach((h, i) => (idx[h] = i));

  const values = data.slice(CFG.HEADER_ROW);
  return { values, idx, headers };
}

function rowToObject_(headers, row) {
  const obj = {};
  for (let i = 0; i < headers.length; i++) obj[headers[i]] = row[i];
  return obj;
}

function toBool_(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v || '').trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1';
}

function getRunById_(runID) {
  const sh = openSheet_(CFG.SHEETS.EXPORT_RUNS);
  const { values, idx } = readSheet_(sh);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    if (String(row[idx[CFG.COLS.RUNID]] || '').trim() === String(runID).trim()) {
      return exportRun_(row, idx, i + CFG.HEADER_ROW + 1);
    }
  }

  return null;
}

function formatReqList_(reqStr, badgeName) {
  if (!reqStr) return '';

  const badge = String(badgeName || '').trim();
  if (!badge) return '';

  const items = String(reqStr)
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);

  // Accept multiple dash types and spacing:
  // "Canoeing – 1a" / "Canoeing-1a" / "Canoeing — 1a"
  const dashClass = "[\\u2012\\u2013\\u2014\\u2015\\-]"; // figure/en/em/horizontal/hyphen
  const re = new RegExp("^" + escapeRegex_(badge) + "\\s*" + dashClass + "\\s*(.+)$");

  const out = [];
  for (const it of items) {
    const m = it.match(re);
    if (m && m[1]) out.push(m[1].trim()); // keep ONLY this badge's req codes
  }

  return out.join(', ');
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function scrubBadgePrefix_(text, badgeName) {
  const t = String(text || '').trim();
  const b = String(badgeName || '').trim();
  if (!b) return t;

  const esc = b.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // Handles "Badge - 1a", "Badge – 1a", "Badge — 1a"
  const re = new RegExp('^' + esc + '\\s*[-–—]\\s*', 'i');
  return t.replace(re, '').trim();
}

function buildScoutMap_() {
  const sh = openSheet_(CFG.SHEETS.SCOUTS);
  const { values, idx } = readSheet_(sh);

  const map = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const scoutId = String(row[idx[CFG.COLS.SCOUTID]] || '').trim();
    if (!scoutId) continue;

    map.set(scoutId, {
      ScoutID: scoutId,
      FirstName: row[idx['FirstName']],
      LastName: row[idx['LastName']],
      UnitNumber: row[idx['UnitNumber']],
      UnitType: row[idx['UnitType']],
      UnitState: row[idx['UnitState']],
    });
  }

  return map;
}

function buildBadgeMap_() {
  const sh = openSheet_(CFG.SHEETS.BADGES);
  const { values, idx } = readSheet_(sh);

  const map = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const badgeId = String(row[idx[CFG.COLS.BADGEID]] || '').trim();
    if (!badgeId) continue;

    map.set(badgeId, {
      BadgeID: badgeId,
      BadgeName: row[idx['BadgeName']]
    });
  }

  return map;
}

function getLatestYearFromRoster_() {
  const sh = openSheet_(CFG.SHEETS.ROSTER);
  const { values, idx } = readSheet_(sh);

  const yIdx = idx[CFG.COLS.ROSTER_YEAR];
  if (yIdx === undefined) throw new Error('Roster missing Year column');

  let maxYear = null;

  for (const row of values) {
    const y = Number(String(row[yIdx] || '').trim());
    if (!Number.isFinite(y)) continue;
    if (maxYear === null || y > maxYear) maxYear = y;
  }

  return maxYear;
}

function promptYearOptionA_(ui) {
  const latest = getLatestYearFromRoster_();
  if (latest !== null) {
    const useLatest = ui.alert(
      'Year',
      `Use latest year found in Roster (${latest})?`,
      ui.ButtonSet.YES_NO
    );
    if (useLatest === ui.Button.YES) return latest;
  }

  // Fallback or user chose NO
  return promptNumber_(ui, 'Year', 'Enter year (ex: 2026):');
}

function getSS_() {
  return SpreadsheetApp.openById(CFG.SS_ID);
}
