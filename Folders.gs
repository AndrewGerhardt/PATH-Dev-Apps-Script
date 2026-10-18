function createTroopFoldersForNextRunOnly() {
  const run = findNextPendingOrRunningRun_();
  if (!run) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No Pending/Running runs found.', 'PATH', 6);
    return;
  }

  const count = createTroopFoldersForRunOnly_(run);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Folders ready: ${count} (Run ${run.RunID})`,
    'PATH',
    8
  );
}

function createTroopFoldersForRunOnly_(run) {
  const ss = SpreadsheetApp.openByUrl(CFG.SS_URL);

  const rosterSh = ss.getSheetByName(CFG.SHEETS.ROSTER);
  const { values: rosterVals, idx: rIdx } = readSheet_(rosterSh);

  const scoutMap = buildScoutMap_(ss);

  const rootId = String(CFG.OUTPUT_FOLDER_ID).trim();

  // Validate folder exists (will throw if invalid)
  Drive.Files.get(rootId, { supportsAllDrives: true });

  const ensured = new Set();
  let ensuredCount = 0;

  for (let i = 0; i < rosterVals.length; i++) {
    const row = rosterVals[i];

    if (String(row[rIdx[CFG.COLS.ROSTER_YEAR]] || '').trim() !== String(run.Year)) continue;
    if (String(row[rIdx[CFG.COLS.ROSTER_SESSION]] || '').trim() !== String(run.Session)) continue;

    const scoutId = String(row[rIdx[CFG.COLS.ROSTER_SCOUTID]] || '').trim();
    const scout = scoutMap.get(scoutId);
    if (!scout) continue;

    const unitNumber = String(scout.UnitNumber || '').trim();
    const unitTypeBG = String(scout.UnitType || '').trim();
    const unitState = normalizeState_(scout.UnitState);

    if (!unitNumber || !unitTypeBG) continue;

    const folderName = `${unitNumber}${unitTypeBG} ${unitState}`;
    if (ensured.has(folderName)) continue;

    ensureFolderExists_(rootId, folderName);

    ensured.add(folderName);
    ensuredCount++;
  }

  return ensuredCount;
}

function ensureFolderExists_(parentId, name) {
  const safeName = name.replace(/'/g, "\\'");

  const query =
    `'${parentId}' in parents ` +
    `and title = '${safeName}' ` +  // ← FIXED
    `and mimeType = 'application/vnd.google-apps.folder' ` +
    `and trashed = false`;

  const results = Drive.Files.list({
    q: query,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });

  if (results.items && results.items.length > 0) {
    return results.items[0];
  }

  return Drive.Files.insert({
    title: name,
    mimeType: 'application/vnd.google-apps.folder',
    parents: [{ id: parentId }]
  }, null, {
    supportsAllDrives: true
  });
}

/**
 * 🗺️ Normalize state to 2-letter abbreviation
 */
function normalizeState_(state) {
  if (!state) return '';

  const s = String(state).trim();

  // Already abbreviation
  if (s.length === 2) {
    return s.toUpperCase();
  }

  const map = {
    'ALABAMA': 'AL',
    'ALASKA': 'AK',
    'ARIZONA': 'AZ',
    'ARKANSAS': 'AR',
    'CALIFORNIA': 'CA',
    'COLORADO': 'CO',
    'CONNECTICUT': 'CT',
    'DELAWARE': 'DE',
    'FLORIDA': 'FL',
    'GEORGIA': 'GA',
    'HAWAII': 'HI',
    'IDAHO': 'ID',
    'ILLINOIS': 'IL',
    'INDIANA': 'IN',
    'IOWA': 'IA',
    'KANSAS': 'KS',
    'KENTUCKY': 'KY',
    'LOUISIANA': 'LA',
    'MAINE': 'ME',
    'MARYLAND': 'MD',
    'MASSACHUSETTS': 'MA',
    'MICHIGAN': 'MI',
    'MINNESOTA': 'MN',
    'MISSISSIPPI': 'MS',
    'MISSOURI': 'MO',
    'MONTANA': 'MT',
    'NEBRASKA': 'NE',
    'NEVADA': 'NV',
    'NEW HAMPSHIRE': 'NH',
    'NEW JERSEY': 'NJ',
    'NEW MEXICO': 'NM',
    'NEW YORK': 'NY',
    'NORTH CAROLINA': 'NC',
    'NORTH DAKOTA': 'ND',
    'OHIO': 'OH',
    'OKLAHOMA': 'OK',
    'OREGON': 'OR',
    'PENNSYLVANIA': 'PA',
    'RHODE ISLAND': 'RI',
    'SOUTH CAROLINA': 'SC',
    'SOUTH DAKOTA': 'SD',
    'TENNESSEE': 'TN',
    'TEXAS': 'TX',
    'UTAH': 'UT',
    'VERMONT': 'VT',
    'VIRGINIA': 'VA',
    'WASHINGTON': 'WA',
    'WEST VIRGINIA': 'WV',
    'WISCONSIN': 'WI',
    'WYOMING': 'WY'
  };

  const key = s.toUpperCase();

  return map[key] || key.substring(0, 2).toUpperCase(); // fallback
}
