function buildQueueForRun_(run) {

  const rosterSh = openSheet_(CFG.SHEETS.ROSTER);
  const { values: rosterVals, idx: rIdx } = readSheet_(rosterSh);

  const queue = [];

  for (let i = 0; i < rosterVals.length; i++) {
    const row = rosterVals[i];

    const year = String(row[rIdx[CFG.COLS.ROSTER_YEAR]] || '').trim();
    const session = String(row[rIdx[CFG.COLS.ROSTER_SESSION]] || '').trim();

    if (year !== String(run.Year) || session !== String(run.Session)) continue;

    const offeringId = String(row[rIdx[CFG.COLS.ROSTER_OFFERINGID]] || '')
      .trim();

    const pr1 = row[rIdx[CFG.COLS.PR_FILE_1]];
    const pr2 = row[rIdx[CFG.COLS.PR_FILE_2]];

    if (!offeringId) continue;
    if (pr1 && pr2) continue;

    queue.push({
      rowNumber: i + CFG.HEADER_ROW + 1
    });
  }

  return queue;
}

function buildOfferingMap_() {
  const sh = openSheet_(CFG.SHEETS.OFFERINGS);
  const { values, idx } = readSheet_(sh);

  const map = new Map();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const offeringId = String(row[idx[CFG.COLS.OFFERINGID]] || '')
      .replace(/\u00A0/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    if (!offeringId) continue;

    map.set(offeringId, {
      BadgeID_1: String(row[idx[CFG.COLS.BADGEID_1]] || '').trim(),
      BadgeID_2: String(row[idx[CFG.COLS.BADGEID_2]] || '').trim(),
      Period: row[idx['Period']],
    });
  }

  return map;
}
