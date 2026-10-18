function processRunChunk_(runID) {
  console.warn("Sequential processor disabled — using parallel only.");
  return;
}

function processOneRosterRow_(
  run,
  rosterSh,
  headers,
  rIdx,
  offeringMap,
  scoutMap,
  badgeMap,
  rosterRowNumber,
  row
) {
  if (!rosterRowNumber) {
    console.error("Invalid rowNumber passed to processOneRosterRow_:", rosterRowNumber);
    return true; // skip safely
  }

  const rosterObj = rowToObject_(headers, row);

  // Ensure still matches run
  if (String(rosterObj[CFG.COLS.ROSTER_YEAR]).trim() !== String(run.Year)) return true;
  if (String(rosterObj[CFG.COLS.ROSTER_SESSION]).trim() !== String(run.Session)) return true;

  const offeringId = String(rosterObj[CFG.COLS.ROSTER_OFFERINGID] || '').trim();

  const offering = offeringMap.get(offeringId);

  if (!offering) {
    return true; // skip safely
  }

  const badgeId1 = String(offering.BadgeID_1 || '').trim();
  const badgeId2 = String(offering.BadgeID_2 || '').trim();
  const isCombo = !!badgeId2;

  const pr1 = rosterObj[CFG.COLS.PR_FILE_1];
  const pr2 = rosterObj[CFG.COLS.PR_FILE_2];

  const mergedCommon = buildMergedRowData_(
    rosterObj,
    scoutMap,
    badgeMap,
    offeringMap
  );

  // --- Badge 1 ---
  if (badgeId1) {

    if (!rosterObj[CFG.COLS.PR_FILE_1]) {
      try {
        const merged1 = {
          ...mergedCommon,
          targetBadgeId: badgeId1,
          targetBadgeSlot: 1
        };

        const pdf1 = retryDriveOperation_(() =>
          generatePdfForMergedRow_(run, merged1)
        );

        rosterSh
          .getRange(rosterRowNumber, rIdx[CFG.COLS.PR_FILE_1] + 1)
          .setValue(pdf1.url);

      } catch (err) {
        console.error(`Row ${rosterRowNumber} badge1 FAILED after retries`, err);
        return false; // retry later
      }
    }
  }

  // --- Badge 2 ---
  if (isCombo) {

    if (!rosterObj[CFG.COLS.PR_FILE_2]) {
      try {
        const merged2 = {
          ...mergedCommon,
          targetBadgeId: badgeId2,
          targetBadgeSlot: 2
        };

        const pdf2 = retryDriveOperation_(() =>
          generatePdfForMergedRow_(run, merged2)
        );

        rosterSh
          .getRange(rosterRowNumber, rIdx[CFG.COLS.PR_FILE_2] + 1)
          .setValue(pdf2.url);

      } catch (err) {
        console.error(`Row ${rosterRowNumber} badge2 FAILED after retries`, err);
        return false; // retry later
      }
    }
  }

  return true; // success
}

function buildMergedRowData_(rosterObj, scoutMap, badgeMap, offeringMap) {

  const scoutId = rosterObj[CFG.COLS.ROSTER_SCOUTID];
  const offeringId = rosterObj[CFG.COLS.ROSTER_OFFERINGID];

  const scout = scoutMap.get(String(scoutId).trim());
  const offering = offeringMap.get(String(offeringId).trim());

  const badge1 = offering
    ? badgeMap.get(String(offering.BadgeID_1 || '').trim())
    : null;

  const badge2 = offering
    ? badgeMap.get(String(offering.BadgeID_2 || '').trim())
    : null;

  return {
    runYear: rosterObj[CFG.COLS.ROSTER_YEAR],
    runSession: rosterObj[CFG.COLS.ROSTER_SESSION],

    roster: rosterObj,
    scout: scout,
    offering: offering,
    badge1: badge1,
    badge2: badge2,

    ScoutName: scout ? `${scout.FirstName} ${scout.LastName}` : '',
    Unit: scout ? `${scout.UnitType} ${scout.UnitNumber}` : '',
    OfferingID: offeringId || '',
    BadgeName_1: badge1 ? badge1.BadgeName : '',
    BadgeName_2: badge2 ? badge2.BadgeName : '',
  };
}

/** Parallel Runs */
function processRunChunkParallel_(workerId) {

  const ss = getSS_();

  const run = findNextPendingOrRunningRun_();
  if (!run) return;

  const runID = run.RunID;

  const rosterSh = openSheet_(CFG.SHEETS.ROSTER);
  const { headers, idx: rIdx } = readHeaders_(rosterSh);

  const data = rosterSh.getDataRange().getValues();

  const lastRow = rosterSh.getLastRow();

  const offeringMap = buildOfferingMap_();
  const scoutMap = buildScoutMap_();
  const badgeMap = buildBadgeMap_();

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  let state;
  try {
    state = loadRunState_(runID) || {};

    if (state.nextIndex === undefined || state.nextIndex === null) {
      state.nextIndex = 0;
    }

  } finally {
    lock.releaseLock();
  }

  const t0 = Date.now();

  while (true) {

    if (Date.now() - t0 > CFG.TIME_BUDGET_MS) break;

    let rowNumber;

    // --- LOCK: safely grab next row ---
    lock.waitLock(5000);
    try {
      const latestState = loadRunState_(runID) || {};

      const nextIndex = Number(latestState.nextIndex ?? 0);

      rowNumber = nextIndex + CFG.HEADER_ROW + 1;

      latestState.nextIndex = nextIndex + 1;

      saveRunState_(runID, latestState);

    } finally {
      lock.releaseLock();
    }

    // Stop if past sheet
    if (rowNumber > lastRow) break;

    let success = true;

    try {
      const row = data[rowNumber - 1]; // adjust for 1-based sheet

      success = processOneRosterRow_(
        run,
        rosterSh,
        headers,
        rIdx,
        offeringMap,
        scoutMap,
        badgeMap,
        rowNumber,
        row
      );
    } catch (err) {
      success = false;
      console.error(`Worker ${workerId} failed at row ${rowNumber}`, err);
    }

    // --- IMPORTANT: do NOT retry forever ---
    if (!success) {
      console.warn(`Skipping failed row ${rowNumber} after error`);
    }
  }

  const finalState = loadRunState_(runID) || {};
  const nextIndex = Number(finalState.nextIndex ?? 0);

  // total data rows (excluding header)
  const totalRows = lastRow - CFG.HEADER_ROW;

  if (nextIndex >= totalRows) {
    console.log("All rows processed. Stopping run.");

    clearRunState_(runID);
    markRunComplete_(runID);
    clearParallelWorkerTriggers_();

  } else {
    scheduleParallelWorker_(workerId, 10000);
  }

}

function retryDriveOperation_(fn, maxAttempts = 5) {
  let attempt = 0;

  while (attempt < maxAttempts) {
    try {
      return fn();
    } catch (err) {
      const msg = String(err.message || '');

      const isRateLimit =
        msg.includes('User rate limit exceeded') ||
        msg.includes('Rate Limit Exceeded') ||
        msg.includes('Service invoked too many times');

      if (!isRateLimit) {
        throw err; // not retryable
      }

      attempt++;

      const waitMs = Math.pow(2, attempt) * 500; // exponential backoff
      console.warn(`Rate limit hit. Retry ${attempt}/${maxAttempts} in ${waitMs}ms`);

      Utilities.sleep(waitMs);
    }
  }

  throw new Error('Max retries exceeded for Drive operation');
}