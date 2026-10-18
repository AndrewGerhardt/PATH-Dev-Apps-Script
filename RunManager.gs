function createExportRun_(year, session) {
  const sh = openSheet_(CFG.SHEETS.EXPORT_RUNS);

  const runId = String(Date.now());

  sh.appendRow([runId, year, session, 'Pending', '']);

  return runId;
}

function startOrContinueRunById_(runId) {
  const run = getRunById_(runId);
  if (!run) return;

  // Ensure status + StartedAt are set the same way as "next run" flow
  if (String(run.Status || '').trim() !== 'Running') {
    markRunRunning_(runId); // also fills StartedAt if blank
  }

  processRunChunk_(runId);
}

function getLatestRun_(ss) {
  const sh = ss.getSheetByName(CFG.SHEETS.EXPORT_RUNS);
  const { values, headers } = readSheet_(sh);

  if (!values.length) return null;

  const last = values[values.length - 1];
  return rowToObject_(headers, last);
}

function startOrContinueNextRun() {
  const run = findNextPendingOrRunningRun_();
  if (!run) return;

  if (String(run.Status || '').trim() !== 'Running') {
    markRunRunning_(run.RunID);
  }

  processRunChunkParallel_(0);
}

function findNextPendingOrRunningRun_() {
  const sh = openSheet_(CFG.SHEETS.EXPORT_RUNS);
  const { values, idx } = readSheet_(sh);

  let pending = null;

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const status = String(row[idx[CFG.COLS.RUN_STATUS]] || '').trim();

    if (status === 'Running') {
      return exportRun_(row, idx, r + CFG.HEADER_ROW + 1);
    }
    if (!pending && status === 'Pending') {
      pending = exportRun_(row, idx, r + CFG.HEADER_ROW + 1);
    }
  }

  return pending;
}

function exportRun_(row, idx, rowNumber) {
  return {
    _rowNumber: rowNumber,
    RunID: String(row[idx[CFG.COLS.RUNID]] || '').trim(),
    Year: String(row[idx[CFG.COLS.RUN_YEAR]] || '').trim(),
    Session: String(row[idx[CFG.COLS.RUN_SESSION]] || '').trim(),
    Status: String(row[idx[CFG.COLS.RUN_STATUS]] || '').trim(),
    StartedAt: row[idx[CFG.COLS.RUN_STARTED_AT]],
  };
}

function markRunRunning_(runID) {
  const sh = openSheet_(CFG.SHEETS.EXPORT_RUNS);
  const { values, idx } = readSheet_(sh);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[idx[CFG.COLS.RUNID]] || '').trim() === String(runID).trim()) {
      const rowNum = i + CFG.HEADER_ROW + 1;

      // Status -> Running
      sh.getRange(rowNum, idx[CFG.COLS.RUN_STATUS] + 1).setValue('Running');

      // StartedAt -> now (only if blank)
      const startedCol = idx[CFG.COLS.RUN_STARTED_AT];
      if (startedCol !== undefined) {
        const startedVal = row[startedCol];
        if (!startedVal) {
          sh.getRange(rowNum, startedCol + 1).setValue(new Date());
        }
      }
      return;
    }
  }
}

function stopCurrentExportRun() {
  if (stopPathPortalSnapshotExport_()) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'The portal export was stopped safely.',
      'PATH',
      5
    );
    return;
  }

  const run = findNextPendingOrRunningRun_();
  if (!run) {
    SpreadsheetApp.getActiveSpreadsheet().toast('No running export found.', 'PATH', 5);
    return;
  }

  // 1️⃣ Stop future triggers
  clearContinuationTrigger_();

  // 2️⃣ Clear saved state (queue + index)
  clearRunState_(run.RunID);

  // 3️⃣ Mark as Stopped in sheet
  markRunStopped_(run.RunID);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Export stopped.\nRunID: ${run.RunID}`,
    'PATH',
    5
  );
}

function markRunStopped_(runID) {
  const sh = openSheet_(CFG.SHEETS.EXPORT_RUNS);
  const { values, idx } = readSheet_(sh);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[idx[CFG.COLS.RUNID]] || '').trim() === String(runID).trim()) {
      const rowNum = i + CFG.HEADER_ROW + 1;
      sh.getRange(rowNum, idx[CFG.COLS.RUN_STATUS] + 1).setValue('Stopped');
      return;
    }
  }
}

function markRunComplete_(runID) {
  const sh = openSheet_(CFG.SHEETS.EXPORT_RUNS);
  const { values, idx } = readSheet_(sh);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    if (String(row[idx[CFG.COLS.RUNID]] || '').trim() === String(runID).trim()) {
      const rowNum = i + CFG.HEADER_ROW + 1;

      // Set Status → Complete
      sh.getRange(rowNum, idx[CFG.COLS.RUN_STATUS] + 1).setValue('Complete');

      // Set Status for AppSheet
      setRunActive_(false);

      return;
    }
  }
}
