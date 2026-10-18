function uiStartPdfRun() {
  const ui = SpreadsheetApp.getUi();

  const ss = getSS_();

  const year = promptYearOptionA_(ui);
  if (year === null) return;

  const session = promptNumber_(ui, 'Session', 'Enter session (ex: 1–6):');
  if (session === null) return;

  const runId = createExportRun_(year, session, ss);
  markRunRunning_(runId);

  const state = {
    i0: 0,
    i1: 1,
    i2: 2,
    i3: 3
  };

  saveRunState_(runId, state);

  clearContinuationTrigger_();
  clearParallelWorkerTriggers_();

  scheduleParallelWorker_(1, 2000);
  scheduleParallelWorker_(2, 4000);
  scheduleParallelWorker_(3, 6000);
  scheduleParallelWorker_(0, 8000);

  ui.alert(
    `Started PARALLEL PDF export.\nRunID: ${runId}\nYear: ${year}\nSession: ${session}`
  );
}

function uiCreateFoldersOnly() {
  const ui = SpreadsheetApp.getUi();

  const year = promptYearOptionA_(ui);
  if (year === null) return;

  const session = promptNumber_(ui, 'Session', 'Enter session (ex: 1–6):');
  if (session === null) return;

  const run = {
    Year: String(year),
    Session: String(session),
  };

  const count = createTroopFoldersForRunOnly_(run);
  ui.alert(`Created/ensured ${count} troop folders.\nYear: ${year}\nSession: ${session}\n`);
}

function uiShowLatestRunStatus() {
  const ui = SpreadsheetApp.getUi();
  const ss = getSS_();
  const latest = getLatestRun_(ss);
  if (!latest) {
    ui.alert('No runs found.');
    return;
  }
  ui.alert(
    `Latest Run:\nRunID: ${latest.RunID}\nYear: ${latest.Year}\nSession: ${latest.Session}\nStatus: ${latest.Status}\nStartedAt: ${latest.StartedAt || ''}`
  );
}

/* ---------- helpers ---------- */

function promptNumber_(ui, title, msg) {
  const r = ui.prompt(title, msg, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  const n = Number(String(r.getResponseText()).trim());
  if (!Number.isFinite(n)) {
    ui.alert('Please enter a valid number.');
    return null;
  }
  return n;
}

function promptYesNo_(ui, title, msg) {
  const r = ui.alert(title, msg, ui.ButtonSet.YES_NO);
  if (r === ui.Button.YES) return true;
  if (r === ui.Button.NO) return false;
  return null;
}