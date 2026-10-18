function startRunCore_(session, year) {
  Logger.log("CORE START");

  const ss = getSS_();

  const runId = createExportRun_(year, session, ss);

  setRunActive_(true);

  Logger.log("Before markRunRunning_");
  markRunRunning_(runId);

  const run = {
    RunID: runId,
    Session: session,
    Year: year
  };

  Logger.log("Before buildQueueForRun_");
  buildQueueForRun_(run);

  Logger.log("Scheduling workers");

  clearContinuationTrigger_();
  clearParallelWorkerTriggers_();

  scheduleParallelWorker_(1, 2000);
  scheduleParallelWorker_(2, 4000);
  scheduleParallelWorker_(3, 6000);
  scheduleParallelWorker_(0, 8000);

  Logger.log("CORE END");

  return runId;
}