function loadRunState_(runID) {
  const raw = PropertiesService.getScriptProperties().getProperty(CFG.PROP_PREFIX + runID);
  if (!raw) return {};
  return JSON.parse(raw);
}


function saveRunState_(runID, state) {
  PropertiesService.getScriptProperties().setProperty(CFG.PROP_PREFIX + runID, JSON.stringify(state));
}

function clearRunState_(runID) {
  PropertiesService.getScriptProperties().deleteProperty(CFG.PROP_PREFIX + runID);
}

// Advanced: if you need to clear state for the currently running run
function resetCurrentRunState() {
  const run = findNextPendingOrRunningRun_();
  if (!run) return;
  clearRunState_(run.RunID);
  clearContinuationTrigger_();
}
