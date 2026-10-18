/*************************************************
 * Sequential Continuation Trigger (legacy)
 *************************************************/

function parallelWorker0_() { processRunChunkParallel_(0); }
function parallelWorker1_() { processRunChunkParallel_(1); }
function parallelWorker2_() { processRunChunkParallel_(2); }
function parallelWorker3_() { processRunChunkParallel_(3); }

function scheduleContinuationTrigger_() {
  clearContinuationTrigger_();
  ScriptApp.newTrigger('startOrContinueNextRun')
    .timeBased()
    .after(10 * 1000)
    .create();
}

function clearContinuationTrigger_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'startOrContinueNextRun') {
      ScriptApp.deleteTrigger(t);
    }
  });
}


/*************************************************
 * Parallel Worker Trigger System (4 workers)
 *************************************************/

function clearParallelWorkerTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (
      fn === 'parallelWorker0_' ||
      fn === 'parallelWorker1_' ||
      fn === 'parallelWorker2_' ||
      fn === 'parallelWorker3_'
    ) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function scheduleParallelWorker_(workerId, delayMs) {

  const handler = 'parallelWorker' + workerId + '_';

  // Delete existing trigger for this worker only
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .after(delayMs || 2000)
    .create();
}