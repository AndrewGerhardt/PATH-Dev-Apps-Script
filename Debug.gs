function wipeAllRunState() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();

  for (const key in all) {
    if (key.startsWith(CFG.PROP_PREFIX)) {
      props.deleteProperty(key);
    }
  }
}

function testDriveAccessToast() {
  const ss = getSS_();

  const outId = String(CFG.OUTPUT_FOLDER_ID).trim();
  const tplId = String(CFG.TEMPLATE_DOC_ID).trim();

  const outFolderName = DriveApp.getFolderById(outId).getName();
  const templateName = DriveApp.getFileById(tplId).getName();

  console.log(`Output folder OK: ${outFolderName}`);
  console.log(`Template doc OK: ${templateName}`);

  ss.toast(
    `Drive OK ✅\nOutput: ${outFolderName}\nTemplate: ${templateName}`,
    'PATH Export Test',
    10
  );
}


function debugConfigIdsToast() {
  const ss = getSS_();

  const out = String(CFG.OUTPUT_FOLDER_ID ?? '');
  const tpl = String(CFG.TEMPLATE_DOC_ID ?? '');

  ss.toast(
    `OUTPUT_FOLDER_ID: "${out}" (len ${out.length})\n` +
    `TEMPLATE_DOC_ID: "${tpl}" (len ${tpl.length})`,
    'CFG Debug',
    12
  );
}

function debugConfigIdsLog() {
  const out = String(CFG.OUTPUT_FOLDER_ID ?? '');
  const tpl = String(CFG.TEMPLATE_DOC_ID ?? '');

  console.log(`OUTPUT_FOLDER_ID="${out}" len=${out.length}`);
  console.log(`TEMPLATE_DOC_ID="${tpl}" len=${tpl.length}`);
}

function testReqFormat() {
  console.log(formatReqList_('Canoeing – 1a , Canoeing – 1b , Canoeing – 1c', 'Canoeing'));
}

function testDocsApi() {
  const doc = DocumentApp.create("Docs API Test");
  const docId = doc.getId();

  Docs.Documents.batchUpdate(
    {
      requests: [
        {
          insertText: {
            location: { index: 1 },
            text: "Docs API is working!"
          }
        }
      ]
    },
    docId
  );

  Logger.log("Created doc: https://docs.google.com/document/d/" + docId);
}

function testParallelSimple() {
  ScriptApp.newTrigger('testWorkerA')
    .timeBased()
    .after(1000)
    .create();

  ScriptApp.newTrigger('testWorkerB')
    .timeBased()
    .after(1000)
    .create();

  Logger.log('Scheduled two workers.');
}

function testWorkerA() {
  Logger.log('Worker A running');
}

function testWorkerB() {
  Logger.log('Worker B running');
}

function testParallelQueue() {

  const run = findNextPendingOrRunningRun_();
  if (!run) {
    Logger.log('No run found.');
    return;
  }

  const queue = buildQueueForRun_(run);
  Logger.log('Queue length: ' + queue.length);

  PropertiesService.getScriptProperties().setProperty('TEST_QUEUE', JSON.stringify(queue));
  PropertiesService.getScriptProperties().setProperty('TEST_RUN', run.RunID);

  ScriptApp.newTrigger('testQueueWorker0')
    .timeBased()
    .after(1000)
    .create();

  ScriptApp.newTrigger('testQueueWorker1')
    .timeBased()
    .after(1000)
    .create();
}
function testQueueWorker0() { runTestQueueWorker_(0); }
function testQueueWorker1() { runTestQueueWorker_(1); }

function runTestQueueWorker_(workerId) {

  const queue = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('TEST_QUEUE') || '[]'
  );

  const workerCount = 2;

  Logger.log('Worker ' + workerId + ' starting.');

  for (let i = workerId; i < queue.length; i += workerCount) {
    Logger.log('Worker ' + workerId + ' processing rowNumber ' + queue[i].rowNumber);
  }

  Logger.log('Worker ' + workerId + ' done.');
}
