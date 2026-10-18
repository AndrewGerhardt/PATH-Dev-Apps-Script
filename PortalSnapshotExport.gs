/**
 * Sends finalized PATH report data to the records portal without creating
 * Google Drive PDFs. The browser generates one PDF when an authorized user
 * opens or downloads one report.
 */
const PATH_PORTAL_EXPORT = Object.freeze({
  STATE_PROPERTY: 'PATH_PORTAL_SNAPSHOT_EXPORT_STATE_V1',
  LAST_RESULT_PROPERTY: 'PATH_PORTAL_SNAPSHOT_EXPORT_LAST_RESULT_V1',
  CONTINUATION_FUNCTION: 'continuePathPortalSnapshotExport',
  MAX_RUNTIME_MS: 4 * 60 * 1000,
  RETRY_DELAY_MS: 15 * 1000,
  TEMPLATE_VERSION: 1,
  SCHEMA_VERSION: 1,
  SOURCE_FORMAT: 'current',
  COMPLETION_EMAIL_ENDPOINT: '/api/portal-export-complete',
  COMPLETION_EMAIL_MAX_ATTEMPTS: 3
});

/**
 * Read-only preview. This validates the selected session and shows how many
 * independent reports will be sent. It does not write to D1 or Cloudflare.
 */
function uiPreviewPortalReportExport() {
  const ui = SpreadsheetApp.getUi();
  const selection = promptPathPortalExportSelection_(ui);

  if (!selection) return;

  try {
    const packageData = buildPathPortalExportPackage_(
      selection.year,
      selection.session
    );

    ui.alert(
      'Portal Export Preview',
      buildPathPortalPreviewMessage_(packageData),
      ui.ButtonSet.OK
    );

    return packageData.summary;
  } catch (error) {
    ui.alert(
      'Portal Export Blocked',
      String(error && error.message || error),
      ui.ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Confirms that the Cloudflare endpoint, Resend configuration, sender, and
 * admin recipient are working before a full-session export is started.
 */
function uiTestPathPortalExportCompletionEmail() {
  const ui = SpreadsheetApp.getUi();

  try {
    const result = postPathPortalCompletionEmail_({
      test: true,
      requestId: Utilities.getUuid(),
      generatedAt: new Date().toISOString()
    });

    ui.alert(
      'Completion Email Test Sent',
      result.message ||
        'The portal accepted the test and sent it through Resend.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Completion Email Test Failed',
      'The portal could not send the test through Resend.\n\n' +
        String(error && error.message || error),
      ui.ButtonSet.OK
    );
    throw error;
  }
}

/**
 * Starts the production export from the spreadsheet menu.
 */
function uiStartPortalReportExport() {
  const ui = SpreadsheetApp.getUi();
  const selection = promptPathPortalExportSelection_(ui);

  if (!selection) return;

  const packageData = buildPathPortalExportPackage_(
    selection.year,
    selection.session
  );

  const confirmation = ui.alert(
    'Send Reports to Portal',
    buildPathPortalPreviewMessage_(packageData) +
      '\n\nNo Google Drive PDFs will be created. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) return;

  const runId = startPathPortalSnapshotExport_(
    selection.session,
    selection.year,
    packageData
  );

  ui.alert(
    'Portal Export Started',
    'RunID: ' + runId + '\n\n' +
      packageData.summary.reportCount +
      ' separate reports passed validation. The export will continue automatically.',
    ui.ButtonSet.OK
  );

  return runId;
}

/**
 * Production pilot for one Roster row. A combo offering still creates and
 * uploads both independent badge reports for that row.
 */
function uiStartPortalReportPilot() {
  const ui = SpreadsheetApp.getUi();
  const selection = promptPathPortalExportSelection_(ui);

  if (!selection) return;

  const response = ui.prompt(
    'Production Pilot',
    'Paste the exact RosterID to send:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const rosterId = cleanPathPortalText_(response.getResponseText());
  if (!rosterId) {
    ui.alert('A RosterID is required.');
    return;
  }

  const packageData = buildPathPortalExportPackage_(
    selection.year,
    selection.session,
    { rosterId: rosterId }
  );
  const confirmation = ui.alert(
    'Send Production Pilot',
    buildPathPortalPreviewMessage_(packageData) +
      '\n\nOnly RosterID ' + rosterId + ' will be sent. Continue?',
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) return;

  const runId = startPathPortalSnapshotExport_(
    selection.session,
    selection.year,
    packageData,
    { rosterId: rosterId }
  );

  ui.alert(
    'Production Pilot Started',
    'RunID: ' + runId + '\nRosterID: ' + rosterId + '\nReports: ' +
      packageData.summary.reportCount,
    ui.ButtonSet.OK
  );

  return runId;
}

/**
 * AppSheet-safe entry point. It validates the data before scheduling work.
 */
function startPathPortalSnapshotExport_(session, year, preparedPackage, options) {
  const normalizedYear = Number(year);
  const normalizedSession = Number(session);

  if (!Number.isInteger(normalizedYear) || !Number.isInteger(normalizedSession)) {
    throw new Error('A whole-number year and session are required.');
  }

  if (isRunActive_()) {
    throw new Error('A PATH export is already in progress.');
  }

  if (loadPathPortalExportState_()) {
    throw new Error(
      'A prior portal export is waiting to be resumed or stopped.'
    );
  }

  const packageData = preparedPackage || buildPathPortalExportPackage_(
    normalizedYear,
    normalizedSession,
    options || {}
  );
  const runId = createExportRun_(normalizedYear, normalizedSession);
  const finalizedAt = new Date().toISOString();
  const batchId = createIncrementalD1DeterministicId_(
    'batch',
    'portal|' + normalizedYear + '|' + normalizedSession + '|' +
      finalizedAt + '|' + Utilities.getUuid()
  );

  markRunRunning_(runId);
  setRunActive_(true);

  const state = {
    runId: runId,
    year: normalizedYear,
    session: normalizedSession,
    stage: 'metadata',
    cursor: 0,
    finalizedAt: finalizedAt,
    batchId: batchId,
    expectedReportCount: packageData.summary.reportCount,
    sourceFingerprint: packageData.sourceFingerprint,
    filterRosterId: cleanPathPortalText_(options && options.rosterId),
    metadataBatchCreated: false,
    metadata: {
      newRecords: 0,
      updatedRecords: 0,
      unchanged: 0
    },
    snapshots: emptyPathPortalSnapshotSummary_(),
    startedAt: finalizedAt,
    lastError: ''
  };

  savePathPortalExportState_(state);
  clearPathPortalExportTriggers_();
  schedulePathPortalExportContinuation_(1000);

  return runId;
}

/**
 * Time-trigger handler. Work is resumable and source data is fingerprinted so
 * a session cannot be partly exported from two different Roster states.
 */
function continuePathPortalSnapshotExport() {
  const startedAt = Date.now();
  const state = loadPathPortalExportState_();

  if (!state) return;

  clearPathPortalExportTriggers_();

  try {
    const packageData = buildPathPortalExportPackage_(
      state.year,
      state.session,
      { rosterId: state.filterRosterId || '' }
    );

    assertPathPortalSourceUnchanged_(state, packageData);

    if (state.stage === 'metadata') {
      processPathPortalMetadataStage_(state, packageData);

      if (state.stage === 'waiting_metadata') {
        savePathPortalExportState_(state);
        schedulePathPortalExportContinuation_(PATH_PORTAL_EXPORT.RETRY_DELAY_MS);
        return;
      }
    }

    if (state.stage === 'waiting_metadata') {
      const bookmark = PropertiesService
        .getScriptProperties()
        .getProperty(D1_INCREMENTAL_IMPORT_BOOKMARK_PROPERTY);

      if (!bookmark) {
        throw new Error('The pending D1 metadata import bookmark is missing.');
      }

      const metadataResult = pollIncrementalD1Import_(
        bookmark,
        Math.min(PATH_PORTAL_EXPORT.MAX_RUNTIME_MS, 3 * 60 * 1000)
      );

      if (!metadataResult.complete) {
        savePathPortalExportState_(state);
        schedulePathPortalExportContinuation_(PATH_PORTAL_EXPORT.RETRY_DELAY_MS);
        return;
      }

      state.stage = 'snapshots';
      savePathPortalExportState_(state);
    }

    if (state.stage === 'snapshots') {
      processPathPortalSnapshotStage_(state, packageData, startedAt);

      if (state.cursor < packageData.reports.length) {
        savePathPortalExportState_(state);
        schedulePathPortalExportContinuation_(PATH_PORTAL_EXPORT.RETRY_DELAY_MS);
        return;
      }

      state.stage = 'notifications';
      savePathPortalExportState_(state);
    }

    if (state.stage === 'notifications') {
      let notificationResult = null;
      const hasPendingNotifications = Boolean(
        PropertiesService.getScriptProperties().getProperty(
          D1_RECORD_ACTIVITY_PENDING_BATCH_PROPERTY
        )
      );

      if (state.metadataBatchCreated || hasPendingNotifications) {
        try {
          notificationResult = sendPendingRecordActivityNotifications();
        } catch (notificationError) {
          notificationResult = {
            complete: false,
            error: String(notificationError.message || notificationError)
          };
          Logger.log(
            'Report snapshots are complete; record activity emails remain pending: ' +
              notificationResult.error
          );
        }
      }

      completePathPortalExport_(state, packageData, notificationResult);
    }
  } catch (error) {
    failPathPortalExport_(state, error);
    throw error;
  }
}

function processPathPortalMetadataStage_(state, packageData) {
  const existingState = loadExistingD1State_(state.year);

  packageData.records.forEach(function(record) {
    const existing = existingState.recordsBySourceKey[record.sourceKey];
    if (existing && existing.fileUrl) {
      record.fileUrl = existing.fileUrl;
    }
  });

  const comparison = compareCurrentRecordsToD1_(
    packageData.records,
    existingState
  );
  const changedRecords = comparison.newRecords.concat(
    comparison.updatedRecords
  );

  state.metadata = {
    newRecords: comparison.newRecords.length,
    updatedRecords: comparison.updatedRecords.length,
    unchanged: comparison.unchangedCount
  };

  if (!changedRecords.length) {
    state.stage = 'snapshots';
    savePathPortalExportState_(state);
    return;
  }

  const sqlResult = buildIncrementalD1Sql_(
    changedRecords,
    existingState,
    state.year,
    state.batchId
  );

  if (!sqlResult.statementCount) {
    throw new Error('Metadata changes were found but no D1 writes were built.');
  }

  const importResult = uploadIncrementalSqlToD1_(sqlResult.sql);

  savePendingRecordActivityBatch_(state.batchId, state.year);
  state.metadataBatchCreated = true;
  state.stage = importResult.complete ? 'snapshots' : 'waiting_metadata';
  savePathPortalExportState_(state);
}

function processPathPortalSnapshotStage_(state, packageData, startedAt) {
  const existingState = loadExistingD1State_(state.year);
  const secret = PropertiesService
    .getScriptProperties()
    .getProperty(PATH_REPORT_SNAPSHOT_UPLOAD.SECRET_PROPERTY);

  if (!secret) {
    throw new Error(
      'Missing Script Property: ' +
        PATH_REPORT_SNAPSHOT_UPLOAD.SECRET_PROPERTY
    );
  }

  let uploadedAChunk = false;

  while (state.cursor < packageData.reports.length) {
    if (
      uploadedAChunk &&
      Date.now() - startedAt >= PATH_PORTAL_EXPORT.MAX_RUNTIME_MS
    ) {
      break;
    }

    const drafts = packageData.reports.slice(
      state.cursor,
      state.cursor + PATH_REPORT_SNAPSHOT_UPLOAD.CHUNK_SIZE
    );
    const envelopes = drafts.map(function(draft) {
      const existing = existingState.recordsBySourceKey[draft.sourceKey];

      if (!existing || !existing.recordId) {
        throw new Error(
          'Metadata record not found for ' + draft.fileName +
          '. The snapshot was not uploaded.'
        );
      }

      return {
        recordId: existing.recordId,
        templateVersion: PATH_PORTAL_EXPORT.TEMPLATE_VERSION,
        finalizedAt: state.finalizedAt,
        snapshot: draft.snapshot
      };
    });

    const result = postPathSnapshotChunk_(
      envelopes,
      state.batchId,
      secret
    );

    addPathPortalSnapshotSummary_(state.snapshots, result);
    state.cursor += envelopes.length;
    state.lastProgressAt = new Date().toISOString();
    savePathPortalExportState_(state);
    uploadedAChunk = true;
  }
}

function completePathPortalExport_(state, packageData, notificationResult) {
  const completedAt = new Date().toISOString();
  const result = {
    complete: true,
    runId: state.runId,
    year: state.year,
    session: state.session,
    reportCount: packageData.summary.reportCount,
    metadata: state.metadata,
    snapshots: state.snapshots,
    notifications: notificationResult,
    startedAt: state.startedAt,
    completedAt: completedAt
  };

  result.completionEmail = state.filterRosterId
    ? {
        sent: false,
        skipped: true,
        reason: 'Production pilots do not send completion emails.'
      }
    : sendPathPortalExportCompletionEmail_(result);

  PropertiesService.getScriptProperties().setProperty(
    PATH_PORTAL_EXPORT.LAST_RESULT_PROPERTY,
    JSON.stringify(result)
  );

  clearPathPortalExportTriggers_();
  clearPathPortalExportState_();
  markRunComplete_(state.runId);

  const message =
    packageData.summary.reportCount + ' separate reports sent to the portal. ' +
    state.snapshots.created + ' created, ' +
    state.snapshots.versioned + ' versioned, ' +
    state.snapshots.restored + ' restored, and ' +
    state.snapshots.unchanged + ' unchanged.';

  Logger.log(message);
  getSS_().toast(message, 'Portal Export Complete', 12);
}

function failPathPortalExport_(state, error) {
  const message = String(error && error.message || error);

  state.lastError = message;
  state.failedAt = new Date().toISOString();
  savePathPortalExportState_(state);
  clearPathPortalExportTriggers_();
  setRunActive_(false);
  markRunStopped_(state.runId);

  Logger.log('Portal export stopped: ' + message);
  getSS_().toast(
    'Portal export stopped safely. Open Portal Export Status for details.',
    'Portal Export Stopped',
    12
  );
}

function resumePathPortalSnapshotExport() {
  const state = loadPathPortalExportState_();

  if (!state) {
    SpreadsheetApp.getUi().alert('There is no portal export to resume.');
    return;
  }

  if (isRunActive_()) {
    SpreadsheetApp.getUi().alert('A PATH export is already running.');
    return;
  }

  state.lastError = '';
  delete state.failedAt;
  savePathPortalExportState_(state);
  markRunRunning_(state.runId);
  setRunActive_(true);
  clearPathPortalExportTriggers_();
  schedulePathPortalExportContinuation_(1000);

  SpreadsheetApp.getUi().alert(
    'Portal export resumed from report ' + state.cursor + '.'
  );
}

function uiShowPortalExportStatus() {
  const properties = PropertiesService.getScriptProperties();
  const state = loadPathPortalExportState_();
  const lastText = properties.getProperty(
    PATH_PORTAL_EXPORT.LAST_RESULT_PROPERTY
  );

  if (state) {
    SpreadsheetApp.getUi().alert(
      'Portal Export Status',
      [
        'RunID: ' + state.runId,
        'Year: ' + state.year,
        'Session: ' + state.session,
        'Stage: ' + state.stage,
        'Reports uploaded: ' + state.cursor + ' / ' + state.expectedReportCount,
        'Last error: ' + (state.lastError || 'None')
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  if (lastText) {
    const last = JSON.parse(lastText);
    SpreadsheetApp.getUi().alert(
      'Last Portal Export',
      [
        'RunID: ' + last.runId,
        'Year: ' + last.year,
        'Session: ' + last.session,
        'Reports: ' + last.reportCount,
        'Completed: ' + last.completedAt,
        'Completion email: ' +
          describePathPortalCompletionEmailStatus_(last.completionEmail)
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  SpreadsheetApp.getUi().alert('No portal export status is available.');
}

function stopPathPortalSnapshotExport_() {
  const state = loadPathPortalExportState_();
  if (!state) return false;

  clearPathPortalExportTriggers_();
  clearPathPortalExportState_();
  setRunActive_(false);
  markRunStopped_(state.runId);
  return true;
}

function sendPathPortalExportCompletionEmail_(result) {
  try {
    const response = postPathPortalCompletionEmail_(result);

    return {
      sent: true,
      provider: 'Resend',
      message: response.message || '',
      sentAt: new Date().toISOString()
    };
  } catch (error) {
    const message = String(error && error.message || error);
    Logger.log('Portal export completion email failed: ' + message);
    return {
      sent: false,
      provider: 'Resend',
      error: message
    };
  }
}

function postPathPortalCompletionEmail_(result) {
  const secret = PropertiesService
    .getScriptProperties()
    .getProperty(PATH_REPORT_SNAPSHOT_UPLOAD.SECRET_PROPERTY);

  if (!secret) {
    throw new Error(
      'Missing Script Property: ' +
        PATH_REPORT_SNAPSHOT_UPLOAD.SECRET_PROPERTY
    );
  }

  const url =
    PATH_REPORT_SNAPSHOT_UPLOAD.PORTAL_ORIGIN +
    PATH_PORTAL_EXPORT.COMPLETION_EMAIL_ENDPOINT;
  const payload = JSON.stringify(result);
  let lastError = null;

  for (
    let attempt = 1;
    attempt <= PATH_PORTAL_EXPORT.COMPLETION_EMAIL_MAX_ATTEMPTS;
    attempt += 1
  ) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'X-Import-Secret': secret
        },
        payload: payload,
        muteHttpExceptions: true
      });
      const status = response.getResponseCode();
      const text = response.getContentText();
      let data = {};

      try {
        data = text ? JSON.parse(text) : {};
      } catch (error) {
        const invalidResponseError = new Error(
          'Completion email returned HTTP ' + status +
            ' with an invalid response.'
        );
        invalidResponseError.pathDoNotRetry = status < 500;
        throw invalidResponseError;
      }

      if (status >= 200 && status < 300 && data.success) {
        return data;
      }

      const retryable = status === 429 || status >= 500;
      const message = data.message ||
        'Completion email failed with HTTP ' + status + '.';

      if (!retryable) {
        const requestError = new Error(message);
        requestError.pathDoNotRetry = true;
        throw requestError;
      }

      lastError = new Error(message);
    } catch (error) {
      if (error && error.pathDoNotRetry) throw error;
      lastError = error;
    }

    if (attempt < PATH_PORTAL_EXPORT.COMPLETION_EMAIL_MAX_ATTEMPTS) {
      Utilities.sleep(1000 * Math.pow(2, attempt - 1));
    }
  }

  throw lastError || new Error('The completion email could not be sent.');
}

function describePathPortalCompletionEmailStatus_(completionEmail) {
  if (!completionEmail) return 'Not recorded';
  if (completionEmail.sent) return 'Sent through Resend';
  if (completionEmail.skipped) return 'Not sent for pilot';
  return 'Failed - ' + cleanPathPortalText_(
    completionEmail.error || 'No error details were returned.'
  );
}

function buildPathPortalExportPackage_(year, session, options) {
  const normalizedYear = Number(year);
  const normalizedSession = Number(session);
  const filterRosterId = cleanPathPortalText_(options && options.rosterId);
  const rosterSheet = openSheet_(CFG.SHEETS.ROSTER);
  const rosterData = readSheet_(rosterSheet);
  const scoutMap = buildPathPortalScoutMap_();
  const offeringMap = buildPathPortalOfferingMap_();
  const badgeMap = buildPathPortalBadgeMap_();
  const templateData = buildPathPortalRequirementTemplate_();
  const errors = [];
  const builtItems = [];
  let rosterCount = 0;
  let comboRosterCount = 0;

  requirePathHeaders_(
    CFG.SHEETS.ROSTER,
    rosterData.idx,
    [
      CFG.COLS.ROSTERID,
      CFG.COLS.ROSTER_YEAR,
      CFG.COLS.ROSTER_SESSION,
      CFG.COLS.ROSTER_OFFERINGID,
      CFG.COLS.ROSTER_SCOUTID,
      'M', 'T', 'W', 'TH', 'F',
      'ReqComplete', 'ReqOptional', 'ReqIncomplete',
      'Comments', 'Earned_1', 'Earned_2'
    ]
  );

  rosterData.values.forEach(function(row, index) {
    const roster = rowToObject_(rosterData.headers, row);
    const currentRosterId = cleanPathPortalText_(
      roster[CFG.COLS.ROSTERID]
    );

    if (
      Number(roster[CFG.COLS.ROSTER_YEAR]) !== normalizedYear ||
      Number(roster[CFG.COLS.ROSTER_SESSION]) !== normalizedSession
    ) {
      return;
    }

    const selected = !filterRosterId || currentRosterId === filterRosterId;
    if (selected) rosterCount += 1;
    const rowNumber = index + CFG.HEADER_ROW + 1;
    const rosterId = currentRosterId || 'row ' + rowNumber;
    const scoutId = cleanPathPortalText_(
      roster[CFG.COLS.ROSTER_SCOUTID]
    );
    const offeringId = cleanPathPortalText_(
      roster[CFG.COLS.ROSTER_OFFERINGID]
    );
    const scout = scoutMap.get(scoutId);
    const offering = offeringMap.get(offeringId);

    if (!scout) {
      errors.push(rosterId + ': ScoutID ' + scoutId + ' was not found.');
      return;
    }

    if (!offering) {
      errors.push(rosterId + ': OfferingID ' + offeringId + ' was not found.');
      return;
    }

    const badgeIds = [offering.BadgeID_1, offering.BadgeID_2]
      .map(cleanPathPortalText_)
      .filter(Boolean);

    if (!badgeIds.length) {
      errors.push(rosterId + ': the offering has no badge or activity.');
      return;
    }

    if (selected && badgeIds.length > 1) comboRosterCount += 1;

    badgeIds.forEach(function(badgeId, badgeIndex) {
      try {
        const badge = badgeMap.get(badgeId);

        if (!badge || !badge.BadgeName) {
          throw new Error('BadgeID ' + badgeId + ' was not found in Badges.');
        }

        const built = buildPathPortalReport_(
          normalizedYear,
          normalizedSession,
          roster,
          scout,
          offering,
          badge,
          badgeIndex + 1,
          badgeIds,
          badgeMap,
          templateData
        );

        built.rosterId = currentRosterId;
        built.rowNumber = rowNumber;
        built.selected = selected;
        built.period = cleanPathPortalText_(offering.Period);
        built.legacyFileUrl = cleanPathPortalText_(
          roster[
            badgeIndex === 1
              ? CFG.COLS.PR_FILE_2
              : CFG.COLS.PR_FILE_1
          ]
        );
        builtItems.push(built);
      } catch (error) {
        errors.push(
          rosterId + ' / ' + badgeId + ': ' +
            String(error.message || error)
        );
      }
    });
  });

  if (!rosterCount) {
    errors.push(
      'No Roster rows matched Year ' + normalizedYear +
      ', Session ' + normalizedSession +
      (filterRosterId ? ', and RosterID ' + filterRosterId : '') + '.'
    );
  }

  if (errors.length) {
    throw new Error(buildPathPortalValidationError_(errors));
  }

  finalizePathPortalReportIdentities_(builtItems);

  const selectedItems = builtItems.filter(function(item) {
    return item.selected;
  });
  const reports = selectedItems.map(function(item) {
    return item.report;
  });
  const records = selectedItems.map(function(item) {
    return item.record;
  });
  const prerequisiteCount = selectedItems.reduce(function(total, item) {
    return total + item.prerequisiteCount;
  }, 0);

  const sourceFingerprint = pathPortalSha256Hex_(
    JSON.stringify(reports.map(function(report) {
      return {
        sourceKey: report.sourceKey,
        snapshot: report.snapshot
      };
    }))
  );

  return {
    reports: reports,
    records: records,
    sourceFingerprint: sourceFingerprint,
    summary: {
      year: normalizedYear,
      session: normalizedSession,
      rosterCount: rosterCount,
      reportCount: reports.length,
      comboRosterCount: comboRosterCount,
      prerequisiteCount: prerequisiteCount
    }
  };
}

/**
 * A Scout can have more than one Roster row for the same badge and session.
 * Keep the legacy identity for one row so existing portal records are updated,
 * then give every additional row a stable RosterID-based identity. All rows in
 * the repeated group receive distinguishable filenames and portal metadata.
 */
function finalizePathPortalReportIdentities_(builtItems) {
  const groups = {};

  builtItems.forEach(function(item) {
    if (!groups[item.sourceKey]) groups[item.sourceKey] = [];
    groups[item.sourceKey].push(item);
  });

  Object.keys(groups).forEach(function(baseSourceKey) {
    const group = groups[baseSourceKey];
    if (group.length < 2) return;

    const missingRosterId = group.find(function(item) {
      return !item.rosterId;
    });
    if (missingRosterId) {
      throw new Error(
        'Repeated reports require a RosterID on every matching Roster row.'
      );
    }

    const periodCounts = {};
    group.forEach(function(item) {
      const periodKey = normalizeKeyPart_(item.period || 'Period unavailable');
      periodCounts[periodKey] = (periodCounts[periodKey] || 0) + 1;
    });

    group.sort(function(a, b) {
      const legacyFileDifference =
        Number(Boolean(b.legacyFileUrl)) - Number(Boolean(a.legacyFileUrl));
      if (legacyFileDifference) return legacyFileDifference;

      return normalizeKeyPart_(a.rosterId).localeCompare(
        normalizeKeyPart_(b.rosterId)
      );
    });

    const assignedSourceKeys = {};

    group.forEach(function(item, index) {
      const period = item.period || 'Period unavailable';
      const periodKey = normalizeKeyPart_(period);
      const repeatedPeriod = periodCounts[periodKey] > 1;
      const qualifier = repeatedPeriod
        ? period + ' - ' + item.rosterId
        : period;
      const sourceKey = index === 0
        ? baseSourceKey
        : baseSourceKey + '|roster:' + normalizeKeyPart_(item.rosterId);

      if (assignedSourceKeys[sourceKey]) {
        throw new Error(
          'Repeated report rows must have different RosterIDs: ' +
            item.rosterId + '.'
        );
      }
      assignedSourceKeys[sourceKey] = true;

      const fileName = appendPathPortalFileQualifier_(
        item.report.fileName,
        qualifier
      );

      item.sourceKey = sourceKey;
      item.record.sourceKey = sourceKey;
      item.record.legacyTime = period;
      item.record.legacyCode = repeatedPeriod ? item.rosterId : null;
      item.record.fileName = fileName;
      item.report.sourceKey = sourceKey;
      item.report.fileName = fileName;
    });
  });
}

function appendPathPortalFileQualifier_(fileName, qualifier) {
  const cleanQualifier = cleanPathPortalFileComponent_(qualifier);
  if (!cleanQualifier) return fileName;

  return String(fileName || '').replace(
    /\.pdf$/i,
    ' - ' + cleanQualifier + '.pdf'
  );
}

function buildPathPortalReport_(
  year,
  session,
  roster,
  scout,
  offering,
  badge,
  badgeSlot,
  offeringBadgeIds,
  badgeMap,
  templateData
) {
  const badgeId = cleanPathPortalText_(badge.BadgeID);
  const badgeName = cleanPathPortalText_(badge.BadgeName);
  const recordClassName = cleanPathPortalFileComponent_(badgeName);
  const troopKey = buildPathPortalTroopKey_(scout);
  const firstName = cleanPathPortalText_(scout.FirstName);
  const lastName = cleanPathPortalText_(scout.LastName);
  const recordFirstName = cleanPathPortalFileComponent_(firstName);
  const recordLastName = cleanPathPortalFileComponent_(lastName);

  if (!firstName || !lastName) {
    throw new Error('The Scout name is incomplete.');
  }

  const scoutKey = [troopKey, recordLastName, recordFirstName]
    .map(normalizeKeyPart_)
    .join('|');
  const sourceKey = createRecordSourceKey_({
    scoutKey: scoutKey,
    className: recordClassName,
    session: session,
    year: year,
    legacyTime: null,
    legacyCode: null,
    sourceFormat: PATH_PORTAL_EXPORT.SOURCE_FORMAT
  });
  const fileName = buildPathPortalFileName_(
    scout,
    badgeName,
    session,
    year
  );
  const earned = normalizePathPortalEarned_(
    roster[badgeSlot === 2 ? 'Earned_2' : 'Earned_1']
  );
  const requirements = buildPathPortalRequirements_(
    roster,
    badgeId,
    badgeName,
    offeringBadgeIds,
    badgeMap,
    templateData
  );
  const notes = cleanPathPortalText_(roster.Comments);
  const snapshot = {
    schemaVersion: PATH_PORTAL_EXPORT.SCHEMA_VERSION,
    scout: {
      firstName: firstName,
      lastName: lastName,
      troopKey: troopKey
    },
    offering: {
      name: badgeName,
      area: cleanPathPortalText_(badge.Area || offering.Area),
      period: cleanPathPortalText_(offering.Period),
      session: session,
      year: year
    },
    attendance: {
      Monday: normalizePathPortalAttendance_(roster.M),
      Tuesday: normalizePathPortalAttendance_(roster.T),
      Wednesday: normalizePathPortalAttendance_(roster.W),
      Thursday: normalizePathPortalAttendance_(roster.TH),
      Friday: normalizePathPortalAttendance_(roster.F)
    },
    report: {
      status: earned
    },
    activities: [
      {
        name: badgeName,
        earned: earned,
        requirements: requirements
      }
    ]
  };

  if (notes) snapshot.notes = notes;

  return {
    sourceKey: sourceKey,
    prerequisiteCount: requirements.filter(function(item) {
      return item.prerequisite;
    }).length,
    record: {
      troop: troopKey,
      firstName: recordFirstName,
      lastName: recordLastName,
      scoutKey: scoutKey,
      sourceKey: sourceKey,
      className: recordClassName,
      session: session,
      year: year,
      legacyTime: null,
      legacyCode: null,
      sourceFormat: PATH_PORTAL_EXPORT.SOURCE_FORMAT,
      fileName: fileName,
      fileUrl: null
    },
    report: {
      sourceKey: sourceKey,
      fileName: fileName,
      snapshot: snapshot
    }
  };
}

function buildPathPortalRequirements_(
  roster,
  badgeId,
  badgeName,
  offeringBadgeIds,
  badgeMap,
  templateData
) {
  const statuses = [
    { field: 'ReqComplete', status: 'Complete' },
    { field: 'ReqOptional', status: 'Optional' },
    { field: 'ReqIncomplete', status: 'Not complete' }
  ];
  const byCode = {};
  const isCombo = offeringBadgeIds.length > 1;

  statuses.forEach(function(item) {
    splitPathPortalRequirementList_(roster[item.field])
      .forEach(function(token) {
        const parsed = parsePathPortalRequirementToken_(
          token,
          badgeId,
          badgeName,
          offeringBadgeIds,
          badgeMap,
          isCombo
        );

        if (!parsed.applies) return;

        const code = parsed.code;
        const normalizedCode = normalizePathPortalRequirementCode_(code);

        if (!normalizedCode) {
          throw new Error('A blank requirement code was found.');
        }

        if (
          byCode[normalizedCode] &&
          byCode[normalizedCode].status !== item.status
        ) {
          throw new Error(
            'Requirement ' + code + ' appears in both ' +
            byCode[normalizedCode].status + ' and ' + item.status + '.'
          );
        }

        byCode[normalizedCode] = {
          code: code,
          status: item.status,
          prerequisite: Boolean(
            templateData.prerequisites[
              buildPathPortalRequirementKey_(badgeId, code)
            ]
          ),
          order: templateData.order[
            buildPathPortalRequirementKey_(badgeId, code)
          ]
        };
      });
  });

  return Object.keys(byCode)
    .map(function(key) {
      return byCode[key];
    })
    .sort(function(left, right) {
      const leftOrder = Number.isFinite(left.order) ? left.order : 1000000;
      const rightOrder = Number.isFinite(right.order) ? right.order : 1000000;

      if (leftOrder !== rightOrder) return leftOrder - rightOrder;

      return left.code.localeCompare(right.code, undefined, {
        numeric: true,
        sensitivity: 'base'
      });
    })
    .map(function(item) {
      return {
        code: item.code,
        status: item.status,
        prerequisite: item.prerequisite
      };
    });
}

function parsePathPortalRequirementToken_(
  token,
  badgeId,
  badgeName,
  offeringBadgeIds,
  badgeMap,
  isCombo
) {
  const text = cleanPathPortalText_(token);

  if (!text) return { applies: false, code: '' };

  const descriptors = offeringBadgeIds.map(function(offeringBadgeId) {
    const badge = badgeMap.get(offeringBadgeId) || {};
    return {
      badgeId: offeringBadgeId,
      badgeName: cleanPathPortalText_(badge.BadgeName)
    };
  });

  for (const descriptor of descriptors) {
    const prefixes = [descriptor.badgeName, descriptor.badgeId]
      .filter(Boolean);

    for (const prefix of prefixes) {
      const match = text.match(
        new RegExp(
          '^' + escapeRegex_(prefix) + '\\s*[-\\u2012\\u2013\\u2014\\u2015:]\\s*(.+)$',
          'i'
        )
      );

      if (match) {
        return {
          applies: descriptor.badgeId === badgeId,
          code: cleanPathPortalText_(match[1])
        };
      }
    }
  }

  if (isCombo) {
    throw new Error(
      'Combo requirement "' + text +
      '" is missing its badge prefix. No reports were uploaded.'
    );
  }

  return {
    applies: true,
    code: scrubBadgePrefix_(text, badgeName)
  };
}

function buildPathPortalScoutMap_() {
  const sheet = openSheet_(CFG.SHEETS.SCOUTS);
  const data = readSheet_(sheet);

  requirePathHeaders_(CFG.SHEETS.SCOUTS, data.idx, [
    'ScoutID', 'FirstName', 'LastName', 'UnitNumber', 'UnitType', 'UnitState'
  ]);

  const map = new Map();
  data.values.forEach(function(row) {
    const item = rowToObject_(data.headers, row);
    const id = cleanPathPortalText_(item.ScoutID);
    if (id) map.set(id, item);
  });
  return map;
}

function buildPathPortalOfferingMap_() {
  const sheet = openSheet_(CFG.SHEETS.OFFERINGS);
  const data = readSheet_(sheet);

  requirePathHeaders_(CFG.SHEETS.OFFERINGS, data.idx, [
    'OfferingID', 'BadgeID_1', 'BadgeID_2', 'Period'
  ]);

  const map = new Map();
  data.values.forEach(function(row) {
    const item = rowToObject_(data.headers, row);
    const id = cleanPathPortalText_(item.OfferingID);
    if (id) map.set(id, item);
  });
  return map;
}

function buildPathPortalBadgeMap_() {
  const sheet = openSheet_(CFG.SHEETS.BADGES);
  const data = readSheet_(sheet);

  requirePathHeaders_(CFG.SHEETS.BADGES, data.idx, [
    'BadgeID', 'BadgeName'
  ]);

  const map = new Map();
  data.values.forEach(function(row) {
    const item = rowToObject_(data.headers, row);
    const id = cleanPathPortalText_(item.BadgeID);
    if (id) map.set(id, item);
  });
  return map;
}

function buildPathPortalRequirementTemplate_() {
  const sheet = openSheet_(CFG.SHEETS.BADGE_REQ_TEMPLATE);
  const data = readSheet_(sheet);

  requirePathHeaders_(CFG.SHEETS.BADGE_REQ_TEMPLATE, data.idx, [
    'BadgeID', 'ReqCode', 'PreReq'
  ]);

  const prerequisites = {};
  const order = {};

  data.values.forEach(function(row, index) {
    const item = rowToObject_(data.headers, row);
    const badgeId = cleanPathPortalText_(item.BadgeID);
    const code = cleanPathPortalText_(item.ReqCode);

    if (!badgeId || !code) return;

    const key = buildPathPortalRequirementKey_(badgeId, code);
    order[key] = index;

    if (toBool_(item.PreReq)) {
      prerequisites[key] = true;
    }
  });

  return {
    prerequisites: prerequisites,
    order: order
  };
}

function buildPathPortalTroopKey_(scout) {
  const number = cleanPathPortalText_(scout.UnitNumber);
  const type = cleanPathPortalText_(scout.UnitType).toUpperCase();
  const state = cleanPathPortalText_(stateToAbbrev_(scout.UnitState)).toUpperCase();

  if (!number || !type || !state) {
    throw new Error('Unit number, type, and state are required.');
  }

  return number + type + ' ' + state;
}

function buildPathPortalFileName_(scout, badgeName, session, year) {
  return (
    cleanPathPortalFileComponent_(scout.UnitNumber) +
    cleanPathPortalFileComponent_(scout.UnitType) + ' ' +
    cleanPathPortalFileComponent_(scout.LastName) + ', ' +
    cleanPathPortalFileComponent_(scout.FirstName) + ' - ' +
    cleanPathPortalFileComponent_(badgeName) + ' - S' +
    cleanPathPortalFileComponent_(session) + ' ' +
    cleanPathPortalFileComponent_(year) + '.pdf'
  );
}

function cleanPathPortalFileComponent_(value) {
  return String(value || '').replace(/[\\/:*?"<>|]/g, '').trim();
}

function normalizePathPortalEarned_(value) {
  const text = cleanPathPortalText_(value);
  const normalized = text.toLowerCase();

  if (
    value === true || value === 1 ||
    ['1', 'true', 'yes', 'complete', 'completed', 'earned'].includes(normalized)
  ) {
    return 'Earned';
  }

  if (
    value === false || value === 0 || !text ||
    ['0', 'false', 'no', 'incomplete', 'not complete', 'not earned'].includes(normalized)
  ) {
    return 'Not earned';
  }

  return text;
}

function normalizePathPortalAttendance_(value) {
  const text = cleanPathPortalText_(value);

  if (value === true || text.toLowerCase() === 'true') return 'Present';
  if (value === false || text.toLowerCase() === 'false') return 'Absent';
  return text || 'Not recorded';
}

function splitPathPortalRequirementList_(value) {
  return String(value || '')
    .split(/\s*,\s*/)
    .map(cleanPathPortalText_)
    .filter(Boolean);
}

function buildPathPortalRequirementKey_(badgeId, code) {
  return (
    cleanPathPortalText_(badgeId).toLowerCase() + '|' +
    normalizePathPortalRequirementCode_(code)
  );
}

function normalizePathPortalRequirementCode_(value) {
  return cleanPathPortalText_(value).toLowerCase();
}

function promptPathPortalExportSelection_(ui) {
  const year = promptYearOptionA_(ui);
  if (year === null) return null;

  const session = promptNumber_(
    ui,
    'Session',
    'Enter session (ex: 1-6):'
  );
  if (session === null) return null;

  return {
    year: Number(year),
    session: Number(session)
  };
}

function buildPathPortalPreviewMessage_(packageData) {
  const summary = packageData.summary;
  return [
    'Year: ' + summary.year,
    'Session: ' + summary.session,
    'Roster rows: ' + summary.rosterCount,
    'Separate reports: ' + summary.reportCount,
    'Combo-class roster rows split: ' + summary.comboRosterCount,
    'Prerequisite rows identified: ' + summary.prerequisiteCount,
    '',
    'Every report has one badge or activity and one independent portal record.'
  ].join('\n');
}

function buildPathPortalValidationError_(errors) {
  const shown = errors.slice(0, 20);
  const remaining = errors.length - shown.length;

  return (
    'The portal export found ' + errors.length +
    ' issue' + (errors.length === 1 ? '' : 's') +
    ' and made no changes:\n\n' +
    shown.join('\n') +
    (remaining > 0 ? '\n\n' + remaining + ' additional issues were omitted.' : '')
  );
}

function requirePathHeaders_(sheetName, index, required) {
  const missing = required.filter(function(header) {
    return index[header] === undefined;
  });

  if (missing.length) {
    throw new Error(
      sheetName + ' is missing required columns: ' + missing.join(', ')
    );
  }
}

function assertPathPortalSourceUnchanged_(state, packageData) {
  if (
    packageData.summary.reportCount !== Number(state.expectedReportCount) ||
    packageData.sourceFingerprint !== state.sourceFingerprint
  ) {
    throw new Error(
      'The source PATH data changed after this export started. ' +
      'Stop this run and start a new portal export so every report uses one consistent final dataset.'
    );
  }
}

function emptyPathPortalSnapshotSummary_() {
  return {
    processed: 0,
    created: 0,
    versioned: 0,
    restored: 0,
    unchanged: 0
  };
}

function addPathPortalSnapshotSummary_(target, result) {
  ['processed', 'created', 'versioned', 'restored', 'unchanged']
    .forEach(function(key) {
      target[key] = Number(target[key] || 0) + Number(result[key] || 0);
    });
}

function loadPathPortalExportState_() {
  const text = PropertiesService.getScriptProperties().getProperty(
    PATH_PORTAL_EXPORT.STATE_PROPERTY
  );
  return text ? JSON.parse(text) : null;
}

function savePathPortalExportState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    PATH_PORTAL_EXPORT.STATE_PROPERTY,
    JSON.stringify(state)
  );
}

function clearPathPortalExportState_() {
  PropertiesService.getScriptProperties().deleteProperty(
    PATH_PORTAL_EXPORT.STATE_PROPERTY
  );
}

function schedulePathPortalExportContinuation_(delayMs) {
  clearPathPortalExportTriggers_();
  ScriptApp.newTrigger(PATH_PORTAL_EXPORT.CONTINUATION_FUNCTION)
    .timeBased()
    .after(Math.max(1000, Number(delayMs) || 1000))
    .create();
}

function clearPathPortalExportTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (
      trigger.getHandlerFunction() ===
      PATH_PORTAL_EXPORT.CONTINUATION_FUNCTION
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function pathPortalSha256Hex_(value) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  ).map(function(byte) {
    const unsigned = byte < 0 ? byte + 256 : byte;
    return unsigned.toString(16).padStart(2, '0');
  }).join('');
}

function cleanPathPortalText_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
