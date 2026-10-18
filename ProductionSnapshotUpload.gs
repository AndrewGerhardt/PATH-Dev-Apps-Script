/**
 * PATH Records report-snapshot transport.
 *
 * Add this file to the existing Apps Script export project. The existing
 * record-metadata import must run first so every recordId already exists in D1.
 * Store the shared import secret in Script Properties as
 * PATH_RECORD_PRODUCTION_IMPORT_SECRET; do not place it directly in source
 * code. Keep the preview-test secret in a separate Script Property.
 */
const PATH_REPORT_SNAPSHOT_UPLOAD = Object.freeze({
  PORTAL_ORIGIN: 'https://records.campgeigerscouting.org',
  SECRET_PROPERTY: 'PATH_RECORD_PRODUCTION_IMPORT_SECRET',
  CHUNK_SIZE: 40,
  MAX_ATTEMPTS: 3
});

/**
 * Uploads finalized snapshot envelopes in chunks accepted by the portal.
 *
 * Each element must contain:
 * {
 *   recordId,
 *   templateVersion: 1,
 *   finalizedAt: ISO string or Unix timestamp,
 *   snapshot: { schemaVersion: 1, scout, offering, ... }
 * }
 *
 * Create one envelope per badge. If an AppSheet offering teaches two badges
 * during the same period, emit two record IDs and two snapshot envelopes.
 */
function uploadPathReportSnapshots(snapshotEnvelopes, batchId) {
  if (!Array.isArray(snapshotEnvelopes) || !snapshotEnvelopes.length) {
    throw new Error('No report snapshots were supplied.');
  }

  const secret = PropertiesService
    .getScriptProperties()
    .getProperty(PATH_REPORT_SNAPSHOT_UPLOAD.SECRET_PROPERTY);

  if (!secret) {
    throw new Error(
      `Missing Script Property: ${PATH_REPORT_SNAPSHOT_UPLOAD.SECRET_PROPERTY}`
    );
  }

  const summary = {
    processed: 0,
    created: 0,
    versioned: 0,
    restored: 0,
    unchanged: 0
  };

  for (
    let start = 0;
    start < snapshotEnvelopes.length;
    start += PATH_REPORT_SNAPSHOT_UPLOAD.CHUNK_SIZE
  ) {
    const chunk = snapshotEnvelopes.slice(
      start,
      start + PATH_REPORT_SNAPSHOT_UPLOAD.CHUNK_SIZE
    );
    const result = postPathSnapshotChunk_(chunk, batchId, secret);

    summary.processed += Number(result.processed || 0);
    summary.created += Number(result.created || 0);
    summary.versioned += Number(result.versioned || 0);
    summary.restored += Number(result.restored || 0);
    summary.unchanged += Number(result.unchanged || 0);
  }

  return summary;
}

function postPathSnapshotChunk_(snapshots, batchId, secret) {
  const url =
    PATH_REPORT_SNAPSHOT_UPLOAD.PORTAL_ORIGIN +
    '/api/import-snapshots';
  const payload = JSON.stringify({
    batchId: String(batchId || '').trim() || null,
    snapshots
  });
  let lastError = null;

  for (
    let attempt = 1;
    attempt <= PATH_REPORT_SNAPSHOT_UPLOAD.MAX_ATTEMPTS;
    attempt += 1
  ) {
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'X-Import-Secret': secret
        },
        payload,
        muteHttpExceptions: true
      });
      const status = response.getResponseCode();
      const text = response.getContentText();
      let data = {};

      try {
        data = text ? JSON.parse(text) : {};
      } catch (error) {
        const invalidResponseError = new Error(
          `Snapshot import returned HTTP ${status} with an invalid response.`
        );
        invalidResponseError.pathDoNotRetry = status < 500;
        throw invalidResponseError;
      }

      if (status >= 200 && status < 300 && data.success) {
        return data;
      }

      const retryable = status === 429 || status >= 500;
      const message =
        data.message ||
        `Snapshot import failed with HTTP ${status}.`;

      if (!retryable) {
        const requestError = new Error(message);
        requestError.pathDoNotRetry = true;
        throw requestError;
      }

      lastError = new Error(message);

    } catch (error) {
      if (error && error.pathDoNotRetry) {
        throw error;
      }

      lastError = error;
    }

    if (attempt < PATH_REPORT_SNAPSHOT_UPLOAD.MAX_ATTEMPTS) {
      Utilities.sleep(1000 * Math.pow(2, attempt - 1));
    }
  }

  throw lastError || new Error('Snapshot import failed.');
}
