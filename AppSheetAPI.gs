/**
 * AppSheet Entry Point
 * This is the function AppSheet calls
 */
function startExportFromAppSheet(input) {
  try {
    return _startExportFromAppSheet(input);
  } catch (err) {
    Logger.log("TOP LEVEL ERROR:");
    Logger.log(err);
    Logger.log(err.stack);

    return {
      status: "error",
      message: err.message
    };
  }
}


/**
 * Internal Logic (separated for debugging + clarity)
 */
function _startExportFromAppSheet(input) {
  Logger.log("START");

  // ✅ Validate input
  if (!input || !input.session || !input.year) {
    throw new Error("Missing session or year");
  }

  // ✅ Prevent overlapping runs
  if (isRunActive_()) {
    throw new Error("A run is already in progress");
  }

  Logger.log("Before startPathPortalSnapshotExport_");

  const runId = startPathPortalSnapshotExport_(input.session, input.year);

  Logger.log("After startPathPortalSnapshotExport_");

  return {
    status: "started",
    runId: runId
  };
}


/**
 * Read-only test. This validates and counts reports but does not upload them.
 */
function test_startExport() {
  try {
    const result = buildPathPortalExportPackage_(2026, 1).summary;

    Logger.log("RESULT:");
    Logger.log(result);

  } catch (err) {
    Logger.log("TEST ERROR:");
    Logger.log(err);
    Logger.log(err.stack);
    throw err;
  }
}
