/**
 * CAMP GEIGER CURRENT-YEAR INCREMENTAL D1 IMPORT
 *
 * Replaces the prior D1Import.gs.
 *
 * Main function:
 *   buildAndImportCurrentYearChangesToD1()
 *
 * It:
 *   1. Finds the folder named for the current year under ROOT_FOLDER_ID.
 *   2. Scans only that year's PDFs.
 *   3. Reads existing D1 records for that year.
 *   4. Skips records whose stored values are unchanged.
 *   5. Generates SQL only for new or changed records.
 *   6. Uploads that SQL directly through Cloudflare's D1 import API.
 *
 * Existing Script Properties used:
 *   CLOUDFLARE_ACCOUNT_ID
 *   CLOUDFLARE_D1_DATABASE_ID
 *   CLOUDFLARE_D1_API_TOKEN
 *
 * Existing functions expected from the records-index project:
 *   collectPdfEntries_()
 *   parseRecordPdf_()
 *   compareRecords_()
 *
 * Optional existing function:
 *   buildCanonicalTroopIndex_()
 */

const D1_INCREMENTAL_IMPORT_BOOKMARK_PROPERTY =
  "CLOUDFLARE_D1_INCREMENTAL_IMPORT_BOOKMARK";

const D1_INCREMENTAL_POLL_LIMIT_MS =
  4 * 60 * 1000;

const D1_INCREMENTAL_POLL_DELAY_MS =
  2000;

const D1_RECORD_ACTIVITY_PENDING_BATCH_PROPERTY =
  "CLOUDFLARE_D1_RECORD_ACTIVITY_PENDING_BATCH";

const D1_RECORD_ACTIVITY_NOTIFICATION_SECRET_PROPERTY =
  "PATH_RECORDS_IMPORT_NOTIFICATION_SECRET";

const D1_RECORD_ACTIVITY_SITE_ORIGIN_PROPERTY =
  "PATH_RECORDS_SITE_ORIGIN";


/**
 * Normal post-session import.
 *
 * Run this after each camp session.
 */
function buildAndImportCurrentYearChangesToD1() {
  const importYear =
    new Date().getFullYear();

  return buildAndImportYearChangesToD1_(
    importYear
  );
}


/**
 * Use this only when you intentionally need to import
 * a different year's folder.
 */
function buildAndImportSpecificYearChangesToD1() {
  const ui =
    SpreadsheetApp.getUi();

  const response =
    ui.prompt(
      "Import Records to D1",
      "Enter the four-digit year to scan:",
      ui.ButtonSet.OK_CANCEL
    );

  if (
    response.getSelectedButton() !==
    ui.Button.OK
  ) {
    return;
  }

  const year =
    Number(
      response.getResponseText().trim()
    );

  if (
    !Number.isInteger(year) ||
    year < 2000 ||
    year > 2100
  ) {
    throw new Error(
      "Enter a valid four-digit year."
    );
  }

  return buildAndImportYearChangesToD1_(
    year
  );
}


/**
 * Main incremental import implementation.
 */
function buildAndImportYearChangesToD1_(
  importYear
) {
  const startedAt =
    Date.now();

  showD1IncrementalToast_(
    "Scanning the " +
      importYear +
      " records folder...",
    "D1 Import",
    10
  );

  const output =
    buildRecordsDataForYear_(
      importYear
    );

  Logger.log(
    "Parsed records: " +
      output.records.length
  );

  showD1IncrementalToast_(
    "Reading existing " +
      importYear +
      " records from D1...",
    "D1 Import",
    10
  );

  const existingState =
    loadExistingD1State_(
      importYear
    );

  const comparison =
    compareCurrentRecordsToD1_(
      output.records,
      existingState
    );

  Logger.log(
    "New: " +
      comparison.newRecords.length +
      ", updated: " +
      comparison.updatedRecords.length +
      ", unchanged: " +
      comparison.unchangedCount +
      ", duplicate source keys in Drive: " +
      comparison.duplicateSourceKeyCount
  );

  const changedRecords =
    comparison.newRecords.concat(
      comparison.updatedRecords
    );

  if (!changedRecords.length) {
    const noChangesMessage =
      output.records.length +
      " current-year records scanned. " +
      "No new or changed records were found.";

    showD1IncrementalToast_(
      noChangesMessage,
      "D1 Already Current",
      12
    );

    Logger.log(noChangesMessage);

    return {
      complete:
        true,

      importYear:
        importYear,

      scanned:
        output.records.length,

      newRecords:
        0,

      updatedRecords:
        0,

      unchanged:
        comparison.unchangedCount,

      duplicateSourceKeys:
        comparison.duplicateSourceKeyCount,

      skippedFiles:
        output.skippedFiles.length
    };
  }

  showD1IncrementalToast_(
    "Building SQL for " +
      changedRecords.length +
      " changed records...",
    "D1 Import",
    10
  );

  const batchId =
    createIncrementalD1DeterministicId_(
      "batch",
      importYear +
        "|" +
        new Date().toISOString() +
        "|" +
        Utilities.getUuid()
    );

  const sqlResult =
    buildIncrementalD1Sql_(
      changedRecords,
      existingState,
      importYear,
      batchId
    );

  if (!sqlResult.statementCount) {
    const noSqlMessage =
      "Changes were detected, but no database writes were required.";

    Logger.log(noSqlMessage);

    showD1IncrementalToast_(
      noSqlMessage,
      "D1 Already Current",
      12
    );

    return {
      complete:
        true,

      importYear:
        importYear,

      scanned:
        output.records.length,

      newRecords:
        comparison.newRecords.length,

      updatedRecords:
        comparison.updatedRecords.length,

      unchanged:
        comparison.unchangedCount
    };
  }

  const sqlBytes =
    Utilities
      .newBlob(sqlResult.sql)
      .getBytes();

  const sqlSizeMb =
    sqlBytes.length /
    1024 /
    1024;

  Logger.log(
    "Incremental SQL size: " +
      sqlSizeMb.toFixed(2) +
      " MB"
  );

  Logger.log(
    "Expected write statements: " +
      sqlResult.statementCount
  );

  showD1IncrementalToast_(
    "Uploading " +
      comparison.newRecords.length +
      " new and " +
      comparison.updatedRecords.length +
      " updated records...",
    "D1 Import",
    10
  );

  const importResult =
    uploadIncrementalSqlToD1_(
      sqlResult.sql
    );

  savePendingRecordActivityBatch_(
    batchId,
    importYear
  );

  let notificationResult = null;

  if (importResult.complete) {
    try {
      notificationResult =
        sendPendingRecordActivityNotifications();
    } catch (notificationError) {
      notificationResult = {
        complete:
          false,
        error:
          notificationError.message
      };

      Logger.log(
        "Record activity notifications remain pending: " +
        notificationError.message
      );

      showD1IncrementalToast_(
        "The records import completed, but activity emails remain pending. Run sendPendingRecordActivityNotifications after checking the configuration.",
        "Import Complete; Email Pending",
        12
      );
    }
  }

  const elapsedSeconds =
    Math.round(
      (Date.now() - startedAt) /
      1000
    );

  const summary = {
    complete:
      importResult.complete,

    importYear:
      importYear,

    scanned:
      output.records.length,

    newRecords:
      comparison.newRecords.length,

    updatedRecords:
      comparison.updatedRecords.length,

    unchanged:
      comparison.unchangedCount,

    duplicateSourceKeys:
      comparison.duplicateSourceKeyCount,

    skippedFiles:
      output.skippedFiles.length,

    troopWrites:
      sqlResult.troopWriteCount,

    scoutWrites:
      sqlResult.scoutWriteCount,

    scoutKeyWrites:
      sqlResult.scoutKeyWriteCount,

    recordWrites:
      sqlResult.recordWriteCount,

    elapsedSeconds:
      elapsedSeconds,

    batchId:
      batchId,

    notifications:
      notificationResult,

    cloudflare:
      importResult
  };

  if (importResult.complete) {
    const completeMessage =
      comparison.newRecords.length +
      " new and " +
      comparison.updatedRecords.length +
      " updated records imported. " +
      comparison.unchangedCount +
      " unchanged records were skipped.";

    Logger.log(
      completeMessage
    );

    showD1IncrementalToast_(
      completeMessage,
      "D1 Import Complete",
      12
    );
  } else {
    showD1IncrementalToast_(
      "Cloudflare accepted the import and is still processing it. " +
      "Run checkCurrentYearD1ImportStatus.",
      "D1 Import Running",
      12
    );
  }

  return summary;
}


/**
 * Checks an import that continued after the first Apps Script run.
 */
function checkCurrentYearD1ImportStatus() {
  const properties =
    PropertiesService
      .getScriptProperties();

  const bookmark =
    properties.getProperty(
      D1_INCREMENTAL_IMPORT_BOOKMARK_PROPERTY
    );

  if (!bookmark) {
    throw new Error(
      "There is no saved incremental D1 import to check."
    );
  }

  const result =
    pollIncrementalD1Import_(
      bookmark,
      D1_INCREMENTAL_POLL_LIMIT_MS
    );

  let notificationResult = null;

  if (result.complete) {
    try {
      notificationResult =
        sendPendingRecordActivityNotifications();
    } catch (notificationError) {
      notificationResult = {
        complete:
          false,
        error:
          notificationError.message
      };

      Logger.log(
        "Record activity notifications remain pending: " +
        notificationError.message
      );
    }
  }

  showD1IncrementalToast_(
    result.complete
      ? (
          notificationResult &&
          notificationResult.complete === false
            ? "The D1 import is complete, but record activity emails remain pending."
            : "The Cloudflare D1 import is complete. Record activity notifications were processed."
        )
      : "The D1 import is still running. Check again shortly.",
    result.complete
      ? "D1 Import Complete"
      : "D1 Import Running",
    10
  );

  result.notifications =
    notificationResult;

  return result;
}


/**
 * Finds the year folder immediately under ROOT_FOLDER_ID
 * and scans only that folder.
 */
function buildRecordsDataForYear_(
  importYear
) {
  if (
    typeof ROOT_FOLDER_ID ===
    "undefined"
  ) {
    throw new Error(
      "ROOT_FOLDER_ID is not defined in this Apps Script project."
    );
  }

  const rootFolder =
    DriveApp.getFolderById(
      ROOT_FOLDER_ID
    );

  const yearFolders =
    rootFolder.getFoldersByName(
      String(importYear)
    );

  if (!yearFolders.hasNext()) {
    throw new Error(
      "A folder named " +
      importYear +
      " was not found directly under the records root folder."
    );
  }

  const yearFolder =
    yearFolders.next();

  if (yearFolders.hasNext()) {
    throw new Error(
      "More than one folder named " +
      importYear +
      " exists under the records root folder."
    );
  }

  if (
    typeof filesProcessed !==
    "undefined"
  ) {
    filesProcessed = 0;
  }

  const pdfEntries = [];
  const skippedFiles = [];

  collectPdfEntries_(
    yearFolder,
    pdfEntries,
    [String(importYear)]
  );

  Logger.log(
    "PDFs found in " +
      importYear +
      ": " +
      pdfEntries.length
  );

  if (
    typeof buildCanonicalTroopIndex_ ===
    "function"
  ) {
    buildCanonicalTroopIndex_(
      pdfEntries
    );
  }

  const records = [];

  for (
    let index = 0;
    index < pdfEntries.length;
    index += 1
  ) {
    const entry =
      pdfEntries[index];

    const record =
      parseRecordPdf_(
        entry.file,
        entry.folderPath
      );

    if (record) {
      records.push(
        normalizeIncrementalD1Record_(
          record,
          index
        )
      );
    } else {
      skippedFiles.push({
        path:
          entry.folderPath.join(" / ") +
          " / " +
          entry.file.getName(),

        reason:
          "Filename did not match a supported format."
      });
    }

    if (
      index > 0 &&
      index % 1000 === 0
    ) {
      Logger.log(
        "Parsed " +
          index +
          " of " +
          pdfEntries.length +
          " PDFs."
      );
    }
  }

  if (
    typeof compareRecords_ ===
    "function"
  ) {
    records.sort(
      compareRecords_
    );
  }

  return {
    generatedAt:
      new Date().toISOString(),

    year:
      importYear,

    records:
      records,

    skippedFiles:
      skippedFiles
  };
}


/**
 * Reads the existing database state using Cloudflare's D1 query API.
 * These are read queries and do not consume D1 row writes.
 */
function loadExistingD1State_(
  importYear
) {
  const recordsRows =
    queryD1AllRows_(
      [
        "SELECT",
        "  r.record_id,",
        "  r.scout_id,",
        "  r.source_key,",
        "  r.original_scout_key,",
        "  r.troop_at_time,",
        "  rto.troop_key AS troop_override_key,",
        "  COALESCE(",
        "    override_merge.target_troop_key,",
        "    rto.troop_key,",
        "    imported_merge.target_troop_key,",
        "    r.troop_at_time",
        "  ) AS effective_troop_key,",
        "  r.class_name,",
        "  r.session,",
        "  r.record_year,",
        "  r.legacy_time,",
        "  r.legacy_code,",
        "  r.source_format,",
        "  r.file_name,",
        "  r.file_url",
        "FROM records AS r",
        "LEFT JOIN record_troop_overrides AS rto",
        "  ON rto.record_id = r.record_id",
        "LEFT JOIN troop_merges AS override_merge",
        "  ON override_merge.source_troop_key = rto.troop_key",
        "LEFT JOIN troop_merges AS imported_merge",
        "  ON imported_merge.source_troop_key = r.troop_at_time",
        "WHERE r.record_year = ?"
      ].join("\n"),
      [importYear]
    );

  const scoutRows =
    queryD1AllRows_(
      [
        "SELECT",
        "  sk.scout_key,",
        "  sk.scout_id,",
        "  sk.troop_key,",
        "  sk.first_seen_year,",
        "  sk.last_seen_year,",
        "  s.first_name,",
        "  s.last_name,",
        "  s.current_troop_key,",
        "  s.active",
        "FROM scout_keys sk",
        "JOIN scouts s",
        "  ON s.scout_id = sk.scout_id"
      ].join("\n"),
      []
    );

  const troopRows =
    queryD1AllRows_(
      [
        "SELECT troop_key, active",
        "FROM troops"
      ].join("\n"),
      []
    );

  const troopMergeRows =
    queryD1AllRows_(
      [
        "SELECT source_troop_key, target_troop_key",
        "FROM troop_merges"
      ].join("\n"),
      []
    );

  const recordsBySourceKey = {};
  const scoutsByScoutKey = {};
  const troopsByKey = {};
  const troopMergesBySource = {};

  recordsRows.forEach(
    function(row) {
      const sourceKey =
        cleanIncrementalD1Text_(
          row.source_key
        )
          .toLowerCase();

      if (sourceKey) {
        recordsBySourceKey[sourceKey] =
          normalizeExistingRecordRow_(row);
      }
    }
  );

  scoutRows.forEach(
    function(row) {
      const scoutKey =
        cleanIncrementalD1Text_(
          row.scout_key
        )
          .toLowerCase();

      if (scoutKey) {
        scoutsByScoutKey[scoutKey] = {
          scoutKey:
            scoutKey,

          scoutId:
            cleanIncrementalD1Text_(
              row.scout_id
            ),

          troopKey:
            cleanIncrementalD1Text_(
              row.troop_key
            ),

          firstSeenYear:
            nullableIncrementalD1Number_(
              row.first_seen_year
            ),

          lastSeenYear:
            nullableIncrementalD1Number_(
              row.last_seen_year
            ),

          firstName:
            cleanIncrementalD1Text_(
              row.first_name
            ),

          lastName:
            cleanIncrementalD1Text_(
              row.last_name
            ),

          currentTroop:
            cleanIncrementalD1Text_(
              row.current_troop_key
            ),

          active:
            Number(row.active || 0)
        };
      }
    }
  );

  troopRows.forEach(
    function(row) {
      const troopKey =
        cleanIncrementalD1Text_(
          row.troop_key
        );

      if (troopKey) {
        troopsByKey[troopKey] = {
          troopKey:
            troopKey,

          active:
            Number(row.active || 0)
        };
      }
    }
  );

  troopMergeRows.forEach(
    function(row) {
      const source =
        cleanIncrementalD1Text_(
          row.source_troop_key
        );

      const target =
        cleanIncrementalD1Text_(
          row.target_troop_key
        );

      if (source && target) {
        troopMergesBySource[source] = target;
      }
    }
  );

  return {
    recordsBySourceKey:
      recordsBySourceKey,

    scoutsByScoutKey:
      scoutsByScoutKey,

    troopsByKey:
      troopsByKey,

    troopMergesBySource:
      troopMergesBySource
  };
}


/**
 * Keeps one Drive record per source key and classifies it.
 */
function compareCurrentRecordsToD1_(
  currentRecords,
  existingState
) {
  const uniqueCurrent = {};
  let duplicateSourceKeyCount = 0;

  currentRecords.forEach(
    function(record) {
      if (
        uniqueCurrent[record.sourceKey]
      ) {
        duplicateSourceKeyCount += 1;
      }

      /*
       * Preserve the last matching item, consistent with the
       * previous full-import behavior.
       */
      uniqueCurrent[record.sourceKey] =
        record;
    }
  );

  const newRecords = [];
  const updatedRecords = [];
  let unchangedCount = 0;

  Object.keys(uniqueCurrent)
    .forEach(
      function(sourceKey) {
        const current =
          uniqueCurrent[sourceKey];

        const existing =
          existingState
            .recordsBySourceKey[
              sourceKey
            ];

        if (!existing) {
          newRecords.push(current);
          return;
        }

        if (
          incrementalRecordMatches_(
            current,
            existing
          )
        ) {
          unchangedCount += 1;
        } else {
          updatedRecords.push(current);
        }
      }
    );

  return {
    newRecords:
      newRecords,

    updatedRecords:
      updatedRecords,

    unchangedCount:
      unchangedCount,

    duplicateSourceKeyCount:
      duplicateSourceKeyCount
  };
}


function incrementalRecordMatches_(
  current,
  existing
) {
  return (
    normalizeComparisonText_(
      current.scoutKey
    ) ===
      normalizeComparisonText_(
        existing.scoutKey
      ) &&

    normalizeComparisonText_(
      current.troop
    ) ===
      normalizeComparisonText_(
        existing.troop
      ) &&

    normalizeComparisonText_(
      current.className
    ) ===
      normalizeComparisonText_(
        existing.className
      ) &&

    normalizeComparisonNumber_(
      current.session
    ) ===
      normalizeComparisonNumber_(
        existing.session
      ) &&

    Number(current.year) ===
      Number(existing.year) &&

    normalizeComparisonNullableText_(
      current.legacyTime
    ) ===
      normalizeComparisonNullableText_(
        existing.legacyTime
      ) &&

    normalizeComparisonNullableText_(
      current.legacyCode
    ) ===
      normalizeComparisonNullableText_(
        existing.legacyCode
      ) &&

    normalizeComparisonNullableText_(
      current.sourceFormat
    ) ===
      normalizeComparisonNullableText_(
        existing.sourceFormat
      ) &&

    normalizeComparisonText_(
      current.fileName
    ) ===
      normalizeComparisonText_(
        existing.fileName
      ) &&

    normalizeComparisonNullableText_(
      current.fileUrl
    ) ===
      normalizeComparisonNullableText_(
        existing.fileUrl
      )
  );
}


/**
 * Creates SQL only for records that are new or changed.
 *
 * Troops, scouts, and scout_keys are also written only when
 * their stored state actually needs to change.
 */
function buildIncrementalD1Sql_(
  changedRecords,
  existingState,
  importYear,
  batchId
) {
  const unixTime =
    Math.floor(
      Date.now() /
      1000
    );

  const desiredTroops = {};
  const desiredScouts = {};

  changedRecords.forEach(
    function(record) {
      desiredTroops[record.troop] =
        record.troop;

      let scout =
        desiredScouts[record.scoutKey];

      const existingScout =
        existingState
          .scoutsByScoutKey[
            record.scoutKey
          ];

      if (!scout) {
        scout = {
          scoutKey:
            record.scoutKey,

          scoutId:
            existingScout &&
            existingScout.scoutId
              ? existingScout.scoutId
              : createIncrementalD1DeterministicId_(
                  "scout",
                  record.scoutKey
                ),

          firstName:
            record.firstName,

          lastName:
            record.lastName,

          currentTroop:
            record.troop,

          firstYear:
            record.year,

          lastYear:
            record.year
        };

        desiredScouts[record.scoutKey] =
          scout;
      } else {
        scout.firstYear =
          Math.min(
            scout.firstYear,
            record.year
          );

        if (
          record.year >=
          scout.lastYear
        ) {
          scout.lastYear =
            record.year;

          scout.firstName =
            record.firstName;

          scout.lastName =
            record.lastName;

          scout.currentTroop =
            record.troop;
        }
      }
    }
  );

  const activityRows =
    buildRecordActivityRows_(
      changedRecords,
      existingState,
      batchId,
      unixTime
    );

  const activitySummary =
    summarizeRecordActivityRows_(
      activityRows
    );

  const lines = [
    "-- Camp Geiger incremental D1 import",
    "-- Import year: " +
      importYear,
    "-- Generated: " +
      new Date().toISOString(),
    "-- Changed records: " +
      changedRecords.length,
    "",
    "PRAGMA defer_foreign_keys = true;",
    ""
  ];

  lines.push(
    "INSERT INTO record_import_batches (" +
    "batch_id, import_year, status, started_at, " +
    "new_record_count, updated_record_count, moved_record_count, " +
    "affected_scout_count, affected_unit_count, created_at, updated_at" +
    ") VALUES (" +
    incrementalD1SqlValue_(batchId) +
    ", " + importYear +
    ", 'processing', " + unixTime +
    ", " + activitySummary.addedCount +
    ", " + activitySummary.updatedCount +
    ", " + activitySummary.movedCount +
    ", " + activitySummary.affectedScoutCount +
    ", " + activitySummary.affectedUnitCount +
    ", " + unixTime +
    ", " + unixTime +
    ") ON CONFLICT(batch_id) DO NOTHING;"
  );

  let troopWriteCount = 0;
  let scoutWriteCount = 0;
  let scoutKeyWriteCount = 0;
  let recordWriteCount = 0;

  Object.keys(desiredTroops)
    .forEach(
      function(troopKey) {
        const existingTroop =
          existingState
            .troopsByKey[
              troopKey
            ];

        if (
          existingTroop &&
          existingTroop.active === 1
        ) {
          return;
        }

        if (existingTroop) {
          lines.push(
            "UPDATE troops " +
            "SET active = 1, " +
            "updated_at = " +
            unixTime +
            " WHERE troop_key = " +
            incrementalD1SqlValue_(
              troopKey
            ) +
            " AND active IS NOT 1;"
          );
        } else {
          lines.push(
            "INSERT INTO troops (" +
            "troop_key, active, created_at, updated_at" +
            ") VALUES (" +
            incrementalD1SqlValue_(
              troopKey
            ) +
            ", 1, " +
            unixTime +
            ", " +
            unixTime +
            ");"
          );
        }

        troopWriteCount += 1;
      }
    );

  Object.keys(desiredScouts)
    .forEach(
      function(scoutKey) {
        const desired =
          desiredScouts[scoutKey];

        const existing =
          existingState
            .scoutsByScoutKey[
              scoutKey
            ];

        if (!existing) {
          lines.push(
            "INSERT INTO scouts (" +
            "scout_id, first_name, last_name, current_troop_key, " +
            "active, created_at, updated_at" +
            ") VALUES (" +
            incrementalD1SqlValue_(
              desired.scoutId
            ) +
            ", " +
            incrementalD1SqlValue_(
              desired.firstName
            ) +
            ", " +
            incrementalD1SqlValue_(
              desired.lastName
            ) +
            ", " +
            incrementalD1SqlValue_(
              desired.currentTroop
            ) +
            ", 1, " +
            unixTime +
            ", " +
            unixTime +
            ");"
          );

          lines.push(
            "INSERT INTO scout_keys (" +
            "scout_key, scout_id, troop_key, first_seen_year, " +
            "last_seen_year, created_at, updated_at" +
            ") VALUES (" +
            incrementalD1SqlValue_(
              desired.scoutKey
            ) +
            ", " +
            incrementalD1SqlValue_(
              desired.scoutId
            ) +
            ", " +
            incrementalD1SqlValue_(
              desired.currentTroop
            ) +
            ", " +
            desired.firstYear +
            ", " +
            desired.lastYear +
            ", " +
            unixTime +
            ", " +
            unixTime +
            ");"
          );

          scoutWriteCount += 1;
          scoutKeyWriteCount += 1;
          return;
        }

        const scoutNeedsUpdate =
          normalizeComparisonText_(
            desired.firstName
          ) !==
            normalizeComparisonText_(
              existing.firstName
            ) ||

          normalizeComparisonText_(
            desired.lastName
          ) !==
            normalizeComparisonText_(
              existing.lastName
            ) ||

          normalizeComparisonText_(
            desired.currentTroop
          ) !==
            normalizeComparisonText_(
              existing.currentTroop
            ) ||

          existing.active !== 1;

        if (scoutNeedsUpdate) {
          lines.push(
            "UPDATE scouts SET " +
            "first_name = " +
            incrementalD1SqlValue_(
              desired.firstName
            ) +
            ", last_name = " +
            incrementalD1SqlValue_(
              desired.lastName
            ) +
            ", current_troop_key = " +
            incrementalD1SqlValue_(
              desired.currentTroop
            ) +
            ", active = 1, " +
            "updated_at = " +
            unixTime +
            " WHERE scout_id = " +
            incrementalD1SqlValue_(
              existing.scoutId
            ) +
            ";"
          );

          scoutWriteCount += 1;
        }

        const desiredFirstYear =
          existing.firstSeenYear === null
            ? desired.firstYear
            : Math.min(
                existing.firstSeenYear,
                desired.firstYear
              );

        const desiredLastYear =
          existing.lastSeenYear === null
            ? desired.lastYear
            : Math.max(
                existing.lastSeenYear,
                desired.lastYear
              );

        const scoutKeyNeedsUpdate =
          normalizeComparisonText_(
            desired.currentTroop
          ) !==
            normalizeComparisonText_(
              existing.troopKey
            ) ||

          desiredFirstYear !==
            existing.firstSeenYear ||

          desiredLastYear !==
            existing.lastSeenYear;

        if (scoutKeyNeedsUpdate) {
          lines.push(
            "UPDATE scout_keys SET " +
            "troop_key = " +
            incrementalD1SqlValue_(
              desired.currentTroop
            ) +
            ", first_seen_year = " +
            desiredFirstYear +
            ", last_seen_year = " +
            desiredLastYear +
            ", updated_at = " +
            unixTime +
            " WHERE scout_key = " +
            incrementalD1SqlValue_(
              desired.scoutKey
            ) +
            ";"
          );

          scoutKeyWriteCount += 1;
        }
      }
    );

  changedRecords.forEach(
    function(record) {
      const existingScout =
        existingState
          .scoutsByScoutKey[
            record.scoutKey
          ];

      const desiredScout =
        desiredScouts[
          record.scoutKey
        ];

      const scoutId =
        existingScout &&
        existingScout.scoutId
          ? existingScout.scoutId
          : desiredScout.scoutId;

      const existingRecord =
        existingState
          .recordsBySourceKey[
            record.sourceKey
          ];

      if (existingRecord) {
        lines.push(
          "UPDATE records SET " +
          "scout_id = " +
          incrementalD1SqlValue_(
            scoutId
          ) +
          ", original_scout_key = " +
          incrementalD1SqlValue_(
            record.scoutKey
          ) +
          ", troop_at_time = " +
          incrementalD1SqlValue_(
            record.troop
          ) +
          ", class_name = " +
          incrementalD1SqlValue_(
            record.className
          ) +
          ", session = " +
          incrementalD1SqlNumber_(
            record.session
          ) +
          ", record_year = " +
          record.year +
          ", legacy_time = " +
          incrementalD1SqlValue_(
            record.legacyTime
          ) +
          ", legacy_code = " +
          incrementalD1SqlValue_(
            record.legacyCode
          ) +
          ", source_format = " +
          incrementalD1SqlValue_(
            record.sourceFormat
          ) +
          ", file_name = " +
          incrementalD1SqlValue_(
            record.fileName
          ) +
          ", file_url = " +
          incrementalD1SqlValue_(
            record.fileUrl
          ) +
          ", updated_at = " +
          unixTime +
          " WHERE source_key = " +
          incrementalD1SqlValue_(
            record.sourceKey
          ) +
          ";"
        );
      } else {
        lines.push(
          "INSERT INTO records (" +
          "record_id, scout_id, original_scout_key, troop_at_time, " +
          "class_name, session, record_year, legacy_time, legacy_code, " +
          "source_format, file_name, file_url, source_key, " +
          "created_at, updated_at" +
          ") VALUES (" +
          incrementalD1SqlValue_(
            createIncrementalD1DeterministicId_(
              "record",
              record.sourceKey
            )
          ) +
          ", " +
          incrementalD1SqlValue_(
            scoutId
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.scoutKey
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.troop
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.className
          ) +
          ", " +
          incrementalD1SqlNumber_(
            record.session
          ) +
          ", " +
          record.year +
          ", " +
          incrementalD1SqlValue_(
            record.legacyTime
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.legacyCode
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.sourceFormat
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.fileName
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.fileUrl
          ) +
          ", " +
          incrementalD1SqlValue_(
            record.sourceKey
          ) +
          ", " +
          unixTime +
          ", " +
          unixTime +
          ");"
        );
      }

      recordWriteCount += 1;
    }
  );

  activityRows.forEach(
    function(activity) {
      lines.push(
        "INSERT INTO record_activity (" +
        "activity_id, batch_id, record_id, scout_id, troop_key, " +
        "source_troop_key, previous_troop_key, class_name, session, " +
        "record_year, activity_type, is_new_scout, file_name, created_at" +
        ") VALUES (" +
        incrementalD1SqlValue_(activity.activityId) +
        ", " + incrementalD1SqlValue_(activity.batchId) +
        ", " + incrementalD1SqlValue_(activity.recordId) +
        ", " + incrementalD1SqlValue_(activity.scoutId) +
        ", " + incrementalD1SqlValue_(activity.troopKey) +
        ", " + incrementalD1SqlValue_(activity.sourceTroopKey) +
        ", " + incrementalD1SqlValue_(activity.previousTroopKey) +
        ", " + incrementalD1SqlValue_(activity.className) +
        ", " + incrementalD1SqlNumber_(activity.session) +
        ", " + activity.recordYear +
        ", " + incrementalD1SqlValue_(activity.activityType) +
        ", " + (activity.isNewScout ? 1 : 0) +
        ", " + incrementalD1SqlValue_(activity.fileName) +
        ", " + unixTime +
        ") ON CONFLICT(batch_id, record_id) DO NOTHING;"
      );
    }
  );

  lines.push(
    "UPDATE record_import_batches SET " +
    "status = 'complete', " +
    "completed_at = " + unixTime +
    ", updated_at = " + unixTime +
    " WHERE batch_id = " +
    incrementalD1SqlValue_(batchId) +
    ";"
  );

  const statementCount =
    troopWriteCount +
    scoutWriteCount +
    scoutKeyWriteCount +
    recordWriteCount +
    activityRows.length +
    2;

  return {
    sql:
      lines.join("\n"),

    statementCount:
      statementCount,

    troopWriteCount:
      troopWriteCount,

    scoutWriteCount:
      scoutWriteCount,

    scoutKeyWriteCount:
      scoutKeyWriteCount,

    recordWriteCount:
      recordWriteCount,

    batchId:
      batchId,

    activityWriteCount:
      activityRows.length
  };
}


function buildRecordActivityRows_(
  changedRecords,
  existingState,
  batchId,
  unixTime
) {
  return changedRecords.map(
    function(record) {
      const existingRecord =
        existingState
          .recordsBySourceKey[
            record.sourceKey
          ];

      const existingScout =
        existingState
          .scoutsByScoutKey[
            record.scoutKey
          ];

      const desiredScoutId =
        existingScout && existingScout.scoutId
          ? existingScout.scoutId
          : createIncrementalD1DeterministicId_(
              "scout",
              record.scoutKey
            );

      const recordId =
        existingRecord && existingRecord.recordId
          ? existingRecord.recordId
          : createIncrementalD1DeterministicId_(
              "record",
              record.sourceKey
            );

      const previousTroop =
        existingRecord
          ? cleanIncrementalD1Text_(
              existingRecord.effectiveTroop ||
              existingRecord.troop
            )
          : "";

      let targetTroop;

      if (
        existingRecord &&
        existingRecord.troopOverride
      ) {
        targetTroop =
          resolveIncrementalTroopMerge_(
            existingRecord.troopOverride,
            existingState.troopMergesBySource
          );
      } else {
        targetTroop =
          resolveIncrementalTroopMerge_(
            record.troop,
            existingState.troopMergesBySource
          );
      }

      let activityType = "added";

      if (existingRecord) {
        activityType =
          normalizeComparisonText_(previousTroop) !==
          normalizeComparisonText_(targetTroop)
            ? "moved"
            : "updated";
      }

      return {
        activityId:
          createIncrementalD1DeterministicId_(
            "activity",
            batchId + "|" + recordId
          ),

        batchId:
          batchId,

        recordId:
          recordId,

        scoutId:
          desiredScoutId,

        troopKey:
          targetTroop,

        sourceTroopKey:
          record.troop,

        previousTroopKey:
          previousTroop || null,

        className:
          record.className,

        session:
          record.session,

        recordYear:
          record.year,

        activityType:
          activityType,

        isNewScout:
          !existingScout,

        fileName:
          record.fileName,

        createdAt:
          unixTime
      };
    }
  );
}


function summarizeRecordActivityRows_(rows) {
  const scoutIds = {};
  const troopKeys = {};
  let addedCount = 0;
  let updatedCount = 0;
  let movedCount = 0;

  rows.forEach(
    function(row) {
      scoutIds[row.scoutId] = true;
      troopKeys[row.troopKey] = true;

      if (row.activityType === "added") {
        addedCount += 1;
      } else if (
        row.activityType === "moved"
      ) {
        movedCount += 1;
      } else {
        updatedCount += 1;
      }
    }
  );

  return {
    addedCount:
      addedCount,

    updatedCount:
      updatedCount,

    movedCount:
      movedCount,

    affectedScoutCount:
      Object.keys(scoutIds).length,

    affectedUnitCount:
      Object.keys(troopKeys).length
  };
}


function resolveIncrementalTroopMerge_(
  troopKey,
  mergeMap
) {
  let current =
    cleanIncrementalD1Text_(troopKey);

  const visited = {};

  while (
    current &&
    mergeMap &&
    mergeMap[current] &&
    !visited[current]
  ) {
    visited[current] = true;
    current =
      cleanIncrementalD1Text_(
        mergeMap[current]
      );
  }

  return current;
}


function savePendingRecordActivityBatch_(
  batchId,
  importYear
) {
  PropertiesService
    .getScriptProperties()
    .setProperty(
      D1_RECORD_ACTIVITY_PENDING_BATCH_PROPERTY,
      JSON.stringify({
        batchId:
          batchId,
        importYear:
          importYear
      })
    );
}


function sendPendingRecordActivityNotifications() {
  const properties =
    PropertiesService
      .getScriptProperties();

  const pendingText =
    properties.getProperty(
      D1_RECORD_ACTIVITY_PENDING_BATCH_PROPERTY
    );

  if (!pendingText) {
    return {
      complete:
        true,
      message:
        "There is no pending record activity batch."
    };
  }

  let pending;

  try {
    pending = JSON.parse(pendingText);
  } catch (error) {
    throw new Error(
      "The pending record activity batch is not valid JSON."
    );
  }

  const secret =
    cleanIncrementalD1Text_(
      properties.getProperty(
        D1_RECORD_ACTIVITY_NOTIFICATION_SECRET_PROPERTY
      )
    );

  if (!secret) {
    throw new Error(
      "Set the Apps Script property " +
      D1_RECORD_ACTIVITY_NOTIFICATION_SECRET_PROPERTY +
      " before sending record activity emails."
    );
  }

  const siteOrigin =
    cleanIncrementalD1Text_(
      properties.getProperty(
        D1_RECORD_ACTIVITY_SITE_ORIGIN_PROPERTY
      )
    ) ||
    "https://records.path-systems.org";

  const endpoint =
    siteOrigin.replace(/\/$/, "") +
    "/api/import-batches/" +
    encodeURIComponent(pending.batchId) +
    "/complete";

  let totalSent = 0;
  let totalFailed = 0;
  let permanentFailures = 0;
  let result = null;

  for (
    let page = 0;
    page < 50;
    page += 1
  ) {
    const response =
      UrlFetchApp.fetch(
        endpoint,
        {
          method:
            "post",
          headers: {
            "X-Import-Secret":
              secret,
            "X-Requested-With":
              "PATH-Records"
          },
          muteHttpExceptions:
            true
        }
      );

    const code =
      response.getResponseCode();

    const text =
      response.getContentText();

    let payload;

    try {
      payload = JSON.parse(text);
    } catch (error) {
      throw new Error(
        "The portal returned HTTP " +
        code +
        " instead of JSON while sending record activity emails."
      );
    }

    if (
      code < 200 ||
      code >= 300 ||
      !payload.success
    ) {
      throw new Error(
        payload.message ||
        "The portal returned HTTP " + code +
        " while sending record activity emails."
      );
    }

    result = payload;
    totalSent +=
      Number(payload.sent || 0);
    totalFailed +=
      Number(payload.failed || 0);
    permanentFailures =
      Number(payload.permanentFailures || 0);

    if (payload.complete) {
      properties.deleteProperty(
        D1_RECORD_ACTIVITY_PENDING_BATCH_PROPERTY
      );

      const message =
        totalSent +
        " record activity emails sent." +
        (permanentFailures
          ? " " + permanentFailures +
            " recipients could not be reached after three attempts."
          : "");

      Logger.log(message);

      showD1IncrementalToast_(
        message,
        "Record Activity Emails",
        12
      );

      return {
        complete:
          true,
        batchId:
          pending.batchId,
        sent:
          totalSent,
        failedAttempts:
          totalFailed,
        permanentFailures:
          permanentFailures
      };
    }

    Utilities.sleep(500);
  }

  throw new Error(
    "Record activity emails are still pending after 50 portal requests. Run sendPendingRecordActivityNotifications again."
  );
}


/**
 * D1 REST query.
 */
function queryD1AllRows_(
  sql,
  params
) {
  const config =
    getIncrementalD1Config_();

  const url =
    "https://api.cloudflare.com/client/v4/accounts/" +
    encodeURIComponent(
      config.accountId
    ) +
    "/d1/database/" +
    encodeURIComponent(
      config.databaseId
    ) +
    "/query";

  const response =
    UrlFetchApp.fetch(
      url,
      {
        method:
          "post",

        contentType:
          "application/json",

        headers: {
          Authorization:
            "Bearer " +
            config.apiToken
        },

        payload:
          JSON.stringify({
            sql:
              sql,

            params:
              params || []
          }),

        muteHttpExceptions:
          true
      }
    );

  const code =
    response.getResponseCode();

  const text =
    response.getContentText();

  let data;

  try {
    data =
      JSON.parse(text);
  } catch (error) {
    throw new Error(
      "Cloudflare returned invalid JSON for a D1 query. HTTP " +
      code +
      ": " +
      text
    );
  }

  if (
    code < 200 ||
    code >= 300 ||
    data.success === false
  ) {
    throw new Error(
      buildIncrementalCloudflareError_(
        "D1 query failed",
        code,
        data,
        text
      )
    );
  }

  const resultList =
    Array.isArray(data.result)
      ? data.result
      : [data.result];

  const rows = [];

  resultList.forEach(
    function(result) {
      if (
        result &&
        Array.isArray(result.results)
      ) {
        result.results.forEach(
          function(row) {
            rows.push(row);
          }
        );
      }
    }
  );

  return rows;
}


/**
 * Cloudflare D1 import:
 * init -> signed upload -> ingest -> poll.
 */
function uploadIncrementalSqlToD1_(
  sql
) {
  const config =
    getIncrementalD1Config_();

  const importUrl =
    "https://api.cloudflare.com/client/v4/accounts/" +
    encodeURIComponent(
      config.accountId
    ) +
    "/d1/database/" +
    encodeURIComponent(
      config.databaseId
    ) +
    "/import";

  const sqlBlob =
    Utilities.newBlob(
      sql,
      "application/sql",
      "records-incremental-import.sql"
    );

  const sqlBytes =
    sqlBlob.getBytes();

  const etag =
    incrementalMd5Hex_(
      sqlBytes
    );

  const headers = {
    Authorization:
      "Bearer " +
      config.apiToken
  };

  const initData =
    incrementalCloudflareJsonRequest_(
      importUrl,
      headers,
      {
        action:
          "init",

        etag:
          etag
      }
    );

  const initResult =
    initData.result || {};

  const uploadUrl =
    initResult.upload_url;

  const filename =
    initResult.filename;

  if (
    !uploadUrl ||
    !filename
  ) {
    throw new Error(
      "Cloudflare did not return an upload URL and filename."
    );
  }

  const uploadResponse =
    UrlFetchApp.fetch(
      uploadUrl,
      {
        method:
          "put",

        contentType:
          "application/sql",

        payload:
          sqlBytes,

        muteHttpExceptions:
          true,

        followRedirects:
          true
      }
    );

  const uploadCode =
    uploadResponse.getResponseCode();

  if (
    uploadCode < 200 ||
    uploadCode >= 300
  ) {
    throw new Error(
      "Cloudflare SQL upload failed. HTTP " +
      uploadCode +
      ": " +
      uploadResponse.getContentText()
    );
  }

  const ingestData =
    incrementalCloudflareJsonRequest_(
      importUrl,
      headers,
      {
        action:
          "ingest",

        etag:
          etag,

        filename:
          filename
      }
    );

  const ingestResult =
    ingestData.result || {};

  const bookmark =
    ingestResult.at_bookmark;

  if (!bookmark) {
    if (
      ingestResult.success === true ||
      ingestResult.status ===
        "complete"
    ) {
      return {
        complete:
          true,

        result:
          ingestResult
      };
    }

    throw new Error(
      "Cloudflare started the import but did not return a bookmark."
    );
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      D1_INCREMENTAL_IMPORT_BOOKMARK_PROPERTY,
      bookmark
    );

  return pollIncrementalD1Import_(
    bookmark,
    D1_INCREMENTAL_POLL_LIMIT_MS
  );
}


function pollIncrementalD1Import_(
  bookmark,
  maxWaitMs
) {
  const config =
    getIncrementalD1Config_();

  const importUrl =
    "https://api.cloudflare.com/client/v4/accounts/" +
    encodeURIComponent(
      config.accountId
    ) +
    "/d1/database/" +
    encodeURIComponent(
      config.databaseId
    ) +
    "/import";

  const headers = {
    Authorization:
      "Bearer " +
      config.apiToken
  };

  const startedAt =
    Date.now();

  let latestResult =
    null;

  while (
    Date.now() - startedAt <
    maxWaitMs
  ) {
    const pollData =
      incrementalCloudflareJsonRequest_(
        importUrl,
        headers,
        {
          action:
            "poll",

          current_bookmark:
            bookmark
        }
      );

    latestResult =
      pollData.result || {};

    Logger.log(
      "D1 import poll: " +
      JSON.stringify(
        latestResult
      )
    );

    if (
      latestResult.status ===
        "error" ||
      latestResult.error
    ) {
      throw new Error(
        "Cloudflare D1 import failed: " +
        (
          latestResult.error ||
          "Unknown import error."
        )
      );
    }

    if (
      latestResult.success ===
        true ||
      latestResult.status ===
        "complete" ||
      latestResult.result
    ) {
      PropertiesService
        .getScriptProperties()
        .deleteProperty(
          D1_INCREMENTAL_IMPORT_BOOKMARK_PROPERTY
        );

      return {
        complete:
          true,

        result:
          latestResult
      };
    }

    Utilities.sleep(
      D1_INCREMENTAL_POLL_DELAY_MS
    );
  }

  return {
    complete:
      false,

    result:
      latestResult,

    bookmark:
      bookmark
  };
}


function incrementalCloudflareJsonRequest_(
  url,
  headers,
  body
) {
  const response =
    UrlFetchApp.fetch(
      url,
      {
        method:
          "post",

        contentType:
          "application/json",

        headers:
          headers,

        payload:
          JSON.stringify(
            body
          ),

        muteHttpExceptions:
          true
      }
    );

  const code =
    response.getResponseCode();

  const text =
    response.getContentText();

  let data;

  try {
    data =
      JSON.parse(text);
  } catch (error) {
    throw new Error(
      "Cloudflare returned invalid JSON. HTTP " +
      code +
      ": " +
      text
    );
  }

  if (
    code < 200 ||
    code >= 300 ||
    data.success === false
  ) {
    throw new Error(
      buildIncrementalCloudflareError_(
        "Cloudflare request failed",
        code,
        data,
        text
      )
    );
  }

  return data;
}


function buildIncrementalCloudflareError_(
  prefix,
  code,
  data,
  fallbackText
) {
  const messages = [];

  if (
    data &&
    Array.isArray(data.errors)
  ) {
    data.errors.forEach(
      function(error) {
        messages.push(
          error.message ||
          JSON.stringify(error)
        );
      }
    );
  }

  if (
    data &&
    Array.isArray(data.messages)
  ) {
    data.messages.forEach(
      function(message) {
        messages.push(
          message.message ||
          JSON.stringify(message)
        );
      }
    );
  }

  return (
    prefix +
    ". HTTP " +
    code +
    ": " +
    (
      messages.length
        ? messages.join(" | ")
        : fallbackText
    )
  );
}


function getIncrementalD1Config_() {
  const properties =
    PropertiesService
      .getScriptProperties();

  const accountId =
    properties.getProperty(
      "CLOUDFLARE_ACCOUNT_ID"
    );

  const databaseId =
    properties.getProperty(
      "CLOUDFLARE_D1_DATABASE_ID"
    );

  const apiToken =
    properties.getProperty(
      "CLOUDFLARE_D1_API_TOKEN"
    );

  const missing = [];

  if (!accountId) {
    missing.push(
      "CLOUDFLARE_ACCOUNT_ID"
    );
  }

  if (!databaseId) {
    missing.push(
      "CLOUDFLARE_D1_DATABASE_ID"
    );
  }

  if (!apiToken) {
    missing.push(
      "CLOUDFLARE_D1_API_TOKEN"
    );
  }

  if (missing.length) {
    throw new Error(
      "Missing Apps Script properties: " +
      missing.join(", ")
    );
  }

  return {
    accountId:
      accountId.trim(),

    databaseId:
      databaseId.trim(),

    apiToken:
      apiToken.trim()
  };
}


function normalizeIncrementalD1Record_(
  input,
  index
) {
  const troop =
    cleanIncrementalD1Text_(
      input &&
      input.troop
    );

  const firstName =
    cleanIncrementalD1Text_(
      input &&
      input.firstName
    );

  const lastName =
    cleanIncrementalD1Text_(
      input &&
      input.lastName
    );

  const className =
    cleanIncrementalD1Text_(
      input &&
      input.className
    );

  const fileName =
    cleanIncrementalD1Text_(
      input &&
      input.fileName
    );

  const year =
    Number(
      input &&
      input.year
    );

  const rawSession =
    input &&
    input.session;

  const session =
    rawSession === null ||
    rawSession === undefined ||
    rawSession === ""
      ? null
      : Number(rawSession);

  const scoutKey =
    cleanIncrementalD1Text_(
      input &&
      input.scoutKey
        ? input.scoutKey
        : [
            troop,
            lastName,
            firstName
          ]
            .map(function(value) {
              return cleanIncrementalD1Text_(
                value
              )
                .toLowerCase();
            })
            .join("|")
    )
      .toLowerCase();

  const fileId =
    cleanIncrementalD1Text_(
      input &&
      input.fileId
    );

  const sourceKey =
    cleanIncrementalD1Text_(
      input &&
      input.sourceKey
        ? input.sourceKey
        : fileId ||
          [
            scoutKey,
            year,
            session === null
              ? ""
              : session,
            className,
            fileName
          ].join("|")
    )
      .toLowerCase();

  const fileUrl =
    cleanIncrementalD1Text_(
      input &&
      (
        input.downloadPath ||
        input.fileUrl ||
        input.filePath
      )
    );

  if (
    !troop ||
    !firstName ||
    !lastName ||
    !scoutKey ||
    !className ||
    !sourceKey ||
    !fileName ||
    !Number.isInteger(year)
  ) {
    throw new Error(
      "Missing required data in " +
      (
        fileName ||
        "record " +
        (index + 1)
      ) +
      "."
    );
  }

  if (
    session !== null &&
    !Number.isInteger(session)
  ) {
    throw new Error(
      "Session is not a whole number in " +
      fileName +
      "."
    );
  }

  return {
    troop:
      troop,

    firstName:
      firstName,

    lastName:
      lastName,

    scoutKey:
      scoutKey,

    className:
      className,

    sourceKey:
      sourceKey,

    fileName:
      fileName,

    fileUrl:
      fileUrl || null,

    year:
      year,

    session:
      session,

    legacyTime:
      nullableIncrementalD1Text_(
        input &&
        input.legacyTime
      ),

    legacyCode:
      nullableIncrementalD1Text_(
        input &&
        input.legacyCode
      ),

    sourceFormat:
      nullableIncrementalD1Text_(
        input &&
        input.sourceFormat
      )
  };
}


function normalizeExistingRecordRow_(
  row
) {
  return {
    recordId:
      cleanIncrementalD1Text_(
        row.record_id
      ),

    scoutId:
      cleanIncrementalD1Text_(
        row.scout_id
      ),

    sourceKey:
      cleanIncrementalD1Text_(
        row.source_key
      )
        .toLowerCase(),

    scoutKey:
      cleanIncrementalD1Text_(
        row.original_scout_key
      )
        .toLowerCase(),

    troop:
      cleanIncrementalD1Text_(
        row.troop_at_time
      ),

    troopOverride:
      cleanIncrementalD1Text_(
        row.troop_override_key
      ),

    effectiveTroop:
      cleanIncrementalD1Text_(
        row.effective_troop_key
      ),

    className:
      cleanIncrementalD1Text_(
        row.class_name
      ),

    session:
      nullableIncrementalD1Number_(
        row.session
      ),

    year:
      Number(row.record_year),

    legacyTime:
      nullableIncrementalD1Text_(
        row.legacy_time
      ),

    legacyCode:
      nullableIncrementalD1Text_(
        row.legacy_code
      ),

    sourceFormat:
      nullableIncrementalD1Text_(
        row.source_format
      ),

    fileName:
      cleanIncrementalD1Text_(
        row.file_name
      ),

    fileUrl:
      nullableIncrementalD1Text_(
        row.file_url
      )
  };
}


function createIncrementalD1DeterministicId_(
  prefix,
  value
) {
  const source =
    prefix +
    "|" +
    value;

  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      source,
      Utilities.Charset.UTF_8
    );

  const hex =
    digest
      .map(function(byte) {
        const unsigned =
          byte < 0
            ? byte + 256
            : byte;

        return unsigned
          .toString(16)
          .padStart(2, "0");
      })
      .join("");

  return (
    prefix +
    "_" +
    hex.substring(0, 32)
  );
}


function incrementalMd5Hex_(
  bytes
) {
  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      bytes
    );

  return digest
    .map(function(value) {
      const unsigned =
        value < 0
          ? value + 256
          : value;

      return unsigned
        .toString(16)
        .padStart(2, "0");
    })
    .join("");
}


function incrementalD1SqlValue_(
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "NULL";
  }

  return (
    "'" +
    String(value)
      .replace(/'/g, "''") +
    "'"
  );
}


function incrementalD1SqlNumber_(
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "NULL";
  }

  const number =
    Number(value);

  return Number.isFinite(number)
    ? String(number)
    : "NULL";
}


function cleanIncrementalD1Text_(
  value
) {
  return String(
    value === null ||
    value === undefined
      ? ""
      : value
  )
    .trim()
    .replace(/\s+/g, " ");
}


function nullableIncrementalD1Text_(
  value
) {
  const cleaned =
    cleanIncrementalD1Text_(
      value
    );

  return cleaned || null;
}


function nullableIncrementalD1Number_(
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return null;
  }

  const number =
    Number(value);

  return Number.isFinite(number)
    ? number
    : null;
}


function normalizeComparisonText_(
  value
) {
  return cleanIncrementalD1Text_(
    value
  );
}


function normalizeComparisonNullableText_(
  value
) {
  const cleaned =
    cleanIncrementalD1Text_(
      value
    );

  return cleaned || "";
}


function normalizeComparisonNumber_(
  value
) {
  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return "";
  }

  const number =
    Number(value);

  return Number.isFinite(number)
    ? String(number)
    : "";
}


function showD1IncrementalToast_(
  message,
  title,
  seconds
) {
  Logger.log(
    title +
    ": " +
    message
  );

  try {
    SpreadsheetApp
      .getActiveSpreadsheet()
      .toast(
        message,
        title,
        seconds
      );
  } catch (error) {
    /*
     * Logger output is sufficient for a standalone project.
     */
  }
}
