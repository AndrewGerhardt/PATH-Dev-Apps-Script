/**
 * Safeguards in v5 (updated from v4):
 * - Never changes badge codes or advancement tracking.
 * - Does not silently drop rows belonging to master badges with parse failures.
 * - Blocks LIVE sync when tracking codes are absent from source, or source
 *   requirements have no tracker on the item or any of its parent instructions.
 * - Allows extra, unoffered Scouting America badges to be ignored.
 * - Preserves requirement codes as plain text on output.
 * - Preserves existing destination text for activities without a source match.
 * - Reports tracking problems before the informational ignored-badge list.
 */
/**
 * PATH 2.0 — Merit Badge Requirement Text Sync
 * Install in the PATH MASTER spreadsheet's bound Apps Script project.
 * Source: scraper spreadsheet tab BadgeRequirementText:
 *   BadgeName | Slug | ReqCode | ParentCode | DisplayOrder | Description | TrackProgress
 * Destination: master tab BadgeRequirementText:
 *   BadgeID | ReqCode | ParentCode | DisplayOrder | Description | TrackProgress
 *
 * Only existing Badges.BadgeName entries are eligible. BadgeID comes from Badges.
 * TrackProgress comes from Badges.ReqCodes, not the scraper's heuristic.
 * Run previewPATHRequirementTextSync first; then syncPATHRequirementText.
 */
const PATH_REQ_SYNC = {
  SOURCE_SPREADSHEET_ID: "1MOBC7E2Tf2MNSUu1CcUHM7MFEvGuGbAkg2-YN1h_AGU",
  SOURCE_SHEET: "BadgeRequirementText",
  BADGES_SHEET: "Badges",
  DEST_SHEET: "BadgeRequirementText",
  LOG_SHEET: "RequirementSyncLog",
  DEST_HEADERS: [
    "BadgeID",
    "ReqCode",
    "ParentCode",
    "DisplayOrder",
    "Description",
    "TrackProgress",
  ],
  // Name variations can be added deliberately here after previewing unmatched badges.
  NAME_ALIASES: {
    "Fish Wildlife Management": "Fish and Wildlife Management",
  },
  BLOCK_LIVE_SYNC_ON_TRACKING_MISMATCH: true,
  BLOCK_LIVE_SYNC_ON_SOURCE_FAILURE: true,
};

function previewPATHRequirementTextSync() {
  return runPATHRequirementTextSync_(true);
}
function syncPATHRequirementText() {
  return runPATHRequirementTextSync_(false);
}

function runPATHRequirementTextSync_(preview) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(30000))
    throw new Error("Another requirement sync is running.");
  try {
    const cfg = PATH_REQ_SYNC;
    if (
      !cfg.SOURCE_SPREADSHEET_ID ||
      cfg.SOURCE_SPREADSHEET_ID.includes("PASTE_")
    ) {
      throw new Error(
        "Set PATH_REQ_SYNC.SOURCE_SPREADSHEET_ID to the scraper spreadsheet ID first.",
      );
    }
    const master = SpreadsheetApp.getActiveSpreadsheet();
    const badgeSheet = master.getSheetByName(cfg.BADGES_SHEET);
    const dest = master.getSheetByName(cfg.DEST_SHEET);
    if (!badgeSheet || !dest)
      throw new Error(
        "Master requires Badges and BadgeRequirementText sheets.",
      );
    const sourceBook = SpreadsheetApp.openById(cfg.SOURCE_SPREADSHEET_ID);
    const source = sourceBook.getSheetByName(cfg.SOURCE_SHEET);
    if (!source)
      throw new Error(
        "Source is missing BadgeRequirementText. Build its hierarchy first.",
      );

    const bRows = badgeSheet.getDataRange().getValues();
    const sRows = source.getDataRange().getValues();
    if (bRows.length < 2 || sRows.length < 2)
      throw new Error(
        "Badges or source hierarchy is empty. Refusing to clear destination.",
      );
    const bCol = pathHeaderMap_(bRows[0]);
    const sCol = pathHeaderMap_(sRows[0]);
    pathRequire_(bCol, ["BadgeID", "BadgeName", "ReqCodes"], "Badges");
    pathRequire_(
      sCol,
      [
        "BadgeName",
        "Slug",
        "ReqCode",
        "ParentCode",
        "DisplayOrder",
        "Description",
      ],
      "Source hierarchy",
    );

    const byName = new Map();
    const byId = new Map();
    const warnings = [];
    for (const row of bRows.slice(1)) {
      const id = String(row[bCol.BadgeID] || "").trim();
      const name = String(row[bCol.BadgeName] || "").trim();
      if (!id || !name) continue;
      const key = pathNormalizeName_(name);
      if (byName.has(key) && byName.get(key).id !== id) {
        warnings.push(
          "Duplicate master badge name: " +
            name +
            " (" +
            byName.get(key).id +
            ", " +
            id +
            ")",
        );
        byName.get(key).ambiguous = true;
        continue;
      }
      const reqCodes = new Set(
        String(row[bCol.ReqCodes] || "")
          .split(",")
          .map((x) => x.trim().toLowerCase())
          .filter(Boolean),
      );
      const badge = { id, name, reqCodes, ambiguous: false };
      byName.set(key, badge);
      byId.set(id, badge);
    }

    const matched = new Map();
    const missingTracked = [];
    const uncoveredSource = [];
    const trackingMismatchIds = new Set();
    const sourceFailures = [];
    const failureSheet = sourceBook.getSheetByName("HierarchyValidationLog");
    if (failureSheet && failureSheet.getLastRow() > 1) {
      const failureRows = failureSheet
        .getRange(2, 1, failureSheet.getLastRow() - 1, 4)
        .getValues();
      const failedBySlug = new Map();
      for (const failure of failureRows) {
        if (failure[2])
          failedBySlug.set(String(failure[2]).toLowerCase(), failure);
      }
      for (const failure of failedBySlug.values()) {
        const name = String(failure[1] || "");
        const badge = byName.get(
          pathNormalizeName_(cfg.NAME_ALIASES[name] || name),
        );
        if (badge)
          sourceFailures.push(badge.id + " " + name + ": " + failure[3]);
      }
    }
    const unmatched = new Set();
    const seen = new Map();
    let duplicatesSkipped = 0;
    const output = [];
    const sourceCodes = new Map();
    let invalid = 0;
    for (const row of sRows.slice(1)) {
      const name = String(row[sCol.BadgeName] || "").trim();
      const code = String(row[sCol.ReqCode] || "").trim();
      const description = String(row[sCol.Description] || "").trim();
      if (!name && !code && !description) continue;
      const normalized = pathNormalizeName_(cfg.NAME_ALIASES[name] || name);
      const badge = byName.get(normalized);
      if (!badge || badge.ambiguous) {
        unmatched.add(name || "(missing name)");
        continue;
      }
      if (!code || !description) {
        invalid++;
        warnings.push("Blank code or description: " + name + " / " + code);
        continue;
      }
      const key = badge.id + "\u0000" + code.toLowerCase();
      const order = Number(row[sCol.DisplayOrder]);
      const parent = String(row[sCol.ParentCode] || "").trim();
      const signature = JSON.stringify([
        parent.toLowerCase(),
        order,
        description,
      ]);
      if (seen.has(key)) {
        if (seen.get(key) === signature) {
          duplicatesSkipped++;
          continue;
        }
        invalid++;
        warnings.push(
          "Conflicting duplicate requirement: " + badge.id + " / " + code,
        );
        continue;
      }
      seen.set(key, signature);
      matched.set(badge.id, badge);
      if (!sourceCodes.has(badge.id)) sourceCodes.set(badge.id, new Set());
      sourceCodes.get(badge.id).add(code.toLowerCase());
      if (!Number.isFinite(order) || order <= 0) {
        invalid++;
        warnings.push("Invalid order: " + badge.id + " / " + code);
        continue;
      }
      output.push([
        badge.id,
        code,
        parent,
        order,
        description,
        badge.reqCodes.has(code.toLowerCase()),
      ]);
    }
    for (const [id, badge] of matched) {
      const codes = sourceCodes.get(id) || new Set();
      const missing = [...badge.reqCodes].filter((code) => !codes.has(code));
      if (missing.length) {
        const issue =
          id +
          " " +
          badge.name +
          ": tracked codes missing from source: " +
          missing.join(", ");
        warnings.push(issue);
        missingTracked.push(issue);
        trackingMismatchIds.add(id);
      }
    }
    // A parent can deliberately be the tracked requirement, with its numbered
    // examples/choices shown as reference children. That covers its descendants.
    // An entire missing section (such as Personal Fitness 6) does not.
    const rowsByBadge = new Map();
    for (const row of output) {
      if (!rowsByBadge.has(row[0])) rowsByBadge.set(row[0], new Map());
      rowsByBadge.get(row[0]).set(row[1].toLowerCase(), row);
    }
    for (const [id, items] of rowsByBadge) {
      const badge = matched.get(id);
      const parents = new Set(
        [...items.values()]
          .map((r) => String(r[2] || "").toLowerCase())
          .filter(Boolean),
      );
      const uncovered = [];
      for (const [code, row] of items) {
        let current = code,
          covered = false;
        const visited = new Set();
        while (current) {
          if (visited.has(current) || !items.has(current)) {
            invalid++;
            warnings.push("Invalid source parent chain: " + id + " / " + code);
            break;
          }
          visited.add(current);
          if (badge.reqCodes.has(current)) covered = true;
          current = String(items.get(current)[2] || "")
            .trim()
            .toLowerCase();
        }
        if (!parents.has(code) && !covered) uncovered.push(row[1]);
      }
      if (uncovered.length) {
        const issue =
          id +
          " " +
          badge.name +
          ": source requirements missing from tracking (including parent coverage): " +
          uncovered.join(", ");
        uncoveredSource.push(issue);
        warnings.push(issue);
        trackingMismatchIds.add(id);
      }
    }
    output.sort(
      (a, b) =>
        a[0].localeCompare(b[0]) || a[3] - b[3] || a[1].localeCompare(b[1]),
    );
    const report = [
      "Mode: " + (preview ? "PREVIEW (no changes)" : "SYNC"),
      "Master badges: " + byId.size,
      "Matched badges: " + matched.size,
      "Requirement rows: " + output.length,
      "Extra source badge names ignored (expected): " + unmatched.size,
      "Master activities without source match: " + (byId.size - matched.size),
      "Invalid or conflicting rows: " + invalid,
      "Identical duplicates skipped: " + duplicatesSkipped,
      "Warnings: " + warnings.length,
      "Tracking mismatches (badges): " + trackingMismatchIds.size,
      "Failed source badges in master: " + sourceFailures.length,
    ];
    if (!output.length)
      throw new Error(
        "No matched requirements. Nothing was changed. Check spreadsheet ID and badge names.",
      );
    if (invalid)
      throw new Error(
        "Validation failed; nothing changed. " +
          report.join(" | ") +
          "\n" +
          warnings.slice(0, 15).join("\n"),
      );
    if (
      !preview &&
      cfg.BLOCK_LIVE_SYNC_ON_TRACKING_MISMATCH &&
      trackingMismatchIds.size
    ) {
      throw new Error(
        "Live sync blocked: requirement tracking definitions do not cover the source hierarchy. Nothing changed.\n" +
          missingTracked.concat(uncoveredSource).slice(0, 25).join("\n"),
      );
    }
    if (
      !preview &&
      cfg.BLOCK_LIVE_SYNC_ON_SOURCE_FAILURE &&
      sourceFailures.length
    ) {
      throw new Error(
        "Live sync blocked: some badges in your master have not passed source parsing. Nothing changed.\n" +
          sourceFailures.slice(0, 25).join("\n"),
      );
    }

    // All-or-nothing prevalidation of source, destination layout, and required data.
    const existingHeaders = dest
      .getRange(1, 1, 1, 6)
      .getValues()[0]
      .map(String);
    const expected = cfg.DEST_HEADERS;
    if (existingHeaders.some((h, i) => h.trim() !== expected[i])) {
      throw new Error(
        "Destination headers must be exactly: " + expected.join(" | "),
      );
    }
    // Replace only badges matched in this run. Keep reference text maintained
    // locally for Pioneers, COPE, other custom activities, or an absent source.
    const preserved =
      dest.getLastRow() > 1
        ? dest
            .getRange(2, 1, dest.getLastRow() - 1, 6)
            .getValues()
            .filter(
              (r) =>
                r.some((v) => String(v || "").trim()) &&
                !matched.has(String(r[0] || "").trim()),
            )
        : [];
    const finalOutput = output.concat(preserved);
    report.push(
      "Existing unmatched-activity rows preserved: " + preserved.length,
    );
    if (!preview) {
      // Keep existing destination intact until the entire source is validated.
      const oldRows = Math.max(0, dest.getLastRow() - 1);
      if (dest.getMaxRows() < finalOutput.length + 1)
        dest.insertRowsAfter(
          dest.getMaxRows(),
          finalOutput.length + 1 - dest.getMaxRows(),
        );
      // Keep requirement codes as text (e.g., 4e-1).
      dest
        .getRange(2, 2, Math.max(finalOutput.length, 1), 2)
        .setNumberFormat("@");
      // Write new values first, then clear any obsolete trailing rows.
      for (let i = 0; i < finalOutput.length; i += 500) {
        const batch = finalOutput.slice(i, i + 500);
        dest.getRange(i + 2, 1, batch.length, 6).setValues(batch);
      }
      if (oldRows > finalOutput.length)
        dest
          .getRange(finalOutput.length + 2, 1, oldRows - finalOutput.length, 6)
          .clearContent();
      dest.setFrozenRows(1);
      const log =
        master.getSheetByName(cfg.LOG_SHEET) ||
        master.insertSheet(cfg.LOG_SHEET);
      if (!log.getLastRow()) log.appendRow(["Timestamp", "Status", "Detail"]);
      log.appendRow([new Date(), "SYNC", report.join(" | ")]);
    }
    const details = report.concat(
      warnings,
      ...sourceFailures.map((n) => "Source parsing failed: " + n),
      [...byId.values()]
        .filter((badge) => !matched.has(badge.id))
        .map(
          (badge) =>
            "No source match; existing text kept: " +
            badge.id +
            " " +
            badge.name,
        ),
      [...unmatched].sort().map((n) => "Extra source badge ignored: " + n),
    );
    Logger.log(details.join("\n"));
    // Do not use SpreadsheetApp.getUi() here: the sync can be run in contexts
    // without an interactive spreadsheet UI (editor, trigger, API, etc.).
    // The full report is already in the execution log; return a concise summary.
    return report.join(" | ");
  } finally {
    lock.releaseLock();
  }
}
function pathHeaderMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    const key = String(h || "").trim();
    if (key) map[key] = i;
  });
  return map;
}
function pathRequire_(map, names, sheet) {
  const missing = names.filter(
    (name) => !Object.prototype.hasOwnProperty.call(map, name),
  );
  if (missing.length)
    throw new Error(sheet + " missing headers: " + missing.join(", "));
}
function pathNormalizeName_(value) {
  return String(value || "")
    .normalize("NFKC")
    .replace(/&/g, " and ")
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[^a-z0-9]+/gi, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}
