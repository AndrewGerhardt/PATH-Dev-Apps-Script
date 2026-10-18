/**
 * PATH 2.0 — Formatted Badge Requirement Pages for AppSheet
 *
 * INSTALL: Add this as a NEW .gs file to the PATH 2.0 master spreadsheet's
 * bound Apps Script project. Keep your existing scraper and sync scripts.
 *
 * SOURCE SHEETS (read-only):
 *   Badges — BadgeID, BadgeName, ReqCodes (only ID and name used here)
 *   BadgeRequirementText — BadgeID, ReqCode, ParentCode, DisplayOrder,
 *                          Description, TrackProgress
 *
 * OUTPUT SHEET (created automatically): BadgeRequirementPages
 *   BadgeID | BadgeName | AutoMarkdown | CustomMarkdown
 *
 * Each badge has ONE row, keyed by the existing BadgeID. AutoMarkdown is
 * regenerated from the source hierarchy. CustomMarkdown is NEVER overwritten
 * by this script; use it for your own formatted version of a badge's text.
 * Existing rows stay in place. Only BadgeID, BadgeName, and AutoMarkdown are
 * written; CustomMarkdown and any extra columns are left for you to maintain.
 * Pages for badges removed from Badges are kept until you delete them yourself.
 * Configure alphabetical display order in AppSheet rather than sorting rows here.
 *
 * AppSheet: add a LongText virtual column named DisplayMarkdown with:
 *   IF(ISNOTBLANK([CustomMarkdown]), [CustomMarkdown], [AutoMarkdown])
 * Set DisplayMarkdown formatting to Markdown and use it in a Detail view.
 * AppSheet rich text may require Settings > Views: General > Preview new
 * features. Don't display AutoMarkdown/CustomMarkdown in other view types.
 *
 * Run previewPATHRequirementPages() before buildPATHRequirementPages().
 * Running buildPATHRequirementPages() after syncing the source requirement text
 * refreshes only the generated column, preserving custom work.
 */
const PATH_BADGE_PAGES = {
  BADGES_SHEET: "Badges",
  SOURCE_SHEET: "BadgeRequirementText",
  OUTPUT_SHEET: "BadgeRequirementPages",
  HEADERS: ["BadgeID", "BadgeName", "AutoMarkdown", "CustomMarkdown"],
  // Google Sheets cell limit is 50,000 characters. Keep a safety margin.
  MAX_PAGE_LENGTH: 48000,
};

function previewPATHRequirementPages() {
  return runPATHRequirementPages_(true);
}

function buildPATHRequirementPages() {
  return runPATHRequirementPages_(false);
}

function runPATHRequirementPages_(preview) {
  // Take the same document lock used by the text sync before reading either
  // source or destination, so cooperating script updates cannot interleave.
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error("Another sheet update is running.");
  try {
    return pathPagesRunLocked_(preview);
  } finally {
    lock.releaseLock();
  }
}

function pathPagesRunLocked_(preview) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const badges = ss.getSheetByName(PATH_BADGE_PAGES.BADGES_SHEET);
  const reqs = ss.getSheetByName(PATH_BADGE_PAGES.SOURCE_SHEET);
  if (!badges || !reqs) {
    throw new Error(
      "Missing Badges or BadgeRequirementText sheet in PATH master.",
    );
  }

  const bData = badges.getDataRange().getValues();
  const rData = reqs.getDataRange().getValues();
  if (bData.length < 2 || rData.length < 2) {
    throw new Error(
      "Badges or BadgeRequirementText contains no data. Nothing changed.",
    );
  }
  const bCols = pathPagesHeaders_(bData[0]);
  const rCols = pathPagesHeaders_(rData[0]);
  pathPagesRequire_(bCols, ["BadgeID", "BadgeName"], "Badges");
  pathPagesRequire_(
    rCols,
    ["BadgeID", "ReqCode", "ParentCode", "DisplayOrder", "Description"],
    "BadgeRequirementText",
  );

  // Only Badges in the master sheet are eligible, even if requirement text
  // exists for an unrelated badge.
  const byBadge = new Map();
  for (const row of bData.slice(1)) {
    const id = String(row[bCols.BadgeID] || "").trim();
    const name = String(row[bCols.BadgeName] || "").trim();
    if (!id || !name) continue;
    if (byBadge.has(id)) throw new Error("Duplicate BadgeID in Badges: " + id);
    byBadge.set(id, { id, name, items: [] });
  }
  for (const row of rData.slice(1)) {
    const id = String(row[rCols.BadgeID] || "").trim();
    const badge = byBadge.get(id);
    if (!badge) continue;
    const code = String(row[rCols.ReqCode] || "").trim();
    const parent = String(row[rCols.ParentCode] || "").trim();
    const description = String(row[rCols.Description] || "").trim();
    const order = Number(row[rCols.DisplayOrder]);
    if (!code && !parent && !description) continue;
    badge.items.push({ code, parent, description, order });
  }

  const output = [];
  const failures = [];
  let withoutRequirements = 0;
  for (const badge of byBadge.values()) {
    if (!badge.items.length) {
      withoutRequirements++;
      continue;
    }
    try {
      const markdown = pathPagesMarkdown_(badge);
      if (markdown.length > PATH_BADGE_PAGES.MAX_PAGE_LENGTH) {
        throw new Error(
          "Page is " +
            markdown.length +
            " characters, too long for one Sheet cell.",
        );
      }
      output.push([badge.id, badge.name, markdown]);
    } catch (err) {
      failures.push(
        badge.id + " " + badge.name + ": " + (err.message || String(err)),
      );
    }
  }
  if (!output.length)
    throw new Error(
      "No valid badge pages could be generated. Nothing changed.",
    );
  // Keep a stable BadgeID -> physical row map. Do not copy and later rewrite
  // custom text: a person can edit it without acquiring an Apps Script lock.
  const pageSheet = ss.getSheetByName(PATH_BADGE_PAGES.OUTPUT_SHEET);
  const retained = new Map();
  let pageColumns = pathPagesHeaders_(PATH_BADGE_PAGES.HEADERS);
  if (pageSheet && pageSheet.getLastRow() > 0) {
    const current = pageSheet.getDataRange().getValues();
    pageColumns = pathPagesHeaders_(current[0]);
    pathPagesRequire_(
      pageColumns,
      PATH_BADGE_PAGES.HEADERS,
      PATH_BADGE_PAGES.OUTPUT_SHEET,
    );
    for (let i = 1; i < current.length; i++) {
      const row = current[i];
      const id = String(row[pageColumns.BadgeID] || "").trim();
      if (!id) {
        if (row.some((value) => String(value == null ? "" : value).trim())) {
          throw new Error(
            "Page row " +
              (i + 1) +
              " has content but no BadgeID. Nothing changed.",
          );
        }
        continue;
      }
      if (retained.has(id)) {
        throw new Error(
          "Duplicate BadgeID in " +
            PATH_BADGE_PAGES.OUTPUT_SHEET +
            ": " +
            id +
            " (rows " +
            retained.get(id).row +
            " and " +
            (i + 1) +
            "). Nothing changed.",
        );
      }
      retained.set(id, {
        row: i + 1,
        custom: String(row[pageColumns.CustomMarkdown] || ""),
      });
    }
  }

  const goodIds = new Set(output.map((r) => r[0]));
  const existing = [...retained.entries()];
  const overrides = existing.filter(([, item]) => item.custom.trim()).length;
  const preserved = existing.filter(([id]) => !goodIds.has(id)).length;
  const orphaned = existing.filter(([id]) => !byBadge.has(id)).length;
  const added = output.filter((row) => !retained.has(row[0])).length;
  const summary = [
    preview ? "PREVIEW (no changes)" : "BUILD",
    "Master badges: " + byBadge.size,
    "Generated pages: " + output.length,
    "Without requirement text: " + withoutRequirements,
    "Failed: " + failures.length,
    "Custom overrides preserved: " + overrides,
    "Existing pages retained without regeneration: " + preserved,
    "Retained pages no longer in Badges: " + orphaned,
    "New pages: " + added,
    "Existing pages to update: " + (output.length - added),
  ].join(" | ");
  Logger.log(summary + (failures.length ? "\n" + failures.join("\n") : ""));
  if (preview) return summary;
  if (failures.length) {
    // Avoid a partially refreshed set of pages if even one badge has malformed
    // hierarchy. Fix the source before rebuilding; existing pages are intact.
    throw new Error(
      "Page validation failed; no changes made.\n" + failures.join("\n"),
    );
  }

  const sh = pageSheet || ss.insertSheet(PATH_BADGE_PAGES.OUTPUT_SHEET);
  if (!sh.getLastRow()) {
    sh.getRange(1, 1, 1, 4).setValues([PATH_BADGE_PAGES.HEADERS]);
  }
  let nextRow = Math.max(2, sh.getLastRow() + 1);
  const updates = output
    .slice()
    .sort((a, b) => a[1].localeCompare(b[1]))
    .map((values) => ({
      row: retained.has(values[0]) ? retained.get(values[0]).row : nextRow++,
      values: values.slice(0, 3),
    }));
  const lastNeeded = Math.max(...updates.map((item) => item.row));
  if (sh.getMaxRows() < lastNeeded) {
    sh.insertRowsAfter(sh.getMaxRows(), lastNeeded - sh.getMaxRows());
  }
  pathPagesWriteGenerated_(sh, pageColumns, updates);
  sh.setFrozenRows(1);
  sh.setColumnWidth(pageColumns.BadgeID + 1, 115);
  sh.setColumnWidth(pageColumns.BadgeName + 1, 210);
  sh.setColumnWidth(pageColumns.AutoMarkdown + 1, 300);
  sh.setColumnWidth(pageColumns.CustomMarkdown + 1, 300);
  Logger.log(
    "BadgeRequirementPages updated. Existing row positions and custom text preserved.",
  );
  return summary;
}

/** Batch adjacent rows, writing only the three generated fields by column name. */
function pathPagesWriteGenerated_(sheet, columns, updates) {
  const sorted = updates.slice().sort((a, b) => a.row - b.row);
  const fields = ["BadgeID", "BadgeName", "AutoMarkdown"];
  for (let i = 0; i < sorted.length;) {
    const first = sorted[i].row;
    const batch = [sorted[i++]];
    while (
      i < sorted.length &&
      batch.length < 100 &&
      sorted[i].row === first + batch.length
    ) {
      batch.push(sorted[i++]);
    }
    fields.forEach((field, index) => {
      const range = sheet.getRange(first, columns[field] + 1, batch.length, 1);
      if (field === "BadgeID") range.setNumberFormat("@");
      range.setValues(batch.map((item) => [item.values[index]]));
    });
  }
}

function pathPagesMarkdown_(badge) {
  const nodes = new Map();
  for (const item of badge.items) {
    if (
      !item.code ||
      !item.description ||
      !Number.isFinite(item.order) ||
      item.order <= 0
    ) {
      throw new Error(
        "Blank/invalid requirement code, description, or display order.",
      );
    }
    const key = item.code.toLowerCase();
    if (nodes.has(key))
      throw new Error("Duplicate requirement code: " + item.code);
    nodes.set(key, { ...item, children: [], key });
  }
  const roots = [];
  for (const node of nodes.values()) {
    if (!node.parent) {
      roots.push(node);
    } else {
      const parent = nodes.get(node.parent.toLowerCase());
      if (!parent)
        throw new Error("Missing parent " + node.parent + " for " + node.code);
      if (parent === node)
        throw new Error("Self-parenting requirement: " + node.code);
      parent.children.push(node);
    }
  }
  function compare(a, b) {
    return a.order - b.order || a.code.localeCompare(b.code);
  }
  roots.sort(compare);
  for (const node of nodes.values()) node.children.sort(compare);
  const emitted = new Set();
  function render(node, depth) {
    if (emitted.has(node.key))
      throw new Error("Cycle or duplicate branch at " + node.code);
    emitted.add(node.key);
    const code = pathPagesEscape_(node.code);
    const desc = pathPagesEscape_(node.description).replace(/\s*\n\s*/g, " ");
    if (depth === 0) {
      // Parent instructions stay visible; no completion status is inferred.
      return ["## Requirement " + code, "", desc]
        .concat(
          node.children.length
            ? ["", ...node.children.map((c) => render(c, 1))]
            : [],
        )
        .join("\n");
    }
    const indentation = "  ".repeat(depth - 1);
    const line = indentation + "- **" + code + "** — " + desc;
    return [line]
      .concat(node.children.map((c) => render(c, depth + 1)))
      .join("\n");
  }
  const body = roots.map((r) => render(r, 0)).join("\n\n");
  if (emitted.size !== nodes.size)
    throw new Error("Disconnected or cyclic requirement hierarchy.");
  return "# " + pathPagesEscape_(badge.name) + "\n\n" + body;
}

function pathPagesEscape_(text) {
  // Requirement descriptions are content, not Markdown instructions. Protect
  // punctuation that may otherwise be interpreted as formatting.
  return String(text || "")
    .replace(/\\/g, "\\\\")
    .replace(/([*_[\]`])/g, "\\$1")
    .replace(/\r?\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function pathPagesHeaders_(headers) {
  const out = {};
  headers.forEach((value, index) => {
    const key = String(value || "").trim();
    if (key) {
      if (Object.prototype.hasOwnProperty.call(out, key)) {
        throw new Error("Duplicate column header: " + key);
      }
      out[key] = index;
    }
  });
  return out;
}

function pathPagesRequire_(columns, required, sheetName) {
  const missing = required.filter(
    (h) => !Object.prototype.hasOwnProperty.call(columns, h),
  );
  if (missing.length)
    throw new Error(sheetName + " missing headers: " + missing.join(", "));
}
