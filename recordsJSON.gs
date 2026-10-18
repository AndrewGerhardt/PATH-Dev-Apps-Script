const ROOT_FOLDER_ID = "1u9xdvAe7D5fqGZNK54kkuAOM9YFBWpNQ";
const OUTPUT_FILE_NAME = "records-index.json";

let filesProcessed = 0;

/**
 * Builds records-index.json from every PDF inside the root folder
 * and all of its subfolders.
 */

function buildRecordsIndex() {

  filesProcessed = 0;

  const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);

  const pdfEntries = [];
  const skippedFiles = [];

  collectPdfEntries_(
    rootFolder,
    pdfEntries,
    []
  );

  Logger.log("Finished scanning.");
  Logger.log("Total PDFs found: " + pdfEntries.length);

  const records = [];

  for (const entry of pdfEntries) {
    const record = parseRecordPdf_(
      entry.file,
      entry.folderPath
    );

    if (record) {
      records.push(record);
    } else {
      skippedFiles.push({
        path:
          entry.folderPath.join(" / ") +
          " / " +
          entry.file.getName(),
        reason: "Filename did not match a supported format."
      });
    }
  }

  records.sort(compareRecords_);

  const output = {
    generatedAt: new Date().toISOString(),
    recordCount: records.length,
    skippedFileCount: skippedFiles.length,
    records: records
  };

  saveRecordsIndex_(
    rootFolder,
    JSON.stringify(output, null, 2)
  );

  Logger.log("Indexed " + records.length + " records.");

  for (const skipped of skippedFiles) {
    Logger.log(
      "Skipped: " +
      skipped.path +
      " | " +
      skipped.reason
    );
  }

  showCompletionMessage_(
    records.length,
    skippedFiles.length
  );
}

/**
 * Collects PDFs before parsing them so the script can first discover
 * the current canonical troop names.
 */

function collectPdfEntries_(folder, pdfEntries, folderPath) {
  const currentPath = folderPath.concat(folder.getName());

  /* Logger.log("Scanning: " + currentPath.join(" / ")); */

  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();

    if (
      file.getMimeType() === MimeType.PDF &&
      file.getName() !== OUTPUT_FILE_NAME
    ) {
      pdfEntries.push({
        file: file,
        folderPath: currentPath
      });

      filesProcessed++;

      if (filesProcessed % 500 === 0) {
        Logger.log("Found " + filesProcessed + " PDFs...");
      }
    }
  }

  const subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    collectPdfEntries_(
      subfolders.next(),
      pdfEntries,
      currentPath
    );
  }
}

/**
 * Recursively scans a folder and all subfolders.
 *
 * folderPath contains the names of the parent folders. This lets
 * the script detect a year or session from the folder structure
 * when it is not included in the filename.
 */

/**
 * Determines which historical filename format a PDF uses.
 */
function parseRecordPdf_(file, folderPath) {
  const fileName = String(file.getName()).trim();

  return (
    parseCurrentFormat_(file, fileName, folderPath) ||
    parse2025Format_(file, fileName, folderPath) ||
    parse2024Format_(file, fileName, folderPath)
  );
}


/**
 * Current format:
 *
 * 71B Felts, Hunter - Beginning Swimmer Instruction - S1 2026.pdf
 *
 * Also supports a troop containing a state:
 *
 * 71B MO Felts, Hunter - Beginning Swimmer Instruction - S1 2026.pdf
 */
function parseCurrentFormat_(file, fileName, folderPath) {
  const pattern =
    /^(.+?)\s+([^,]+),\s*(.+?)\s+-\s+(.+?)\s+-\s+S(\d+)\s+(\d{4})\.pdf$/i;

  const match = fileName.match(pattern);

  if (!match) {
    return null;
  }

  return createRecord_({
    file: file,
    originalTroop: match[1],
    lastName: match[2],
    firstName: match[3],
    className: match[4],
    session: Number(match[5]),
    year: Number(match[6]),
    legacyTime: null,
    legacyCode: null,
    sourceFormat: "current",
    folderPath: folderPath
  });
}


/**
 * 2025 format:
 *
 * 71B Mound City - Killin, Brody - Camping 330.pdf
 *
 * The final number is retained as legacyCode. It is not treated as
 * a session because its exact meaning may differ from the current
 * filename format.
 */

function parse2025Format_(file, fileName, folderPath) {
  /*
   * Supports:
   *
   * 71B Mound City - Killin, Brody - Camping 330.pdf
   * 71B Mound City - Wallace-Cocoran, Gabriel - Sustainability.pdf
   * 71Boy Seneca - Ganstrom, Luke - Wilderness Survival.pdf
   * 71Boy Mound City - Young, Hayden - Signs, Signals, and Codes.pdf
   */

  const pattern =
    /^(.+?)\s+-\s+([^,]+),\s*(.+?)\s+-\s+(.+?)(?:\s+(\d+))?\.pdf$/i;

  const match = fileName.match(pattern);

  if (!match) {
    return null;
  }

  return createRecord_({
    file: file,
    originalTroop: normalizeHistoricalTroopText_(match[1]),
    lastName: match[2],
    firstName: match[3],
    className: match[4],
    session: detectSessionFromFolders_(folderPath),
    year: detectYearFromFolders_(folderPath) || 2025,
    legacyTime: null,
    legacyCode: match[5] || null,
    sourceFormat: "2025",
    folderPath: folderPath
  });
}

/**
 * 2024 format:
 *
 * 7B Clark, Lucas - Indian Lore 10:00.pdf
 *
 * The time is retained as legacyTime.
 */

function parse2024Format_(file, fileName, folderPath) {
  /*
   * Supports:
   *
   * 7B Clark, Lucas - Indian Lore 10:00.pdf
   * 30B Murphy, Daniel - American Cultures.pdf
   * 6067G McKee, Samantha - American Cultures.pdf
   */

  const pattern =
    /^(.+?)\s+([^,]+),\s*(.+?)\s+-\s+(.+?)(?:\s+(\d{1,2}:\d{2}))?\.pdf$/i;

  const match = fileName.match(pattern);

  if (!match) {
    return null;
  }

  return createRecord_({
    file: file,
    originalTroop: normalizeHistoricalTroopText_(match[1]),
    lastName: match[2],
    firstName: match[3],
    className: match[4],
    session: detectSessionFromFolders_(folderPath),
    year: detectYearFromFolders_(folderPath) || 2024,
    legacyTime: match[5] || null,
    legacyCode: null,
    sourceFormat: "2024",
    folderPath: folderPath
  });
}

/**
 * Creates the final record stored in the JSON.
 */
function createRecord_(data) {
  const troop = detectTroopFromFolders_(data.folderPath);

  if (!troop) {
    return null;
  }

  const firstName = cleanText_(data.firstName);
  const lastName = cleanText_(data.lastName);
  const className = cleanText_(data.className);

  const fileId = data.file.getId();

  const scoutKey = [
    troop,
    lastName,
    firstName
  ]
    .map(normalizeKeyPart_)
    .join("|");

  const session =
    data.session !== null &&
      data.session !== undefined
      ? Number(data.session)
      : null;

  const year = Number(data.year);
  const legacyTime = data.legacyTime || null;
  const legacyCode = data.legacyCode || null;
  const sourceFormat = data.sourceFormat;

  const sourceKey = createRecordSourceKey_({
    scoutKey: scoutKey,
    className: className,
    session: session,
    year: year,
    legacyTime: legacyTime,
    legacyCode: legacyCode,
    sourceFormat: sourceFormat
  });

  return {
    troop: troop,

    firstName: firstName,
    lastName: lastName,

    scoutKey: scoutKey,
    sourceKey: sourceKey,

    className: className,

    session: session,
    year: year,

    legacyTime: legacyTime,
    legacyCode: legacyCode,
    sourceFormat: sourceFormat,

    fileName: data.file.getName(),
    fileId: fileId,

    filePath: data.file.getUrl(),

    downloadPath:
      "https://drive.google.com/uc?export=download&id=" +
      encodeURIComponent(fileId)
  };
}

function createRecordSourceKey_(record) {
  return [
    record.scoutKey,
    record.className,
    record.year,
    record.session === null
      ? ""
      : record.session,
    record.legacyTime || "",
    record.legacyCode || "",
    record.sourceFormat || ""
  ]
    .map(normalizeKeyPart_)
    .join("|");
}

function detectYearFromFolders_(folderPath) {

  // Expected:
  // [Progress Reports, 2025, Week 1, 71B MO, Scout]

  return Number(folderPath[1]);
}

function detectSessionFromFolders_(folderPath) {

  // Week 1
  const match = folderPath[2].match(/(\d+)/);

  return match ? Number(match[1]) : null;
}

/**
 * Sort order:
 * troop
 * last name
 * first name
 * newest year first
 * class name
 * session
 */
function compareRecords_(a, b) {
  const troopCompare = a.troop.localeCompare(
    b.troop,
    undefined,
    {
      numeric: true,
      sensitivity: "base"
    }
  );

  if (troopCompare !== 0) {
    return troopCompare;
  }

  const lastNameCompare = a.lastName.localeCompare(
    b.lastName,
    undefined,
    {
      sensitivity: "base"
    }
  );

  if (lastNameCompare !== 0) {
    return lastNameCompare;
  }

  const firstNameCompare = a.firstName.localeCompare(
    b.firstName,
    undefined,
    {
      sensitivity: "base"
    }
  );

  if (firstNameCompare !== 0) {
    return firstNameCompare;
  }

  if (a.year !== b.year) {
    return b.year - a.year;
  }

  const classCompare = a.className.localeCompare(
    b.className,
    undefined,
    {
      sensitivity: "base"
    }
  );

  if (classCompare !== 0) {
    return classCompare;
  }

  const sessionA =
    a.session === null ? 999 : Number(a.session);

  const sessionB =
    b.session === null ? 999 : Number(b.session);

  return sessionA - sessionB;
}


/**
 * Replaces the existing records-index.json in the root folder.
 */
function saveRecordsIndex_(folder, json) {
  const existingFiles =
    folder.getFilesByName(OUTPUT_FILE_NAME);

  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }

  folder.createFile(
    OUTPUT_FILE_NAME,
    json,
    MimeType.PLAIN_TEXT
  );
}


/**
 * Normalizes spacing without changing capitalization.
 */
function cleanText_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ");
}


/**
 * Normalizes values used in scout keys.
 */
function normalizeKeyPart_(value) {
  return cleanText_(value).toLowerCase();
}


/**
 * Displays completion information when the script is attached
 * to a spreadsheet. Otherwise, it writes the result to the log.
 */
function showCompletionMessage_(
  recordCount,
  skippedCount
) {
  const message =
    recordCount +
    " records indexed. " +
    skippedCount +
    " file" +
    (skippedCount === 1 ? "" : "s") +
    " skipped.";

  try {
    SpreadsheetApp
      .getActiveSpreadsheet()
      .toast(
        message,
        "Records Portal",
        8
      );
  } catch (error) {
    Logger.log(message);
  }
}

function normalizeHistoricalTroopText_(troop) {
  return cleanText_(troop)
    .replace(/^(\d+)\s*BOY\b/i, "$1B")
    .replace(/^(\d+)\s*GIRL\b/i, "$1G");
}

function detectTroopFromFolders_(folderPath) {
  for (let i = folderPath.length - 1; i >= 0; i--) {
    const folderName = cleanText_(folderPath[i]);

    const match = folderName.match(
      /^(\d+)\s*([BGF])\s+([A-Z]{2}(?:\/[A-Z]{2})?)$/i
    );

    if (!match) {
      continue;
    }

    return (
      match[1] +
      match[2].toUpperCase() +
      " " +
      match[3].toUpperCase()
    );
  }

  return null;
}

// Already ran, fields are stored in script project settings
function configureRecordsPortalSync() {
  PropertiesService
    .getScriptProperties()
    .setProperties({
      RECORDS_SYNC_URL:
        "https://records.campgeigerscouting.org/api/sync",

      RECORDS_SYNC_KEY:
        "PASTE_THE_SAME_RECORDS_SYNC_KEY_HERE"
    });

  Logger.log("Records portal sync configured.");
}

// Already ran, fields are stored in script project settings
function configureCloudflareD1Import() {
  PropertiesService
    .getScriptProperties()
    .setProperties({
      CLOUDFLARE_ACCOUNT_ID:
        "PASTE_ACCOUNT_ID_HERE",

      CLOUDFLARE_D1_DATABASE_ID:
        "PASTE_DATABASE_ID_HERE",

      CLOUDFLARE_D1_API_TOKEN:
        "PASTE_API_TOKEN_HERE"
    });

  Logger.log(
    "Cloudflare D1 import settings saved."
  );
}