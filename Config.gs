const CFG = {
  // ===== IMPORT STAGING (separate spreadsheet) =====
  STAGING_SS_URL: 'https://docs.google.com/spreadsheets/d/1S-PAnkL6HtGG4OOnlQ4ia2Fh0QRWWsjWw7AMZhL68Oo/edit?gid=300661597#gid=300661597',

  STAGING_SHEETS: {
    ROSTER_RAW: 'Staging_RosterRaw',     // tab where you paste the export
    OFFERING_ALIASES: 'OfferingAliases', // Period | ExportClassName | OfferingID
    BADGE_ALIASES: 'BadgeAliases',       // ExportBadgeName | BadgeID (not required in this import)
    IMPORT_LOG: 'Import_Log',            // optional
  },

  SPREADSHEET_ID: '12pdiSs1nAghefPbJ6iD2o_G5tXb8m_pNVi0OUMiMNbI',

  SS_ID: '12pdiSs1nAghefPbJ6iD2o_G5tXb8m_pNVi0OUMiMNbI',

  SS_URL: 'https://docs.google.com/spreadsheets/d/12pdiSs1nAghefPbJ6iD2o_G5tXb8m_pNVi0OUMiMNbI/edit?gid=229795838#gid=229795838',

  TEMPLATE_DOC_ID: '1Z_wz1MsvLApZygfgStdZCSlZGWE9FQkBgtLIWrc5DMk',

  OUTPUT_FOLDER_ID: '1u0AVMHnzcqlrryh8tqmPD1xYHoYExrTe',

  SHEETS: {
    BADGES: 'Badges',
    BADGE_REQ_TEMPLATE: 'BadgeReqTemplate',
    OFFERINGS: 'Offerings',
    SCOUTS: 'Scouts',
    ROSTER: 'Roster',
    EXPORT_RUNS: 'ExportRuns',
  },

  HEADER_ROW: 1,

  // Chunking / time limit safety
  MAX_ITEMS_PER_EXECUTION: 1000,
  TIME_BUDGET_MS: 5 * 60 * 1000, // stop before 6 min

  // Properties keys
  PROP_PREFIX: 'PATH_EXPORT_',

  // Roster columns (exact header names)
  COLS: {
    // ExportRuns
    RUNID: 'RunID',
    RUN_YEAR: 'Year',
    RUN_SESSION: 'Session',
    RUN_STATUS: 'Status',
    RUN_STARTED_AT: 'StartedAt',
    PR_FILE_1: 'PR_File_1',
    PR_FILE_2: 'PR_File_2',
    EARNED_1: 'Earned_1',
    EARNED_2: 'Earned_2',


    // Roster
    ROSTERID: 'RosterID',
    ROSTER_YEAR: 'Year',
    ROSTER_SESSION: 'Session',
    ROSTER_OFFERINGID: 'OfferingID',
    ROSTER_SCOUTID: 'ScoutID',
    ROSTER_PR_FILE_1_COL: 'PR_File_1',
    ROSTER_PR_FILE_2_COL: 'PR_File_2',
    ROSTER_EARNED_1_COL: 'Earned_1',
    ROSTER_EARNED_2_COL: 'Earned_2',

    // Offerings
    OFFERINGID: 'OfferingID',
    BADGEID_1: 'BadgeID_1',
    BADGEID_2: 'BadgeID_2',

    // Scouts
    SCOUTID: 'ScoutID',

    // Badges
    BADGEID: 'BadgeID',
    BADGE_NAME: 'BadgeName',
    BADGE_AREA: 'Area',
  }
};
