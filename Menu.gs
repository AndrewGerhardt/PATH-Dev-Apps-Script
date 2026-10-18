function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('PATH Functions')
    .addSubMenu(
      ui.createMenu('Importing')
        .addItem('1) Import Scouts', 'importScoutsFromRosterExport')
        .addItem('2) Import Rosters', 'importRosterFromRosterExport')
    )
    .addSubMenu(
      ui.createMenu('Exporting')
        .addItem('1) Preview portal export', 'uiPreviewPortalReportExport')
        .addItem('2) Send one RosterID pilot', 'uiStartPortalReportPilot')
        .addItem('3) Send reports to portal', 'uiStartPortalReportExport')
        
        .addSeparator()
        .addItem('Legacy: Create troop folders', 'uiCreateFoldersOnly')
        .addItem('Legacy: Generate Drive PDFs', 'uiStartPdfRun')
    )
    .addItem('Stop Current Export', 'stopCurrentExportRun')
    .addItem('Resume Portal Export', 'resumePathPortalSnapshotExport')
    .addSeparator()
    .addItem('Remove Blank Rows', 'removeBlankRows')
    .addSeparator()
    .addItem('Portal Export Status', 'uiShowPortalExportStatus')
    .addItem('View Latest Run Status', 'uiShowLatestRunStatus')
    .addToUi();
}
