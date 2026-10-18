function removeBlankRows() {
  const sheet = SpreadsheetApp.getActiveSheet();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  for (let r = lastRow; r >= 1; r--) {
    const isBlank = values[r - 1].every(cell => cell === '');

    if (isBlank) {
      sheet.deleteRow(r);
    }
  }
}