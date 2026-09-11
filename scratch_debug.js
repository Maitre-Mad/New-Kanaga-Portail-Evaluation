function debugSubTasks() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ProjectTasks');
  const data = sheet.getRange(2, 1, Math.min(10, sheet.getLastRow()), 8).getValues();
  Logger.log(JSON.stringify(data));
}