function triggerProcessCallCenter() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Call Center").activate
  processRows
}

function triggerProcessFollowUp() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Follow Up").activate
  processRows
}
