function copy() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A176').activate();
  spreadsheet.getRange('A141:K174').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
};