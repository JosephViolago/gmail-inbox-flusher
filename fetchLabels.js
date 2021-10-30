// https://webapps.stackexchange.com/a/102153
// Portability:
// Unfortunately, Container-bound scripts are locked to a particular gDoc
// To reuse, Open "Tools" >> "Script Editor" in the working gDoc to create a new Container-bound Script space
// https://developers.google.com/apps-script/guides/bound

function countLabels() {
  var labels = GmailApp.getUserLabels();
  var labelArray = [];
  for (var i = 0; i < labels.length; i++) {
    var labelName = labels[i].getName();
    var labelCount = GmailApp.getUserLabelByName(labelName).getThreads().length;
    labelArray.push([labelName,labelCount]);;
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("GMail Labels");
  sheet.clear();
  sheet.appendRow(["Label Name","Label Count"]);
  sheet.getRange(1, 1).setFontWeight("bold");
  sheet.getRange(1, 2).setFontWeight("bold");
  sheet.getRange(sheet.getLastRow() + 1, 1, labelArray.length, 2).setValues(labelArray);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("GMail")
    .addItem("Count Labels", "countLabels")
    .addToUi();
}