function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Menu")
    .addItem("New Sheet", "addNewSheet")
    .addSeparator()
    .addItem("Clear Current Sheet", "clear")
    .addToUi();
}

//Add New Sheet Prompt
function addNewSheet() {
  var addSheet = HtmlService.createHtmlOutputFromFile("index.html")
    .setWidth(400)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(addSheet, "Add New Sheet");
}

//Add New Sheet onClick function (Duplicate Template Sheet)
function duplicate(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var templateSheet = ss.getSheetByName("Template");

  if (formData.newTabName == "") {
    ss.insertSheet(formData.newTabName, { template: templateSheet });
    templateSheet.hideSheet();
    ss.getSheetByName(formData.newTabName).getRange("D2").setValue("Paducah");
    ss.getSheetByName(formData.newTabName).activate();
    ss.moveActiveSheet(1);
  } else {
    return;
  }
}

//Clear Sheet onClick function
function clear() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var date = new Date();
  var month = date.getMonth();
  month += 1;
  var day = date.getDate();
  var hours = date.getHours();
  hours += -1;
  var minutes = date.getMinutes();
  var seconds = date.getSeconds();
  var quoteNumber = sheet.getRange("N6");

  //Set MC# for Sales Rep
  var quoteID = `MCx${day}${hours}${minutes}${seconds}`;
  var name = "Sales Rep Name Here";

  sheet.getRange("B14").setValue(name);

  quoteNumber.clearContent();
  quoteNumber.setValue(quoteID);

  // Cells to be cleared
  sheet.getRange("C19:C43").clearContent();
  sheet.getRange("B10").clearContent();
  sheet.getRange("B12").clearContent();
  sheet.getRange("N45").clearContent();
  sheet.getRange("F44:K47").clearContent();

  setFormula();
  time();

  //Set Cell Values
  sheet.getRange("B5").setValue("48 Remington Way").mergeAcross();
  sheet.getRange("B6").setValue("Hickory, KY 42051");
  sheet.getRange("L19:L38").setValue(1);
}
