function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Menu")
    .addItem("Add Customer", "addPrompt")
    .addItem("New Sheet", "addNewSheet")
    .addSeparator()
    .addItem("Clear Current Sheet", "clearPrompt")
    .addToUi();
}

//Add Customer Prompt
function addPrompt() {
  var form = HtmlService.createHtmlOutputFromFile("index.html")
    .setWidth(400)
    .setHeight(475);
  SpreadsheetApp.getUi().showModalDialog(form, "Add New Customer");
}

//Add Customer onClick function
function addCustomer(form) {
  var row = [
    form.customerName,
    form.customerID,
    form.streetAddress,
    form.streetAddress2,
    form.city,
    form.state,
    form.zipCode,
    form.phoneNumber,
    form.emailAddress,
    form.companyName,
    form.term,
    form.customerRep,
    form.customerType,
    form.commissionType,
    form.priceList,
  ];
  SpreadsheetApp.getActive().getSheetByName("New_Customers").appendRow(row);
}

//Add New Sheet Prompt
function addNewSheet() {
  var addSheet = HtmlService.createHtmlOutputFromFile("newSheet.html")
    .setWidth(400)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(addSheet, "Add New Sheet");
}

//Add New Sheet onClick function (Duplicate Template Sheet)
function duplicate() {
  var ss = "Template";

  SpreadsheetApp.getActive().getSheetByName(ss).showSheet();
  SpreadsheetApp.setActiveSheet(
    SpreadsheetApp.getActive().getSheetByName(ss),
    true
  );
  SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();
  SpreadsheetApp.getActive().getSheetByName(ss).hideSheet();
  clearInvoice();
}

//Clear Contents Prompt
function clearPrompt() {
  var clear = HtmlService.createHtmlOutputFromFile("clear.html")
    .setWidth(550)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(clear, "Clear Current Sheet");
}

//Clear Sheet onClick function
function clearInvoice() {
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
