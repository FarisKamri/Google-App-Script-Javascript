function copyAndUpdateCells() {
  // Open the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get the sheet named "Main Dashboard"
  var sheet = spreadsheet.getSheetByName("Main Dashboard");
  
  if (!sheet) {
    // If the sheet does not exist, show an alert and exit
    SpreadsheetApp.getUi().alert('Sheet "Main Dashboard" not found.');
    return;
  }

  // Copy the range D20:D50 to E20:E50
  var sourceRange = sheet.getRange("D20:D50");
  var destinationRange = sheet.getRange("E20:E50");
  var values = sourceRange.getValues();
  destinationRange.setValues(values);
  
  // Copy the single cell H10 to I23
  var singleSourceCell = sheet.getRange("H10");
  var singleDestinationCell = sheet.getRange("I23");
  var singleValue = singleSourceCell.getValue();
  singleDestinationCell.setValue(singleValue);
  
  // Write today's date in cell E18
  var today = new Date();
  sheet.getRange("E18").setValue(today);
}
