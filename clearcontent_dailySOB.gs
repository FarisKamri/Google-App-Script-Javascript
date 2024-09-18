function clearContent_DailySOB() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Define all the ranges to clear
  var rangesToClear = {
    "1. Daily Monitoring":["K35:M40","V57:W57"],
    "2. Topline/CIR Calculator":["ae35:af60","Ah35:Ah60", "ae32"],
    "4.a Weekly Actuals":["o40","o42","o94","o96","o154","o156","o210","o212","o272","o274"]


  };

  // Loop through each sheet and clear its ranges
  Object.keys(rangesToClear).forEach(function(sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    var ranges = rangesToClear[sheetName];
    
    ranges.forEach(function(range) {
      sheet.getRange(range).clearContent();
    });
  });
}
