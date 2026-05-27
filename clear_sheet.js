// Clears the sheet
function clearSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  // 1. Handle B5: Clear anything before " - "
  var b5Range = sheet.getRange("B5");
  var b5Value = b5Range.getValue().toString();
  var suffix = " - ";
  
  if (b5Value.includes(suffix)) {
    // Find the starting position of " - "
    var index = b5Value.indexOf(suffix);
    // Extract everything from that position to the end of the string
    var keptText = b5Value.substring(index); 
    b5Range.setValue(keptText);
  } else {
    b5Range.clearContent(); 
  }

  // 2. Clear specific individual cells
  var cellsToClear = [
    'B7', 'B9', 'E5', 'E7', 'B22', 
    'A27', 'A30', 'C27', 'C30', 'E27'
  ];
  sheet.getRangeList(cellsToClear).clearContent();

  // 3. Clear everything from Row 33 downwards
  var lastRow = sheet.getLastRow();

  if (lastRow >= 33) {
    // getRange(row, column, numRows, numColumns)
    sheet.getRange(33, 2, (lastRow - 32), 1).clearContent();
  }
  
  console.log("Sheet cleared successfully.");
}
