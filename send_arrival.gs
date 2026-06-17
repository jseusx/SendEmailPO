// Main function:
// Sends email with info, uses library POEmailLibrary.
function emailPO() {
  POEmailLibrary.emailPO();
}

function authorizePOEmail() {
  var spreadsheetId = "";

  var ss = SpreadsheetApp.openById(spreadsheetId);
  Logger.log("Spreadsheet name: " + ss.getName());

  var file = DriveApp.getFileById(spreadsheetId);
  Logger.log("Drive file name: " + file.getName());

  var testResponse = UrlFetchApp.fetch("https://www.google.com", {
    muteHttpExceptions: true
  });
  Logger.log("UrlFetchApp test response: " + testResponse.getResponseCode());

  Logger.log("Mail quota remaining: " + MailApp.getRemainingDailyQuota());

  Logger.log("Authorization check complete.");
}
