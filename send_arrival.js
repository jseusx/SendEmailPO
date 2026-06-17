// extracts text before " - " from the given string
function extractBeforeDash(text) {
  var dashIndex = text.indexOf(" - ");
  if (dashIndex !== -1) {
    return text.substring(0, dashIndex); // Returns the substring before " - "
  }
  return text; // Return the whole text if " - " is not found
}

// cleans time to remove the extra bits not needed
function cleanTime(text) {
  //converts to String data type
  text = String(text);
  
  var zeroIndex = text.indexOf(" 00:00:00");

  if(zeroIndex !== -1) {
    return text.substring(0, zeroIndex); // Returns the substring before the " 00:00:00"
  }
  return text;
}

// Main function:
// Sends email with info
function emailPO() {
  // email recipient goes here: ↓
  const emailRecipient = "accountspayable@pitt.k12.nc.us";
  const currentUser = Session.getActiveUser().getEmail();
  const ccEmail = currentUser;

  // Gets the spreadsheet and currently active sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error("No active spreadsheet found. Run this from the spreadsheet/button, not directly from the library.");
  }

  var sheet = ss.getActiveSheet();

  if (!sheet) {
    throw new Error("No active sheet found.");
  }

    var spreadsheetId = ss.getId();
    var sheetID = sheet.getSheetId();

  // values[Row][Column] in sheet
  var values = sheet.getDataRange().getValues();

  var poNumber = extractBeforeDash(values[4][1]);

  // arrival variable is the time the PO arrived
  var arrival = cleanTime(values[2][2]);

  //gets the last row with data. Should always be the last s/n unless layout of spreadsheet is changed
  //var lastRow = SpreadsheetApp.getActiveSheet().getLastRow();
  var lastRow = sheet.getLastRow();

  // creates the pdf url to be exported/sent off to recipient
  //var outputURL = "https://docs.google.com/spreadsheets/d/" + SpreadsheetApp.getActiveSpreadsheet().getId() + "/export?range=a1:f"+ range + "&format=pdf&gid=" + sheetID
  var outputURL =
    "https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/export" +
    "?format=pdf" +
    "&gid=" + sheetID +
    "&range=" + encodeURIComponent("A1:F" + lastRow) +
    "&portrait=true" +
    "&fitw=true" +
    "&sheetnames=false" +
    "&printtitle=false" +
    "&pagenumbers=false" +
    "&gridlines=false" +
    "&fzr=false";

  Logger.log("Export URL:");
  Logger.log(outputURL);

 // Fetch the PDF file as a blob
  var response = UrlFetchApp.fetch(outputURL, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
      muteHttpExceptions: true
  });

  Logger.log("PDF export response code:");
  Logger.log(response.getResponseCode());

  if (response.getResponseCode() !== 200) {
    Logger.log("PDF export response body:");
    Logger.log(response.getContentText());
    throw new Error("PDF export failed. HTTP " + response.getResponseCode());
  }

  var pdfBlob = response.getBlob().setName("PO_" + poNumber + ".pdf");
  
  MailApp.sendEmail({
    to: emailRecipient,
    cc: ccEmail,
    subject: "PO " + poNumber + " has arrived",
    body: "Date of Arrival: "+ arrival + "\n\nAttached is the PDF to the PO.",
    attachments: [pdfBlob]

  });

  console.log("Email sent")
}
