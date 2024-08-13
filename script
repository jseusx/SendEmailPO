// Gets URL
function sheetGid() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  return sheet.getSheetId();
}

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
  const emailRecipient = " ";
  const ccEmail = " ";

  //gets the current sheet gid
  var sheetID = sheetGid();

  // values[Row][Colomn] in sheet
  var values = SpreadsheetApp.getActiveSheet().getDataRange().getValues();

  var poNumber = extractBeforeDash(values[4][1]);

  // arrival variable is the time the PO arrived
  var arrival = cleanTime(values[2][2]);

  //gets the last row with data. Should always be the last s/n unless layout of spreadsheet is changed
  var range = SpreadsheetApp.getActiveSheet().getLastRow();

  // creates the pdf url to be exported/sent off to recipient
  var outputURL = "https://docs.google.com/spreadsheets/d/" + SpreadsheetApp.getActiveSpreadsheet().getId() + "/export?range=a1:f"+ range + "&format=pdf&gid=" + sheetID

 // Fetch the PDF file as a blob
  var response = UrlFetchApp.fetch(outputURL, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });

  var pdfBlob = response.getBlob().setName("PO_" + poNumber + ".pdf");
  
  MailApp.sendEmail({
    to: emailRecipient,
    cc: ccEmail,
    subject: "PO " + poNumber + " has arrived",
    body: "Date of Arrival: "+ arrival + "\n\nAttached is the PDF to the PO.",
    attachments: [pdfBlob]

  });

  console.log("Email sent");
}
