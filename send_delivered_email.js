function emailDelivered() {
  // values[Row][Colomn] in sheet
  var values = SpreadsheetApp.getActiveSheet().getDataRange().getValues();

  var poNumber = extractBeforeDash(values[4][1]);

  // Original subject line (Searches for it as well)
  var query = 'subject:"PO ' + poNumber + ' has arrived"';
  var threads = GmailApp.search(query);

   // Check if any thread was found
  if (threads.length === 0) {
    console.log("No email found with the subject: " + query);
    return;
  }

  // Get message from first matching thread
  var messages = threads[0].getMessages();
  // Chooses the last message in the thread (most recent)
  var lastMessage = messages[messages.length - 1];
  

  // delivery variable is the time/day the PO was taken to be delivered
  //var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  rawDate = cleanTime(values[17][4]);
  
  // Convert rawDate to a Date object if it isn’t one already. If cleanTime returns a Date object, you can omit new Date().
  var formattedDate = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), "M/d/yyyy");

  
  lastMessage.reply("Date of delivery: " + formattedDate);

  console.log("Reply sent")

}
