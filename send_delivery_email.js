// Main function:
// Sends email with info of when PO has been taken for delivery
function emailDelivery() {
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
  var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
  
  lastMessage.reply("Date of taken for delivery: " + date);

  console.log("Reply sent")
}
