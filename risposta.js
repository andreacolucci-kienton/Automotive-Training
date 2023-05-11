function RISPOSTA() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Richiesta");
  // Get the last row in the sheet that has data.
  var numRows = sheet.getLastRow();
  var numColumns = sheet.getLastColumn();
  var mails = sheet.getRange(3, 1, 1, numColumns).getValues();
  mails = mails[0]
  var rangeRows = sheet.getRange(2, 1, numRows - 1, numColumns).getValues();
  // Use a for loop to process each row of data
  var row = rangeRows[0];
  var topic = row[0];
  for (var ind in mails) {
    var mail = mails[parseInt(ind)];
    var subject = "Formazione Interna - SLOT BLOCCATO: " + topic;
    var body = "Lo slot per la lezione su: " + topic + " è stato bloccato. Grazie"
    if (mail != "") {
      MailApp.sendEmail(mail, subject, body);
    }
  }
}
