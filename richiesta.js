function RICHIESTA() {
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
  var ui = SpreadsheetApp.getUi();
  var id = ui.prompt("Inserisci l'ID della richiesta:");
  var cw = ui.prompt("Inserisci la disponibilità (CW):");
  for (var ind in mails) {
    var mail = mails[parseInt(ind)];
    var subject = "Richiesta Formazione Interna: " + topic;
    var body = "Mail automatica.Per favore comunicare a Pasquale Perrotta la disponibilità per (CW " + cw.getResponseText() + ") per la lezione di formazione interna su: " + topic + ".\nNome meeting: [" + id.getResponseText() + "] - " + topic;
    if (mail != "") {
      MailApp.sendEmail(mail, subject, body);
    }
  }
}
