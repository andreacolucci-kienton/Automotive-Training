function read_csv_AT_v04() {
  var ui = SpreadsheetApp.getUi();
  var Q2result = ui.prompt("Inserisci il nome della lezione:");
  nome_cartella = Q2result.getResponseText().toString();
  var folders = DriveApp.getFoldersByName(nome_cartella);
  var i = 0;
  var file_in_cartella = [];
  var fold = folders.next();
  var f = fold.getFiles();
  
  while (f.hasNext()) {
    file_in_cartella.push(f.next());
    i++;
  }
  var al = "";
  for (var i in file_in_cartella) {
    al = al + file_in_cartella[i] + "\n"
  }
  SpreadsheetApp.getUi().alert(al);
  var ui = SpreadsheetApp.getUi();
  var Q1result = ui.prompt("Nome del file csv per le presenze:");
  nome_csv = Q1result.getResponseText().toString();
  var file = fold.getFilesByName(nome_csv).next();
  var row_done = -1
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ",");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Partecipanti");
  var sheet_risorse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formazione interna");
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  sheet.deleteRows(2, 4);
  sheet.deleteColumns(2, 2);
  var A1 = sheet.getRange("A1:A1").getValue();
  var A1String = A1.toString().replace("\"*     ", "").replace("\"", "");
  var fi = A1String.split("] ");
  var fi_id = fi[0].replace("[", "");
  fi_id = fi_id.substring(1, fi_id.length);
  Logger.log("ID Lesson: " + fi_id);
  sheet.getRange("A1:A1").setValue(A1String);
  ids = sheet_risorse.getRange(2, 1, sheet_risorse.getLastRow() - 1, 1).getValues();
  for (var i in ids) {
    if (ids[i].toString() == fi_id) {
      row_done = i;
    }
  }
  Logger.log("ID found on row: " + row_done);
  row_write_presence = parseInt(row_done, 10) + 2;
  var owner = sheet_risorse.getRange(row_write_presence,13).getValues();
  owner = owner.toString();
  if (row_done != -1) {
    var nomi = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    for (var i = 2; i < nomi.length + 1; i = i + 1) {
      nomi[i - 1] = nomi[i - 1].toString().toUpperCase().replace("(KINETON)", "")
      nomi[i - 1] = nomi[i - 1].toString().split(" ")
      nomi[i - 1] = nomi[i - 1][1] + " " + nomi[i - 1][0]
      sheet.getRange(i, 1).setValue(nomi[i - 1].toString().toUpperCase())
    }
    var numColumns_ris = sheet_risorse.getLastColumn();
    var lista_risorse = sheet_risorse.getRange(1, 1, 1, numColumns_ris).getValues();
    for (var i_part = 1; i_part < nomi.length; i_part = i_part + 1) {
      for (var i_ris = 0; i_ris <= numColumns_ris - 1; i_ris = i_ris + 1) {
        var ris = lista_risorse[0][i_ris].toString().toUpperCase();
        var part = nomi[i_part].toString().toUpperCase();
        if (ris != "") {
          if (ris.trim() == part.trim()  & ris.trim() != owner.toString().toLocaleUpperCase().trim() ) {
            sheet_risorse.getRange(row_write_presence, i_ris + 2).setValue("D")
          }
        }
      }
    }
  }
}
