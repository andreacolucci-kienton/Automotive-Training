function read_csv_Albania_v01() {
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
  var Q2result = ui.prompt("Inserisci ID:");
  fi_id = Q2result.getResponseText().toString();
  SpreadsheetApp.getUi().alert(al);
  var ui = SpreadsheetApp.getUi();
  var Q1result = ui.prompt("Nome del file csv per le presenze:");
  nome_csv = Q1result.getResponseText().toString();
  var file = fold.getFilesByName(nome_csv).next();
  var row_done = -1
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString(), ",");
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Partecipanti");
  var sheet_risorse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formazione interna");
  ids = sheet_risorse.getRange(2, 1, sheet_risorse.getLastRow() - 1, 1).getValues();
  for (var i in ids) {
    if (ids[i].toString() == fi_id) {
      row_done = i;
    }
  }
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  sheet.deleteRows(2, 4);
  sheet.deleteColumns(2, 2);
  row_write_presence = parseInt(row_done, 10) + 2;
  var owner = sheet_risorse.getRange(row_write_presence,13).getValues();
  owner = owner.toString().toUpperCase().split(" ").sort();
  n_own = owner.length
  temp = ""
  for (j = n_own-1; j >= 0; j = j-1){
    temp = temp + owner[j]  
  }
  owner = temp
  if (row_done != -1) {
    var nomi = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    for (var i = 2; i < nomi.length + 1; i = i + 1) {
      if (nomi[i - 1].toString().includes("KINETON")) {
        nomi[i - 1] = nomi[i - 1].toString().toUpperCase()
        nomi[i - 1] = nomi[i - 1].replace(" (KINETON)", "").split(" ")
        nomi[i - 1] = nomi[i - 1].sort()
        n_str = nomi[i-1].length
        temp = ""
        for (j = n_str-1; j >= 0; j = j-1){
          temp = temp + nomi[i-1][j]  
        }
      }else{
        nomi[i - 1] = nomi[i - 1].toString().toUpperCase().split(" ")
        nomi[i - 1] = nomi[i - 1].sort()
        n_str = nomi[i-1].length
        temp = ""
        for (j = n_str-1; j >= 0; j = j-1){
          temp = temp + nomi[i-1][j]  
        }
      }
      nomi[i - 1] = temp
      sheet.getRange(i, 1).setValue(nomi[i - 1].toString())
    }
    var numColumns_ris = sheet_risorse.getLastColumn();
    var lista_risorse = sheet_risorse.getRange(1, 1, 1, numColumns_ris).getValues();
    for (var i_part = 1; i_part < nomi.length; i_part = i_part + 1) {
      for (var i_ris = 0; i_ris <= numColumns_ris - 1; i_ris = i_ris + 1) {
        var ris = lista_risorse[0][i_ris].toString().toUpperCase();
        var part = nomi[i_part].toString().toUpperCase();
        ris = ris.split(" ").sort()
        n_str = ris.length
        temp = ""
        for (j = n_str-1; j >= 0; j = j-1){
          temp = temp + ris[j]  
        }
        ris = temp
        if (ris != "") {
          if (ris == part & part != owner) {
            sheet_risorse.getRange(row_write_presence, i_ris + 2).setValue("D")
          }
        }
      }
    }
  }
}
