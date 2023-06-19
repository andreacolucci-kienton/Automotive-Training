function add_person() {
  var foglioAttivo = SpreadsheetApp.getActive().getSheetByName("Overview");
  var foglioAttivo2 = SpreadsheetApp.getActive().getSheetByName("Formazione interna");
  var ui = SpreadsheetApp.getUi();
  var Q1result = ui.prompt("Inserisci il nome della risorsa da aggiungere:");
  nome_ris = Q1result.getResponseText().toString();
  var Q2result = ui.prompt("Inserisci il nome del manager:");
  nome_manager = Q2result.getResponseText().toString();
  var datiCelle = foglioAttivo.getDataRange().getValues();
  var datiCelle2 = foglioAttivo2.getDataRange().getValues();
  foglioAttivo2.insertColumnsAfter(foglioAttivo2.getLastColumn(), 2)
  var numRows = foglioAttivo2.getLastRow();
  var valori = new Array(numRows - 2).fill(["NR"]);
  var valori1 = new Array(numRows - 2).fill(["ND"]);
  var req = new Array(1).fill(["Req"])
  var done = new Array(1).fill(["Done"])
  valori = req.concat(valori);
  valori1 = done.concat(valori1);
  foglioAttivo2.getRange(2, foglioAttivo2.getLastColumn() + 1, numRows - 1, 1).setValues(valori);
  foglioAttivo2.getRange(2, foglioAttivo2.getLastColumn() + 1, numRows - 1, 1).setValues(valori1);
  sourceRange = foglioAttivo2.getRange(3, foglioAttivo2.getLastColumn() - 1, numRows, 1);
  foglioAttivo2.getRange(1, foglioAttivo2.getLastColumn() - 1).setValue(nome_ris);
  var lettere = "";
  var lettere_ex = "";
  var lettere_ex2 = "";

  indice = foglioAttivo2.getLastColumn() - 1
  while (indice > 0) {
    var resto = (indice - 1) % 26;
    lettere = String.fromCharCode(65 + resto) + lettere;
    indice = Math.floor((indice - 1) / 26);
  }
  col_req = lettere;
  indice2 = foglioAttivo2.getLastColumn();
  lettere = "";
  while (indice2 > 0) {
    var resto3 = (indice2 - 1) % 26;
    lettere = String.fromCharCode(65 + resto3) + lettere;
    indice2 = Math.floor((indice2 - 1) / 26);
  }
  col_done = lettere;
  for (var riga = 0; riga < datiCelle.length; riga++) {
    for (var colonna = 0; colonna < datiCelle[riga].length; colonna++) {
      var valoreCella = datiCelle[riga][colonna];
      if (valoreCella === nome_manager) {
        var datiCelle = foglioAttivo.getRange(1, colonna + 1, foglioAttivo.getLastRow(), 1).getValues();
        var ultimaRiga = datiCelle.filter(String).length;
        cella_add_person = foglioAttivo.getRange(ultimaRiga + 1, colonna + 1);
        origin = foglioAttivo.getRange(ultimaRiga, colonna + 1);
        cella_add_person.setValue(nome_ris)
        origin.copyTo(cella_add_person, { formatOnly: true });
        ex_range = foglioAttivo.getRange(ultimaRiga, colonna + 1);
        example_name = foglioAttivo.getRange(ultimaRiga, colonna + 1).getValue();
        example = foglioAttivo.getRange(ultimaRiga, colonna + 2);
        example2 = foglioAttivo.getRange(ultimaRiga, colonna + 3);
        example3 = foglioAttivo.getRange(ultimaRiga, colonna + 4);
        example4 = foglioAttivo.getRange(ultimaRiga, colonna + 5);
        example5 = foglioAttivo.getRange(ultimaRiga, colonna + 6);

        target_for1 = foglioAttivo.getRange(ultimaRiga + 1, colonna + 2);
        target_for2 = foglioAttivo.getRange(ultimaRiga + 1, colonna + 3);
        target_for3 = foglioAttivo.getRange(ultimaRiga + 1, colonna + 4);
        target_for4 = foglioAttivo.getRange(ultimaRiga + 1, colonna + 5);
        target_for5 = foglioAttivo.getRange(ultimaRiga + 1, colonna + 6);
        break;
      }
    }
  }
  var riga = 0;
  var ex_col = 0;
  var ex_col2 = 0;
  for (var colonna = 0; colonna < datiCelle2[riga].length; colonna++) {
    var valoreCella2 = datiCelle2[riga][colonna];
    if (valoreCella2 === example_name) {
      ex_col = colonna + 1;
      ex_col2 = colonna + 2;
      break;
    }
  }
  while (ex_col > 0) {
    var resto = (ex_col - 1) % 26;
    lettere_ex = String.fromCharCode(65 + resto) + lettere_ex;
    ex_col = Math.floor((ex_col - 1) / 26);
  }
  while (ex_col2 > 0) {
    var resto2 = (ex_col2 - 1) % 26;
    lettere_ex2 = String.fromCharCode(65 + resto2) + lettere_ex2;
    ex_col2 = Math.floor((ex_col2 - 1) / 26);
  }

  ex_string = lettere_ex;
  ex_string2 = lettere_ex2;
  last_row = foglioAttivo2.getLastRow()
  str1_for1 = lettere_ex + "$3:" + lettere_ex + "$" + last_row;
  str1_for2 = col_req + "$3:" + col_req + "$" + last_row;

  str2_for1 = lettere_ex2 + "$3:" + lettere_ex2 + "$" + last_row;
  str2_for2 = col_done + "$3:" + col_done + "$" + last_row;
  for1 = example.getFormula();
  for1 = for1.replace(str1_for1, str1_for2);
  target_for1.setFormula(for1);
  for2 = example2.getFormula();
  for2 = for2.replace(str2_for1, str2_for2);
  target_for2.setFormula(for2);
  for3 = example3.getFormula().toString();
  for3 = for3.replace(str2_for1, str2_for2);
  target_for3.setFormula(for3);
  for4 = example4.getFormula().toString();
  for4 = for4.replace(ex_range.getA1Notation(), cella_add_person.getA1Notation());
  target_for4.setFormula(for4);
  for5 = example5.getFormula().toString();
  for5 = for5.replace(ex_range.getA1Notation(), cella_add_person.getA1Notation());
  target_for5.setFormula(for5);
  example.copyTo(target_for1, { formatOnly: true });
  example.copyTo(target_for2, { formatOnly: true });
  example.copyTo(target_for3, { formatOnly: true });
  example.copyTo(target_for4, { formatOnly: true });
  example.copyTo(target_for5, { formatOnly: true });
}