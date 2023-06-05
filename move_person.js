function spostamento_risorse() {
  var foglioAttivo = SpreadsheetApp.getActive().getSheetByName("prova");
  var ui = SpreadsheetApp.getUi();
  var Q1result = ui.prompt("Inserisci il nome della risorsa da riallocare:");
  nome_ris = Q1result.getResponseText().toString();
  var Q1result = ui.prompt("Inserisci il nome del nuovo manager:");
  nome_nuovo_manager = Q1result.getResponseText().toString();

  var datiCelle = foglioAttivo.getDataRange().getValues();
  for (var riga = 0; riga < datiCelle.length; riga++) {
    for (var colonna = 0; colonna < datiCelle[riga].length; colonna++) {
      var valoreCella = datiCelle[riga][colonna];
      if (valoreCella === nome_ris) {
        var num_ris = foglioAttivo.getRange(riga + 1, colonna + 1, 1, 1).getA1Notation();
        var nome_ris = foglioAttivo.getRange(riga + 1, colonna + 1, 1, 1).getValue();
        var num_lesson = foglioAttivo.getRange(riga + 1, colonna + 2, 1, 1).getFormula();
        var num_less_received = foglioAttivo.getRange(riga + 1, colonna + 3, 1, 1).getFormula();
        var h_less_rec = foglioAttivo.getRange(riga + 1, colonna + 4, 1, 1).getFormula();
        var num_req_open = foglioAttivo.getRange(riga + 1, colonna + 5, 1, 1).getFormula();
        var h_less_prov = foglioAttivo.getRange(riga + 1, colonna + 6, 1, 1).getFormula();
        rig = riga
        col = colonna
        break
      }
    }
  }
  for (var riga = 0; riga < datiCelle.length; riga++) {
    for (var colonna = 0; colonna < datiCelle[riga].length; colonna++) {
      var valoreCella = datiCelle[riga][colonna];
      if (valoreCella === nome_nuovo_manager) {
        var datiCelle = foglioAttivo.getRange(1, colonna + 1, foglioAttivo.getLastRow(), 1).getValues();
        var ultimaRiga = datiCelle.filter(String).length;
        foglioAttivo.getRange(ultimaRiga + 1, colonna + 1, 1, 1).setValue(nome_ris);
        foglioAttivo.getRange(ultimaRiga + 1, colonna + 2, 1, 1).setFormula(num_lesson);
        foglioAttivo.getRange(ultimaRiga + 1, colonna + 3, 1, 1).setFormula(num_less_received);
        foglioAttivo.getRange(ultimaRiga + 1, colonna + 4, 1, 1).setFormula(h_less_rec);
        example_nome = foglioAttivo.getRange(ultimaRiga, colonna + 1, 1, 1);
        example = foglioAttivo.getRange(ultimaRiga, colonna + 2, 1, 1);
        for1 = num_req_open.toString();
        str1 = foglioAttivo.getRange(ultimaRiga + 1, colonna + 1, 1, 1).getA1Notation();
        for1 = for1.replace(num_ris, str1);
        foglioAttivo.getRange(ultimaRiga + 1, colonna + 5, 1, 1).setFormula(for1);

        for2 = h_less_prov.toString();
        for2 = for2.replace(num_ris, str1);
        foglioAttivo.getRange(ultimaRiga + 1, colonna + 6, 1, 1).setFormula(for2);

        foglioAttivo.getRange(ultimaRiga + 1, colonna + 6, 1, 1)
        u_rig = ultimaRiga;
        col2 = colonna;
        break
      }
    }
  }
  var nome_ris = foglioAttivo.getRange(rig + 1, col + 1, 1, 6).clearContent();
  foglioAttivo.getRange(rig + 2, col + 1, foglioAttivo.getLastRow(), 6).moveTo(foglioAttivo.getRange(rig + 1,col + 1));
  example_nome.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 1, 1, 1), { formatOnly: true });
  example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 2, 1, 1), { formatOnly: true });
  example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 3, 1, 1), { formatOnly: true });
  example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 4, 1, 1), { formatOnly: true });
  example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 5, 1, 1), { formatOnly: true });
  example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 6, 1, 1), { formatOnly: true });

}