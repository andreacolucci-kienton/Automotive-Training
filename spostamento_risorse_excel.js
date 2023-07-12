function spostamento_risorse_excel() {
  var foglioAttivo = SpreadsheetApp.getActive().getSheetByName("Overview");
  var foglioAssegnazioni = SpreadsheetApp.getActive().getSheetByName("Foglio29");
  var assegnazioni = foglioAssegnazioni.getDataRange().getValues();
  var managers = foglioAssegnazioni.getRange(2, 1, foglioAssegnazioni.getLastRow()-1, 1).getValues().toString()
  var colonnaManager = 0;
  var colonnaRis = 1;
  for (var rigaAssegn = 1; rigaAssegn < assegnazioni.length; rigaAssegn++){
    nome_ris = assegnazioni[rigaAssegn][colonnaRis];
    nome_nuovo_manager = assegnazioni[rigaAssegn][colonnaManager];
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
    var col2 = 0;
    for (var riga = 0; riga < datiCelle.length; riga++) {
      while(col2 < datiCelle[riga].length){
        var valoreCella = datiCelle[riga][col2];
        if (valoreCella === nome_nuovo_manager) {
          var datiCelle = foglioAttivo.getRange(1, col2 + 1, foglioAttivo.getLastRow(), 1).getValues();
          var u_rig = datiCelle.filter(String).length;
          if (col != col2 && !(managers.includes(nome_ris))) {
            foglioAttivo.getRange(u_rig + 1, col2 + 1, 1, 1).setValue(nome_ris);
            foglioAttivo.getRange(u_rig + 1, col2 + 2, 1, 1).setFormula(num_lesson);
            foglioAttivo.getRange(u_rig + 1, col2 + 3, 1, 1).setFormula(num_less_received);
            foglioAttivo.getRange(u_rig + 1, col2 + 4, 1, 1).setFormula(h_less_rec);
            example_nome = foglioAttivo.getRange(u_rig, col2 + 1, 1, 1);
            example = foglioAttivo.getRange(u_rig, col2 + 2, 1, 1);
            for1 = num_req_open.toString();
            str1 = foglioAttivo.getRange(u_rig + 1, col2 + 1, 1, 1).getA1Notation();
            for1 = for1.replace(num_ris, str1);
            foglioAttivo.getRange(u_rig + 1, col2 + 5, 1, 1).setFormula(for1);
            for2 = h_less_prov.toString();
            for2 = for2.replace(num_ris, str1);
            foglioAttivo.getRange(u_rig + 1, col2 + 6, 1, 1).setFormula(for2);
            example_nome.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 1, 1, 1), { formatOnly: true });
            example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 2, 1, 1), { formatOnly: true });
            example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 3, 1, 1), { formatOnly: true });
            example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 4, 1, 1), { formatOnly: true });
            example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 5, 1, 1), { formatOnly: true });
            example.copyTo(foglioAttivo.getRange(u_rig + 1, col2 + 6, 1, 1), { formatOnly: true });
            foglioAttivo.getRange(rig + 1, col + 1, 1, 6).clearContent();
            foglioAttivo.getRange(rig + 2, col + 1, foglioAttivo.getLastRow(), 6).moveTo(foglioAttivo.getRange(rig + 1, col + 1));
            Logger.log("Moved! " + nome_ris)
            break
          } else {
            Logger.log("Already present or manager! " + nome_ris);
            break
          }
        }
        col2 = col2 + 1;
      }
    }
  }
}