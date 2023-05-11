function mail_nuova_richiesta() {
    var gruppo = "automotive.training@kineton.it";
    var mail = "andrea.colucci@kineton.it"
    var subject = "Nuova Richiesta Formazione Interna";
    var body = "https://docs.google.com/spreadsheets/d/1__vsS1auloqpAScLqoIi-439qfhewJhl6dPbp_vWLKo/edit#gid=1662919666"
    MailApp.sendEmail(mail, subject, body);
    MailApp.sendEmail(gruppo, subject, body);
}
