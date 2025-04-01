function subtractDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() - days);
  return result;
}
function ChecaDataMandaEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var currentDate = new Date();

  for (var i = 2; i <= lastRow; i++) {
    var date = sheet.getRange("O" + i).getValue();
    if (date instanceof Date && date.toDateString() === currentDate.toDateString()) {
      var nomeBolsista = sheet.getRange("A"+i).getValue();
      var emailAdress = "jovem.aprendiz2@sif.org.br";
      var subject = "Contrato chegou ao prazo final";
      var message = "Boa tarde, Juliano. \n\nO contrato do bolsista: " + nomeBolsista +" chegou ao prazo final. \nFineza checar se há aditivo. Caso haja, marcar a coluna 'Quitação' como 'OK'. Caso contrário, encaminhar esse email para apoio.embrapii@sif.org.br. \n\nEncarecidamente,\nApoio Administrativo EMBRAPII.";
      MailApp.sendEmail(emailAdress, subject, message);
    }

    var data_retorno = sheet.getRange("AF" + i).getValue();
    var date_retorno = new Date(data_retorno)
    var threeDaysBefore = subtractDays(date_retorno, 3)
    var data_formatada = date_retorno.toLocaleDateString('pt-BR');
    if (threeDaysBefore instanceof Date && threeDaysBefore.toDateString() === currentDate.toDateString()) {
      Logger.log("E-mail enviado")
      var nomeBolsista = sheet.getRange("A"+i).getValue();
      var emailAdress = "genteegestao@sif.org.br";
      var subject = "Retorno esperado de "+ nomeBolsista
      var message = "Boa tarde, Lara. \n\nO retorno do bolsista: " + nomeBolsista + " está agendado para o dia " + data_formatada + "\n\nAtenciosamente, \nApoio Administrativo EMBRAPII.";
      MailApp.sendEmail(emailAdress, subject, message);
    }
    }
  
}
