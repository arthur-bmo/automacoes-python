function fractionToMinutes(fraction){
  return fraction * 24 * 60;
}

function relatorio_de_horas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // Define o mês anterior (por exemplo, "fevereiro")
  var data1 = new Date();
  data1.setMonth(data1.getMonth() - 1);
  var mesAnterior = data1.toLocaleString('pt-BR', { month: 'long' });

  var timezone = Session.getScriptTimeZone();

  // Arrays para armazenar os dados
  var employees = [];
  var summary = []; // Cada linha: [nome, tipo, duracao_total]
  var globalTotalMinutes = 0;

  // Itera por cada aba (funcionário)
  for (var s = 0; s < sheets.length; s++){
    var sheet = sheets[s];

    // Lê informações do funcionário
    var nomeFuncionario = sheet.getRange("B1").getValue();
    var cargoFuncionario = sheet.getRange("B2").getValue();
    var tipoFuncionario = sheet.getRange("B3").getValue();

    // Coleta os registros (a partir da linha 8, por exemplo)
    var lastrow = sheet.getLastRow();
    var lastcolumn = sheet.getLastColumn();
    var dados = sheet.getRange(8, 1, lastrow - 7, lastcolumn).getValues();
    
    // Cria a tabela de atividades com cabeçalho
    var tabelaAtividades = [];
    tabelaAtividades.push(["Data", "Início", "Término", "Projeto", "Descrição", "Duração"]);
    
    var duracao_total = 0;
    
    // Itera sobre os registros (a partir da segunda linha de 'dados')
    for (var i = 1; i < dados.length; i++) {
      var linha = dados[i];
      // Se a primeira célula estiver vazia, assume que é a linha de total
      if (linha[0] === "" || linha[0] === null) {
        break;
      }
      
      // Filtra pelo mês: converte a data da linha e obtém o nome do mês
      var data = linha[0];
      var dataObj = new Date(data);
      var mesDaLinha = dataObj.toLocaleString('pt-BR', { month: 'long' });
      if (mesDaLinha.toLowerCase() !== mesAnterior.toLowerCase()){
        continue;
      }
      
      // Formata as informações da linha
      var horarioInicio = linha[1];
      var horarioTermino = linha[2];
      var projeto = linha[3];
      var descricao = linha[4];
      var duracao = linha[5];

      var minutosduracao = duracao.getHours()*60 + duracao.getMinutes();
      duracao_total += minutosduracao;
      
      var dataFormatada = Utilities.formatDate(dataObj, timezone, "dd/MM/yyyy");
      var horarioInicioFormatado = Utilities.formatDate(new Date(horarioInicio), timezone, "HH:mm");
      var horarioTerminoFormatado = Utilities.formatDate(new Date(horarioTermino), timezone, "HH:mm");
      var duracaoFormatada = Utilities.formatDate(new Date(duracao), timezone, "HH:mm");

      tabelaAtividades.push([
        dataFormatada, 
        horarioInicioFormatado, 
        horarioTerminoFormatado, 
        projeto, 
        descricao, 
        duracaoFormatada
      ]);
    }
    
    // Adiciona a linha de total na tabela de atividades
    tabelaAtividades.push(["", "", "", "", "Total", formatMinutesToHHMM(duracao_total)]);
    
    // Acumula o total global (convertendo HH:mm para minutos)
    globalTotalMinutes += (duracao_total);
    
    summary.push([nomeFuncionario, tipoFuncionario, formatMinutesToHHMM(duracao_total)]);
    
    employees.push({
      nome: nomeFuncionario,
      cargo: cargoFuncionario,
      tipo: tipoFuncionario,
      duracao_total: duracao_total,
      atividades: tabelaAtividades
    });
  }
  
  // Adiciona a linha final com o total global de horas na tabela de resumo
  var globalTotal = formatMinutesToHHMM(globalTotalMinutes);
  summary.push(["TOTAL", "", globalTotal]);
  
  // Prepara os dados para o template HTML
  var template = HtmlService.createTemplateFromFile("relatorioTemplate");
  template.mesAnterior = mesAnterior;
  template.summary = summary;
  template.employees = employees;
  
    // Avalia o template para gerar o HTML final
  var htmlOutput = template.evaluate().getContent();
  
  // Cria um arquivo HTML no Drive e obtém o link
  var htmlFile = DriveApp.createFile("Relatorio_horas_" + mesAnterior + ".html", htmlOutput, MimeType.HTML);
  var htmlUrl = htmlFile.getUrl();
  
  // Converte o HTML para um Google Docs (usando o Advanced Drive Service)
  var blob = Utilities.newBlob(htmlOutput, 'text/html', 'relatorio.html');
  var resource = {
    title: "Relatório de Horas " + mesAnterior,
    mimeType: MimeType.GOOGLE_DOCS
  };
  var docFile = Drive.Files.insert(resource, blob, {convert: true});
  var docUrl = "https://docs.google.com/document/d/" + docFile.id;

  var pdfBlob = blob.getAs('application/pdf');
  var pdfFile = DriveApp.createFile(pdfBlob);
  
  Logger.log("Arquivo HTML criado: " + htmlUrl);
  Logger.log("Documento Google Docs criado: " + docUrl);
  Logger.log("Documento PDF criado: " + pdfFile.getUrl());
}

// Função auxiliar: converte "HH:mm" para minutos
function parseTimeToMinutes(timeStr) {
  var parts = timeStr.split(":");
  var hours = parseInt(parts[0], 10);
  var minutes = parseInt(parts[1], 10);
  return hours * 60 + minutes;
}

// Função auxiliar: converte minutos para "HH:mm"
function formatMinutesToHHMM(totalMinutes) {
  var hours = Math.floor(totalMinutes / 60);
  var minutes = totalMinutes % 60;
  return ("0" + hours).slice(-2) + ":" + ("0" + minutes).slice(-2);
}
