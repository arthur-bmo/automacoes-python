# 🧩 Processos Automatizados com Python

Este repositório reúne scripts criados para automatizar tarefas administrativas repetitivas, especialmente em contextos de gestão de projetos, prestação de contas e organização documental. A ideia é simples: usar a programação para eliminar tarefas que não precisam mais ser feitas por humanos.

## ✨ Destaques

- Geração e preenchimento automático de ~1500 documentos semestrais
- Extração e tratamento de dados a partir de plataformas como SRInfo, Conveniar e Omie
- Download automatizado de arquivos e relatórios com uso de Selenium
- Manipulação de PDFs (divisão, junção, anotações dinâmicas)
- Integração com Google Sheets e APIs externas
- Scripts com interfaces Tkinter para uso por pessoas sem experiência com código

## 🧠 Filosofia

Cada script aqui foi criado para resolver um problema real, com foco em:
- Reduzir o tempo gasto com tarefas mecânicas
- Garantir consistência nos processos
- Tornar acessível o uso de ferramentas automatizadas para quem não programa

## 💻 Exemplo de Script (Google Apps Script)

```javascript
function relatorio_de_horas() {
  // Percorre todas as abas da planilha, organiza os dados por colaborador,
  // calcula horas mensais e exporta para HTML, PDF e Google Docs.
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
```

Para ver o script completo e comentado, acesse: `scripts/relatorio_de_horas.gs`

## 📁 Estrutura Sugerida

- `docs/`: Modelos e documentos gerados
- `pdf_tools/`: Scripts de manipulação de PDF
- `emails/`: Envio de e-mails com base em planilhas
- `omie/`: Integração com a API do Omie
- `conveniar/`: Download automatizado e tratamento de prestações de contas
- `tk_gui/`: Scripts com interfaces visuais
- `utils/`: Funções auxiliares reutilizáveis

## 🛠 Requisitos

- Python 3.10+
- `selenium`, `pandas`, `xlwings`, `PyPDF2`, `python-docx`, `gspread`, `reportlab`
- (Opcional) Credenciais Google API para integração com Google Sheets

---

> “I choose a lazy person to do a hard job. Because a lazy person will find an easy way to do it.” -Bill Gates
