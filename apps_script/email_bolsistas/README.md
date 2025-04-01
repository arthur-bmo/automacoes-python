# 📬 Google Apps Script – Alerta de vencimento de contratos e retorno de colaboradores

Este script automatiza o envio de e-mails com base em datas registradas em uma planilha do Google Sheets. Ele foi criado para apoiar o setor administrativo na gestão de prazos relacionados a contratos e retornos de estagiários, bolsistas ou outros colaboradores.

## 🚀 O que o script faz

- Verifica a **coluna O** da planilha:
  - Se a data registrada for **igual à data atual**, envia um e-mail avisando que o contrato chegou ao fim.
- Verifica a **coluna AF**:
  - Se **faltarem 3 dias** para a data registrada, envia um e-mail lembrando sobre o retorno previsto do colaborador.

As mensagens incluem o nome do colaborador, extraído da **coluna A**, e são enviadas automaticamente para os setores responsáveis.

## 📫 Exemplos de mensagens

**1. Vencimento de contrato:**
```
Assunto: Contrato chegou ao prazo final

Boa tarde,

O contrato do(a) colaborador(a): [NOME] chegou ao prazo final.

Por favor, verifique se há necessidade de renovação ou outro encaminhamento.

Atenciosamente,  
Equipe Administrativa
```

**2. Lembrete de retorno próximo:**
```
Assunto: Retorno esperado de [NOME]

Boa tarde,

O retorno do(a) colaborador(a): [NOME] está previsto para o dia [DATA].

Atenciosamente,  
Equipe Administrativa
```

## 🛠️ Tecnologias utilizadas

- Google Apps Script
- `MailApp.sendEmail()` – envio automático de e-mails
- `SpreadsheetApp.getActiveSpreadsheet()` – leitura da planilha ativa
- `Logger.log()` – marcação de execução no log

## 📋 Estrutura da planilha esperada

| A (Nome)       | ... | O (Fim do Contrato) | ... | AF (Data de Retorno) |
|----------------|-----|----------------------|-----|------------------------|
| Fulano da Silva|     | 01/04/2025           |     | 04/04/2025             |

## 🔁 Recomendações de uso

- Agendar o script para rodar diariamente usando o menu de **Triggers** no Apps Script.
- Adaptar os endereços de e-mail no código para os setores ou responsáveis reais.
- Manter a consistência dos dados e formatos na planilha para garantir o bom funcionamento.

---

> Um lembrete simples, mas essencial, para manter a gestão de contratos em dia sem depender da memória humana.
