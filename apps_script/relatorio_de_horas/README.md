# ⏱️ Google Apps Script – Relatório de Horas (Google Sheets)

Este projeto automatiza a geração de relatórios de horas trabalhadas com base em uma planilha do Google Sheets. Foi criado para ser integrado diretamente a um Google Spreadsheet, permitindo que equipes visualizem, exportem e assinem relatórios organizados por mês.

## 🚀 O que o sistema faz

- Lê as abas da planilha (uma por colaborador)
- Extrai dados de atividades e filtra automaticamente pelo mês anterior
- Calcula a carga horária de cada funcionário e o total geral
- Gera um relatório HTML personalizado
- Converte o relatório em:
  - Documento Google Docs
  - Arquivo PDF
  - Arquivo HTML
- Cria áreas de assinatura no relatório, inclusive para o funcionário, coordenadora de contratos e direção

## 📁 Arquivos

- `preenche_relatorio_de_horas.js`: Script principal do Google Apps Script
- `relatorioTemplate.html`: Modelo HTML usado para compor o relatório final  
  *(desenvolvido com apoio do ChatGPT)*

## 🧠 Observações

- Os dados devem começar na linha 8 de cada aba
- Campos esperados nas células:  
  - `B1` = nome, `B2` = cargo, `B3` = tipo
- Funciona diretamente a partir do menu do Google Sheets ao ser instalado como script do documento

## 🛠 Tecnologias utilizadas

- Google Apps Script (JavaScript)
- HTML + templating interno do GAS
- Google Drive API (para criação de Docs e PDFs)

---

> Automatizar a geração desses relatórios tornou mais rápida e confiável a gestão de horas em projetos com múltiplos colaboradores.
