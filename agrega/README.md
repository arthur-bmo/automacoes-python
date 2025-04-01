# 📊 Agrega – Automação de acompanhamento de projetos

Este script automatiza a coleta de dados financeiros no sistema AGREGA (Fundação Arthur Bernardes) e atualiza planilhas do Google Sheets com essas informações. Foi criado para facilitar o compartilhamento de dados entre a equipe de projetos e os professores responsáveis por cada convênio.

## 🚀 O que o script faz

- Lê códigos de projetos listados em diferentes abas de uma planilha no Google Sheets.
- Acessa o site do AGREGA automaticamente.
- Faz login e pesquisa os dados de cada projeto.
- Extrai a tabela de rubricas aprovadas.
- Atualiza a planilha com os dados coletados, aplicando formatação de moeda.

## 🛠️ Tecnologias utilizadas

- Python
- Selenium (navegação automatizada)
- Google Sheets API (via `gspread`)
- Pandas (manipulação de dados)
- BeautifulSoup (extração de HTML)

## ⚠️ Observações

- É necessário configurar a chave de autenticação da API do Google (Service Account).
- O script depende de acesso ao sistema AGREGA com login e senha válidos.
- Algumas configurações como o ID da planilha e credenciais estão hardcoded (recomenda-se o uso de variáveis de ambiente em versões públicas).

## 📁 Estrutura

- `puxa contas empresa agrega online.py` – Script principal com toda a lógica de automação

---

> Automatizar esse processo poupou tempo, evitou erros manuais e facilitou o processo de acompanhamento de recursos de projetos.
