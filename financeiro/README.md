# 💸 Financeiro – Quebra e organização de comprovantes bancários

Script desenvolvido para facilitar a gestão de comprovantes financeiros. Ele processa um único arquivo PDF contendo diversos comprovantes bancários, separa cada página em arquivos individuais propriamente renomeados, e tenta associá-los automaticamente aos registros corretos de uma planilha de pagamentos.

## 🚀 O que o script faz

- Abre um PDF contendo vários comprovantes de transferências, pagamentos ou boletos.
- Divide o PDF em páginas individuais, salvando cada comprovante como um arquivo separado.
- Lê automaticamente o nome do recebedor, valor e/ou CNPJ de cada página.
- Compara as informações com uma planilha do Excel que contém os registros de pagamento.
- Marca os pagamentos encontrados na planilha e associa os PDFs corretos a cada linha.
- Evita nomes duplicados com numeração incremental (Ex: `nome.pdf`, `nome 1.pdf`).

## 🛠️ Tecnologias utilizadas

- Python
- `PyPDF2` (para leitura e separação do PDF)
- `xlwings` (para leitura e atualização da planilha Excel)
- `tkinter` (para seleção interativa de arquivos e pastas)

## 🧪 Exemplo de uso

Ideal para organizações que realizam pagamentos em massa e recebem um único PDF com todos os comprovantes do banco. A automação substitui o processo manual de "abrir, renomear, associar e organizar" comprovante por comprovante.

## ⚠️ Observações

- O script requer que o modelo de planilha tenha colunas bem definidas (ex: nome, CNPJ, valor, status).
- As comparações de valores e nomes são flexíveis, mas dependem da consistência nos dados.
- Ao final da execução, a planilha é atualizada com os comprovantes identificados.

---

> Automatizar esse processo economizou horas de trabalho manual e reduziu o risco de erros em conferência de pagamentos.
