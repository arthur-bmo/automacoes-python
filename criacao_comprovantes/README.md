# 🧾 Criação de Comprovantes – Geração automática de recibos de bolsas

Script simples criado para uma demanda pontual: gerar comprovantes de pagamento de bolsas com base em uma planilha de dados preenchida.

## 🚀 O que o script faz

- Lê dados de bolsistas a partir de uma planilha `.xlsm`
- Preenche automaticamente um modelo `.docx` com as informações:
  - Nome, CPF, RG, endereço, valor, data, projeto, coordenador, etc.
- Salva o documento preenchido em `.docx`
- Converte automaticamente para `.pdf`
- Nomeia os arquivos de forma organizada por bolsista, mês e ano

## 🛠️ Tecnologias utilizadas

- Python
- `pandas` (leitura da planilha)
- `python-docx` (edição de documentos Word)
- `docx2pdf` (conversão de DOCX para PDF)
- `locale` (formatação monetária)

## ⚠️ Observações

- O modelo `.docx` precisa conter os marcadores como `NBOLSISTO`, `CEPEEFE`, `VALUEDABOLSA`, etc.
- O caminho para o modelo e planilha deve ser ajustado no script conforme o ambiente local.

---

> Apesar de simples, esse script automatizou uma tarefa que antes exigia preenchimento manual de milhares de recibos.
