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

> Nada disso seria necessário se os sistemas fossem bem-feitos. Como não são, o jeito é automatizar.
