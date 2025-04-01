# 📑 Prestação de Contas – Automação de registros no SRInfo com geração de comprovantes

Este foi meu primeiro projeto de automação robusto — e talvez o mais impactante em termos de economia de tempo e esforço. Apesar de ter sido desenvolvido com estrutura inicial simples e escrita ainda imatura, este script realiza uma tarefa essencial que anteriormente consumia dezenas de horas manuais por ano.

## 🚀 O que o script faz

- Lê planilhas mensais de uso de equipamentos e bolsas vinculadas a projetos.
- Extrai automaticamente os dados necessários (nome, valor, finalidade, macroentrega, etc.).
- Gera comprovantes em PDF personalizados, preenchidos a partir de modelos.
- Associa notas fiscais e comprovantes de renda às pessoas e equipamentos.
- Realiza login no sistema **SRInfo** e cria registros automáticos para cada entrada.
- Faz o **upload automático dos documentos gerados diretamente na plataforma**.

## 📌 Impacto

- Mais de **80 horas economizadas anualmente** em tarefas manuais de geração e upload de comprovantes.
- Redução de falhas humanas em processos repetitivos.
- Fluxo contínuo de prestação de contas com mais organização, rastreabilidade e agilidade.

## 🛠️ Tecnologias utilizadas

- Python
- Selenium (automatização do navegador)
- `PyPDF2`, `reportlab`, `io` (geração e manipulação de PDFs)
- `xlwings` (leitura avançada de Excel)
- `tkinter` (interface gráfica para seleção de arquivos e mês)
- `locale` (formatação monetária)

## 🔧 Sobre o desenvolvimento

Este código foi escrito em um momento em que eu estava aprendendo na prática. Ele ainda carrega heranças da minha inexperiência na época (ex: pouca modularização, blocos longos e repetitivos). Mesmo assim, **funciona bem e entregou valor real**, o que mostra que a utilidade vem antes da perfeição.

## ⚠️ Observações

- Os modelos de comprovantes em `.pdf` devem estar em pastas específicas, com nomes padronizados.
- O script usa autenticação direta no navegador — o usuário e senha do SRInfo devem ser atualizados no código.
- A interface foi feita para facilitar o uso por pessoas não técnicas.

---

> Este projeto me ensinou que a melhor forma de aprender programação é resolvendo um problema real que ninguém mais quer resolver.
