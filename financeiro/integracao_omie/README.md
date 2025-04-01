# 🔄 Financeiro – Automação de inclusão de pedidos no Omie a partir do Conveniar

Este conjunto de scripts foi desenvolvido para automatizar a etapa de **pagamento de pedidos** em projetos gerenciados via Conveniar. Ele busca diariamente os pedidos aprovados, baixa os documentos relacionados, e os envia para a plataforma Omie em forma de lote de pagamentos.

## 🚀 O que o sistema faz

- Acessa o sistema **Conveniar** via Selenium e extrai todos os pedidos aprovados do dia.
- Baixa os comprovantes e documentos anexos dos pedidos.
- Coleta dados bancários, históricos e valores associados.
- Verifica se o fornecedor já existe na base de clientes do **Omie**.
- Caso não exista, registra o novo cliente automaticamente via API.
- Cria um objeto `pedido` com os dados financeiros, bancários e de categoria.
- Envia os pedidos organizados para um **lote de pagamento diário no Omie**.

## 🧠 Lógica modular

A automação foi dividida em módulos para facilitar a manutenção:

- `conveniar_utils.py` – login, extração de dados, download de arquivos do Conveniar
- `omie_api_utils.py` – consulta e registro de clientes via API do Omie
- `baixa_planilha_clientes.py` – obtenção da planilha de clientes atualizada
- `pedidos aprovados conveniar.py` – script principal que orquestra todas as etapas

## 🔧 Em desenvolvimento

Este código está em constante evolução. As próximas otimizações incluem:

- Tratamento automatizado de pagamentos via boletos com código de barras
- Integração para leitura e validação de **PIX via QR Code**
- Otimizações de desempenho e padronização de erros

## 🛠️ Tecnologias utilizadas

- Python
- Selenium + WebDriverManager
- Requests (para chamadas de API)
- Pandas (manipulação de planilhas)
- BeautifulSoup (leitura de HTML)
- PyAutoGUI + Pyperclip (automação de salvamento de arquivos)
- API REST do Omie

## 💡 Benefícios

- Redução drástica de erros humanos na etapa de cadastro de pagamentos.
- Agilidade no fluxo entre aprovação de pedidos e envio para pagamento.
- Eliminação de tarefas manuais como login, download, preenchimento e envio.
- Organização padronizada de documentos por data e número de pedido.

## ⚠️ Observações

- É necessário configurar as credenciais de acesso do Omie e do Conveniar nos scripts.
- O caminho para os arquivos (planilhas e pastas de destino) deve ser ajustado conforme o ambiente.
- Idealmente, este processo deve ser executado uma vez por dia.

---

> Essa automação encurta o caminho entre aprovação e execução financeira — trazendo precisão e rastreabilidade para o fluxo de pagamentos.
