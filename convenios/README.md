# 📁 Convênios – Automação de download e organização de comprovantes

Este script automatiza o acesso ao sistema Conveniar para realizar o download, renomeação e organização de comprovantes de projetos de prestação de contas, como compras, bolsas, fretes, viagens, entre outros.

## 🚀 O que o script faz

- Acessa o sistema SIF/Conveniar automaticamente.
- Navega entre os projetos de prestação de contas.
- Faz login e localiza os lançamentos com arquivos anexos.
- Baixa comprovantes em PDF e converte imagens para PDF, se necessário.
- Renomeia os arquivos com data, histórico e valor.
- Organiza os arquivos em pastas nomeadas por projeto.
- Junta arquivos PDF com prefixo similar em um único documento final.

## 📌 Exemplos de uso

Ideal para equipes financeiras e administrativas que lidam com dezenas de comprovantes por projeto e precisam organizar e entregar documentos com consistência.

## 🛠️ Tecnologias utilizadas

- Python
- Selenium (para automação do navegador)
- BeautifulSoup (análise do HTML)
- PyPDF2 (junção de arquivos PDF)
- PIL (conversão de imagens em PDF)
- PyAutoGUI e Pyperclip (automação de salvamento de arquivos)

## ⚠️ Requisitos

- Acesso válido ao sistema SIF/Conveniar.
- Definição do caminho de download e de salvamento no script.
- ChromeDriver instalado e configurado.

---

> Automatizar esse processo eliminou a necessidade de downloads manuais e organização de centenas de arquivos. Resultado: menos erros, mais tempo para o que importa.
