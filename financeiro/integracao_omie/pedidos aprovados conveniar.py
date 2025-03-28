from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import requests
from bs4 import BeautifulSoup
import time
import datetime
import os
import pyperclip
import pyautogui
import base64
import pandas as pd
import json
from baixa_planilha_clientes import baixa_planilha_clientes
import re
import omie_api_utils as omie
import conveniar_utils as conveniar


app_key = "api_key"
app_secret = "api_secret"

pedidos_path = rf"A:\caminho_para_pedidos\{datetime.datetime.now().strftime("%d-%m-%Y")}"
os.makedirs(pedidos_path, exist_ok=True)

driver = conveniar.create_driver()

conveniar.loga_conveniar(driver)
conveniar.configure_table_columns(driver) 
rows = conveniar.tabela_linhas(driver)

lista_categorias = pd.read_excel(rf"a:\caminho_para_lista\lista_categorias.xlsx")
lista_clientes = pd.read_excel(baixa_planilha_clientes(r"caminho_para_planilha_clientes.xlsx"))

n_linha = 0
pedidos_erros = []
for row in rows:
    n_linha += 1
    valor, vencimento, cnpj, codigo_projeto, n_pedido, favorecido, cpfpj, codigo_categoria = conveniar.extrai_valor(row)
    print(f"#INICIANDO EXTRAÇÃO DO PEDIDO: {n_pedido}\nFavorecido: {favorecido}\nCPF/CNPJ: {cpfpj}\nValor: R${valor}", f"\nVencimento: {vencimento}", f"\nCódigo do Projeto: {codigo_projeto}")

    conveniar.baixa_docs(pedidos_path, row, n_linha, n_pedido, valor,driver)
    today_date = datetime.datetime.now().strftime("%d/%m/%Y")

    codigo_fornecedor_omie = omie.encontra_codigo_cliente(favorecido, cpfpj, lista_clientes)
    if codigo_fornecedor_omie is not None:
        print(f"Código do cliente no Omie {codigo_fornecedor_omie}")
    else:
        print("Código do Omie não encontrado. Iniciando registro")
        codigo_fornecedor_omie = omie.registra_cliente_omie(app_key, app_secret, favorecido, cpfpj)
        print(f"Código do cliente no Omie {codigo_fornecedor_omie}")

    if not codigo_categoria:
        codigo_categoria = "1.01.90"

    tipo_conta, conta, agencia, banco = conveniar.extrai_cnab(row)
    print(f"Tipo de conta: {tipo_conta}\nConta: {conta}\nAgência: {agencia}\nBanco: {banco}")

    if "Conta Corrente" in tipo_conta:
        codigo_forma_pagamento = "TRA"
        finalidade = "Transferência por Conta Corrente PIX"
    if "Conta Poupança" in tipo_conta:
        codigo_forma_pagamento = "TRA"
        finalidade = "Transferência por Conta Poupança PIX"
    if "Boleto" in tipo_conta:
        codigo_forma_pagamento = "BOL"
    
    if codigo_forma_pagamento == "TRA":
        cnab_integracao = {
            "codigo_forma_pagamento": codigo_forma_pagamento,
            "banco_transferencia": banco,
            "agencia_transferencia": agencia,
            "conta_corrente_transferencia": conta,
            "finalidade_transferencia": finalidade,
            "cpf_cnpj_transferencia": cpfpj,
            "nome_transferencia": favorecido,
        }

    else:
        pedidos_erros.append(n_pedido)
        continue
    novo_pedido = {
        "codigo_lancamento_integracao": f"{n_pedido.replace("/", '')}",
        "codigo_cliente_fornecedor": codigo_fornecedor_omie,
        "data_vencimento": today_date,
        "valor_documento": valor,  # Usando a variável 'valor' aqui
        "codigo_categoria": codigo_categoria,
        "data_previsao": today_date,    
        "cnab_integracao_bancaria": cnab_integracao
    }
    
    lote = str(today_date).replace("/","").replace(" ","").replace(":","")
    try:
        resposta = omie.inclui_conta_lote(app_key, app_secret,favorecido, cpfpj, valor, codigo_fornecedor_omie, n_pedido, lote, novo_pedido)
        if resposta == "200":
            print("Pedido incluído no Omie")
        else:
            pedidos_erros.append(n_pedido)

    except Exception as e:
        print("Erro ao incluir conta no lote:", e)


print("Pedidos com erro:", pedidos_erros)
driver.quit()


