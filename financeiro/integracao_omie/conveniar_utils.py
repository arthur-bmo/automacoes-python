import selenium
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


def create_driver():
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.prompt_for_download": True,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False
    })
    return webdriver.Chrome(options=chrome_options)

def loga_conveniar(driver):

    driver.get("https://sif.conveniar.com.br/Fundacao/Forms3/Convenio/GestorFinanceiro.aspx")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_UserName')))

    # Localize os campos de entrada
    username_input = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_UserName')
    password_input = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_Password')
    login_button = driver.find_element(By.ID, 'ctl00_ContentPlaceHolder1_ObjWucLoginCaptcha_lgUsuario_btnLogin')

    # Preencha os campos de login
    username_input.send_keys('login')
    password_input.send_keys('senha')

    # Clique no botão de login
    login_button.click()

def configure_table_columns(driver):    
    time.sleep(8)
    today_date = datetime.datetime.now().strftime("%d/%m/%Y")

    # Locate the input element and insert today's date
    date_input = driver.find_element(By.ID, 'txtDataVencimentoInicial')
    date_input.clear()
    date_input.send_keys(today_date)
    consultar_button = driver.find_element(By.ID, "btnConsultarLancamento")
    consultar_button.click()

    time.sleep(8)

    # Wait for the configuration button to be present and click it
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, 'btnConfigurarColunasTabela')))
    config_button = driver.find_element(By.ID, "btnConfigurarColunasTabela")
    driver.execute_script("arguments[0].scrollIntoView(true);", config_button)
    driver.execute_script("arguments[0].click();", config_button)
    time.sleep(3)

    # Wait for the "Marcar Todas" button to be present and click it
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'btnMarcarTodasColunas')))
    mark_all_button = driver.find_element(By.ID, 'btnMarcarTodasColunas')
    mark_all_button.click()

    # Wait for the "Aplicar" button to be present and click it
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'btnConfiguracaoColunaAplicar')))
    apply_button = driver.find_element(By.ID, 'btnConfiguracaoColunaAplicar')
    apply_button.click()
    time.sleep(8)

def tabela_linhas(driver):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, 'table')))

    # Get the page source and parse it with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find the table
    table = soup.find('table', {'id': 'TabelaGestorFinanceiro'})

    WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.TAG_NAME, 'tr')))
    rows = []
    tbody = table.find('tbody')
    rows = tbody.find_all('tr', class_ = 'linha')
    return rows

def extrai_valor(row):
    cells = row.find_all('td')
    valor = cells[6].text.replace(".","").replace(",",".")
    historico = cells[9].text
    if historico is not None:
        if "Categoria: " in historico:
            codigo_categoria = historico.split("Categoria: ")[1].split(",")[0]
        elif "Categoria:" in historico:
            codigo_categoria = historico.split("Categoria:")[1].split(",")[0]
        else:
            codigo_categoria = None
    else:
        codigo_categoria = None
    if not codigo_categoria:
        print("Código de categoria do Omie não encontrado")
    vencimento = cells[10].text
    cnpj = cells[15].text
    codigo_projeto = cells[22].text
    n_pedido = cells[20].text
    favorecido = cells[14].text
    cpfpj = cells[15].text

    return valor, vencimento, cnpj, codigo_projeto, n_pedido, favorecido, cpfpj, codigo_categoria

def extrai_cnab(row):
    cells = row.find_all('td')
    tipo_conta = cells[17].text
    agencia = cells[18].text
    if len(str(agencia.split("-")[0].replace(" ",""))) < 4:
        zeros_faltando = 4 - len(str(agencia.split("-")[0].replace(" ","")))
        agencia = "0" * zeros_faltando + agencia
    conta = cells[19].text
    banco = cells[16].text
    if "Caixa" in banco:
        banco = "104"
    if "Itaú" in banco:
        banco = "341"
    if "Santander" in banco:
        banco = "033"
    if "Bradesco" in banco:
        banco = "237"
    if "Banco do Brasil" in banco:
        banco = "001"
    if "Sicoob" in banco:
        banco = "756"
    if "Nubank" in banco:
        banco = "260"
    if "Inter" in banco:
        banco = "077"
    if "Safra" in banco:
        banco = "422"
    if "BTG Pactual" in banco:
        banco = "208"
    if "Banrisul" in banco:
        banco = "041"
    if "Banestes" in banco:
        banco = "021"
    if "BRB" in banco:
        banco = "070"
    if "Original" in banco:
        banco = "212"
    if "C6 Bank" in banco:
        banco = "336"
    if "XP Investimentos" in banco:
        banco = "102"
    if "Mercado Pago" in banco:
        banco = "323"
    if "PagSeguro" in banco:
        banco = "290"
    if "Neon" in banco:
        banco = "735"
    if "ModalMais" in banco:
        banco = "746"

    return tipo_conta, agencia, conta, banco

def baixa_relatorio(driver, n_linha, pedidos_path, n_pedido, valor):
    botao_relatorio = driver.find_element(By.CLASS_NAME, "ui-state-default")
    botao_relatorio.click()
    time.sleep(5)

    if botao_relatorio:
        relatorio_path = f"{n_linha}.1. PEDIDO {n_pedido.replace("/",".")} - {valor}.pdf"
        time.sleep(4)
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'iModalResponsive'))
            )
    except Exception as e:
        print('imodalresponsive não identificado')

    try:
        iframe = driver.find_element(By.ID, 'iModalResponsive')

        src = iframe.get_attribute('src')

    except Exception as e:
        print('erro reconhecendo imodal e iframe')

    if src.startswith('data:application/pdf;base64,'):
        base64_data = src.split(',')[1]  # Obtemos apenas a parte base64

            # Decodifique os dados
        pdf_data = base64.b64decode(base64_data)

        pdf_path = f'{pedidos_path}/{relatorio_path}'

            # Salve o arquivo
        with open(pdf_path, 'wb') as f:
            f.write(pdf_data)
            print(f"Relatório do pedido {relatorio_path} baixado")

        time.sleep(2)

        try:
            botao_fechar_iframe = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "btnSicoobClose"))
            )
            botao_fechar_iframe.click()
        except Exception:
            print("Erro ao clicar no botão de fechar iFrame. Tentando novamente")
            try:
                botao_fechar_iframe = driver.find_element(By.CLASS_NAME, "btn-secondary")
                print("Botão de fechar iframe identificado")

                if botao_fechar_iframe:
                    botao_fechar_iframe.click()
                    print('Iframe Fechado')
                else:
                    print("erro ao fechar iframe")

            except Exception:
                botao_fechar_iframe = driver.find_element(By.XPATH, "//button[text()='Fechar']")
                print("Botão de fechar iframe identificado")
                if botao_fechar_iframe:
                    botao_fechar_iframe.click()
                    print('Iframe Fechado')
                else:
                    print("erro ao fechar iframe")
        time.sleep(2)
                
def baixa_docs(path, row, n_linha, n_pedido, valor, driver):
    cells = row.find_all('td')


    WebDriverWait(driver, 20).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'tr.linha')))

    selenium_row = driver.find_elements(By.CSS_SELECTOR, 'tr.linha')[n_linha - 1]
    selenium_cells = selenium_row.find_elements(By.TAG_NAME, 'td')
    try:
        download_button = selenium_cells[27].find_element(By.CSS_SELECTOR, 'button.btn.btn-link')

    except:
        print("Botão de download não encontrado")
        pass

    if download_button is not None:
        driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
        try:
            driver.execute_script("arguments[0].click();", download_button)
        except Exception:
            print("Erro ao clicar no botão de download, tentando novamente")
            download_button.click()
            print("Botão clicado")

    else:
        print("Botão de download não encontrado")
    time.sleep(2)
    
    baixa_relatorio(driver, n_linha, path, n_pedido, valor)
    try:
        new_table_rows = driver.find_elements(By.CSS_SELECTOR, 'tbody tr.gridRow, tbody tr.gridAlternateRow')
    except Exception:
        linha_nula = driver.find_element(By.XPATH, "//tr[td[text()='Nenhum arquivo disponível no pedido']]")
        print("Nenhum arquivo no lançamento")
    if new_table_rows is not None:
        print(f"Número de documentos: {len(new_table_rows)}")
        segundo_numero = 1
        for new_row in new_table_rows:
            celulas_tabela_download = new_row.find_elements(By.TAG_NAME, "td")
            if len(celulas_tabela_download) > 3:
                segundo_numero += 1
                pdf_path = f"{path}\\{n_linha}.{segundo_numero} PEDIDO {n_pedido.replace("/",".")} {valor}.pdf"
                try:
                    baixar_button = new_row.find_element(By.CSS_SELECTOR, 'a[title="Baixar Anexo"]')
                    baixar_button.click()
                except Exception:
                    print("Erro ao identificar botao de baixar anexo, tentando novamente")
                    baixar_button = new_row.find_element(By.CLASS_NAME, 'tema baixar')
                    print("2222 Botão de baixar anexo identificado")
                    if baixar_button:
                        print("Botão de baixar anexo identificado")
                        baixar_button.click()
                    else:
                        print("Erro ao clicar no botão de baixar anexo")
                        break
                time.sleep(2)
                pyperclip.copy(pdf_path)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                pyautogui.press('enter')
                print(f"Arquivo Nº {segundo_numero} do pedido {n_pedido} baixado")

                time.sleep(2)
    try:
        botao_fechar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Fechar']]"))
        )
        botao_fechar.click()
        time.sleep(1)
    except Exception as e:
        print("Botão de fechar não encontrado ou não clicável:", e)
