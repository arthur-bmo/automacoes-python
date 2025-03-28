from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import base64
import traceback
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import pyautogui
import pyperclip
from datetime import datetime
from PIL import Image

import requests
from bs4 import BeautifulSoup
import os

import PyPDF2

def merge_pdfs_with_same_prefix(directory):
    # Create a dictionary to store PDFs with the same prefix
    pdf_dict = {}

    # List all files in the directory
    files = sorted(os.listdir(directory))

    # Iterate through each file
    for file in files:
        # Check if it's a PDF file
        if file.endswith('.pdf'):
            # Extract the prefix
            prefix = file[:2]
            prefix1 = file[:1]

            # If prefix already exists in the dictionary, append the file to its list
            if prefix in pdf_dict:
                pdf_dict[prefix].append(file)
            # Otherwise, create a new entry with a list containing the file
            else:
                pdf_dict[prefix] = [file]

    # Iterate through the dictionary and merge PDFs with the same prefix
    for prefix, pdfs in pdf_dict.items():
        # Create a PdfWriter object to write the merged PDF
        pdf_writer = PyPDF2.PdfWriter()
        original_name = os.path.splitext(pdfs[0])[0]


        # Iterate through each PDF file with the same prefix
        for pdf_file in pdfs:
            # Open the PDF file
            with open(os.path.join(directory, pdf_file), 'rb') as pdf_file_obj:
                # Create a PdfReader object
                pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
                # Iterate through each page and add it to the PdfWriter object
                for page in pdf_reader.pages:
                    pdf_writer.add_page(page)

        # Write the merged PDF to a file
        with open(os.path.join(directory, f'{original_name} Completo.pdf'), 'wb') as merged_pdf_file:
            pdf_writer.write(merged_pdf_file)


lista_erros = []

def create_driver(download_dir):
    chrome_options = Options()
    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": download_dir,
        "download.prompt_for_download": True,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": False
    })
    return webdriver.Chrome(options=chrome_options)

def loga_conveniar(driver):
    driver.get("https://sif.conveniar.com.br/Fundacao/Login.aspx?ReturnUrl=%2ffundacao")

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

def download_arquivos_pagina(driver, n_prestacao):
    
    pagina_projetos = driver.page_source

    # Usa BeautifulSoup para analisar o HTML
    soup_projetos = BeautifulSoup(pagina_projetos, 'html.parser')

    tabela_projetos = soup_projetos.find('table', id='ctl00_ContentPlaceHolder1_gvPrincipal')

    # Encontra as linhas da tabela (todas as linhas exceto o cabeçalho)
    linhas_projetos = tabela_projetos.find_all('tr')[n_prestacao]  # Ignora o cabeçalho

    celulas = linhas_projetos.find_all('td')
    if celulas:  # Verifica se há células
        codigo_projeto = str(celulas[4].text.strip())
        nome_prestacao = str(celulas[2].text.strip())
        
        descricao_projeto = celulas[5].text.strip()
        descricao_projeto = descricao_projeto.split('/')[0].strip()  # Extrai a primeira parte e remove espaços

        # Define o caminho do projeto
        projeto_path = fr"A:\caminho_para_documentos\{descricao_projeto} - {codigo_projeto} - {nome_prestacao}"

        time.sleep(3)

        # Tenta criar o diretório
        try:
            os.mkdir(projeto_path)
            print(f"Diretório criado: {projeto_path}")
        except FileExistsError:
            print(f"O diretório já existe: {projeto_path}")
        except Exception as e:
            print(f"Ocorreu um erro ao criar o diretório: {e}")

        driver2 = create_driver(projeto_path)

        loga_conveniar(driver2)

        time.sleep(5)

        driver2.get('https://sif.conveniar.com.br/fundacao/Forms/PrestacaoConta/PrestacaodeConta.aspx?CodUsuarioGestor=2502&TipoRelatorioSolicitacao=36')

        time.sleep(5)
        
        pagina_projetos2 = driver2.page_source

        # Usa BeautifulSoup para analisar o HTML
        soup_projetos2 = BeautifulSoup(pagina_projetos2, 'html.parser')

        tabela_projetos2 = soup_projetos2.find('table', id='ctl00_ContentPlaceHolder1_gvPrincipal')

        # Encontra as linhas da tabela (todas as linhas exceto o cabeçalho)
        linhas_projetos2 = tabela_projetos2.find_all('tr')[:1]  # Ignora o cabeçalho
        celulas2 = linhas_projetos2.find_all('td')
        if celulas2:  # Verifica se há células
            botao_editar2 = celulas2[0].find('a', title='Editar registro')  # Procurando o <a> pelo title
            if botao_editar2:
                # Clica no botão usando Selenium
                editar_registro = WebDriverWait(driver2, 30).until(
                        EC.element_to_be_clickable((By.XPATH, f"//a[@title='Editar registro'][@id='{botao_editar2['id']}']"))
                    )
                editar_registro.click()
                print(f"Clicou no botão de editar registro")

                try:
                    time.sleep(2)
                    
                    # Clique no botão Lançamentos
                    botao_lancamentos = WebDriverWait(driver2, 30).until(
                        EC.element_to_be_clickable((By.XPATH, "//a[span[text()='Lançamentos']]"))
                    )
                    botao_lancamentos.click()

                    print('Botão de lançamentos identificado e clicado')

                    time.sleep(8)

                    contador1 = 0

                    pagina_lancamentos = driver2.page_source

                    soup_lancamentos = BeautifulSoup(pagina_lancamentos, 'html.parser')

                    try:
                        tabela_lancamentos = soup_lancamentos.find('table', id = 'ctl00_ContentPlaceHolder1_PrestacaodeContaLancamentoUserControl1_gvPrincipal')
                        print("tabela de lançamentos identificada")
                    except Exception:
                        print("Erro ao identificar tabela de lançamentos")
                        nenhum_registro = driver.find_element(By.XPATH, "//td[@colspan='34' and text()='Nenhum registro encontrado.']")
                        if nenhum_registro:
                            print('Nenhum registro encontrado para essa prestação de contas.')
                            pass
                                     
                    try:
                        rows2 = tabela_lancamentos.find_all('tr', class_=['gridRow', 'gridAlternateRow'])
                    except Exception:
                        print("Erro ao identificar linhas da tabela")

                    if rows2:
                        print(f"Linhas da tabela de lançamentos identificadas. Número de linhas: {len(rows2)}")

                        for index, row2 in enumerate(rows2):
                            contador2 = 0
                            time.sleep(4)

                            contador1 += 1
                            
                            # Agora você pode acessar a linha atual novamente
                            row2 = rows2[index]
                            celulas_tabela_lancamentos = row2.find_all('td')
                            if len(celulas_tabela_lancamentos)>5:
                                numero_lancamento = str(celulas_tabela_lancamentos[5].text)
                                numero_lancamento = numero_lancamento.replace("/","-")
                                rubrica = str(celulas_tabela_lancamentos[8].text)

                                historico_lancamento = str(celulas_tabela_lancamentos[7].text)
                                historico_lancamento = historico_lancamento.replace("/","-")

                                historico_lancamento_split = historico_lancamento.split()
                                historico_lancamento = " ".join(historico_lancamento_split[:3])

                                data_lancamento = str(celulas_tabela_lancamentos[11].text)
                                data_lancamento = data_lancamento.replace("/","-")
                                data_lancamento = datetime.strptime(data_lancamento, "%d-%m-%Y")
                                data_lancamento = data_lancamento.strftime("%Y-%m-%d")

                                valor_lancamento = str(celulas_tabela_lancamentos[9].text)
                                valor_lancamento = valor_lancamento.replace(".","")
                                valor_lancamento = valor_lancamento.replace(",",".")

                                print(f"Iniciando Download do lançamento: {numero_lancamento} - {historico_lancamento}")


                                try:
                                    time.sleep(3)
                                    botao_arquivos_01 = celulas_tabela_lancamentos[4].find('a', title = 'Visualizar arquivos e relatórios do pedido.')
                                    botao_arquivos_02 = WebDriverWait(driver2, 30).until(
                                            EC.element_to_be_clickable((By.XPATH, f"//a[@id='{botao_arquivos_01['id']}']"))
                                        )
                                    print("Botão de arquivos reconhecido")

                                except Exception as e:
                                    traceback.print_exc()
                                    print("Erro ao selecionar botão de arquivos, tentando de novo")
                                    try:
                                        botao_arquivos = row2.find_element(By.XPATH, ".//a[contains(@onclick, 'ModalAnexosPedido') or contains(@title, 'Visualizar arquivos e relatórios do pedido.')]")
                                        print("Botão de arquivos reconhecido")
                                    except Exception as e:
                                        print("Não foi possível reconhecer o botão de arquivos. Tentando mais uma vez")
                                        try:
                                            botao_arquivos = WebDriverWait(row2, 15).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[onclick*='ModalAnexosPedido'], a[title*='Visualizar arquivos e relatórios do pedido.']"))
                                            )
                                            print("Botão de arquivos reconhecido")
                                        except Exception as e:
                                            erro_botao_arquivo = str(f'Não foi possível reconhecer o botão de arquivos do lançamento: {celulas_tabela_lancamentos[5].text}')
                                            print(erro_botao_arquivo)
                                            lista_erros.append(erro_botao_arquivo)
                                            pass
                                try:
                                    try:
                                        driver2.execute_script("arguments[0].scrollIntoView();", botao_arquivos_02)
                                    except Exception:
                                        driver2.execute_script("arguments[0].scrollIntoView();", botao_arquivos_02)

                                        print('Erro ao scrollar até o botão, tentando novamente')
                                    botao_arquivos_02.click()
                                    print("Botão de arquivos clicado")
                                except Exception as e:
                                    try:
                                        print("problema ao clicar no botão de arquivos, tentando novamente")
                                        driver2.execute_script("arguments[0].scrollIntoView();", botao_arquivos_02)
                                        driver2.execute_script("arguments[0].click();", botao_arquivos_02)
    
                                    except Exception:
                                        print("Não foi possível clicar no botão de arquivos. Tentando mais uma vez")  
                                        try: 
                                            botao_arquivos = WebDriverWait(row2, 15).until(
                                            EC.element_to_be_clickable((By.CSS_SELECTOR, "a[onclick*='ModalAnexosPedido'], a[title*='Visualizar arquivos e relatórios do pedido.']"))
                                            )
                                            print("Botão de arquivos reconhecido novamente")
                                            try:
                                                driver2.execute_script("arguments[0].scrollIntoView();", botao_arquivos)
                                            except Exception:
                                                print("problema ao scrollar até o botão")
                                            try:
                                                botao_arquivos.click()
                                                print("clicou no botão de arquivos")
                                            except Exception:
                                                print("Não estamos conseguindo clicar no botão.")
                                                driver2.execute_script("arguments[0].scrollIntoView();", botao_arquivos)
                                                driver2.execute_script("arguments[0].click();", botao_arquivos)
                                                print("Desista sucumba")
                                        except Exception as e:
                                            driver2.execute_script("arguments[0].scrollIntoView();", botao_arquivos)
                                            driver2.execute_script("arguments[0].click();", botao_arquivos)
                                            traceback.print_exc()
                                            break
                                time.sleep(5)    
                                try:
                                    tabela_downloads = WebDriverWait(driver2, 15).until(
                                        EC.presence_of_element_located((By.ID, "table-modal-anexos"))
                                    )
                                    print("Tabela de downloads identificada")
                                
                                except Exception as e:
                                        print(f"erro ao identificar linhas da tabela de downloads: {e}")

                                try:
                                    time.sleep(2)
                                    WebDriverWait(driver2, 10).until(
                                    EC.presence_of_all_elements_located((By.XPATH, ".//tbody/tr"))
                                    )
                                    linhas3 = tabela_downloads.find_elements(By.XPATH, ".//tbody/tr")
                                    print("Linhas da tabela de download identificadas")

                                    n_linhas3 = 0
                                except Exception as e:
                                    print("Erro ao identificar linhas na tabela de download")
                                
                                try:
                                    linha_nula = tabela_downloads.find_element(By.XPATH, "//tr[td[text()='Nenhum arquivo disponível no pedido']]")
                                    print("Nenhum arquivo no lançamento")
                                    linha_nula = True
                                
                                except Exception:
                                    linha_nula = False
                                    pass

                                if not linha_nula:
                                    for linha3 in linhas3:
                                        celulas_downloads = linha3.find_elements(By.TAG_NAME, "td")
                                        if 'gridHeader' not in linha3.get_attribute('class') and len(celulas_downloads)>=3:
                                            n_linhas3 = n_linhas3 +1                                            

                                print(f'Número de documentos a serem baixados: {n_linhas3}')
                                
                                if not linha_nula:
                                    for linha3 in linhas3:
                                        if 'gridHeader' not in linha3.get_attribute('class'):
                                            celulas_tabela_download = linha3.find_elements(By.TAG_NAME, "td")
                                            
                                            if celulas_tabela_download and len(celulas_tabela_download) >= 4:
                                                try:
                                                    lista_comprovantes_strings = ["Comprovante", "Comprovantes", "COMPROVANTE", "COMPROVANTES"]
                                                    comprovante_found = [palavra for palavra in lista_comprovantes_strings if palavra in celulas_tabela_download[3].text]
                                                except Exception:
                                                    comprovante_found = False
                                                    pass
                                                if comprovante_found:  # Quarta coluna
                                                    # Execute o download prioritário
                                                    contador2 += 1
                                                    print(f"Iniciando download do Comprovante: {celulas_tabela_download[2].text}")

                                                    lancamento_path = f"{contador1}.{contador2}. {data_lancamento} - {historico_lancamento} - {valor_lancamento}"
                                                    print(f'Nome do arquivo: {lancamento_path}')


                                                    pdf_path = fr'{projeto_path}\{lancamento_path}'

                                                    download_link = linha3.find_element(By.CSS_SELECTOR, "a[title='Baixar Anexo']")
                                                    download_link.click()

                                                    time.sleep(3)
                                                    pyperclip.copy(pdf_path)
                                                    pyautogui.hotkey('ctrl', 'v')
                                                    time.sleep(0.5)
                                                    pyautogui.press('enter')

                                                    time.sleep(2)

                                                    print(f"Download do comprovante: {celulas_tabela_download[2].text} concluído")

                                                    comprovante_encontrado = True  # Marca que o download do comprovante foi feito

                                    time.sleep(1)       
                                    botao_relatorios = False
                                    try:
                                        
                                        botao_relatorio = driver2.find_element(By.CLASS_NAME, "ui-state-default")
                                        print('Botão de relatório identificado')
                                        botao_relatorio.click()
                                        print('Botão de relatórios clicado')
                                        botao_relatorios = True
                                        
                                    except Exception:
                                        try:
                                            botao_relatorio = driver2.find_element(By.XPATH, "//button[span[text()='Relatório']]")
                                            print('Botão de relatório identificado')

                                            botao_relatorio.click()
                                            print('Botão de relatórios clicado')
                                            botao_relatorios = True

                                        except Exception:
                                            botao_relatorio = botao_relatorio = driver2.find_element(By.ID, "btnRelPedido")
                                            print('Botão de relatório identificado')

                                            botao_relatorio.click()
                                            print('Botão de relatórios clicado')
                                            botao_relatorios = True
                                    
                                    time.sleep(2)
                                    if botao_relatorios:
                                        contador2 += 1
                                        relatorio_path = f"{contador1}.{contador2}. {data_lancamento} - {historico_lancamento} - {valor_lancamento}.pdf"
                                        time.sleep(4)
                                    try:
                                        WebDriverWait(driver2, 10).until(
                                            EC.presence_of_element_located((By.ID, 'iModalResponsive'))
                                            )
                                    except Exception as e:
                                        print('imodalresponsive não identificado')

                                    try:
                                        iframe = driver2.find_element(By.ID, 'iModalResponsive')

                                        src = iframe.get_attribute('src')

                                    except Exception as e:
                                        print('erro reconhecendo imodal e iframe')

                                    if src.startswith('data:application/pdf;base64,'):
                                        print("src startswith databablabla")
                                        base64_data = src.split(',')[1]  # Obtemos apenas a parte base64

                                            # Decodifique os dados
                                        pdf_data = base64.b64decode(base64_data)

                                        pdf_path = f'{projeto_path}/{relatorio_path}'

                                            # Salve o arquivo
                                        with open(pdf_path, 'wb') as f:
                                            f.write(pdf_data)
                                            print(f"Registro {pdf_path} baixado")

                                        time.sleep(2)

                                        try:
                                            botao_fechar_iframe = WebDriverWait(driver2, 10).until(
                                                EC.element_to_be_clickable((By.ID, "btnSicoobClose"))
                                            )
                                            print("Botão de fechar iFrame identificado")
                                            botao_fechar_iframe.click()
                                            print("Botão clicado")

                                        except Exception:
                                            print("Erro ao clicar no botão de fechar iFrame. Tentando novamente")
                                            try:
                                                botao_fechar_iframe = driver2.find_element(By.CLASS_NAME, "btn-secondary")
                                                print("Botão de fechar iframe identificado")

                                                if botao_fechar_iframe:
                                                    botao_fechar_iframe.click()
                                                    print('Iframe Fechado')
                                                else:
                                                    print("erro ao fechar iframe")

                                            except Exception:
                                                
                                                botao_fechar_iframe = driver2.find_element(By.XPATH, "//button[text()='Fechar']")
                                                print("Botão de fechar iframe identificado")
                                                if botao_fechar_iframe:
                                                    botao_fechar_iframe.click()
                                                    print('Iframe Fechado')
                                                else:
                                                    print("erro ao fechar iframe")
                                                
                                                
                                        time.sleep(2) 

                                    for linha3 in linhas3:
                                        if 'gridHeader' not in linha3.get_attribute('class'):
                                            time.sleep(2)
                                            celulas_tabela_download = linha3.find_elements(By.TAG_NAME, "td")
                                            
                                            if celulas_tabela_download:
                                                if len(celulas_tabela_download) >= 3:
                                                    try:
                                                        lista_comprovantes_strings = ["Comprovante", "Comprovantes", "COMPROVANTE", "COMPROVANTES"]
                                                        comprovante_nome_found = [palavra for palavra in lista_comprovantes_strings if palavra in celulas_tabela_download[3].text]
                                                    except Exception:
                                                        comprovante_nome_found = False
                                                        pass
                                                    if not comprovante_nome_found:
                                                        contador2 += 1
                                                        print(f"Iniciando download do registro: {celulas_tabela_download[2].text}")

                                                        lancamento_path = f"{contador1}.{contador2}. {data_lancamento} - {historico_lancamento} - {valor_lancamento}"
                                                        descricao_download = celulas_tabela_download[4].text

                                                        time.sleep(3)

                                                        pdf_path = fr'{projeto_path}\{lancamento_path}'

                                                        download_link = linha3.find_element(By.CSS_SELECTOR, "a[title='Baixar Anexo']")
                                                        download_link.click()
                                                        time.sleep(2)
                                                        pyperclip.copy(pdf_path)
                                                        pyautogui.hotkey("ctrl", "v")
                                                        time.sleep(0.5)
                                                        pyautogui.press('enter')

                                                        time.sleep(2)
                
                                try:
                                    botao_fechar = WebDriverWait(driver2, 10).until(
                                                    EC.element_to_be_clickable((By.CLASS_NAME, 'ui-button-text'))
                                                )
                                    botao_fechar.click()
                                except Exception as e:
                                    print("Erro ao identificar botão de fechar, tentando novamente")
                                    botao_fechar = WebDriverWait(driver2, 10).until(
                                                    EC.element_to_be_clickable((By.XPATH, "//button[span[text()='Fechar']]"))
                                                )
                                    print("Botão de fechar identificado")
                                    botao_fechar.click()
                                    print('clicou no botão de fechar')

                                print(f"Download de lançamento completo: {numero_lancamento}")

                                time.sleep(5)
                            
                except Exception as e:
                    print(f"aasaas {e}\n\n")
                    traceback.print_exc()

                time.sleep(5)
        
        for file in os.listdir(projeto_path):
            if file.lower().endswith('jpg') or file.lower().endswith('jpeg'):
                caminho_arquivo = os.path.join(projeto_path, file)
        
                imagem = Image.open(caminho_arquivo)
                caminho_pdf = os.path.join(projeto_path, f"{os.path.splitext(file)[0]}.pdf")
                
                imagem.save(caminho_pdf, 'PDF')
                print(f"Convertido: {file} -> {os.path.basename(caminho_pdf)}")

        merge_pdfs_with_same_prefix(projeto_path)
  

        driver2.quit()
        for i in lista_erros:
            print(i)                   
            
try:

    driver1  = webdriver.Chrome()
    # Acesse a página de login
    loga_conveniar(driver1)

    driver1.get('https://sif.conveniar.com.br/fundacao/Forms/PrestacaoConta/PrestacaodeConta.aspx?CodUsuarioGestor=2502&TipoRelatorioSolicitacao=36')

    time.sleep(5)

    pagina_projetos = driver1.page_source

    # Usa BeautifulSoup para analisar o HTML
    soup_projetos = BeautifulSoup(pagina_projetos, 'html.parser')

    tabela_projetos = soup_projetos.find('table', id='ctl00_ContentPlaceHolder1_gvPrincipal')

    # Encontra as linhas da tabela (todas as linhas exceto o cabeçalho)
    linhas_projetos = tabela_projetos.find_all('tr')[:1]  # Ignora o cabeçalho

    linha_prestacao_atual = 0

    for linha in linhas_projetos:

        linha_prestacao_atual = linha_prestacao_atual + 1

        download_arquivos_pagina(driver1, linha_prestacao_atual) 

        time.sleep(5)       

    time.sleep(5)


finally:
    # Feche o navegador
    driver1.quit()
    print("Download de todos os arquivos concluído.")
