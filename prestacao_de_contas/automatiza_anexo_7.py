import tkinter as tk
from tkinter import filedialog, messagebox
import locale
import PyPDF2
import os
import xlwings as xw
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from decimal import Decimal

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.expected_conditions import visibility_of_element_located
import time
from selenium.webdriver.support.ui import Select
from datetime import datetime

import unicodedata
import re

def wait_for_processing_to_finish(driver, timeout):
    try:
        # Wait until the spinner is visible
        WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.ID, "object-list_processing"))
        )
        
        # Wait until the spinner is not visible
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.ID, "object-list_processing"))
        )
        
        print("Processing completed and spinner is gone.")
    except Exception as e:
        print(f"Error or timeout: {e}")

def format_cpf(number):
    # Ensure the number is a string
    number_str = str(number)
    
    # Remove any non-digit characters
    number_str = ''.join(filter(str.isdigit, number_str))
    
    # Ensure the number has exactly 11 digits (CPF number length)
    if len(number_str) != 11:
        raise ValueError("Number must have exactly 11 digits.")
    
    # Format the number into "000.000.000-00"
    formatted_number = "{:03}.{:03}.{:03}-{:02}".format(
        int(number_str[:3]),
        int(number_str[3:6]),
        int(number_str[6:9]),
        int(number_str[9:])
    )
    
    return formatted_number

def normalize_string(s):
    # Normalize the string to NFKD form to separate characters from their accents
    normalized = unicodedata.normalize('NFKD', s)
    
    # Encode to ASCII bytes, ignoring non-ASCII characters, then decode back to string
    ascii_encoded = normalized.encode('ASCII', 'ignore').decode('ASCII')
    
    # Remove all non-alphanumeric characters
    cleaned = re.sub(r'[^a-zA-Z0-9]', '', ascii_encoded)
    
    # Convert to lowercase (or uppercase if you prefer)
    result = cleaned.lower()
    
    return result

def limit_text(canvas, text, x, y, max_width):
    words = text.split()
    lines = []
    current_line = []

    for word in words:
        # Check if adding the next word exceeds the maximum width
        if canvas.stringWidth(' '.join(current_line + [word]), "Helvetica", 11) <= max_width:
            current_line.append(word)
        else:
            lines.append(' '.join(current_line))
            current_line = [word]

    # Add the last line
    lines.append(' '.join(current_line))

    # Draw the lines on the canvas
    for line in lines:
        canvas.drawString(x, y, line)
        y -= 12  # Adjust this value based on the font size and spacing

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Selecionar Arquivos XLSX e Meses")

        # Botão para selecionar a pasta
        tk.Button(root, text="Selecionar Pasta", command=self.selecionar_pasta).pack(padx=20, pady=10)

        # Frame para exibir a lista de arquivos
        self.frame = tk.Frame(root)
        self.frame.pack(padx=20, pady=10)

        # Botão de meses
        self.meses_button = tk.Button(root, text="Selecionar Mês", command=self.selecionar_mes)
        self.meses_button.pack(pady=10)

        self.projetos_selecionados = []  # Armazena projetos selecionados

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a Pasta")
        if pasta:
            self.exibir_arquivos(pasta)
            self.pasta = pasta
        

    def exibir_arquivos(self, pasta):
        # Limpar o frame anterior
        for widget in self.frame.winfo_children():
            widget.destroy()

        arquivos_xlsx = [f for f in os.listdir(pasta) if f.endswith('.xlsx')]

        if not arquivos_xlsx:
            messagebox.showinfo("Nenhum Arquivo", "Nenhum arquivo XLSX encontrado na pasta selecionada.")
            return

        # Adicionar uma caixa de seleção para cada arquivo
        self.check_vars = []
        for arquivo in arquivos_xlsx:
            var = tk.BooleanVar()
            chk = tk.Checkbutton(self.frame, text=arquivo, variable=var)
            chk.pack(anchor='w')
            self.check_vars.append((var, arquivo))

        # Botão para mostrar os arquivos selecionados
        tk.Button(self.frame, text="Mostrar Arquivos Selecionados", command=self.mostrar_projetos_selecionados).pack(side='left', padx=10, pady=10)

    def mostrar_projetos_selecionados(self):
        self.projetos_selecionados = [arquivo for var, arquivo in self.check_vars if var.get()]
        if self.projetos_selecionados:
            messagebox.showinfo("Arquivos Selecionados", "\n".join(self.projetos_selecionados))
        else:
            messagebox.showinfo("Nenhum Selecionado", "Nenhum arquivo selecionado.")

    def selecionar_mes(self):
        meses = ["Jul-2024", "Ago-2024", "Set-2024", "Out-2024", "Nov-2024", "Dez-2024"]
        mes_selecionado = tk.StringVar(self.root)
        mes_selecionado.set(meses[0])  # valor padrão

        def aplicar_mes():
            selecionado = mes_selecionado.get()
            messagebox.showinfo("Mês Selecionado", f"Mês selecionado: {selecionado}")

            projetos_path = self.pasta
            finais_path = os.path.join("c:/Users/Larisa Cazute/Downloads/Prestação de contas/Finais/")
            nfs_path = os.path.join("c:/Users/Larisa Cazute/Downloads/Prestação de contas/Nfs/")
            modelo_pessoal_path = os.path.join("c:/Users/Larisa Cazute/Downloads/Prestação de contas/Modelos Pessoal/")
            comprovantes_path = os.path.join("c:/Users/Larisa Cazute/Downloads/Prestação de contas/Comprovantes/")

            modelos_pessoal = os.listdir(modelo_pessoal_path)

            locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
            if aplicar_mes:
                for projeto in self.projetos_selecionados:
                    projeto_path = os.path.join(projetos_path, projeto)
                    projeto_reduzido = projeto.rstrip(".xlsx") 

                    
                    planilha = xw.Book(projeto_path)

                    print(f"Iniciando registro do projeto: {projeto}")

                    projeto_cel = (planilha.sheets[1].range("J14")).value
                    projeto_nome = str(projeto_cel)
                    
                    codigo_projeto_cel = planilha.sheets[1].range("J15")
                    codigo_projeto = str(codigo_projeto_cel.value)
                    equipamentos_finais_path = os.path.join(finais_path, codigo_projeto, "Equipamentos/")
                    pessoal_finais_path = os.path.join(finais_path, codigo_projeto, "Pessoal/")
                    os.makedirs(pessoal_finais_path, exist_ok=True)
                    os.makedirs(equipamentos_finais_path, exist_ok = True)
                    print(f"Diretórios Criados: {equipamentos_finais_path}  e   {pessoal_finais_path}")
                    coordenador_cel = planilha.sheets[1].range("J12")
                    coordenador_found = False

                    driver = webdriver.Chrome()
                    link_projeto_cel = planilha.sheets[1].range("J16")
                    link_projeto  = str(link_projeto_cel.value)
                    driver.get(link_projeto)

                    username_input = WebDriverWait(driver,10).until(
                        EC.presence_of_element_located((By.ID, 'id_username'))
                    )

                    username_input.send_keys('contas.ufv-fibras')

                    password_input = driver.find_element(By.ID, 'id_password')
                    password_input.send_keys('Embrapii07')

                    entrar_botao = driver.find_element(By.CLASS_NAME, 'btn-primary')
                    entrar_botao.click()

                    driver.implicitly_wait(10)

                    sorting_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//th[@class="sorting_asc"]'))
                    )
                    sorting_button.click()
                    

                    ##para cada planilha no documento de contrapartida:
                    for sheet in planilha.sheets:
                        if sheet.name == selecionado:
                            print("----------- Início do registro MÊS: " +  sheet.name + "  -------------")
                            for cell in sheet.range("B2:B148"):
                                try:
                                    if cell.value is not None:

                                        nome_eq = str(cell.value)
                                        print("EQUIPAMENTO IDENTIFICADO: " + nome_eq)

                                        finalidade_cel = cell.offset(column_offset = 1)
                                        macroentrega_cel = cell.offset(column_offset = -1)
                                        custo_cel = cell.offset(column_offset = 4)
                                        hora_cel = cell.offset(column_offset = 10)
                                        valor_cel = cell.offset(column_offset = 11)
                                        credor_cel = cell.offset(column_offset = 7)
                                        cnpj_cel = cell.offset(column_offset = 8)
                                        nota_fiscal_cel = cell.offset(column_offset = 9)
                                        data_cel = planilha.sheets[sheet].range("A151")

                                        
                                        finalidade = str(finalidade_cel.value)
                                        custo = locale.currency(float(custo_cel.value), grouping = True)
                                        custo = str(custo)
                                        hora = str(hora_cel.value)
                                        hora = hora.replace(".0", "")
                                        hora = str(hora) + " horas"
                                        valor_format = locale.currency(float(valor_cel.value), grouping = True)
                                        mes_cel = planilha.sheets[sheet].range("B151")
                                        mes = str(mes_cel.value)
                                        data = str(data_cel.value)
                                        data = data.replace(" 00:00:00", "")
                                        data = data.replace("-", "/")
                                        
                                        try:
                                            data_format =datetime.strptime(data, '%Y/%m/%d')
                                        except Exception as e:
                                            data_format = datetime.strptime(data, '%d/%m/%Y')
                                            data_ano = data_format.strptime(data, '%B/%Y')
                                            print("Segundo Formato Funcionou")
                                        else:
                                            data_ano = data_format.strftime('%B/%Y')
                                            print("Primeiro Formato Funcionou")


                                        data_ano = data_ano.capitalize()
                                        coordenador = str(coordenador_cel.value)
                                        projeto = str(projeto_cel)
                                        macroentrega = str(macroentrega_cel.value)
                                        n_macroentrega = macroentrega[0]
                                        n_macroentrega = int(n_macroentrega)- 1

                                        credor = str(credor_cel.value)
                                        cnpj = str(cnpj_cel.value)
                                        if nota_fiscal_cel.value is not None:
                                            nota_fiscal = str(nota_fiscal_cel.value)
                                            nota_fiscal = nota_fiscal[:-2]
                                        valor_registro = float(valor_cel.value)
                                        valor_registro = "{:.2f}".format(valor_registro)
                                        data_cel = data_cel.value
                                        data_formatada = data_cel.strftime('%m/%Y')
                    
                                        ##CRIA CÓPIA DE CADA MODELO DA PASTA DE MODELOS
                                        for modelo in modelos_eq:
                                            modelo_reduzido = modelo.replace(".pdf", "")
                                            modelo_reduzido = normalize_string(modelo_reduzido)
                                            ##se o nome do modelo for igual ao do coordenador:
                                            if modelo_reduzido == normalize_string(coordenador):
                                                print(f"Modelo Encontrado. Coordenador: {coordenador}")
                                                coordenador_found = True
                                                print("aaaaaaa")
                                                modelo_path_nome = os.path.join(modelo_eq_path, modelo)
                                                print(modelo_path_nome)
                                                modelo_path = os.path.join(modelo_path_nome)
                                                print("Modelo_path")

                                                with open(modelo_path, 'rb') as source_file:
                                                    print("Modelo Aberto")
                                                    pdf_reader = PyPDF2.PdfReader(source_file, strict = False)
                                                    page_to_copy = pdf_reader.pages[0]
                                                    pdf_writer = PyPDF2.PdfWriter()
                                                    pdf_writer.add_page(page_to_copy)

                                                    # Cria uma pasta pra cada projeto
                                                    
                                                    doc_eq_final_path = equipamentos_finais_path + "/" + nome_eq +" "+ mes  +".pdf"

                                                    with open(doc_eq_final_path, 'wb') as output_file:
                                                        pdf_writer.write(output_file)

                                                #Essa parte é responsável por preencher a cópia do documento
                                                #que criamos previamente
                                                packet = io.BytesIO()
                                                can = canvas.Canvas(packet, pagesize=letter)
                                                font_name = "Helvetica"
                                                font_size = 12
                                                can.setFont(font_name,font_size)

                                                can.drawString(225, 648, nome_eq)
                                                limit_text(can, projeto_nome, 140, 631.9, 278)
                                                novo_data_ano = mes + "/2024"
                                                can.drawString(450, 590, novo_data_ano)
                                                limit_text(can, macroentrega, 92, 489, 388)
                                                limit_text(can, finalidade, 188.5, 351.3, 300)
                                                can.drawString(217, 263, custo)
                                                can.drawString(295, 237, hora)
                                                can.drawString(357, 212, valor_format)
                                                can.save()

                                                packet.seek(0)

                                                new_pdf = PyPDF2.PdfReader(packet, strict = False)

                                                existing_pdf_path = os.path.join(doc_eq_final_path)

                                                existing_pdf = PyPDF2.PdfReader(existing_pdf_path,"wb")
                                                output = PyPDF2.PdfWriter()

                                                page = existing_pdf.pages[0]
                                                page.merge_page(new_pdf.pages[0])
                                                output.add_page(page)

                                                file_found = False

                                                for nf in os.listdir(nfs_path):
                                                    nome_nf = nf.rstrip(".pdf")
                                                    if normalize_string(nome_nf) == normalize_string(nome_eq):
                                                        print(normalize_string(nome_nf) + "  " + normalize_string(nome_eq))
                                                        print("Nota fiscal Encontrada")
                                                        file_found = True
                                                        nf_path = nfs_path + nf
                                                        pdf2 = PyPDF2.PdfReader(nf_path, strict = False)
                                                        for page in pdf2.pages:
                                                            output.add_page(page)

                                                        outputStream = open(doc_eq_final_path, 'wb')
                                                        output.write(outputStream)
                                                        outputStream.close()
                                                        print("Comprovante criado")                                    

                                                        break
                                                if not file_found:
                                                    print("########## Nota FISCAL NÃO ENCONTRADA - EQUIPAMENTO: " + nome_eq + "################")

                    #Agora, faremos o registro no SRInfo:
                                        time.sleep(1.5)
                                        novo_registro_span = WebDriverWait(driver, 10).until(
                                            EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "Novo registro")]'))
                                            )
                                        novo_registro_span.click()

                                        credor = str("UNIDADE EMBRAPII FIBRAS FLORESTAIS/FUNARBE")
                                        credor_input = WebDriverWait(driver, 10).until(
                                            EC.presence_of_element_located((By.ID, 'DTE_Field_creditor'))
                                            )
                                        credor_input.send_keys(credor)

                                        dropdown = Select(driver.find_element(By.ID, 'DTE_Field_wp_related'))
                                        
                                        dropdown.select_by_index(n_macroentrega)

                                        cnpj  = str("20.320.503/0001-51")

                                        cnpj_input = driver.find_element(By.ID, 'DTE_Field_cpf_cnpj')
                                        cnpj_input.send_keys(cnpj)

                                        despesa_input = driver.find_element(By.ID, 'DTE_Field_expense_type')
                                        despesa_input.send_keys('Uso de equipamento laboratorial próprio')

                                        data_input = driver.find_element(By.ID, 'DTE_Field_reference_month')
                                        data_input.send_keys(data_formatada)

                                        if nota_fiscal_cel.value != None:
                                            nf_input = driver.find_element(By.ID, 'DTE_Field_invoice_number')
                                            nf_input.send_keys(nota_fiscal)

                                        valor_input = driver.find_element(By.ID, 'DTE_Field_value')
                                        valor_input.send_keys(valor_registro)

                                        save_button = driver.find_element(By.XPATH, '//button[@class="btn btn-default"]')
                                        save_button.click()

                                        wait_for_processing_to_finish(driver, 8)
                                        try:
                                            rows = WebDriverWait(driver, 5).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, "#object-list tbody tr"))
                                            )
                                            rows.click()
                                            time.sleep(0.5)

                                        except Exception as e:
                                            print("Erro na seleção de registro, tentando novamente.")
                                        
                                            element = WebDriverWait(driver, 10).until(
                                                EC.presence_of_element_located((By.CSS_SELECTOR, "td.select-checkbox input[type='hidden']"))
                                            )

                                            # Find the parent td element and click it
                                            parent_td = element.find_element(By.XPATH, '..')
                                            parent_td.click()
                                            print("Checkbox selected.")
                        
                                        acoes_span = driver.find_element(By.XPATH, "//a[contains(@class, 'actionsButton')]")
                                        acoes_span.click()
                                        time.sleep(0.5)

                                        anexar_button = driver.find_element(By.XPATH, "//li[contains(@class, 'attachButton')]")
                                        anexar_button.click()

                                        time.sleep(2.5)
                    
                                        try:
                                            upload_input = WebDriverWait(driver, 10).until(
                                                EC.presence_of_element_located((By.XPATH, '//input[@class="form-control"]'))
                                                )
                                            upload_input.send_keys(doc_eq_final_path)
                                        except Exception as e:
                                            print("Erro no SRInfo. Tentando novamente.")
                                            fechar_button = driver.find_element(By.XPATH, '//i[@class="fa fa-times"]')
                                            fechar_button.click()
                                            wait_for_processing_to_finish(driver, 8)
                                            
                                            try:
                                                rows = WebDriverWait(driver, 5).until(
                                                    EC.presence_of_element_located((By.CSS_SELECTOR, "#object-list tbody tr"))
                                                    )
                                                rows.click()
                                                time.sleep(0.5)

                                                acoes_span = driver.find_element(By.XPATH, "//a[contains(@class, 'actionsButton')]")
                                                acoes_span.click()
                                                time.sleep(0.5)

                                                anexar_button = driver.find_element(By.XPATH, "//li[contains(@class, 'attachButton')]")
                                                anexar_button.click()

                                                time.sleep(2.5)

                                                upload_input = WebDriverWait(driver, 30).until(
                                                    EC.presence_of_element_located((By.XPATH, '//input[@class="form-control"]'))
                                                    )
                                                upload_input.send_keys(doc_eq_final_path)

                                            except Exception as e:
                                                print("Erro na seleção de registro, tentando novamente.")
                                            
                                                element = WebDriverWait(driver, 10).until(
                                                    EC.presence_of_element_located((By.CSS_SELECTOR, "td.select-checkbox input[type='hidden']"))
                                                )
                                                # Click on the parent td element
                                                parent_td = element.find_element_by_xpath('..')
                                                parent_td.click()
                                                print("Checkbox selected.")
                                                acoes_span = driver.find_element(By.XPATH, "//a[contains(@class, 'actionsButton')]")
                                                acoes_span.click()
                                                time.sleep(0.5)

                                                anexar_button = driver.find_element(By.XPATH, "//li[contains(@class, 'attachButton')]")
                                                anexar_button.click()

                                                time.sleep(2.5)

                                                upload_input = WebDriverWait(driver, 30).until(
                                                    EC.presence_of_element_located((By.XPATH, '//input[@class="form-control"]'))
                                                    )
                                                upload_input.send_keys(doc_eq_final_path)

                                        time.sleep(4)

                                        print("Equipamento registrado")
                                        try:
                                            fechar_button = driver.find_element(By.XPATH, '//i[@class="fa fa-times"]')
                                            fechar_button.click()
                                        except Exception as e:
                                            print("Erro no botão de fechar")
                                            fechar_button = WebDriverWait(driver, 10).until(
                                                EC.visibility_of_element_located((By.CSS_SELECTOR, "button.btn.btn-default"))
                                                )
                                            fechar_button.click()

                                        wait_for_processing_to_finish(driver, 20)
                                        print("")
                                except Exception as e:
                                    print(f"Erro ao processar a célula {cell.address}: {e}")
                                
                            # Registro de bolsistas:
                            for cell in sheet.range("B154:B169"):
                                if cell.value is not None:

                                    nome = str(cell.value)
                                    print("BOLSISTA ENCONTRADO: "+nome)

                                    finalidade_p = str((cell.offset(column_offset = 1)).value)
                                    sheet.range('C154:C169').number_format = '0'
                                    cpf = str((cell.offset(column_offset = 2)).value)
                                    print(cpf)
                                    cpf = format_cpf(cpf)
                                    hora = str((cell.offset(column_offset = 6)).value)
                                    valor1_cel = cell.offset(column_offset = 10)
                                    salario = str((cell.offset(column_offset = 3)).value)
                                    salario = str("R$"+salario)
                                    
                                    valor1_valor = valor1_cel.value
                                    valor1 = locale.currency(valor1_valor, grouping = True)
                                    valor1 = str(valor1)

                                    mes_cel = planilha.sheets[sheet].range("B151")
                                    mes = str(mes_cel.value)
                                    
                                    data_cel = planilha.sheets[sheet].range("A151").value
                                    
                                    data_formatada = data_cel.strftime('%m/%Y')
                                    data_ano = data_cel.strftime('%B/%Y')
                                    data_ano = data_ano.capitalize()


                                    hora = hora.replace(".0", "")
                                    hora = hora + " horas"
                                    
                                    macro_cel = cell.offset(column_offset = -1)
                                    macro = macro_cel.value
                                    n_macro = macro[0]
                                    n_macro = int(n_macro) -1

                                    valor1_registro = float(valor1_cel.value)
                                    valor1_registro = "{:.2f}".format(valor1_registro)

                                    nome_final = nome +" "+ mes + ".pdf"
                                    final_path = pessoal_finais_path  + nome_final

                                    for modelo in os.listdir(modelo_pessoal_path):
                                        modelo_reduzido = modelo.rstrip(".pdf")
                                        ##se o nome do modelo for igual ao do coordenador:
                                        if normalize_string(modelo_reduzido) == normalize_string(str(coordenador_cel.value)):
                                            coordenador_found = True
                                            coordenador = str(coordenador_cel.value)
                                            print(f'Modelo Encontrado. Coordenador: {coordenador}')
                                            
                                            modelo_path = os.path.join(modelo_pessoal_path, modelo)
                                            with open(modelo_path, 'rb') as source_file:

                                                #abre o modelo em que serão inseridas as informações
                                                pdf_reader = PyPDF2.PdfReader(source_file, strict = False)
                                                page_to_copy = pdf_reader.pages[0]
                                                pdf_writer = PyPDF2.PdfWriter()
                                                pdf_writer.add_page(page_to_copy)


                                                with open(final_path, 'wb') as output_file:
                                                    pdf_writer.write(output_file)

                                            #Essa parte é responsável por preencher a cópia do documento
                                            #que criamos previamente
                                            packet = io.BytesIO()
                                            can = canvas.Canvas(packet, pagesize=letter)
                                            font_name = "Helvetica"
                                            font_size = 12
                                            can.setFont(font_name,font_size)


                                            mesano = f"{mes}/2024 "

                                            can.drawString(85, 676.5, nome)
                                            can.drawString(403, 676.5, cpf)
                                            limit_text(can, projeto_nome, 95, 662, 250)
                                            can.drawString(374, 613, mesano)
                                            limit_text(can, macro, 48.5, 527, 465)
                                            limit_text(can, finalidade_p, 199, 389, 315)
                                            can.drawString(195, 251, hora)
                                            can.drawString(350, 223, valor1)
                                            can.save()

                                            packet.seek(0)

                                            new_pdf = PyPDF2.PdfReader(packet, strict = False)

                                            existing_pdf_path = os.path.join(final_path)

                                            existing_pdf = PyPDF2.PdfReader(existing_pdf_path,"wb")
                                            output = PyPDF2.PdfWriter()

                                            page = existing_pdf.pages[0]
                                            page.merge_page(new_pdf.pages[0])
                                            output.add_page(page)

                                            file_found = False

                                            for comprovante in os.listdir(comprovantes_path):
                                                nome_bo = comprovante.replace(".pdf", "")
                                                if normalize_string(nome_bo) == normalize_string(nome):
                                                    print("Comprovante de renda encontrado")
                                                    file_found = True
                                                    comprovante_path = comprovantes_path + comprovante
                                                    pdf2 = PyPDF2.PdfReader(comprovante_path, strict = False)
                                                    for page in pdf2.pages:
                                                        output.add_page(page)

                                                    outputStream = open(final_path, 'wb')
                                                    output.write(outputStream)
                                                    outputStream.close()
                                                    print("Documento criado")                                    

                                                    break
                                            if not file_found:
                                                print("########## COMPROVANTE DE RENDA NÃO ENCONTRADO - BOLSISTA: " + nome + " ################")

                                                
                                                #Registro do pessoal no SRINFO:
                                            if file_found is True:
                    
                                                novo_registro_span = WebDriverWait(driver, 30).until(
                                                    EC.presence_of_element_located((By.XPATH, '//span[contains(text(), "Novo registro")]'))
                                                    )
                                                novo_registro_span.click()
                                                time.sleep(1)

                                                dropdown = Select(driver.find_element(By.ID, 'DTE_Field_wp_related'))
                                                dropdown.select_by_index(n_macro)
                                                time.sleep(0.5)
                                                
                                                cnpj_input = driver.find_element(By.ID, 'DTE_Field_cpf_cnpj')
                                                cnpj_input.send_keys(cpf)
                                                time.sleep(0.5)

                                                credor_input = WebDriverWait(driver, 30).until(
                                                    EC.presence_of_element_located((By.ID, 'DTE_Field_creditor'))
                                                    )
                                                credor_input.send_keys(nome)
                                                time.sleep(0.5)

                                                data_input = driver.find_element(By.ID, 'DTE_Field_reference_month')
                                                data_input.send_keys(data_formatada)
                                                time.sleep(0.5)

                                                valor_input = driver.find_element(By.ID, 'DTE_Field_value')
                                                valor_input.send_keys(valor1_registro)
                                                time.sleep(1)

                                                save_button = driver.find_element(By.XPATH, '//button[@class="btn btn-default"]')
                                                save_button.click()

                                                wait_for_processing_to_finish(driver, 8)
                                                try:
                                                    rows = WebDriverWait(driver, 5).until(
                                                    EC.presence_of_element_located((By.CSS_SELECTOR, "#object-list tbody tr"))
                                                    )
                                                    rows.click()
                                                    time.sleep(0.5)

                                                except Exception as e:
                                                    print("Erro na seleção de registro, tentando novamente.")
                                                
                                                    highest_visible_row = None
                                                    highest_position = float('inf')  # Initialize with a high value

                                                    for row in rows:
                                                        # Get the row's location
                                                        location = row.location
                                                        # Check if this row is more visible or higher in the viewport
                                                        if location['y'] < highest_position:
                                                            highest_visible_row = row
                                                            highest_position = location['y']

                                                    if highest_visible_row:
                                                        # Click the highest visible row
                                                        highest_visible_row.click()
                                                        time.sleep(0.5)
                                                acoes_span = driver.find_element(By.XPATH, "//a[contains(@class, 'actionsButton')]")
                                                acoes_span.click()
                                                time.sleep(0.5)

                                                anexar_button = driver.find_element(By.XPATH, "//li[contains(@class, 'attachButton')]")
                                                anexar_button.click()

                                                time.sleep(2.5)
                            

                                                try:
                                                    upload_input = WebDriverWait(driver, 10).until(
                                                    EC.presence_of_element_located((By.XPATH, '//input[@class="form-control"]'))
                                                    )
                                                    upload_input.send_keys(final_path)

                                                except Exception as e:
                                                
                                                    print("Erro no SRInfo. Tentando novamente.")
                                                    fechar_button = driver.find_element(By.XPATH, '//i[@class="fa fa-times"]')
                                                    fechar_button.click()
                                                    wait_for_processing_to_finish(driver, 8)
                                                    
                                                    rows.click
                                                    time.sleep(0.5)

                                                    acoes_span = driver.find_element(By.XPATH, "//a[contains(@class, 'actionsButton')]")
                                                    acoes_span.click()
                                                    time.sleep(0.5)

                                                    anexar_button = driver.find_element(By.XPATH, "//li[contains(@class, 'attachButton')]")
                                                    anexar_button.click()

                                                    time.sleep(2.5)

                                                    upload_input = WebDriverWait(driver, 30).until(
                                                        EC.presence_of_element_located((By.XPATH, '//input[@class="form-control"]'))
                                                        )
                                                        
                                                    upload_input.send_keys(final_path)

                                                time.sleep(4)

                                                print("Bolsista registrado(a)")

                                                fechar_button = driver.find_element(By.XPATH, '//i[@class="fa fa-times"]')
                                                fechar_button.click()

                                                time.sleep(2)


                            print("--------------- Mês Concluído: " +  sheet.name + "  -----------------")
                    
                    num_docs_eq = len(os.listdir(equipamentos_finais_path))
                    num_docs_pessoal = len(os.listdir(pessoal_finais_path))
                    num_docs = str(num_docs_eq+num_docs_pessoal)

                    print("Foram criados " + str(num_docs_eq) + " documentos de equipamentos e "+str(num_docs_pessoal)+" de pessoal")

                    driver.quit()
                                        
                    planilha.close
                    planilha.app.quit()                      
            if not aplicar_mes:
                print("Nada selecionado ERRO")

        menu = tk.OptionMenu(self.root, mes_selecionado, *meses)
        menu.pack(pady=10)

        # Botão para confirmar a seleção do mês
        tk.Button(self.root, text="Aplicar Mês", command=aplicar_mes).pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()