import PyPDF2
import xlwings as xw
import tkinter as tk
from tkinter import filedialog

def select_pdf():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")],
        title="Selecione o arquivo PDF"
    )
    return file_path

def select_xls():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    planilha_path = filedialog.askopenfilename(
        title="Selecione a planilha de detalhamento",
        filetypes=[("Excel files", "*.xls *.xlsx *.xlsm *.xlsb")],
    )
    return planilha_path

def select_path():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    directory_path = filedialog.askdirectory(
        title="Selecione a pasta de destino"
    )
    return directory_path

# Path to the PDF file
pdf_path = select_pdf()
d_count = 0
repete_count = 0
dest_path = select_path()
planilha_detalhamento = xw.Book(select_xls())
path_list = []
# Open the PDF file
with open(pdf_path, "rb") as file:
    # Create a PDF reader object
    pdf_reader = PyPDF2.PdfReader(file)
    
    # Get the number of pages in the PDF
    num_pages = len(pdf_reader.pages)
    print(range(num_pages))
    # Iterate through each page and extract text
    for page_num in range(num_pages):
        page = pdf_reader.pages[page_num]
        text = page.extract_text()
        if "nome do recebedor: " in text:
            recebedor = text.split('nome do recebedor: ')[1].split(' ')[0]
            cpfpj = text.split("CPF / CNPJ do recebedor: ")[1].split('\n')[0]
            print(recebedor, cpfpj)
            if "valor:" in text:
                valor = text.split('valor: R$ ')[1].split('\n')[0]
            elif "Valor do boleto (R$);" in text:
                valor = text.split('Valor do boleto (R$);\n')[1].split('\n')[0]   
            for sheet in planilha_detalhamento.sheets:
                row = 20
                while sheet.range(f"A{row}").value is not None:
                    nome = sheet.range(f"A{row}").value
                    cnpj = sheet.range(f"B{row}").value
                    valor2 = str(sheet.range(f"F{row}").value)
                    if nome.split(" ")[0] == recebedor or cnpj.replace(".","").replace(" ","").replace(",","") == cpfpj.replace(".","").replace(" ","").replace(",",""):
                        print(f"Nome associado: {nome} {cpfpj} {cnpj} \n{valor} / {valor2}")
                        valor_format = valor.split(",")[0]
                        valor2_format = valor2.split(".")[0]
                        if valor_format.replace(".","").replace(" ","").replace(",","") == valor2_format.replace(" ","").replace(",","") and sheet.range(f"G{row}").value is None:
                            sheet.range(f"G{row}").value = f"{file}"
                            print(f"Pagamento {recebedor} {valor} efetuado")
                            break
                    row += 1             
        elif "Comprovante de pagamento de boleto" in text:
            recebedor = text.split('Beneficiário:  ')[1].split(' CPF/CNPJ')[0]
            print(recebedor)
            if "valor:" in text:
                valor = text.split('valor: R$ ')[1].split('\n')[0]
            elif "Valor do boleto (R$);" in text:
                valor = text.split('Valor do boleto (R$);\n')[1].split('\n')[0]
            for sheet in planilha_detalhamento.sheets:
                row = 20
                while sheet.range(f"A{row}").value is not None:
                    nome = sheet.range(f"A{row}").value
                    cnpj = sheet.range(f"B{row}").value
                    valor2 = str(sheet.range(f"F{row}").value)
                    if nome.replace(" ","") == recebedor.replace(" ",""):
                        print("Nome encontrado")
                        valor_format = valor.split(",")[0]
                        valor2_format = valor2.split(".")[0]
                        if valor_format.replace(".","").replace(" ","").replace(",","") == valor2_format.replace(" ","").replace(",","") and sheet.range(f"G{row}").value is None:
                            sheet.range(f"G{row}").value = "Efetuado"
                            print(f"Pagamento {recebedor} {valor} efetuado")
                            sheet.range(f"I{row}").value = "Efetuado"
                            break
                    row += 1
        elif "Dados da conta creditada:" in text:
            recebedor = text.split('Nome: ')[1].split(' ')[0]
            print(recebedor)
            if "Valor:" in text:
                valor = text.split('Valor: R$ ')[1].split('\n')[0]
            for sheet in planilha_detalhamento.sheets:
                row = 20
                while sheet.range(f"A{row}").value is not None:
                    nome = sheet.range(f"A{row}").value
                    cnpj = sheet.range(f"B{row}").value
                    valor2 = str(sheet.range(f"F{row}").value)
                    if nome.split(" ")[0] == recebedor.split(" ")[0]:
                        print("Nome associado")
                        if valor_format.replace(".","").replace(" ","").replace(",","") == valor2_format.replace(" ","").replace(",","") and sheet.range(f"G{row}").value is None:
                            sheet.range(f"G{row}").value = "Efetuado"
                            print(f"Pagamento {recebedor} {valor} efetuado")
                            break
                    row += 1
        else:
            d_count += 1
            recebedor = f"Desconhecido {d_count}"
        recebedor_raw = recebedor
        pdf_writer = PyPDF2.PdfWriter()
        pdf_writer.add_page(page)
        new_name = fr'{recebedor} {valor}'
        new_path = fr'{dest_path}\{new_name}.pdf'
        if new_name not in path_list:
            with open(new_path, "wb") as output_pdf:
                pdf_writer.write(output_pdf)
                path_list.append(new_name)
        else:
            repete_count += 1
            new_name = f"{new_name} {repete_count}"
            new_path = fr'{dest_path}\{new_name}.pdf'
            with open(new_path, "wb") as output_pdf:
                pdf_writer.write(output_pdf)
                path_list.append(new_name)
        print(f"PDF SALVO em {new_path}")
        planilha_detalhamento.save()

