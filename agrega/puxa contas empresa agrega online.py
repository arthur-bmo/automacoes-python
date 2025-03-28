from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time
from bs4 import BeautifulSoup
import openpyxl
import traceback
import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import format_cell_range, cellFormat, numberFormat
from gspread.utils import rowcol_to_a1

### colocar todos os códigos do projeto em cada aba e mudar a ordem das informações para colocar a tabela no final

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(r'caminho_para_chave_api_google', scopes=SCOPES)
client = gspread.authorize(creds)

spreadsheet = client.open_by_key("id_da_planilha")

all_dataframes = []

# Configurações do WebDriver
chrome_options = Options()
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
# Usar o ChromeDriver do PATH do sistema
driver = webdriver.Chrome(options=chrome_options)

# Acesso ao site
driver.get("https://agrega.funarbe.org.br/")

# Localiza o campo de e-mail e insere o e-mail
email_field = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'loginform-email'))
)
email_field.send_keys('login')

# Localiza o campo de senha e insere a senha
password_field = driver.find_element(By.ID, 'loginform-senha')
password_field.send_keys('senha')

# Localiza o botão "Entrar" e clica nele
login_button = driver.find_element(By.NAME, 'login-button')
login_button.click()

for sheet in spreadsheet.worksheets():
    codigo_conta = sheet.acell("F10").value
    if codigo_conta:
        print(f"Processando código {codigo_conta}...")
        url = f'https://agrega.funarbe.org.br/convenio/view?id={codigo_conta}#convenio'
        driver.get(url)

        try:
            time.sleep(1.5)

            # Localiza o elemento <span> com o texto "Rubricas aprovadas"
            element = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//span[text()="Rubricas aprovadas"]'))
            )
            
            # Clica no elemento pai <a> do <span>
            parent_link = element.find_element(By.XPATH, './ancestor::a')
            parent_link.click()

            # Verifica se a mensagem "Nenhum resultado foi encontrado" está presente
            no_results_elements = driver.find_elements(By.XPATH, '//td[contains(@class, "empty")]/div[text()="Nenhum resultado foi encontrado."]')
            if no_results_elements:
                print(f"Sem resultados para o código {codigo_conta}. Pulando...")
                continue  # Pula para o próximo código

            # Aguarda até que a tabela esteja presente na página
            table = WebDriverWait(driver, 50).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'kv-grid-table'))
            )
            table_html = table.get_attribute('outerHTML')

            # Processar o HTML da tabela usando BeautifulSoup
            soup = BeautifulSoup(table_html, 'html.parser')

            # Extrair o cabeçalho da tabela
            header = [th.get_text(strip=True) for th in soup.find('tr').find_all('th')]

            # Extrair as linhas da tabela
            rows = []
            for roww in soup.find_all('tr')[1:]:  # Ignorar o cabeçalho
                cols = [td.get_text(strip=True) for td in roww.find_all('td')]
                rows.append(cols)

            if len(rows) > 0:
                last_row = rows[-1]
                last_row[2] = 'Saldo'
                rows[-1] = last_row

            # Criar um DataFrame com os dados extraídos
            df = pd.DataFrame(rows, columns=header)
            df = df.map(lambda x: (x.replace("'",'')) if isinstance(x, str) else x)
            start_row = 17
            start_col = 2
            end_row = start_row + len(df)
            end_col = start_col + len(df.columns) - 1

            # Create the cell list for the range
            cell_list = sheet.range(f'{chr(64 + start_col)}{start_row}:{chr(64 + end_col)}{end_row}')

            for i, cell in enumerate(cell_list):
                row = (i // len(df.columns))
                col = (i % len(df.columns))
                if row == 0:
                    cell.value = df.columns[col]
                else:
                    cell.value = df.iloc[row - 1, col]
            range_start_row = start_row + 1  # Começa após o cabeçalho
            range_start_col = start_col + 1  # A partir da segunda coluna
            range_end_row = end_row
            range_end_col = end_col

            # Converta para notação A1
            range_start = rowcol_to_a1(range_start_row, range_start_col)
            range_end = rowcol_to_a1(range_end_row, range_end_col)
            range_a1 = f'{range_start}:{range_end}'  # Notação A1 para o intervalo

            # Aplicar formatação como moeda
            currency_format = cellFormat(
                numberFormat=numberFormat(type="CURRENCY", pattern="R$#,##0.00"),
            )

            # Corrija a chamada para formatar o intervalo
            format_cell_range(sheet, range_a1, currency_format)
            sheet.update_cells(cell_list)

            print(f"Tabela para o código {codigo_conta} processada com sucesso.")

        except Exception as e:
            print(f"{traceback.format_exc()}")
            print(f"Erro ao processar código {codigo_conta}: {e}")

# Fechar o WebDriver
driver.quit()
