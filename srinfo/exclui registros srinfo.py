import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.expected_conditions import visibility_of_element_located
import time
from selenium.webdriver.support.ui import Select
from datetime import datetime

links_projeto = ["https://srinfo.embrapii.org.br/projectfinance/report/12878/attachment/7"]

for link in links_projeto:
    driver = webdriver.Chrome()
    driver.get(link)
    username_input = WebDriverWait(driver,10).until(
        EC.presence_of_element_located((By.ID, 'id_username'))
    )

    username_input.send_keys('contas.ufv-fibras')

    password_input = driver.find_element(By.ID, 'id_password')
    password_input.send_keys('Embrapii07')

    entrar_botao = driver.find_element(By.CLASS_NAME, 'btn-primary')
    entrar_botao.click()

    driver.implicitly_wait(10)

    while True:
        try:
            rows = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "#object-list tbody tr"))
                    )
            rows.click()
            time.sleep(0.5)

            acoes_span = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//a[contains(@class, 'actionsButton')]"))
            )
            acoes_span.click()
            time.sleep(0.5)
            time.sleep(0.5)

            delete_button = driver.find_element(By.XPATH, "//li[contains(@class, 'deleteButton')]")
            delete_button.click()

            time.sleep(1)

            new_button = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'btn btn-default') and .//i[contains(@class, 'fa fa-times')]]"))
                )
            new_button.click()
            time.sleep(2)

        except Exception as e:
            time.sleep(30)
                    
            print(f"An error occurred: {e}")
            break
        if not rows:
            driver.quit()
            break

        time.sleep(2)

    