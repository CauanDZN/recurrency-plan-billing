import os
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.common.exceptions import StaleElementReferenceException
from datetime import datetime
import time

load_dotenv()

head_office = os.getenv("HEAD_OFFICE")
email_address = os.getenv("SPONTE_EMAIL")
password_value = os.getenv("SPONTE_PASSWORD")

def remove_value_attribute(driver, element):
    driver.execute_script("arguments[0].removeAttribute('value')", element)

def set_input_value(driver, element, value):
    driver.execute_script("arguments[0].value = arguments[1]", element, value)

def get_day_of_week(date):
    day_of_week = date.strftime("%A")
    return day_of_week

today = datetime.strptime(datetime.now().strftime("%d/%m/%Y"), "%d/%m/%Y")

driver = webdriver.Chrome()
driver.get("https://www.sponteeducacional.net.br/SpFin/IntegracaoSpontePay.aspx")

email = driver.find_element(By.ID, "txtLogin")
email.send_keys(email_address)
password = driver.find_element(By.ID, "txtSenha")
password.send_keys(password_value)
login_button = driver.find_element(By.ID, "btnok")
login_button.click()
time.sleep(5)

enterprise = driver.find_element(By.ID, "ctl00_spnNomeEmpresa").get_attribute("innerText").strip().replace(" ", "")

if head_office == "Aldeota" and enterprise == "DIGITALCOLLEGESUL-74070":
    print("Primeiro IF")
    driver.execute_script("$('#ctl00_hdnEmpresa').val(1);javascript:__doPostBack('ctl00$lnkChange','');")
    time.sleep(3)
if head_office == "Sul" and enterprise == "DIGITALCOLLEGEALDEOTA-72546":
    print("Segundo IF")
    driver.execute_script("$('#ctl00_hdnEmpresa').val(3);javascript:__doPostBack('ctl00$lnkChange','');")
    time.sleep(3)
elif head_office == "Aldeota" and enterprise == "DIGITALCOLLEGEALDEOTA-72546":
    print("Terceiro IF")
elif head_office == "Sul" and enterprise == "DIGITALCOLLEGESUL-74070":
    print("Quarto IF")
else:
    print("Quinto IF")

initial_due_date = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_tab_tabGrid_tabLogRecorrencia_TabPanel1_wcdVencimentoInicialR_txtData")
remove_value_attribute(driver, initial_due_date)
set_input_value(driver, initial_due_date, today.strftime("%d/%m/%Y"))

final_due_date = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_tab_tabGrid_tabLogRecorrencia_TabPanel1_wcdVencimentoFinalR_txtData")
remove_value_attribute(driver, final_due_date)
set_input_value(driver, final_due_date, today.strftime("%d/%m/%Y"))

filter_button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_tab_tabGrid_tabLogRecorrencia_TabPanel1_btnFiltroRapido_div")
filter_button.click()
time.sleep(5)

rows_to_process_indices = []
while True:
    try:
        table = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_tab_tabGrid_tabLogRecorrencia_TabPanel1_grd")
        rows = table.find_elements(By.TAG_NAME, "tr")

        for index, row in enumerate(rows):
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) > 9:
                status_span = cols[9].find_element(By.TAG_NAME, "span")
                status_text = status_span.text.strip()
                if status_text in ["Recusada", "Erro"]:
                    rows_to_process_indices.append(index)
        break
    except StaleElementReferenceException:
        print("StaleElementReferenceException encountered. Reloading the table...")
        time.sleep(2)
        continue

for index in rows_to_process_indices:
    try:
        table = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_tab_tabGrid_tabLogRecorrencia_TabPanel1_grd")
        row = table.find_elements(By.TAG_NAME, "tr")[index]
        row.click()
        time.sleep(1)
        reenviar_button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_tab_tabGrid_tabLogRecorrencia_TabPanel1_divReenviar")
        reenviar_button.click()
        time.sleep(1)
        alert = Alert(driver)
        alert.accept()
        time.sleep(6)
    except StaleElementReferenceException:
        print("StaleElementReferenceException encountered during processing. Skipping this row...")
        continue

driver.quit()