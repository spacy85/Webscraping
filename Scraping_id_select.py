from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from time import sleep
import requests
from UliPlot.XLSX import auto_adjust_xlsx_column_width
import pandas as pd
import datetime

def check_internet_connection():
    try:
        response = requests.get("http://www.google.com", timeout=5)
        if response.status_code == 200:
            return True
        else:
            return False
    except requests.ConnectionError:
        return False

def initialize_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)
    return driver

def login_to_website(driver, login, pwd):
    sito = 'https://....it'
    driver.get(sito)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#i0116")))
    driver.find_element(By.CSS_SELECTOR, value="#i0116").send_keys(login + '@dominio.it')
    sleep(5)
    driver.find_element(By.CSS_SELECTOR, value="#idSIButton9").click()
    sleep(5)
    # '''
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#passwordInput")))
    driver.find_element(By.CSS_SELECTOR, value="#passwordInput").send_keys(pwd)
    driver.find_element(By.CSS_SELECTOR, value="#submitButton").click()
    # '''
    sleep(7)

def print_list_id_select(id_select):
    num_col = 4
    num_row = (len(id_select) - 1) // num_col + 1
    len_max = max(len(select) for select in id_select)
    for index_row in range(num_row):
        for index_col in range(num_col):
            index = index_row + index_col * num_row + 1
            if index < len(id_select):
                element = f"{index}: {id_select[index]}"
                print(element.ljust(len_max + 10), end="")
        print()
def select_id():
    print(
        'Opzioni: "0" per tutte gli id, "a-b" per definire un intervallo ,  "a,b,f,z" per selezionare gli id desiderate')
    selezione = input('Scegli gli id da salvare: ')
    selezioni = []

    if selezione == '0':
        selezioni = range(1, len(id_campagne))
        print(selezioni)

    elif '-' in selezione:
        start, end = map(int, selezione.split('-'))
        selezioni = range(start, end + 1)
        print(selezioni)

    else:
        selezioni = map(int, selezione.split(','))

    return selezioni

def copy_data_id(driver, lista):
    parole_da_escludere = ['more_vert', 'Dettagli', 'desktop_windows', '%', '(', ')',
                            'people']

    while True:
        try:
            rows = driver.find_elements(By.XPATH,
                                        "//tr[@data-bind='css: {greenBackground: $parent.isDaEvidenziare($data)}']")

            for row in rows:
                cols = row.find_elements(By.TAG_NAME, 'td')
                cliente = []

                for col_index, col in enumerate(cols):
                    text = col.text.strip()
                    if text and all(parola not in text for parola in parole_da_escludere):
                        if col_index > 1:
                            if text != cliente[-1]:
                                cliente.append(text)
                        else:
                            cliente.append(text)
                lista.append(cliente.copy())
                print(lista)

            driver.find_element(By.XPATH, "//a[@class='tui-page-btn tui-next']").click()
            sleep(3)

        except (TimeoutException, WebDriverException) as e:
            return lista

def create_dataframe(lista):
    clienti = pd.DataFrame(lista,
                           columns=['Nome e Cognome', 'Stato', 'Data1', 'Dato2', 'Data',
                                    'Dato3', 'Dato4', 'Dato5', 'Dato6'],
                           index=pd.RangeIndex(start=1, stop=len(lista) + 1))
    clienti['Data'] = pd.to_datetime(clienti['Data ultimo App'], dayfirst=True,
                                                errors='coerce').dt.date
    clienti['Dato4'] = clienti['Dato'].str.replace('.', '').str.replace(',', '.').str.replace('-',
                                                                                                     '0').astype(
        float)
    clienti['Dato'] = clienti['Dato'].apply(lambda value: f'€{value:.2f}')
    return clienti

def main():
    current_date = datetime.datetime.now().strftime('%d-%m-%Y')
    file_name = f'campagne_{current_date}.xlsx'
    id_select = []

    with open('../login.txt', 'r') as file:
        login = file.read().strip()
    pwd = '###'

    if not check_internet_connection():
        print("Connessione internet persa.")
        return

    driver = initialize_driver()
    if driver:
        try:
            login_to_website(driver, login, pwd)

            elementi = Select(driver.find_element(By.CSS_SELECTOR, value="#selectcampagne"))
            sleep(3)

            id_select = [option.text.replace('/', '_') for option in elementi.options]

            print_list_id_select(id_select)

            selezioni = select_id()

            with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                for j in selezioni:
                    elementi.select_by_index(j)
                    sleep(2)
                    lista = []

                    lista= copy_data_id(driver, lista)

                    clienti= create_dataframe(lista)

                    nome_sheet = id_select[j]
                    nome_sheet = ''.join(word.capitalize() for word in nome_sheet.split())
                    max_length = 31
                    nome_sheet = nome_sheet[:max_length]
                    clienti.to_excel(writer, sheet_name=nome_sheet)
                    workbook = writer.book
                    worksheet = writer.sheets[nome_sheet]
                    auto_adjust_xlsx_column_width(clienti, writer, sheet_name=nome_sheet, margin=4)


        except TimeoutException:
            print('timeout')

if __name__ == "__main__":
    main()
