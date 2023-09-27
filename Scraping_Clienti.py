import json
import os
import requests
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from time import sleep
from UliPlot.XLSX import auto_adjust_xlsx_column_width
import pandas as pd

def save_stato_scraping(pagina_corrente, riga_corrente):
    salva = {"pagina_corrente": pagina_corrente, "riga_corrente": riga_corrente}
    with open(file_scraping, 'w') as file:
        json.dump(salva, file)

def load_stato_scraping():
    if os.path.exists(file_scraping):
        with open(file_scraping, 'r') as file:
            salva = json.load(file)
            return salva.get("pagina_corrente", 1), salva.get("riga_corrente", 0), True
    else:
        return 1, 0, False  # Inizializza lo stato a pagina 1 e riga 0 e lo stato a False se il file non esiste

def check_internet_connection():
    try:
        response = requests.get("http://www.google.com", timeout=5)
        if response.status_code == 200:
            return True
        else:
            return False
    except requests.ConnectionError:
        return False

def copia_dati_clienti(driver, pagina_corrente, riga_corrente):
    lista=[]
    parole_da_escludere = ['more_vert', 'Dettagli', 'desktop_windows', '%', '(', ')']
    while True:
        try:
            rows = driver.find_elements(By.XPATH,
                                        "//tr[@data-bind='css: {greenBackground: $parent.isDaEvidenziare($data)}']")
            col_Scadenze = driver.find_elements(By.XPATH,
                                        "//td[@data-bind='visible: !$parent.showPDColumn()']")

            for row_index, row in enumerate(rows[riga_corrente:]):
                cols = row.find_elements(By.TAG_NAME, 'td')
                cliente = []

                for col_index,col in enumerate(cols):
                    text = col.text.strip()
                    if text and all(parola not in text for parola in parole_da_escludere):
                        if col_index > 1:
                            if text != cliente[-1]:
                               cliente.append(text)
                        else:
                            cliente.append(text)

                if not check_internet_connection():
                    print("Connessione internet persa.")
                    save_stato_scraping(pagina_corrente, row_index)
                    return

                try:
                    col_Scadenze[row_index].find_element(By.CSS_SELECTOR, '.clickText').click()
                    sleep(3)
                    if not check_internet_connection():
                        print("Connessione internet persa.")
                        save_stato_scraping(pagina_corrente, row_index)
                        return
                except NoSuchElementException as fine:
                    print('sono alla ultima pagina')
                    if os.path.exists(file_scraping):
                        os.remove(file_scraping)
                    driver.close()
                    return
                except WebDriverException as d:
                    print("Errore WebDriver:", str(d))
                    if "chrome not reachable" in str(d).lower():
                        print("Connessione internet persa.")
                        save_stato_scraping(pagina_corrente, row_indexx)
                    else:
                        print("Sono alla ultima pagina")
                        if os.path.exists(file_scraping):
                            os.remove(file_scraping)
                        driver.close()
                        return
                row_scadenza = driver.find_elements(By.XPATH,
                                                    "//tr[@style='width: 100%; display: inline-table;']")
                for t in range(0, len(row_scadenza)):
                    col_scadenza = row_scadenza[t].find_elements(By.TAG_NAME, 'td')
                    scadenze = cliente.copy()
                    for s in range(0, 4):
                        scadenze.append(col_scadenza[s].text)

                    lista.append(scadenze.copy())
                    print(scadenze)

                driver.find_element(By.XPATH, "//button[@class='ui-button ui-corner-all ui-widget ui-button-icon-only ui-dialog-titlebar-close']").click()
                sleep(3)
            pagina_corrente += 1
            riga_corrente, row_index = 0
            driver.find_element(By.XPATH, "//a[@class='tui-page-btn tui-next']").click()
            sleep(7)

        except (TimeoutException, WebDriverException, Exception, KeyboardInterrupt) as e:
            save_stato_scraping(pagina_corrente,row_index+1)
            return lista

    return lista

def initialize_driver():
    options = webdriver.ChromeOptions()
    #options.add_argument("--headless")
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
    sleep(10)
    driver.find_element(By.XPATH,
                        "/html[1]/body[1]/div[1]/div[2]/div[6]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/table[1]/thead[1]/tr[1]/th[10]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/div[1]").click()
    sleep(10)

def set_page_number(driver, pagina_corrente):
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#vaiapagina")))
    driver.find_element(By.CSS_SELECTOR, value="#vaiapagina").send_keys(pagina_corrente)
    sleep(1)
    driver.find_element(By.XPATH,
                        "//button[@data-bind='click: function() {goToPage(); } ']/label[@class ='label-button']").click()
    sleep(5)

def carica_dati_esistenti(file_excel):
    if os.path.exists(file_excel):
        return pd.read_excel(file_excel, sheet_name='Scadenze', engine='openpyxl', index_col=0)
    else:
        return pd.DataFrame()

def crea_dataframe(lista,stato):
    clienti = pd.DataFrame(lista, columns=['Nome e Cognome', 'Dato1', 'Dato2', 'Dato3', 'Data',
                                           'Dato4', 'Dato5', 'Scadenza', 'Dato6', 'Descrizione', 'Dato7',
                                           'Dato8', 'Data2'])

    clienti['Scadenza'] = clienti['Scadenza'].str.replace('.', '').str.replace(',', '.').str.replace('-', '0').astype(float)
    clienti['Scadenza'] = clienti['Scadenza'].apply(lambda value: F'€{value:.2F}')

    if stato:  # Se vero, carica i dati esistenti da Excel
        clienti_esistenti = carica_dati_esistenti(file_excel)
        clienti = pd.concat([clienti_esistenti, clienti])

    # Converti le colonne 'Data ultimo App' e 'Data Scadenza' nel formato desiderato
    clienti['Data'] = pd.to_datetime(clienti['Data ultimo App'], dayfirst=True, errors='coerce').dt.date
    clienti['Data2'] = pd.to_datetime(clienti['Data Scadenza'], dayfirst=True, errors='coerce').dt.date

    clienti.reset_index(drop=True, inplace=True)
    clienti.index += 1

    return clienti

def salva_dataframe_e_xlsx(clienti):
    with pd.ExcelWriter('Scadenze.xlsx', engine='xlsxwriter') as writer:
        clienti.to_excel(writer, sheet_name='Scadenze')
        workbook = writer.book
        worksheet = writer.sheets['Scadenze']
        auto_adjust_xlsx_column_width(clienti, writer, sheet_name='Scadenze', margin=4)

def main():
    global file_excel
    file_excel = 'Scadenze.xlsx'
    global file_scraping
    file_scraping = file_excel.replace('.xlsx', '_scraping.json')
    pagina_corrente, riga_corrente, stato = load_stato_scraping()

    if not check_internet_connection():
        print("Connessione internet persa.")
        return

    with open('login.txt', 'r') as file:
        login = file.read().strip()
    pwd = '###'

    driver = initialize_driver()
    if driver:
        try:
            login_to_website(driver, login, pwd)

            if stato:
                set_page_number(driver, pagina_corrente)

            lista = copia_dati_clienti(driver, pagina_corrente, riga_corrente)

            clienti = crea_dataframe(lista,stato)
            salva_dataframe_e_xlsx(clienti)

        except TimeoutException as e:
            print("Errore:", str(e))
        finally:
            if driver:
                driver.quit()

if __name__ == "__main__":
    main()
