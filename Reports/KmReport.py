import time
import os
import pandas as pd
import schedule
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import logging


logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


def add_data_to_existing_report(download_folder, report_path):
    global df_existing, df_downloaded
    downloaded_file = os.path.join(download_folder, "Distância Percorrida.xls")

    if os.path.exists(downloaded_file):
        logging.info(f"Arquivo baixado encontrado: {downloaded_file}")

        try:
            df_downloaded = pd.read_excel(downloaded_file, skiprows=4)

            if "Data" in df_downloaded.columns:
                df_downloaded["Data"] = df_downloaded["Data"].str.split(" ").str[0]
                df_downloaded["Data"] = pd.to_datetime(df_downloaded["Data"], format='%d/%m/%Y')
                logging.info(f"' (UTC-3)' removido da coluna 'Data'.")

            df_downloaded = df_downloaded.drop_duplicates(subset="Veículo", keep="last")
            logging.info(f"Duplicados removidos no arquivo baixado com base na coluna 'Veículo'.")

            if os.path.exists(report_path):
                logging.info(f"Arquivo de relatório encontrado: {report_path}")
                df_existing = pd.read_excel(report_path)

                df_downloaded.columns = df_existing.columns

                df_combined = pd.concat([df_existing, df_downloaded], ignore_index=True)
            else:
                logging.info(f"Arquivo de relatório não encontrado, criando um novo arquivo.")
                df_combined = df_downloaded

            df_combined[df_combined.columns[0]] = df_combined[df_combined.columns[0]].astype(str)
            df_combined = df_combined[~df_combined[df_combined.columns[0]].str.contains("Distância Percorrida", na=False)]

            df_combined.to_excel(report_path, index=False)
            logging.info(f"Relatório atualizado salvo em: {report_path}")

        except Exception as e:
            logging.error(f"Erro ao processar as planilhas: {e}")
            if 'df_existing' in globals():
                logging.info(f"Colunas do relatório existente: {df_existing.columns}")
            if 'df_downloaded' in globals():
                logging.info(f"Colunas do relatório baixado: {df_downloaded.columns}")

        os.remove(downloaded_file)
        logging.info(f"Arquivo baixado {downloaded_file} removido após processamento.")
    else:
        logging.error(f"Arquivo baixado não encontrado: {downloaded_file}")


def login():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument("--headless")

    service = Service(ChromeDriverManager(driver_version="131.0.6778.108").install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    logging.info("Abrindo a página de login")
    driver.get("https://")

    try:
        permission_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn.float-right"))
        )
        permission_button.click()

        usuario_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "form:usuario"))
        )
        login_field = driver.find_element(By.ID, "form:login")
        senha_field = driver.find_element(By.ID, "form:senha")

        usuario_field.send_keys("")
        login_field.send_keys("")
        senha_field.send_keys("")

        ok_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "form:btnOk"))
        )
        ok_button.click()

        checkbox_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "controller:opcoes_relatorio:31"))
        )
        driver.execute_script("arguments[0].scrollIntoView();", checkbox_element)

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "controller:opcoes_relatorio:31"))
        )

        checkbox_element.click()
        logging.info("Checkbox ou botão de relatório clicado!")

        checkbox_element_day = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "controller:periodo:1"))
        )
        checkbox_element_day.click()

        generate_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "controller:btnGeraRelatorio"))
        )

        generate_element.click()

        export_register = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="formrelatorio:listagem"]/div[1]/input'))
        )

        export_register.click()

        time.sleep(5)

        download_folder = os.path.expanduser("~/Downloads")

        report_path = r"H:\Relatorio_km.xlsx"

        add_data_to_existing_report(download_folder, report_path)

    except Exception as e:
        logging.error(f"Erro: {e}")

    finally:
        driver.quit()


schedule_time = "07:00"
schedule.every().day.at(schedule_time).do(login)

while True:
    schedule.run_pending()
    time.sleep(1)